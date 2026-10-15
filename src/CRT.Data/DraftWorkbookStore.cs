using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // READING AND WRITING THE BOARD A DRAFT HOLDS
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** THIS IS WHAT REPLACED DraftDataStore. *** That class read and wrote draft.json - a list
    // of row deltas recorded at edit time. A draft is now a real board folder, so persisting an
    // edit means writing the board WORKBOOK, and reading one means parsing it back.
    //
    // *** EVERY EDIT IS READ-MODIFY-WRITE AGAINST THE FILE ON DISK, NEVER AGAINST A CACHED COPY.
    // *** That is the rule this class exists to enforce, and it follows directly from the
    // project owner's requirement that editing in the app and editing in Excel be interchangeable.
    // The contributor may have changed the workbook since the application last read it; writing
    // back something built on a stale copy would silently discard that. So a caller hands over a
    // function that transforms the CURRENT board, and this class is the one that decides what
    // "current" means.
    //
    // WHAT ABOUT THE UNSAVED STATE ON SCREEN? Nothing here addresses it, deliberately. A tab
    // holding an edit the user has not saved is the tab's business; this class refuses to make a
    // WRITE depend on a value read minutes earlier, which is a different and much more damaging
    // problem.
    //
    // THE SIDECAR IS NOT THIS CLASS'S JOB. Component highlights and KiCad calibrations live in the
    // JSON beside the workbook, and BoardComponentHighlightStorage already reads and writes it -
    // in the published format, which a draft folder now uses unchanged. Adding a second path to it
    // here would be the two-mechanism mistake this whole change is removing.
    // ###########################################################################################
    public static class DraftWorkbookStore
    {
        // ###########################################################################################
        // The board a draft currently holds, or null when there is no draft (or its workbook cannot
        // be parsed).
        //
        // *** READ FRESH, NEVER CACHED. *** BoardDataReader has a process-wide cache and this
        // deliberately does not use it: the whole point is to see what is on disk right now,
        // including an edit made in Excel while the application was running. A cached read here
        // would reintroduce exactly the staleness this class exists to prevent.
        // ###########################################################################################
        public static BoardData? LoadDraftBoard(string draftsRoot, string excelDataFile)
        {
            string workbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            if (workbook.Length == 0 || !File.Exists(workbook))
            {
                return null;
            }

            return BoardDataReader.ReadWorkbookUncached(workbook);
        }

        // ###########################################################################################
        // Applies one edit to a draft's board and writes it back.
        //
        // `transform` receives the board as it is on disk RIGHT NOW and returns the board as it
        // should be. It must not mutate its input - every draft writer in this codebase already
        // follows that rule (they return a new object), and honouring it here means a failed write
        // leaves the caller's own copy untouched.
        //
        // Returns false when there is no draft to edit rather than creating one. Creating a draft
        // is DraftSeeder's job and involves copying the published board; quietly starting an empty
        // one here would produce a draft that reads as "every published row deleted".
        //
        // *** THE WRITE IS ATOMIC. *** BoardWorkbookWriter writes the file in place, so a crash
        // mid-write would leave a truncated workbook - and that workbook is now the only copy of
        // the contributor's work. Writing to a temporary file and moving it into place means a
        // failure leaves the previous, complete draft - its sidecar included: a sidecar that cannot
        // be written puts the previous workbook back (EditCore), so false means nothing changed.
        // ###########################################################################################
        public static bool Edit(
            string draftsRoot,
            string excelDataFile,
            Func<BoardData, BoardData> transform) =>
            DraftWorkbookStore.EditCore(draftsRoot, excelDataFile, expectedFingerprint: null, transform)
                == DraftWorkbookEditOutcome.Saved;

        // ###########################################################################################
        // Edit, but ONLY if the workbook is still exactly what the caller read
        // (added 2026-09-24, for the Drafts tab's table editor).
        //
        // *** WHY Edit's OWN READ-MODIFY-WRITE IS NOT ENOUGH FOR THE TABLE. *** Every other draft
        // writer changes a few rows, so re-reading the file and applying the change on top of it
        // keeps whatever Excel did meanwhile. The table cannot work that way: it holds a copy of
        // EVERY sheet, possibly for a long time, and saving it replaces all of them. If the
        // workbook changed underneath - edited in Excel, or by the label editor on another tab -
        // saving would silently put the old content back over those edits.
        //
        // So the caller passes the fingerprint (Fingerprint below) it took before reading, and a
        // save that finds anything else refuses with ChangedOnDisk rather than guessing. The table
        // then offers to reload.
        //
        // There is a window of milliseconds between this check and Edit's own read in which an
        // Excel save would still be overwritten. Closing it would mean parsing the very bytes that
        // were hashed, which BoardDataReader cannot do from memory today - and a contributor saving
        // in two programs within the same few milliseconds is not the case this guards against.
        // ###########################################################################################
        public static DraftWorkbookEditOutcome EditIfUnchanged(
            string draftsRoot,
            string excelDataFile,
            string expectedFingerprint,
            Func<BoardData, BoardData> transform)
        {
            ArgumentNullException.ThrowIfNull(expectedFingerprint);

            return DraftWorkbookStore.EditCore(draftsRoot, excelDataFile, expectedFingerprint, transform);
        }

        // ###########################################################################################
        // A SHA-256 of the draft workbook's bytes, as lowercase hex, or "" when there is no workbook
        // or it cannot be read.
        //
        // The CONTENT, not the modification time: a timestamp can move without a change (a
        // sync tool touching the file) and, on some file systems, fail to move within the same
        // second as one. Opened FileShare.ReadWrite so a workbook Excel has open can still be
        // hashed - Excel locks it against writing, not against reading.
        // ###########################################################################################
        public static string Fingerprint(string draftsRoot, string excelDataFile)
        {
            string workbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);
            if (workbook.Length == 0 || !File.Exists(workbook))
            {
                return string.Empty;
            }

            try
            {
                using var stream = new FileStream(workbook, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read the draft workbook [{workbook}] to fingerprint it - [{ex.Message}]");

                return string.Empty;
            }
        }

        private static DraftWorkbookEditOutcome EditCore(
            string draftsRoot,
            string excelDataFile,
            string? expectedFingerprint,
            Func<BoardData, BoardData> transform,
            Func<BoardData, BoardData, IReadOnlyDictionary<string, string?>>? followFor = null)
        {
            ArgumentNullException.ThrowIfNull(transform);

            string workbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);
            if (workbook.Length == 0 || !File.Exists(workbook))
            {
                CrtLog.Warning(
                    $"Cannot edit the draft for [{excelDataFile}] - no draft workbook at [{workbook}]");

                return DraftWorkbookEditOutcome.NoDraft;
            }

            if (expectedFingerprint is not null)
            {
                string actual = DraftWorkbookStore.Fingerprint(draftsRoot, excelDataFile);

                if (actual.Length == 0)
                {
                    return DraftWorkbookEditOutcome.WriteFailed;
                }

                if (!string.Equals(actual, expectedFingerprint, StringComparison.Ordinal))
                {
                    CrtLog.Info($"Not saving the draft for [{excelDataFile}] - its workbook changed since it was read");

                    return DraftWorkbookEditOutcome.ChangedOnDisk;
                }
            }

            BoardData? current = BoardDataReader.ReadWorkbookUncached(workbook);
            if (current is null)
            {
                CrtLog.Warning($"Cannot edit the draft for [{excelDataFile}] - its workbook could not be read");

                return DraftWorkbookEditOutcome.WriteFailed;
            }

            BoardData updated = transform(current);
            if (updated is null)
            {
                // A transform that declines the edit. Not an error - the label editor's own save
                // path can decide there is nothing to write - so this is silent and successful.
                return DraftWorkbookEditOutcome.Saved;
            }

            // ###########################################################################################
            // *** WHAT HANGS ON A SCHEMATIC'S NAME FOLLOWS IT (owner request, 2026-10-10: renaming an
            // image - "And fix all"). *** The sidecar files every highlight and the KiCad calibration
            // under a schematic's NAME. An edit that renamed a schematic - its "Board schematics" row
            // in the table, or its Name box - used to leave the highlights on the old name (an error
            // the checks refuse at submit, and nothing in the table can fix) and drop the calibration
            // (collected under the names AFTER the edit). One that deleted a schematic - in the table,
            // or with Remove - left its highlights behind, the same error.
            //
            // Now every edit follows them, with the server's own rule for a table edit
            // (SubmissionRowsBoard.FollowSchematics: a removed name and an added one on the same image
            // file are a rename; a removed name with no such twin was deleted). A rename moves them,
            // a deletion drops them. RenameSchematic says outright what it renamed (`followFor`).
            // ###########################################################################################
            IReadOnlyDictionary<string, string?> follow =
                followFor?.Invoke(current, updated) ??
                SubmissionRowsBoard.FollowSchematics(current.Schematics, updated.Schematics);

            if (follow.Count > 0)
            {
                updated = updated.WithComponentHighlights(SchematicRename.FollowHighlights(updated.ComponentHighlights, follow));
            }

            // *** AND WHAT HANGS ON A COMPONENT'S LABEL FOLLOWS IT (code review, 2026-10-13). *** A
            // label fixed in the table ("U8" -> "U08", nothing else changed) left its rectangles on
            // the old label. The rule is the count's (ComponentRelabel - BoardDataDiffer's pairing),
            // and the server's amendment follows it the same way.
            IReadOnlyDictionary<string, string> relabels = ComponentRelabel.Between(current.Components, updated.Components);

            if (relabels.Count > 0)
            {
                updated = updated.WithComponentHighlights(ComponentRelabel.FollowHighlights(updated.ComponentHighlights, relabels));
            }

            // Read BEFORE the workbook is replaced: the sidecar still files the calibrations under
            // the names the board had.
            List<KiCadCalibrationEntry> calibrations = DraftWorkbookStore.CalibrationsAfter(workbook, current, updated, follow);

            // The board's published renames, as the draft's copy has them - passed on as they
            // are. A draft does not add to them: the server records a rename when it publishes
            // one (SchematicRenameLog.ForPublish), and CRT follows a draft's own renames by the
            // image files it remembers instead (LocalSchematicRename).
            List<SchematicRenameRecord> renameLog = BoardComponentHighlightStorage.LoadSchematicRenames(workbook);

            return DraftWorkbookStore.WriteDraft(workbook, updated, calibrations, renameLog);
        }

        // ###########################################################################################
        // Writes a draft's board: the workbook, then its sidecar (highlights, calibrations, renames),
        // as ONE change - the workbook through a temporary file moved into place, and the workbook as
        // it was put back when the sidecar cannot be written. Shared by every edit (EditCore) and by
        // bringing a draft up to date (ReplaceIfUnchanged), so there is one way a draft is written.
        // ###########################################################################################
        private static DraftWorkbookEditOutcome WriteDraft(
            string workbook,
            BoardData updated,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            IReadOnlyList<SchematicRenameRecord> renameLog)
        {
            // A name CRT's bookkeeping check knows (code review, 2026-10-13): "<workbook>.writing.tmp"
            // stranded by a crash between the write and the move counted as the contributor's work.
            string temporary = DraftFolderLayout.TemporaryPathFor(workbook);

            // The workbook as it was, while the sidecar is written - see the catch below.
            byte[]? previousWorkbook = null;

            try
            {
                BoardWorkbookWriter.Write(temporary, updated);

                byte[] before = DraftWorkbookStore.ReadAllBytesShared(workbook);
                File.Move(temporary, workbook, overwrite: true);
                previousWorkbook = before;

                // ###########################################################################################
                // *** HIGHLIGHTS AND CALIBRATIONS LIVE IN THE SIDECAR, NOT THE WORKBOOK. ***
                //
                // BoardWorkbookSchema has no sheet for ComponentHighlights - BoardDataReader loads
                // them from the JSON beside the workbook (BoardComponentHighlightStorage), and
                // BoardWorkbookWriter silently drops them. So writing only the workbook would lose
                // every rectangle the label editor had just saved, with nothing to say why.
                //
                // Written through BoardSidecarWriter, the SAME writer the server publishes with, so
                // a draft sidecar and a published one are produced by one piece of code. It writes
                // the COMPLETE state, which is right here for the same reason it is right there: the
                // board being saved is the whole intended state, so a rectangle absent from it has
                // been deleted.
                //
                // Calibrations are read back off the existing sidecar and written through unchanged.
                // They are edited by their own path (BoardComponentHighlightStorage.SaveKiCadCalibration,
                // straight into this file) and are not part of BoardData at all, so a save that
                // passed none would erase the contributor's calibration work on the next label edit.
                // ###########################################################################################
                BoardSidecarWriter.Write(workbook, updated.ComponentHighlights, calibrations, renameLog);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The usual cause from the table editor is the workbook being open in Excel, which
                // on Windows locks it against the move above.
                CrtLog.Warning($"Could not write the draft workbook [{workbook}] - [{ex.Message}]");

                DraftWorkbookStore.TryDelete(temporary);

                // ###########################################################################################
                // *** A FAILED WRITE CHANGED NOTHING (code review, 2026-10-11). *** The sidecar is
                // written after the workbook, so a sidecar that could not be written (held open by an
                // editor, say) left the NEW workbook beside the OLD sidecar - and WriteFailed, which
                // every caller reads as "nothing written": "Schematic images" then deleted the images
                // it had copied, and the draft's new rows named files that were gone. The workbook as
                // it was is put back, so WriteFailed means what it says. The sidecar writer replaces
                // its file whole (FileReplacer), so a failed sidecar write left the old one in place.
                // ###########################################################################################
                if (previousWorkbook is not null)
                {
                    DraftWorkbookStore.PutBack(workbook, previousWorkbook);
                }

                return DraftWorkbookEditOutcome.WriteFailed;
            }

            return DraftWorkbookEditOutcome.Saved;
        }

        // A workbook's bytes, readable while Excel holds it open (it locks against writing only).
        private static byte[] ReadAllBytesShared(string path)
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var copy = new MemoryStream();
            stream.CopyTo(copy);

            return copy.ToArray();
        }

        // The workbook as it was before a write that did not finish - replaced whole, like the write.
        private static void PutBack(string workbook, byte[] previous)
        {
            try
            {
                FileReplacer.Replace(workbook, temporary => File.WriteAllBytes(temporary, previous));

                CrtLog.Info($"Put the draft workbook [{workbook}] back as it was - its sidecar could not be written");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Critical(
                    $"Could not put the draft workbook [{workbook}] back after its sidecar failed to write - the " +
                    $"workbook holds the new rows and the sidecar the old ones - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // The calibrations a write keeps: the sidecar's, collected under the names the board had
        // BEFORE the edit and moved or dropped with their schematic (`follow`) - plus any filed under
        // a name the edit brought in that the board did not have, which a collection of the old
        // names alone would not see (what Edit always wrote).
        //
        // Both sets come from ONE read of the sidecar (code review, 2026-10-10): it was parsed once
        // per schematic, twice over, on every save.
        // ###########################################################################################
        private static List<KiCadCalibrationEntry> CalibrationsAfter(
            string workbook,
            BoardData current,
            BoardData updated,
            IReadOnlyDictionary<string, string?> follow)
        {
            HashSet<string> NamesOf(BoardData board) => new(
                board.Schematics.Select(entry => entry.SchematicName?.Trim() ?? string.Empty).Where(name => name.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            HashSet<string> before = NamesOf(current);
            HashSet<string> after = NamesOf(updated);

            List<KiCadCalibrationEntry> filed = BoardComponentHighlightStorage.LoadKiCadCalibrations(
                workbook,
                before.Concat(after.Where(name => !before.Contains(name))));

            List<KiCadCalibrationEntry> kept = SchematicRename.FollowCalibrations(
                filed.Where(entry => before.Contains(entry.SchematicName)).OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase),
                follow);

            var named = new HashSet<string>(kept.Select(entry => entry.SchematicName.Trim()), StringComparer.OrdinalIgnoreCase);

            kept.AddRange(filed
                .Where(entry => after.Contains(entry.SchematicName) && !named.Contains(entry.SchematicName))
                .OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase));

            return kept;
        }

        // ###########################################################################################
        // Renames one schematic image of the draft - its "Board schematics" row, the highlights
        // drawn on it and its KiCad calibration together, in ONE write (SchematicRename has the
        // rules; owner request, 2026-10-10).
        //
        // *** IT SAYS WHAT IT RENAMED. *** Any edit follows a rename it can recognise (a removed
        // name and an added one on the same image file - see EditCore). This one does not have to
        // guess: it names the one rename, so an image whose file another image shares still keeps
        // its highlights and calibration - and when ANOTHER image keeps the old name, nothing moves
        // (SchematicRename.Apply). The row is found by its name AND its image file (SchematicRows),
        // so of two images sharing a name it is the one meant that is renamed.
        //
        // Judged against the draft as it is on disk NOW, like every edit here: a name taken
        // meanwhile (in Excel, say) is refused, and an image renamed or removed meanwhile is
        // NotFound. Nothing is written for any outcome but Renamed.
        // ###########################################################################################
        public static SchematicRenameOutcome RenameSchematic(
            string draftsRoot,
            string excelDataFile,
            SchematicRowRef row,
            string newName)
        {
            ArgumentNullException.ThrowIfNull(row);

            string wanted = newName?.Trim() ?? string.Empty;
            SchematicRenameOutcome outcome = SchematicRenameOutcome.Renamed;
            bool moved = false;
            string renamedFrom = string.Empty;

            DraftWorkbookEditOutcome written = DraftWorkbookStore.EditCore(
                draftsRoot,
                excelDataFile,
                expectedFingerprint: null,
                current =>
                {
                    int index = SchematicRows.IndexOf(current.Schematics, row);

                    if (index < 0)
                    {
                        outcome = SchematicRenameOutcome.NotFound;
                        return null!;
                    }

                    if (string.Equals(current.Schematics[index].SchematicName?.Trim(), wanted, StringComparison.Ordinal))
                    {
                        outcome = SchematicRenameOutcome.Unchanged;
                        return null!;
                    }

                    string? problem = SchematicRename.Problem(
                        wanted,
                        current.Schematics.Where((_, other) => other != index).Select(entry => entry.SchematicName));

                    if (problem is not null)
                    {
                        outcome = wanted.Length == 0 ? SchematicRenameOutcome.Blank : SchematicRenameOutcome.NameTaken;
                        return null!;
                    }

                    renamedFrom = current.Schematics[index].SchematicName ?? string.Empty;

                    (BoardData renamed, bool movedItsData) = SchematicRename.Apply(current, index, wanted);
                    moved = movedItsData;

                    return renamed;
                },
                (_, _) =>
                {
                    var follow = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

                    if (moved)
                    {
                        follow[renamedFrom] = wanted;
                    }

                    return follow;
                });

            return written switch
            {
                DraftWorkbookEditOutcome.NoDraft => SchematicRenameOutcome.NoDraft,
                DraftWorkbookEditOutcome.Saved => outcome,
                _ => SchematicRenameOutcome.WriteFailed,
            };
        }

        // ###########################################################################################
        // Replaces the draft's whole board - rows, highlights, calibrations, renames - but ONLY if the
        // workbook is still exactly what the caller read (`expectedFingerprint`, Fingerprint's value).
        // For bringing a draft up to date (DraftRebaser), which computes a new board from the one it
        // read and must not write over an edit made meanwhile in Excel. Written by WriteDraft, the
        // one way a draft is written.
        // ###########################################################################################
        internal static DraftWorkbookEditOutcome ReplaceIfUnchanged(
            string draftsRoot,
            string excelDataFile,
            string expectedFingerprint,
            BoardData board,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            IReadOnlyList<SchematicRenameRecord> renames)
        {
            ArgumentNullException.ThrowIfNull(expectedFingerprint);
            ArgumentNullException.ThrowIfNull(board);

            string workbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            if (workbook.Length == 0 || !File.Exists(workbook))
                return DraftWorkbookEditOutcome.NoDraft;

            string actual = DraftWorkbookStore.Fingerprint(draftsRoot, excelDataFile);

            if (actual.Length == 0)
                return DraftWorkbookEditOutcome.WriteFailed;

            if (!string.Equals(actual, expectedFingerprint, StringComparison.Ordinal))
            {
                CrtLog.Info($"Not bringing the draft for [{excelDataFile}] up to date - its workbook changed since it was read");
                return DraftWorkbookEditOutcome.ChangedOnDisk;
            }

            return DraftWorkbookStore.WriteDraft(workbook, board, calibrations ?? [], renames ?? []);
        }

        // ###########################################################################################
        // The marker after the draft was brought up to date with the published board: what it is now
        // based on (the revision the server is told, and the content the drift check compares), and
        // no "seen" warning left over. DraftRebaser calls this last, after the board and its base.
        // ###########################################################################################
        public static bool RecordBroughtUpToDate(string draftsRoot, string excelDataFile, string revision, string contentHash) =>
            DraftWorkbookStore.UpdateMarker(draftsRoot, excelDataFile, marker => marker with
            {
                BaseRevision = revision?.Trim() ?? string.Empty,
                BaseContentHash = contentHash ?? string.Empty,
                DriftSeenRevision = string.Empty,
                DriftSeenContentHash = string.Empty,
            });

        // ###########################################################################################
        // "I have looked at this" for a draft that cannot be brought up to date (no base - started
        // before CRT kept one). Quiets the warning for this published revision and NOTHING else
        // (code review, 2026-10-13): it used to overwrite BaseRevision, which the submission tells the
        // server, so a stale draft claimed to be made from the newest data.
        // ###########################################################################################
        //
        // `publishedContentHash` is the published files' content at that moment (DraftBaseStore
        // .PublishedContentHash): a draft judged by content is quiet while the published board
        // still holds exactly that (DraftMarker.DriftSeenContentHash).
        public static bool MarkDriftSeen(string draftsRoot, string excelDataFile, string publishedRevision, string publishedContentHash = "") =>
            DraftWorkbookStore.UpdateMarker(draftsRoot, excelDataFile, marker => marker with
            {
                DriftSeenRevision = publishedRevision?.Trim() ?? string.Empty,
                DriftSeenContentHash = publishedContentHash?.Trim() ?? string.Empty,
            });

        // Reads, changes and writes the draft's marker; false when there is none or it cannot be written.
        private static bool UpdateMarker(string draftsRoot, string excelDataFile, Func<DraftMarker, DraftMarker> change)
        {
            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);

            DraftMarker? marker = DraftMarkerStore.Load(markerPath);
            if (marker is null)
            {
                return false;
            }

            try
            {
                DraftMarkerStore.Save(markerPath, change(marker));

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not update the draft marker [{markerPath}] - [{ex.Message}]");

                return false;
            }
        }

        // ###########################################################################################
        // Deletes a draft outright - the whole board folder, workbook, sidecar, copied files and
        // marker together.
        //
        // "One folder is the whole draft", the same model DraftManager.DiscardDraft used and the
        // same one WorklogManager.DeleteWorkbook uses for a workbook. It matters more now than it
        // did: the folder holds copied image bytes as well as rows, and leaving those behind would
        // strand megabytes per discarded draft.
        //
        // *** THE MARKER GOES LAST (code review, 2026-09-27). *** A recursive delete removed
        // ".crt-draft.json" first - it sorts first - and then stopped at a workbook Excel had open.
        // That left the workbook without its marker, and since hand-placed folders are imported at
        // every start (DraftFolderImport), the next start took it in again as a fresh draft: a
        // discarded or retired draft came back. Now everything else goes first; a delete that stops
        // part-way leaves the marker, so the folder is still the same draft - on the Drafts tab,
        // where discarding it again finishes the job.
        //
        // *** AND THE WORKBOOK GOES LAST BUT ONE (owner report, 2026-10-02). *** Everything else was
        // still deleted in NAME order, and "Data ZX 4B v2.0.0.xlsx" sorts before "Issue 4.B.png": a
        // published draft being retired lost its workbook and sidecar, then stopped at an image that
        // something had open for a moment. A marker beside three images is no draft at all - nothing
        // can compare it with the published board, so it was never retired, and it read as "0 rows
        // changed" with the official data moved under it. Now the copied files go first, then the
        // workbook, then its sidecar, then the marker: a stop anywhere before the workbook leaves a
        // whole draft (any file already gone is found in the published tree, DraftFileResolver), and
        // the workbook before the sidecar because Excel holding the workbook is the usual stop - a
        // workbook left without its highlights and calibrations would read as changes.
        // ###########################################################################################
        public static bool Discard(string draftsRoot, string excelDataFile)
        {
            string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile);

            if (folder.Length == 0 || !Directory.Exists(folder))
            {
                return false;
            }

            try
            {
                DraftWorkbookStore.DeleteMarkerLast(folder, DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile));

                CrtLog.Info($"Discarded draft for [{excelDataFile}]");

                DraftWorkbookStore.RemoveEmptyParents(draftsRoot, folder);

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not discard the draft for [{excelDataFile}] - [{ex.Message}]");

                return false;
            }
        }

        // Everything in the folder but the workbook, its sidecar and the marker; then those three in
        // that order; then the (now empty) folder. See Discard for why the order is the design.
        private static void DeleteMarkerLast(string folder, string workbook)
        {
            string marker = Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName);
            string sidecar = workbook.Length > 0 ? BoardComponentHighlightStorage.GetJsonPath(workbook) : string.Empty;

            string[] last = [workbook, sidecar, marker];

            foreach (string entry in Directory.GetFileSystemEntries(folder))
            {
                if (last.Any(path => DraftWorkbookStore.IsSamePath(entry, path)))
                {
                    continue;
                }

                if (Directory.Exists(entry))
                {
                    Directory.Delete(entry, recursive: true);
                }
                else
                {
                    File.Delete(entry);
                }
            }

            foreach (string path in last)
            {
                if (path.Length > 0 && File.Exists(path))
                {
                    File.Delete(path);
                }
            }

            Directory.Delete(folder, recursive: false);
        }

        private static bool IsSamePath(string entry, string path) =>
            path.Length > 0
            && string.Equals(Path.GetFullPath(entry), Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Removes the folders a discarded draft leaves EMPTY above it - "Drafts/Commodore/C128" once
        // "310378 Open128" has gone, then "Drafts/Commodore" if that was all it held (owner report,
        // 2026-09-27: "there are some left-over folders").
        //
        // Only empty folders, and never the drafts root itself: a hardware folder still holding
        // another board, or a manufacturer folder holding anything else at all, stops the walk. An
        // empty folder removed here is recreated by the next draft that needs it. A failure is only
        // untidy, never a failed discard - the draft itself is already gone.
        // ###########################################################################################
        private static void RemoveEmptyParents(string draftsRoot, string removedFolder)
        {
            string root = Path.GetFullPath(draftsRoot);
            string? parent = Path.GetDirectoryName(Path.GetFullPath(removedFolder));

            while (!string.IsNullOrEmpty(parent))
            {
                string relative = Path.GetRelativePath(root, parent);

                // At or above the root: stop, and never touch the root.
                if (relative == "." || relative.StartsWith("..", StringComparison.Ordinal) || Path.IsPathRooted(relative))
                {
                    return;
                }

                try
                {
                    if (!Directory.Exists(parent) || Directory.EnumerateFileSystemEntries(parent).Any())
                    {
                        return;
                    }

                    Directory.Delete(parent, recursive: false);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Warning($"Could not remove the empty drafts folder [{parent}] - [{ex.Message}]");
                    return;
                }

                parent = Path.GetDirectoryName(parent);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A stranded temporary file is untidy, not harmful - the real write either
                // succeeded or has already been reported.
                CrtLog.Warning($"Could not remove the temporary draft file [{path}] - [{ex.Message}]");
            }
        }
    }

    // ###########################################################################################
    // How DraftWorkbookStore.EditIfUnchanged went. Kept apart from Edit's bool because the table
    // editor has to tell the contributor three different things: there is no draft any more, the
    // file changed and needs reloading, or the write itself failed (most often: open in Excel).
    // ###########################################################################################
    public enum DraftWorkbookEditOutcome
    {
        Saved,
        NoDraft,
        ChangedOnDisk,
        WriteFailed,

        // Never the store's answer: the table's save called off before anything was written - the
        // contributor kept the schematic images it would have removed (BoardTableEditor.SaveAsync).
        Cancelled
    }
}
