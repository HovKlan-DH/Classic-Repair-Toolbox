using System;
using System.IO;
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
        // failure leaves the previous, complete draft.
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
            Func<BoardData, BoardData> transform)
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

            string temporary = workbook + ".writing.tmp";

            try
            {
                BoardWorkbookWriter.Write(temporary, updated);
                File.Move(temporary, workbook, overwrite: true);

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
                BoardSidecarWriter.Write(
                    workbook,
                    updated.ComponentHighlights,
                    DraftBoardSource.CollectCalibrations(workbook, updated));

                return DraftWorkbookEditOutcome.Saved;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The usual cause from the table editor is the workbook being open in Excel, which
                // on Windows locks it against the move above.
                CrtLog.Warning($"Could not write the draft workbook [{workbook}] - [{ex.Message}]");

                DraftWorkbookStore.TryDelete(temporary);

                return DraftWorkbookEditOutcome.WriteFailed;
            }
        }

        // ###########################################################################################
        // Updates the draft's marker - used to rebase a draft onto a newer published revision after
        // the contributor has been shown the drift.
        //
        // Separate from Edit because it touches no board rows: rebasing is a statement about WHICH
        // published revision the draft is measured against, not about the board's contents.
        // ###########################################################################################
        public static bool SetBaseRevision(string draftsRoot, string excelDataFile, string baseRevision)
        {
            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);

            DraftMarker? marker = DraftMarkerStore.Load(markerPath);
            if (marker is null)
            {
                return false;
            }

            try
            {
                DraftMarkerStore.Save(markerPath, new DraftMarker
                {
                    SystemKey = marker.SystemKey,
                    BaseRevision = baseRevision?.Trim() ?? string.Empty,
                    NewSystem = marker.NewSystem,
                    CreatedUtc = marker.CreatedUtc,
                });

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not update the draft marker [{markerPath}] - [{ex.Message}]");

                return false;
            }
        }

        // ###########################################################################################
        // Deletes a draft outright - the whole system folder, workbook, sidecar, copied files and
        // marker together.
        //
        // "One folder is the whole draft", the same model DraftManager.DiscardDraft used and the
        // same one WorklogManager.DeleteWorkbook uses for a workbook. It matters more now than it
        // did: the folder holds copied image bytes as well as rows, and leaving those behind would
        // strand megabytes per discarded draft.
        // ###########################################################################################
        public static bool Discard(string draftsRoot, string excelDataFile)
        {
            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);

            if (folder.Length == 0 || !Directory.Exists(folder))
            {
                return false;
            }

            try
            {
                Directory.Delete(folder, recursive: true);

                CrtLog.Info($"Discarded draft for [{excelDataFile}]");

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not discard the draft for [{excelDataFile}] - [{ex.Message}]");

                return false;
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
        WriteFailed
    }
}
