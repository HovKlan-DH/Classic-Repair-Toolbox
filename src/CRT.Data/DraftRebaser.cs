using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    public enum DraftRebaseOutcome
    {
        // The draft now holds the published update with the contributor's changes on it.
        BroughtUpToDate,

        // Nothing to do: the published board is still the one the draft was copied from.
        AlreadyUpToDate,

        // The draft has no base (started before CRT kept one, or a new board) - it cannot be merged.
        NoBase,

        NoDraft,
        DraftUnreadable,
        PublishedUnreadable,

        // The workbook changed while this ran (an Excel save) - nothing was written.
        ChangedOnDisk,

        // The workbook could not be written (usually: open in Excel) - nothing was changed.
        WriteFailed,

        // The draft HAS the update, but its new base could not be kept, so it is not noted as up
        // to date: the warning stays, and running again finishes the job (code review, 2026-10-13).
        NotNoted,
    }

    public sealed class DraftRebaseReport
    {
        public DraftRebaseOutcome Outcome { get; init; }

        public int ChangesKept { get; init; }

        public int FilesUpdated { get; init; }

        public IReadOnlyList<DraftRebaseConflict> Conflicts { get; init; } = [];
    }

    // ###########################################################################################
    // BRINGS ONE DRAFT UP TO DATE WITH THE PUBLISHED BOARD (code review, 2026-10-13) - the files
    // around DraftRebase's merge.
    //
    // Reads the base (what the draft was copied from), the draft and the published board, merges
    // them, and writes - in this order, so a stop anywhere leaves a draft that is still whole:
    //
    //   1. the draft's workbook and sidecar, as ONE guarded write (DraftWorkbookStore
    //      .ReplaceIfUnchanged - refused if Excel saved the workbook meanwhile, the previous
    //      workbook put back if the sidecar cannot be written);
    //   2. the board's own files the contributor did NOT change, replaced with the published bytes
    //      where the update replaced them, and files the update added copied in - exactly the copies
    //      a fresh draft would hold. A file the contributor changed is never touched;
    //   3. the new base (the published board as it is now);
    //   4. the marker - its BaseRevision, which the submission tells the server, and the content the
    //      drift check compares - ONLY once the new base is kept.
    //
    // Stopped after 1, 2 or 3, the next run merges again from the old base: everything the update
    // brought is then in the draft already and merges without a conflict. *** That is why 4 waits
    // for 3 (code review, 2026-10-13). *** A marker noting the new content beside the OLD base made
    // every later run answer "already up to date", and measured the contributor's changes - and the
    // files the update replaced - against the wrong board for good.
    // ###########################################################################################
    public static class DraftRebaser
    {
        public static DraftRebaseReport Rebase(string draftsRoot, string dataRoot, string excelDataFile)
        {
            DraftMarker? marker = DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile));

            if (marker is null)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoDraft };

            if (marker.IsNewBoard)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoBase };

            DraftBase? based = DraftBaseStore.Load(draftsRoot, excelDataFile);

            if (based is null)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoBase };

            string publishedWorkbook = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            string publishedContent = DraftBaseStore.PublishedContentHash(publishedWorkbook);

            if (publishedContent.Length > 0 && string.Equals(publishedContent, marker.BaseContentHash, StringComparison.Ordinal))
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.AlreadyUpToDate };

            string draftWorkbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);
            string fingerprint = DraftWorkbookStore.Fingerprint(draftsRoot, excelDataFile);

            if (fingerprint.Length == 0)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoDraft };

            BoardData? draft = BoardDataReader.ReadWorkbookUncached(draftWorkbook);

            if (draft is null)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.DraftUnreadable };

            BoardData? published = BoardDataReader.ReadWorkbookUncached(publishedWorkbook);

            if (published is null)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.PublishedUnreadable };

            DraftRebaseResult merged = DraftRebase.Merge(
                based.Board,
                draft,
                published,
                based.Calibrations,
                DraftBoardSource.CollectCalibrations(draftWorkbook, draft),
                DraftBoardSource.CollectCalibrations(publishedWorkbook, published));

            // 1. The board, guarded against an edit made while this ran.
            DraftWorkbookEditOutcome written = DraftWorkbookStore.ReplaceIfUnchanged(
                draftsRoot,
                excelDataFile,
                fingerprint,
                merged.Board,
                merged.Calibrations,
                BoardComponentHighlightStorage.LoadSchematicRenames(publishedWorkbook));

            switch (written)
            {
                case DraftWorkbookEditOutcome.Saved:
                    break;
                case DraftWorkbookEditOutcome.ChangedOnDisk:
                    return new DraftRebaseReport { Outcome = DraftRebaseOutcome.ChangedOnDisk };
                case DraftWorkbookEditOutcome.NoDraft:
                    return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoDraft };
                default:
                    return new DraftRebaseReport { Outcome = DraftRebaseOutcome.WriteFailed };
            }

            // 2. The files the contributor left alone follow the update.
            int filesUpdated = DraftRebaser.FollowPublishedFiles(draftsRoot, dataRoot, excelDataFile, based, merged.Board);

            // 3 and 4. The new base, then - only with it kept - the marker.
            bool noted;

            try
            {
                DraftBaseStore.Save(draftsRoot, excelDataFile, DraftBaseStore.Capture(dataRoot, excelDataFile, published, draftsRoot));
                noted = DraftWorkbookStore.RecordBroughtUpToDate(draftsRoot, excelDataFile, published.RevisionDate, publishedContent);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                // The board is merged; only the base is old. The marker is left as it was, so the
                // warning stays and the next run merges again from the old base - everything the
                // update brought is already in the draft, so nothing is lost.
                CrtLog.Warning($"Brought the draft for [{excelDataFile}] up to date, but could not keep its new base - [{ex.Message}]");
                noted = false;
            }

            if (!noted)
            {
                return new DraftRebaseReport
                {
                    Outcome = DraftRebaseOutcome.NotNoted,
                    ChangesKept = merged.ChangesKept,
                    FilesUpdated = filesUpdated,
                    Conflicts = merged.Conflicts,
                };
            }

            CrtLog.Info(
                $"Brought the draft for [{excelDataFile}] up to date with revision [{published.RevisionDate}] - " +
                $"{merged.ChangesKept} change(s) kept, {merged.Conflicts.Count} conflict(s), {filesUpdated} file(s) updated");

            return new DraftRebaseReport
            {
                Outcome = DraftRebaseOutcome.BroughtUpToDate,
                ChangesKept = merged.ChangesKept,
                FilesUpdated = filesUpdated,
                Conflicts = merged.Conflicts,
            };
        }

        // ###########################################################################################
        // The board's OWN files the merged board references: a copy the contributor did not change
        // (its bytes still the base's) takes the published bytes when the update replaced them, and
        // a file the draft has no copy of is copied in when the published tree has it - the copies a
        // fresh draft would hold. Another board's file is never copied (DraftFolderLayout), and a
        // copy the contributor changed is never touched. Returns how many files were written or
        // removed.
        //
        // *** A COPY NOTHING CITES ANY MORE GOES (code review, 2026-10-13). *** The update deleted
        // U8's image row, or pointed it at another file: the draft's untouched copy of the old file
        // stayed, the new base no longer listed it, and the Drafts row counted it as a file the
        // contributor had changed - Submit on, with nothing of theirs to send, and a draft that
        // could never be retired. A copy of a base file the merged board no longer cites is removed
        // while it still holds the base's bytes; one the contributor changed stays theirs.
        // ###########################################################################################
        private static int FollowPublishedFiles(string draftsRoot, string dataRoot, string excelDataFile, DraftBase based, BoardData merged)
        {
            var baseFiles = new Dictionary<string, DraftBaseFile>(StringComparer.Ordinal);

            foreach (DraftBaseFile file in based.Files)
                baseFiles.TryAdd(file.Path, file);

            int written = DraftRebaser.RemoveUncitedCopies(draftsRoot, excelDataFile, based, merged);

            foreach (string relative in SubmissionManifestBuilder.CollectReferencedFiles(merged))
            {
                string draftPath = DraftFolderLayout.GetReferencedFilePath(draftsRoot, excelDataFile, relative);
                string? withinBoard = DraftFolderLayout.RelativeToBoardFolder(excelDataFile, relative);

                if (draftPath.Length == 0 || withinBoard is null)
                    continue;

                string publishedPath = Path.Combine(dataRoot, relative.Replace('/', Path.DirectorySeparatorChar));

                if (!File.Exists(publishedPath))
                    continue;

                try
                {
                    if (File.Exists(draftPath))
                    {
                        // Changed by the contributor (or not from the base at all): theirs, untouched.
                        if (!baseFiles.TryGetValue(withinBoard, out DraftBaseFile? baseFile) ||
                            !string.Equals(DraftFingerprint.CachedFileHash(draftPath), baseFile.Sha256, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        if (string.Equals(DraftFingerprint.CachedFileHash(publishedPath), baseFile.Sha256, StringComparison.Ordinal))
                            continue; // The update did not replace it either.
                    }

                    string? folder = Path.GetDirectoryName(draftPath);

                    if (!string.IsNullOrWhiteSpace(folder))
                        Directory.CreateDirectory(folder);

                    string temporary = DraftFolderLayout.TemporaryPathFor(draftPath);
                    File.Copy(publishedPath, temporary, overwrite: true);
                    File.Move(temporary, draftPath, overwrite: true);
                    written++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // One file not refreshed is a file that still holds the older copy - shown, sent
                    // and fixable like any other. Not a reason to undo the merge.
                    CrtLog.Warning($"Could not bring the draft's copy of [{relative}] up to date - [{ex.Message}]");
                }
            }

            return written;
        }

        // The base files the merged board no longer cites whose draft copy still holds the base's
        // bytes - removed. Returns how many.
        private static int RemoveUncitedCopies(string draftsRoot, string excelDataFile, DraftBase based, BoardData merged)
        {
            var cited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string relative in SubmissionManifestBuilder.CollectReferencedFiles(merged))
            {
                if (DraftFolderLayout.RelativeToBoardFolder(excelDataFile, relative) is string withinBoard)
                    cited.Add(withinBoard);
            }

            string boardFolder = DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile);

            if (string.IsNullOrWhiteSpace(boardFolder))
                return 0;

            int removed = 0;

            foreach (DraftBaseFile file in based.Files)
            {
                if (cited.Contains(file.Path))
                    continue;

                string copy = Path.GetFullPath(Path.Combine(boardFolder, file.Path.Replace('/', Path.DirectorySeparatorChar)));

                // Inside the board's own draft folder, and still exactly what was copied.
                if (!copy.StartsWith(Path.GetFullPath(boardFolder) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                    !File.Exists(copy))
                {
                    continue;
                }

                try
                {
                    if (!string.Equals(DraftFingerprint.CachedFileHash(copy), file.Sha256, StringComparison.Ordinal))
                        continue; // The contributor's own bytes - theirs, untouched.

                    File.Delete(copy);
                    removed++;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Left behind it reads as a changed file - fixable by hand, not a reason to stop.
                    CrtLog.Warning($"Could not remove the draft's copy of [{file.Path}], which the board no longer cites - [{ex.Message}]");
                }
            }

            return removed;
        }
    }
}
