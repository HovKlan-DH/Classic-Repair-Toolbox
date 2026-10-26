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
    // *** A CRT 4.0 DRAFT (2026-10-24). *** Reads the base (the published board file as it was
    // copied), the draft's board file and the published one, merges the rows (DraftRebase) and the
    // file lists (DraftRebaseFiles), and writes - in this order, so a stop anywhere leaves a draft
    // that is still whole:
    //
    //   1. the draft's board file, as ONE guarded write (DraftWorkbookStore.ReplaceIfUnchanged -
    //      refused if the draft changed meanwhile). Its file list follows the update where the
    //      contributor left a file as copied - the bytes stay in the tree's store, nothing is copied;
    //   2. the new base (the published board file as it is now);
    //   3. the marker - its BaseRevision, which the submission tells the server, and the content the
    //      drift check compares - ONLY once the new base is kept.
    //
    // Stopped after 1 or 2, the next run merges again from the old base: everything the update
    // brought is then in the draft already and merges without a conflict. *** That is why 3 waits
    // for 2 (code review, 2026-10-13). *** A marker noting the new content beside the OLD base made
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

            BoardDocument? baseDocument = DraftBaseStore.LoadDocument(draftsRoot, excelDataFile);

            if (baseDocument is null)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoBase };

            string publishedPath = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            string publishedContent = DraftBaseStore.PublishedContentHash(publishedPath);

            if (publishedContent.Length > 0 && string.Equals(publishedContent, marker.BaseContentHash, StringComparison.Ordinal))
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.AlreadyUpToDate };

            string draftPath = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);
            string fingerprint = DraftWorkbookStore.Fingerprint(draftsRoot, excelDataFile);

            if (fingerprint.Length == 0)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.NoDraft };

            if (BoardFiles.ReadDocument(draftPath) is not BoardDocument draftDocument)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.DraftUnreadable };

            if (BoardFiles.ReadDocument(publishedPath) is not BoardDocument publishedDocument)
                return new DraftRebaseReport { Outcome = DraftRebaseOutcome.PublishedUnreadable };

            BoardContent based = BoardDocumentMapping.ToContent(baseDocument);
            BoardContent draft = BoardDocumentMapping.ToContent(draftDocument);
            BoardContent published = BoardDocumentMapping.ToContent(publishedDocument);

            // A file the contributor replaced - in the draft's list, or lying in its folder as a new
            // file not yet taken in - is theirs, and so are the rows citing it.
            string draftFolder = Path.GetDirectoryName(draftPath)!;
            HashSet<string> replaced = DraftRebaseFiles.Replaced(
                baseDocument.Files,
                draftDocument.Files,
                name => File.Exists(Path.Combine(draftFolder, name.Replace('/', Path.DirectorySeparatorChar))));

            DraftRebaseResult merged = DraftRebase.Merge(
                DraftRebaseFiles.MarkReplaced(based.Board, replaced),
                draft.Board,
                published.Board,
                based.Calibrations,
                draft.Calibrations,
                published.Calibrations);

            List<BoardDocument.StoredFileRow> files = DraftRebaseFiles.Merge(
                baseDocument.Files,
                draftDocument.Files,
                publishedDocument.Files,
                replaced,
                out int filesUpdated);

            // 1. The board, guarded against an edit made while this ran.
            DraftWorkbookEditOutcome written = DraftWorkbookStore.ReplaceIfUnchanged(
                draftsRoot,
                excelDataFile,
                fingerprint,
                merged.Board,
                merged.Calibrations,
                published.Renames,
                files);

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

            // 2 and 3. The new base, then - only with it kept - the marker.
            bool noted;

            try
            {
                DraftBaseStore.SaveDocument(draftsRoot, excelDataFile, publishedDocument);
                noted = DraftWorkbookStore.RecordBroughtUpToDate(draftsRoot, excelDataFile, publishedDocument.RevisionDate, publishedContent);
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
                $"Brought the draft for [{excelDataFile}] up to date with revision [{publishedDocument.RevisionDate}] - " +
                $"{merged.ChangesKept} change(s) kept, {merged.Conflicts.Count} conflict(s), {filesUpdated} file(s) updated");

            return new DraftRebaseReport
            {
                Outcome = DraftRebaseOutcome.BroughtUpToDate,
                ChangesKept = merged.ChangesKept,
                FilesUpdated = filesUpdated,
                Conflicts = merged.Conflicts,
            };
        }
    }
}
