namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A DRAFT CHANGES, READ FROM WHERE IT IS (CRT 4.0 step 1's last slice, 2026-10-25) - for
    // "Review changes" and the Submit and Discard draft dialogs, which show one summary
    // (DraftChangeSummary):
    //
    //   - the draft as every editor shows it: its session's content when one is open
    //     (DraftManager.Sessions - a change the table holds unsaved is in it), else its board file;
    //   - the published board file (none for a board that exists only as a draft), with both board
    //     files' lists of files, so a picture replaced under its own name and the KiCad data are said.
    //
    // Reads files: run it off the UI thread.
    // ###########################################################################################
    internal static class DraftChangeSummaryReader
    {
        public static DraftChangeSummary Read(
            string dataRoot,
            string draftsRoot,
            string excelDataFile,
            bool draftOnly,
            BoardData? publishedOverride = null)
        {
            string draftBoardFile = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            BoardContent? draft = DraftManager.Sessions.Find(draftsRoot, excelDataFile)?.Content
                                  ?? BoardFiles.ReadContent(draftBoardFile);

            if (draft is null)
            {
                return DraftChangeSummary.Empty;
            }

            if (publishedOverride is not null)
            {
                return DraftChangeSummary.Build(new BoardContent(publishedOverride, [], []), draft);
            }

            string publishedPath = draftOnly ? string.Empty : DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            BoardContent? published = publishedPath.Length == 0 ? null : BoardFiles.ReadContent(publishedPath);

            return DraftChangeSummary.Build(
                published,
                draft,
                published is null ? null : DraftFiles.ListedIn(publishedPath),
                DraftFiles.ListedIn(draftBoardFile));
        }

        // The published board of a draft - null for one that exists only as a draft, or with no
        // published board file - for "Undo changes to <component>" (ComponentRevert).
        public static BoardData? PublishedBoard(string dataRoot, string excelDataFile, bool draftOnly)
        {
            if (draftOnly)
            {
                return null;
            }

            string publishedPath = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            return publishedPath.Length == 0 ? null : BoardFiles.ReadContent(publishedPath)?.Board;
        }
    }
}
