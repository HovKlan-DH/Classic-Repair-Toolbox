using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHICH file holds a board's published board (NewContributeStrategy.md Phase 5, task 3).
    //
    // Split out of PublishedBoardReader because this is the part worth testing: resolving a
    // board folder from untrusted identity values, and deciding what "there is no board" means.
    // The reader around it only opens the file.
    //
    // *** THE IDENTITY IS UNTRUSTED, EVEN HERE. *** Manufacturer/Hardware/Board arrive in a
    // submission, and a manufacturer of "../.." would otherwise relocate the board folder
    // outside the data tree - which on a READ means disclosing an arbitrary file's contents to a
    // maintainer, and would make this endpoint a file-disclosure hole rather than a review screen.
    // So every part goes through SubmissionPathRules exactly as the write paths do.
    //
    // *** IN THE CRT 4.0 TREE A BOARD IS ITS BOARD FILE (2026-10-25). *** "<board id>/board.json"
    // (DataTreeLayout.BoardFilePath), always that name - no generations to choose between, no
    // workbook names that differ from the folder's. `WorkbookPath` is that file: the shared readers
    // and writers read and write a board.json as a 4.0 board (DataTreeLayout.IsBoardFile), so every
    // caller that used to hand them the workbook hands them this.
    // ###########################################################################################
    public static class PublishedBoardLocator
    {
        // ###########################################################################################
        // Where this board's published board is, and whether it is there at all.
        //
        // A BLANK OR UNSAFE IDENTITY YIELDS "does not exist" rather than throwing. The manifest's
        // own validation already reports a bad identity with a message the contributor can act
        // on; throwing here would turn a maintainer opening that submission into an error page
        // instead of showing them the findings that explain it.
        // ###########################################################################################
        public static PublishedBoardLocation Locate(string? dataTreeRoot, SubmissionManifest manifest)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return PublishedBoardLocator.Locate(dataTreeRoot, manifest.Manufacturer, manifest.Hardware, manifest.Board);
        }

        // ###########################################################################################
        // The same, from a board ID ("Commodore/C64/250407") - what the review QUEUE holds, which
        // loads no manifest (2026-09-26: the queue's "New board" badge). An id that is not a
        // well-formed one locates nothing.
        // ###########################################################################################
        public static PublishedBoardLocation LocateBoard(string? dataTreeRoot, string? boardId)
        {
            if (!BoardDescriptorRules.IsValidBoardId(boardId))
                return PublishedBoardLocation.None;

            string[] parts = boardId!.Split('/');

            return PublishedBoardLocator.Locate(dataTreeRoot, parts[0], parts[1], parts[2]);
        }

        private static PublishedBoardLocation Locate(string? dataTreeRoot, string? manufacturer, string? hardware, string? board)
        {
            if (string.IsNullOrWhiteSpace(dataTreeRoot))
                return PublishedBoardLocation.None;

            string[] parts = new[] { manufacturer, hardware, board }
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .Select(part => part!)
                .ToArray();

            if (parts.Length != 3)
                return PublishedBoardLocation.None;

            // The tree's own folders are never a board's (DataTreeLayout.ReservedTopFolders).
            if (DataTreeLayout.ReservedTopFolders.Contains(parts[0], StringComparer.OrdinalIgnoreCase))
                return PublishedBoardLocation.None;

            string relative = string.Join('/', parts);

            if (!SubmissionPathRules.TryResolve(dataTreeRoot, DataTreeLayout.BoardFilePath(relative), out string boardFile, out _))
                return PublishedBoardLocation.None;

            if (!File.Exists(boardFile))
                return PublishedBoardLocation.None;

            return new PublishedBoardLocation(
                boardFile,
                Path.GetDirectoryName(boardFile)!,
                Generation: null,
                Exists: true);
        }
    }

    // ###########################################################################################
    // Where a published board is - or that there is none, which is a NEW BOARD and a first-class
    // answer rather than a failure. `WorkbookPath` is the board file ("<board id>/board.json");
    // Generation is the Excel trees' and always null in the 4.0 tree.
    // ###########################################################################################
    public sealed record PublishedBoardLocation(
        string WorkbookPath,
        string BoardFolder,
        Version? Generation,
        bool Exists)
    {
        public static readonly PublishedBoardLocation None =
            new(string.Empty, string.Empty, null, Exists: false);
    }
}
