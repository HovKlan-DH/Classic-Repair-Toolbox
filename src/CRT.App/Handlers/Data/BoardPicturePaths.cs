using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE A BOARD'S PICTURES LIE ON DISK - the Schematics tab's own rule, for every surface that
    // draws or exports a board's pictures: the Schematics tab's thumbnails (Main.BoardSelection.cs),
    // the Workbooks tab's board pane, a workbook's export and a device's history.
    //
    // In CRT 4.0 a picture row NAMES a file of its board ("Images/top.png"), its bytes stored under
    // their checksum - found by DraftFileResolver; an Excel-era path under the data root still
    // resolves. The Workbooks tab joined the name to the data root by itself and, on 4.0 data, found
    // no picture at all: no board on the pane, none in an export (bug found 2026-10-29).
    // ###########################################################################################
    public static class BoardPicturePaths
    {
        // ###########################################################################################
        // `file` of the board `boardIdentity` ("<id>/board.json") as the board is VIEWED -
        // `draftBoardFolder` is the viewed draft's (DraftManager.GetViewedBoardFolder), empty with
        // none - else where an Excel-era row put it under the data root, there or not.
        // ###########################################################################################
        public static string PathOf(string dataRoot, string draftBoardFolder, string file, string? boardIdentity, BoardFileListings? listings = null) =>
            DraftFileResolver.ResolveWithSource(dataRoot, draftBoardFolder, file, boardIdentity, listings)?.FullPath
                ?? Path.Combine(dataRoot, file.Replace('/', Path.DirectorySeparatorChar));

        // A picture of a board, where it lies.
        public sealed record PlacedPicture(string Name, string Path);

        // ###########################################################################################
        // Every picture of `board` with a file, in the board's order, each where it lies - resolved
        // ONCE each, for a caller that needs them in more than one shape (Main's board load: the
        // Schematics tab's thumbnails and the Workbooks tab's map - code review, 2026-10-29). Its
        // board file is read once for them all - or not at all, when `listings` read it already.
        // ###########################################################################################
        public static List<PlacedPicture> Place(
            BoardData? board,
            string dataRoot,
            string draftBoardFolder,
            string? boardIdentity,
            BoardFileListings? listings = null)
        {
            var placed = new List<PlacedPicture>();
            listings ??= new BoardFileListings();

            foreach (BoardSchematicEntry picture in board?.Schematics ?? [])
            {
                if (!string.IsNullOrWhiteSpace(picture.SchematicImageFile))
                {
                    placed.Add(new PlacedPicture(
                        picture.SchematicName,
                        PathOf(dataRoot, draftBoardFolder, picture.SchematicImageFile, boardIdentity, listings)));
                }
            }

            return placed;
        }

        // ###########################################################################################
        // Every picture with a file, by its name - the first row of a name, as the board pane takes
        // it (a board's pictures are not checked for two of one name).
        // ###########################################################################################
        public static Dictionary<string, string> ByName(IEnumerable<PlacedPicture> placed)
        {
            var paths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (PlacedPicture picture in placed)
            {
                if (!string.IsNullOrWhiteSpace(picture.Name))
                {
                    paths.TryAdd(picture.Name, picture.Path);
                }
            }

            return paths;
        }

        // The same, worked out from the board.
        public static Dictionary<string, string> ByName(
            BoardData? board,
            string dataRoot,
            string draftBoardFolder,
            string? boardIdentity,
            BoardFileListings? listings = null) =>
            ByName(Place(board, dataRoot, draftBoardFolder, boardIdentity, listings));
    }
}
