using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE PICTURES OF A BOARD MAIN LOADED, where each lies on disk (BoardPicturePaths' rule) - worked
    // out by that load, off the UI thread, against what it read: the board, its board file and the
    // draft folder it was viewed from (code reviews, 2026-10-29).
    //
    // Kept by Main and handed to the Workbooks tab (its board pane, a workbook's export, a device's
    // history). The map answers only for the board it was made for (PathsFor): drawn against
    // another board's data - a later load landing first, a field set without the other - a picture
    // would be named by one board and found through another. Worked out again from the drop-downs
    // at draw time, that is what happened between a board change and its load.
    // ###########################################################################################
    public sealed class LoadedBoardPictures
    {
        private readonly BoardData thisBoard;
        private readonly IReadOnlyDictionary<string, string> thisPaths;

        public LoadedBoardPictures(BoardData board, IReadOnlyDictionary<string, string> paths)
        {
            this.thisBoard = board ?? throw new ArgumentNullException(nameof(board));
            this.thisPaths = paths ?? throw new ArgumentNullException(nameof(paths));
        }

        // Every picture of `board` with a file, by its name - when `board` is the one these were
        // worked out for; else null, no answer.
        public IReadOnlyDictionary<string, string>? PathsFor(BoardData? board) =>
            ReferenceEquals(board, this.thisBoard) ? this.thisPaths : null;
    }
}
