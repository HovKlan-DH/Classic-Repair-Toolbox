using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // A BOARD'S SCHEMATIC IMAGES RENAMED - this machine's worklogs and settings that name them
    // follow (owner decisions, 2026-10-10: renaming an image; a published rename "must also change
    // for everyone updating to this data"; the image file kept with it "to be unambiguous").
    //
    // Whenever a board is SHOWN, LocalSchematicRename.CatchUp follows every name the user's own
    // data uses to wherever its image FILE is now, by the board as CRT last showed it - which
    // covers a rename published, the user's own draft rename, and the board shown without that
    // draft again. The first time, the board's published list (its .json) is used instead.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // The board just loaded for display: the user's worklogs and settings caught up with its
        // images, BEFORE the thumbnails, the hidden images and the worklogs are read from them. The
        // published renames are in the PUBLISHED board's .json - the list every CRT that syncs the
        // board reads - whichever workbook (draft or published) is being shown. That file is read
        // only when CatchUp needs it, not on every board load.
        // ###########################################################################################
        private void CatchUpOnSchematicRenames(HardwareBoardEntry entry, BoardData boardData) =>
            LocalSchematicRename.CatchUp(
                entry.HardwareName,
                entry.BoardName,
                SchematicRenameLog.ImagesOf(boardData),
                () => BoardComponentHighlightStorage.LoadSchematicRenames(this.GetCurrentBoardExcelPath()));
    }
}
