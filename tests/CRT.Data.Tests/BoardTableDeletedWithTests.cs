using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardTableDeletedWith - the sentence the board table shows when a deleted component took rows on
// other sheets and its highlights with it (owner request, 2026-09-25). Those sheets are not
// the one on screen, so this line is the only place the user learns what else went.
// ###########################################################################################
public sealed class BoardTableDeletedWithTests
{
    [Fact]
    public void Rows_on_several_sheets_and_highlights_are_all_named()
    {
        var deleted = new BoardTableDeletedWith(
            "U8",
            [new BoardTableSheetCount("Component images", 3), new BoardTableSheetCount("Component local files", 1), new BoardTableSheetCount("Component links", 2)],
            Highlights: 2);

        Assert.Equal(
            "Also deleted with U8: 3 rows on Component images, 1 on Component local files and 2 on Component links. " +
            "Its 2 highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back.",
            deleted.Describe());
    }

    [Fact]
    public void One_row_and_one_highlight_read_in_the_singular()
    {
        var deleted = new BoardTableDeletedWith("C7", [new BoardTableSheetCount("Component images", 1)], Highlights: 1);

        Assert.Equal(
            "Also deleted with C7: 1 row on Component images. Its 1 highlight on the schematics goes when you save. Undo (Ctrl+Z) brings it all back.",
            deleted.Describe());
    }

    // A component with nothing on the other sheets but rectangles on a schematic.
    [Fact]
    public void Only_highlights_are_named_by_the_component()
    {
        var deleted = new BoardTableDeletedWith("R12", [], Highlights: 3);

        Assert.Equal(
            "R12's 3 highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back.",
            deleted.Describe());
    }
}
