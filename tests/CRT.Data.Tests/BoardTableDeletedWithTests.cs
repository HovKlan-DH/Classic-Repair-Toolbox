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

    // ------------------------------------------------------------------ Several components at once (2026-10-02)

    // Two components deleted together are one sentence: the labels joined, each sheet's rows added
    // up in the order the sheets came, "Their" highlights.
    [Fact]
    public void Several_components_are_one_sentence_with_each_sheets_rows_added_up()
    {
        BoardTableDeletedWith? deleted = BoardTableDeletedWith.Combine(
        [
            new BoardTableDeletedWith("U8", [new BoardTableSheetCount("Component images", 2), new BoardTableSheetCount("Component links", 1)], Highlights: 2),
            new BoardTableDeletedWith("U9", [new BoardTableSheetCount("Component images", 1), new BoardTableSheetCount("Component local files", 1)], Highlights: 1),
            new BoardTableDeletedWith("U10", [], Highlights: 0)
        ]);

        Assert.Equal(
            "Also deleted with U8, U9 and U10: 3 rows on Component images, 1 on Component links and 1 on Component local files. " +
            "Their 3 highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back.",
            deleted!.Describe());
    }

    // Two regional variants of one component deleted together: the first finds the rows to take,
    // both count its highlights - which are counted once, and the component named once.
    [Fact]
    public void A_component_met_twice_is_named_once_and_its_highlights_counted_once()
    {
        BoardTableDeletedWith? deleted = BoardTableDeletedWith.Combine(
        [
            new BoardTableDeletedWith("U8", [new BoardTableSheetCount("Component images", 3)], Highlights: 2),
            new BoardTableDeletedWith("u8", [], Highlights: 2)
        ]);

        Assert.Equal(
            "Also deleted with U8: 3 rows on Component images. Its 2 highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back.",
            deleted!.Describe());
    }

    // Several components with only highlights between them; one part is returned as it is; none is null.
    [Fact]
    public void Several_with_only_highlights_name_the_components_and_one_or_none_is_unchanged()
    {
        BoardTableDeletedWith? highlightsOnly = BoardTableDeletedWith.Combine(
        [
            new BoardTableDeletedWith("R12", [], Highlights: 1),
            new BoardTableDeletedWith("R13", [], Highlights: 2)
        ]);

        Assert.Equal(
            "The 3 highlights of R12 and R13 on the schematics go when you save. Undo (Ctrl+Z) brings it all back.",
            highlightsOnly!.Describe());

        var one = new BoardTableDeletedWith("R12", [], Highlights: 1);
        Assert.Same(one, BoardTableDeletedWith.Combine([one]));
        Assert.Null(BoardTableDeletedWith.Combine([]));
    }
}
