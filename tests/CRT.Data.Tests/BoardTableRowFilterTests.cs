using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which rows the table shows (owner request, 2026-10-02: the colour key's counts are clickable,
// "and even do remove the "Show changes only" so it works in the same unified way") -
// BoardTableRowFilter.
// ###########################################################################################
public sealed class BoardTableRowFilterTests
{
    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private static BoardTableRow Row(BoardTableSheet sheet, string label, bool deleted = false) =>
        sheet.Rows.First(row => row.IsDeleted == deleted && row.Cells[0].Text == label);

    // U1 modified, U2 deleted, U3 unchanged, U4 added, a second U4 a duplicate. No highlights, so
    // every live component also carries the "not marked on any schematic" warning.
    private static BoardTableSheet Sheet()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2"), Component("U3")),
            Board(Component("U1", "CPU 6510"), Component("U3"), Component("U4"), Component("U4")));

        return document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
    }

    [Fact]
    public void A_row_is_every_kind_it_is()
    {
        BoardTableSheet sheet = Sheet();

        Assert.Equal(BoardTableRowKinds.Modified | BoardTableRowKinds.Warnings, BoardTableRowFilter.KindsOf(Row(sheet, "U1")));
        Assert.Equal(BoardTableRowKinds.Deleted, BoardTableRowFilter.KindsOf(Row(sheet, "U2", deleted: true)));
        Assert.Equal(BoardTableRowKinds.Warnings, BoardTableRowFilter.KindsOf(Row(sheet, "U3")));
        // The first U4 is new - counted as added - and carries the duplicate's error too (every
        // row of a duplicate since 2026-10-03, owner decision: "it should show all rows"). The
        // second is not a change at all, only the server's error (a duplicate on Components);
        // "Flagged" was a kind of its own until 2026-10-03.
        Assert.Equal(
            BoardTableRowKinds.Added | BoardTableRowKinds.Errors | BoardTableRowKinds.Warnings,
            BoardTableRowFilter.KindsOf(sheet.Rows.First(row => row.Cells[0].Text == "U4")));
        Assert.Equal(
            BoardTableRowKinds.Errors | BoardTableRowKinds.Warnings,
            BoardTableRowFilter.KindsOf(sheet.Rows.Last(row => row.Cells[0].Text == "U4")));
    }

    [Fact]
    public void Nothing_picked_shows_every_row_and_a_pick_shows_rows_of_any_picked_kind()
    {
        BoardTableSheet sheet = Sheet();

        Assert.All(sheet.Rows, row => Assert.True(BoardTableRowFilter.Shows(BoardTableRowKinds.None, row)));

        Assert.True(BoardTableRowFilter.Shows(BoardTableRowKinds.Added | BoardTableRowKinds.Deleted, Row(sheet, "U2", deleted: true)));
        Assert.False(BoardTableRowFilter.Shows(BoardTableRowKinds.Added | BoardTableRowKinds.Modified, Row(sheet, "U3")));
    }

    // A new row nothing has been typed into is the contributor's work in progress - hiding it the
    // moment it appeared would read as the insert having failed.
    [Fact]
    public void A_new_empty_row_is_shown_whatever_is_picked()
    {
        BoardTableSheet sheet = Sheet();
        BoardTableRow blank = sheet.InsertRow(null);

        Assert.True(BoardTableRowFilter.Shows(BoardTableRowKinds.Errors, blank));
        Assert.True(sheet.HasRowsShownBy(BoardTableRowKinds.Errors));
    }

    // "Show changes only" without Flagged, which is warnings since 2026-10-03 (case 17 holds the rest).
    [Fact]
    public void The_changes_kinds_are_what_show_changes_only_showed()
    {
        Assert.Equal(
            BoardTableRowKinds.Added | BoardTableRowKinds.Modified | BoardTableRowKinds.Deleted,
            BoardTableRowFilter.Changes);
    }

    // A pick is saved by NAME, so a kind added later cannot shift what an old setting meant.
    [Fact]
    public void A_pick_is_written_by_name_and_read_back_and_an_unknown_name_is_skipped()
    {
        BoardTableRowKinds kinds = BoardTableRowKinds.Added | BoardTableRowKinds.Warnings;

        Assert.Equal("Added,Warnings", BoardTableRowFilter.Format(kinds));
        Assert.Equal(kinds, BoardTableRowFilter.Parse("Added,Warnings"));
        Assert.Equal(kinds, BoardTableRowFilter.Parse(" warnings , ADDED , Sparkly "));
        Assert.Equal(string.Empty, BoardTableRowFilter.Format(BoardTableRowKinds.None));
        Assert.Equal(BoardTableRowKinds.None, BoardTableRowFilter.Parse(null));
        Assert.Equal(BoardTableRowKinds.None, BoardTableRowFilter.Parse("63"));
    }

    // Case 17 (owner request, 2026-10-03; cases agreed with the project owner): "Flagged" became
    // warnings, so a Maintainer tab filter remembered with it picks Warnings - and "Show changes
    // only", read from an older setting, is the three kinds of change.
    [Fact]
    public void A_remembered_Flagged_filter_reads_back_as_Warnings()
    {
        Assert.Equal(BoardTableRowKinds.Warnings, BoardTableRowFilter.Parse("Flagged"));
        Assert.Equal(BoardTableRowKinds.Added | BoardTableRowKinds.Warnings, BoardTableRowFilter.Parse("Added,Flagged"));
        Assert.Equal("Added,Warnings", BoardTableRowFilter.Format(BoardTableRowFilter.Parse("Added, flagged")));

        Assert.Equal(BoardTableRowKinds.Added | BoardTableRowKinds.Modified | BoardTableRowKinds.Deleted, BoardTableRowFilter.Changes);
    }
}
