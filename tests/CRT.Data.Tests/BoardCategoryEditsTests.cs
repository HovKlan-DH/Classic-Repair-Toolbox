using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// "BOARD: CATEGORIES" (owner request, 2026-11-01): rename, add and delete - and no component left with
// none. A category is the word its components carry, so these are changes of those rows.
// ###########################################################################################
public sealed class BoardCategoryEditsTests
{
    private static ComponentEntry Component(string label, string category, string region = "") =>
        new() { BoardLabel = label, Category = category, Region = region, FriendlyName = label + " name" };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    // ###########################################################################################
    // Their order IS which is seen first in the rows (owner request, 2026-11-01: "rearrange the category
    // up/down, alike Files or Links"): a category's rows move together, in the order they had, and what
    // every list of categories reads follows.
    // ###########################################################################################
    [Fact]
    public void Categories_put_in_another_order_move_their_components_rows_together()
    {
        BoardData board = Board(
            Component("U1", "IC"),
            Component("C1", "Capacitor"),
            Component("U2", "IC"),
            Component("R1", "Resistor"),
            Component("C2", "Capacitor"));

        BoardData moved = BoardCategoryEdits.InOrder(board, ["Resistor", "IC", "Capacitor"]);

        Assert.Equal(["R1", "U1", "U2", "C1", "C2"], moved.Components.Select(row => row.BoardLabel));
        Assert.Equal(["Resistor", "IC", "Capacitor"], BoardCategoryEdits.Of(moved).Select(row => row.Name));
        Assert.Equal(["Resistor", "IC", "Capacitor"], ComponentListBuilder.BuildDistinctCategories(moved));

        // Rows of one component stay as they stood - a regional twin beside its twin.
        BoardData twins = BoardCategoryEdits.InOrder(
            Board(Component("U9", "IC", "PAL"), Component("U9", "IC", "NTSC"), Component("C1", "Capacitor")),
            ["Capacitor", "IC"]);

        Assert.Equal(["C1|", "U9|PAL", "U9|NTSC"], twins.Components.Select(row => $"{row.BoardLabel}|{row.Region}"));
    }

    // No step when nothing moves - even with a category's rows standing apart, which only a change of
    // order is a reason to gather.
    [Fact]
    public void Categories_already_in_that_order_are_the_same_board()
    {
        BoardData board = Board(Component("U1", "IC"), Component("C1", "Capacitor"), Component("U2", "IC"));

        Assert.Same(board, BoardCategoryEdits.InOrder(board, ["IC", "Capacitor"]));
        Assert.Same(board, BoardCategoryEdits.InOrder(board, ["ic", " capacitor "]));
    }

    // The components with none stay first, as Of lists them; a category not named comes last.
    [Fact]
    public void Components_with_no_category_stay_first_and_a_category_not_named_comes_last()
    {
        BoardData board = Board(Component("U1", "IC"), Component("X1", ""), Component("C1", "Capacitor"), Component("R1", "Resistor"));

        BoardData moved = BoardCategoryEdits.InOrder(board, ["Capacitor", "IC"]);

        Assert.Equal(["X1", "C1", "U1", "R1"], moved.Components.Select(row => row.BoardLabel));
    }

    [Fact]
    public void Every_category_is_listed_once_with_how_many_components_carry_it_in_first_seen_order()
    {
        List<BoardCategoryRow> rows = BoardCategoryEdits.Of(Board(
            Component("C1", "Capacitor"),
            Component("U1", "IC"),
            Component("C2", "capacitor")));

        Assert.Equal(
            [new BoardCategoryRow("Capacitor", 2), new BoardCategoryRow("IC", 1)],
            rows);
    }

    // The VIC-II has a row per region, and is one component.
    [Fact]
    public void A_component_with_a_row_per_region_is_counted_once()
    {
        List<BoardCategoryRow> rows = BoardCategoryEdits.Of(Board(Component("U19", "IC", "PAL"), Component("U19", "IC", "NTSC")));

        Assert.Equal([new BoardCategoryRow("IC", 1)], rows);
    }

    // What needs doing comes first: the checks refuse a component with none.
    [Fact]
    public void The_components_with_no_category_are_listed_first()
    {
        List<BoardCategoryRow> rows = BoardCategoryEdits.Of(Board(Component("U1", "IC"), Component("U2", "  ")));

        Assert.Equal([new BoardCategoryRow(string.Empty, 1), new BoardCategoryRow("IC", 1)], rows);
    }

    [Fact]
    public void A_rename_moves_every_component_under_it_and_nothing_else()
    {
        BoardData renamed = BoardCategoryEdits.Renamed(Board(Component("U1", "IC"), Component("U2", "ic"), Component("C1", "Capacitor")), "IC", "Chip");

        Assert.Equal(["Chip", "Chip", "Capacitor"], renamed.Components.Select(row => row.Category));
        Assert.Equal(["U1", "U2", "C1"], renamed.Components.Select(row => row.BoardLabel));
        Assert.Equal("U1 name", renamed.Components[0].FriendlyName);
    }

    // No split into "Capacitor" and "capacitor" - the component list groups by the exact word.
    [Fact]
    public void Named_as_another_category_the_two_become_one_under_its_spelling()
    {
        BoardData merged = BoardCategoryEdits.Renamed(Board(Component("C1", "Capacitor"), Component("C2", "Caps")), "Caps", "capacitor");

        Assert.Equal(["Capacitor", "Capacitor"], merged.Components.Select(row => row.Category));
    }

    [Fact]
    public void Capitals_alone_are_a_change_of_its_own_spelling()
    {
        BoardData renamed = BoardCategoryEdits.Renamed(Board(Component("U1", "ic")), "ic", "IC");

        Assert.Equal("IC", renamed.Components[0].Category);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("IC")]
    public void A_blank_name_or_the_same_name_changes_nothing(string to)
    {
        BoardData board = Board(Component("U1", "IC"));

        Assert.Same(board, BoardCategoryEdits.Renamed(board, "IC", to));
    }

    [Fact]
    public void The_components_with_no_category_go_under_the_one_typed()
    {
        BoardData fixedUp = BoardCategoryEdits.Renamed(Board(Component("U1", ""), Component("U2", "IC")), "", "Misc");

        Assert.Equal(["Misc", "IC"], fixedUp.Components.Select(row => row.Category));
    }

    [Fact]
    public void A_rename_of_a_category_nothing_carries_is_the_same_board()
    {
        BoardData board = Board(Component("U1", "IC"));

        Assert.Same(board, BoardCategoryEdits.Renamed(board, "Gone", "Other"));
    }

    // Only the word changes: a category is no order, and Submit counts a moved row as a change.
    [Fact]
    public void Rows_stay_where_they_stand()
    {
        BoardData board = Board(Component("U1", "IC"), Component("C1", "Capacitor"), Component("U2", "IC"));

        BoardData renamed = BoardCategoryEdits.Renamed(board, "Capacitor", "IC");

        Assert.Equal(["U1", "C1", "U2"], renamed.Components.Select(row => row.BoardLabel));
        Assert.False(BoardRowOrder.DiffersFrom(board, renamed));
    }
}
