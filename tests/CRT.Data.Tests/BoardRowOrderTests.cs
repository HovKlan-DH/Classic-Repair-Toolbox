using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardRowOrder - whether a draft holds the rows it shares with the published board in another
// order: the change BoardDataDiffer does not count, which a draft made only to look must still not
// be thrown away with (code review, 2026-10-28).
// ###########################################################################################
public sealed class BoardRowOrderTests
{
    private static BoardData Board(params string[] pictures) => new()
    {
        Schematics = [.. pictures.Select(name => new BoardSchematicEntry { SchematicName = name, SchematicImageFile = $"{name}.png" })],
        Components = [new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "U2" }],
    };

    [Fact]
    public void Pictures_moved_are_another_order()
    {
        Assert.True(BoardRowOrder.DiffersFrom(Board("Top", "Bottom"), Board("Bottom", "Top")));
    }

    [Fact]
    public void The_same_rows_in_the_same_order_are_no_other_order()
    {
        Assert.False(BoardRowOrder.DiffersFrom(Board("Top", "Bottom"), Board("top", "Bottom")));
    }

    // A row added or removed is a change the differ counts - not an order of its own.
    [Fact]
    public void A_row_added_or_removed_leaves_the_others_order_alone()
    {
        Assert.False(BoardRowOrder.DiffersFrom(Board("Top", "Bottom"), Board("Top", "Side", "Bottom")));
        Assert.False(BoardRowOrder.DiffersFrom(Board("Top", "Side", "Bottom"), Board("Top", "Bottom")));
    }

    // Every sheet counts, not only the pictures - a component row dragged in the table too.
    [Fact]
    public void Components_moved_are_another_order_too()
    {
        BoardData moved = Board("Top");
        moved.Components.Reverse();

        Assert.True(BoardRowOrder.DiffersFrom(Board("Top"), moved));
    }
}
