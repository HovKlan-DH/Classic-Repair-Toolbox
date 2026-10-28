using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardPictureEdits - the board's pictures changed on the Contribute tab (CRT 4.0 step 3 - the plan:
// "Pictures: today's Schematic images content, moved inline"): pictures added under names no picture
// has, and a picture moved one place. The free-name rule came from the "Schematic images" window
// with its tests.
// ###########################################################################################
public sealed class BoardPictureEditsTests
{
    private static BoardSchematicEntry Picture(string name, string file) => new() { SchematicName = name, SchematicImageFile = file };

    private static BoardData Board(params BoardSchematicEntry[] pictures)
    {
        var board = new BoardData();
        board.Schematics.AddRange(pictures);
        return board;
    }

    // ---- names --------------------------------------------------------------------------------

    [Fact]
    public void A_free_name_is_the_desired_one_when_nothing_has_taken_it()
    {
        Assert.Equal("Main board", BoardPictureEdits.FreeName("Main board", []));
        Assert.Equal("Main board", BoardPictureEdits.FreeName("  Main board  ", []));
    }

    [Fact]
    public void A_taken_name_gets_a_number_climbing_past_the_names_taken()
    {
        Assert.Equal("Main board (2)", BoardPictureEdits.FreeName("Main board", ["Main board"]));
        Assert.Equal("Main board (4)", BoardPictureEdits.FreeName("Main board", ["Main board", "Main board (2)", "Main board (3)"]));
    }

    [Fact]
    public void Capitals_are_no_difference_when_a_name_is_taken()
    {
        Assert.Equal("Main board (2)", BoardPictureEdits.FreeName("Main board", ["MAIN BOARD"]));
    }

    // A file called ".png" has no name to take - it still needs one, or its row would have a blank key
    // that could never be found or removed again.
    [Fact]
    public void A_blank_name_falls_back_to_a_usable_one()
    {
        Assert.Equal("Board image", BoardPictureEdits.FreeName("   ", []));
        Assert.Equal("Board image (2)", BoardPictureEdits.FreeName(null, ["Board image"]));
    }

    // Pictures added together never share a name - two "top.png" from two folders included.
    [Fact]
    public void Pictures_added_together_are_named_after_their_files_each_one_free()
    {
        Assert.Equal(
            ["top", "top (2)", "Bottom", "Board image"],
            BoardPictureEdits.NamesFor(["Scans/top.png", "Old/top.JPG", "Bottom.webp", ".png"], []));

        Assert.Equal(["Top (2)"], BoardPictureEdits.NamesFor(["Top.png"], ["top"]));
    }

    // ---- adding -------------------------------------------------------------------------------

    [Fact]
    public void Pictures_added_come_after_the_last()
    {
        BoardData board = Board(Picture("Top", "top.png"));

        BoardData result = BoardPictureEdits.WithAdded(board, [("Bottom", "bottom.png"), (" Sheet 1 ", " sheet1.png ")]);

        Assert.Equal(["Top", "Bottom", "Sheet 1"], result.Schematics.Select(row => row.SchematicName));
        Assert.Equal(["top.png", "bottom.png", "sheet1.png"], result.Schematics.Select(row => row.SchematicImageFile));
        Assert.Same(board, BoardPictureEdits.WithAdded(board, []));
    }

    // ---- moving -------------------------------------------------------------------------------

    [Fact]
    public void A_picture_moves_one_place_up_or_down()
    {
        BoardData board = Board(Picture("Top", "top.png"), Picture("Bottom", "bottom.png"), Picture("Sheet", "sheet.png"));

        Assert.Equal(
            ["Top", "Sheet", "Bottom"],
            BoardPictureEdits.Moved(board, new SchematicRowRef("Sheet", "sheet.png"), -1).Schematics.Select(row => row.SchematicName));

        Assert.Equal(
            ["Bottom", "Top", "Sheet"],
            BoardPictureEdits.Moved(board, new SchematicRowRef("Top", "top.png"), +1).Schematics.Select(row => row.SchematicName));
    }

    // At an end, or a picture the draft no longer has: the very board - no step.
    [Fact]
    public void A_picture_at_an_end_or_gone_does_not_move()
    {
        BoardData board = Board(Picture("Top", "top.png"), Picture("Bottom", "bottom.png"));

        Assert.Same(board, BoardPictureEdits.Moved(board, new SchematicRowRef("Top", "top.png"), -1));
        Assert.Same(board, BoardPictureEdits.Moved(board, new SchematicRowRef("Bottom", "bottom.png"), +1));
        Assert.Same(board, BoardPictureEdits.Moved(board, new SchematicRowRef("Gone", "gone.png"), +1));
        Assert.Same(board, BoardPictureEdits.Moved(board, new SchematicRowRef("Top", "top.png"), 0));
    }

    // Of two pictures sharing a name (an error a table edit can make), the one moved is the one asked.
    [Fact]
    public void Of_two_pictures_sharing_a_name_the_one_asked_moves()
    {
        BoardData board = Board(Picture("Top", "a.png"), Picture("Bottom", "bottom.png"), Picture("Top", "b.png"));

        BoardData result = BoardPictureEdits.Moved(board, new SchematicRowRef("Top", "b.png"), -1);

        Assert.Equal(["a.png", "b.png", "bottom.png"], result.Schematics.Select(row => row.SchematicImageFile));
    }
}
