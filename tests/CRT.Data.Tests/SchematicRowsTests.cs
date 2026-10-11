using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SchematicRows - finding the "Board schematics" row a change in the "Schematic images" window is
// for, by its name AND its image file (code review, 2026-10-10). By name alone, a rename, a CAD name
// or a removal on the second of two images sharing a name went to the first - or, for a removal,
// to both.
// ###########################################################################################
public sealed class SchematicRowsTests
{
    private static BoardSchematicEntry Entry(string name, string file, string cadName = "") =>
        new() { SchematicName = name, SchematicImageFile = $"Commodore/C64/250407/{file}", CadName = cadName };

    private static BoardData Board(params BoardSchematicEntry[] schematics)
    {
        var board = new BoardData { HardwareName = "C64", BoardName = "250407" };
        board.Schematics.AddRange(schematics);
        return board;
    }

    private static readonly BoardSchematicEntry[] TwoTops =
        [Entry("Top", "top-a.png"), Entry("Bottom", "bottom.png"), Entry("Top", "top-b.png")];

    [Fact]
    public void A_row_is_found_by_its_name_and_its_image_file()
    {
        Assert.Equal(2, SchematicRows.IndexOf(TwoTops, new SchematicRowRef("top", "Commodore\\C64\\250407\\TOP-B.png")));
        Assert.Equal(0, SchematicRows.IndexOf(TwoTops, new SchematicRowRef("Top", "Commodore/C64/250407/top-a.png")));

        // The name right and the file not, or the other way round: not this board's row.
        Assert.Equal(-1, SchematicRows.IndexOf(TwoTops, new SchematicRowRef("Top", "Commodore/C64/250407/bottom.png")));
        Assert.Equal(-1, SchematicRows.IndexOf(TwoTops, new SchematicRowRef("Gone", "Commodore/C64/250407/top-a.png")));
    }

    // Rows sharing their name AND their file are told apart by their place among them.
    [Fact]
    public void Rows_sharing_name_and_file_are_told_apart_by_their_place()
    {
        BoardSchematicEntry[] twins = [Entry("Top", "top.png", "First"), Entry("Clock", "clock.png"), Entry("Top", "top.png", "Second")];
        var shown = twins.Select(entry => (entry.SchematicName, entry.SchematicImageFile)).ToList();

        SchematicRowRef second = SchematicRows.RefOf(shown, 2);

        Assert.Equal(1, second.Occurrence);
        Assert.Equal(2, SchematicRows.IndexOf(twins, second));
        Assert.Equal(0, SchematicRows.IndexOf(twins, SchematicRows.RefOf(shown, 0)));

        // Two rows naming no file at all are the same "file" here.
        Assert.True(SchematicRows.SameImageFile("", "  "));
    }

    [Fact]
    public void A_CAD_name_goes_to_the_row_meant_and_every_other_value_stays()
    {
        (BoardData board, bool found) = SchematicRows.WithCadName(
            Board(TwoTops),
            new SchematicRowRef("Top", "Commodore/C64/250407/top-b.png"),
            " Open128 - PCB Top ");

        Assert.True(found);
        Assert.Equal(["", "", "Open128 - PCB Top"], board.Schematics.Select(entry => entry.CadName));
        Assert.Equal("C64", board.HardwareName);
        Assert.Equal("250407", board.BoardName);
    }

    // Removing one of two rows sharing a name took both.
    [Fact]
    public void Removing_a_row_removes_that_row_only()
    {
        (BoardData board, bool found) = SchematicRows.Without(
            Board(TwoTops),
            new SchematicRowRef("Top", "Commodore/C64/250407/top-b.png"));

        Assert.True(found);
        Assert.Equal(
            ["Commodore/C64/250407/top-a.png", "Commodore/C64/250407/bottom.png"],
            board.Schematics.Select(entry => entry.SchematicImageFile));
        Assert.Equal("250407", board.BoardName);
    }

    [Fact]
    public void A_row_the_board_does_not_have_changes_nothing()
    {
        BoardData board = Board(TwoTops);
        var gone = new SchematicRowRef("Gone", "Commodore/C64/250407/gone.png");

        Assert.Equal((board, false), SchematicRows.WithCadName(board, gone, "root"));
        Assert.Equal((board, false), SchematicRows.Without(board, gone));
    }
}
