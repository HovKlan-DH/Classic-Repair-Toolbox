using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PictureColours - how a marked component is drawn on one of the board's pictures (CRT 4.0 step 2,
// slice 4b), as the board file keeps it: colours by name or "#RRGGBB", opacities as "30%" - and the
// edit that sets them (BoardPartEdits.WithPictureColours), picture by picture, by name AND image file.
// ###########################################################################################
public sealed class PictureColoursTests
{
    private static BoardData Board() => new()
    {
        Schematics =
        [
            new BoardSchematicEntry { SchematicName = "Top", CadName = "PCB Top", SchematicImageFile = "top.png", SchematicHighlightColor = "Red", SchematicHighlightOpacity = "30%", ThumbnailHighlightColor = "Red", ThumbnailHighlightOpacity = "100%" },
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top-rev7.png", SchematicHighlightColor = "Red" },
            new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "bottom.png" },
        ],
        Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
    };

    // A fraction or a percentage, written as a whole percentage, clamped.
    [Theory]
    [InlineData(0.3, "30%")]
    [InlineData(0.255, "26%")]
    [InlineData(1.0, "100%")]
    [InlineData(0.0, "0%")]
    [InlineData(45.0, "45%")]
    [InlineData(250.0, "100%")]
    [InlineData(-0.5, "0%")]
    public void An_opacity_is_written_as_a_whole_percentage(double value, string expected)
    {
        Assert.Equal(expected, PictureColours.OpacityText(value));
    }

    // A picture's colours, read and written back, keep every other column of the picture.
    [Fact]
    public void A_pictures_colours_are_read_and_written_as_one_value()
    {
        BoardSchematicEntry top = Board().Schematics[0];

        Assert.Equal(new PictureColours("Red", "30%", "", "Red", "100%"), top.ReadColours());

        BoardSchematicEntry recoloured = top.WithColours(new PictureColours("#FF8000", "40%", "DodgerBlue", "Orange", "50%"));

        Assert.Equal(("Top", "PCB Top", "top.png"), (recoloured.SchematicName, recoloured.CadName, recoloured.SchematicImageFile));
        Assert.Equal(new PictureColours("#FF8000", "40%", "DodgerBlue", "Orange", "50%"), recoloured.ReadColours());
    }

    // Only the picture of that name AND image file changes - two pictures may share a name.
    [Fact]
    public void Only_the_picture_of_that_name_and_file_is_recoloured()
    {
        BoardData board = Board();

        BoardData changed = BoardPartEdits.WithPictureColours(board, [("top", "TOP-REV7.png", new PictureColours("Blue", "20%", "", "", ""))]);

        Assert.Equal("Red", changed.Schematics[0].SchematicHighlightColor);
        Assert.Equal("Blue", changed.Schematics[1].SchematicHighlightColor);
        Assert.Equal("20%", changed.Schematics[1].SchematicHighlightOpacity);
        Assert.Same(board.Schematics[2], changed.Schematics[2]);
        Assert.Same(board.Components, changed.Components);
    }

    // Colours picked as they were - spaces aside - are no change: the same board, so no step.
    [Fact]
    public void Colours_as_they_were_return_the_same_board()
    {
        BoardData board = Board();

        Assert.Same(board, BoardPartEdits.WithPictureColours(board, [("Top", "top.png", new PictureColours(" Red ", "30%", "", "Red", "100%"))]));
        Assert.Same(board, BoardPartEdits.WithPictureColours(board, [("Nowhere", "x.png", new PictureColours("Blue", "", "", "", ""))]));
    }
}
