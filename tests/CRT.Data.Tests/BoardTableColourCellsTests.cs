using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardTableColourCells - which cells of the table hold a color, and what a changed one changed from
// and to (owner request, 2026-11-01: "If the color has changed, then I would like a popup showing the
// colors 'Changed from' and 'Changed to'").
// ###########################################################################################
public sealed class BoardTableColourCellsTests
{
    private static BoardSchematicEntry Picture(string color, string name = "Top") =>
        new() { SchematicName = name, SchematicImageFile = "Images/top.png", SchematicHighlightColor = color };

    private static BoardTableSheet SchematicsOf(BoardData? published, BoardData draft) =>
        BoardTableDocument.Create(published, draft).FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

    private static BoardTableCell Cell(BoardTableRow row, string column) =>
        row.Cells[row.Sheet.Columns.ToList().IndexOf(column)];

    [Theory]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, BoardWorkbookSchema.ColSchematicHighlightColor, true)]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, BoardWorkbookSchema.ColOppositeTraceHighlightColor, true)]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, BoardWorkbookSchema.ColThumbnailHighlightColor, true)]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, BoardWorkbookSchema.ColTraceHighlightColor, true)]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, BoardWorkbookSchema.ColSchematicHighlightOpacity, false)]
    [InlineData(BoardWorkbookSchema.SheetComponents, BoardWorkbookSchema.ColSchematicHighlightColor, false)]
    public void Only_the_four_color_columns_hold_a_color(string sheet, string column, bool isColour)
    {
        Assert.Equal(isColour, BoardTableColourCells.IsColourColumn(sheet, column));
    }

    [Fact]
    public void A_changed_color_cell_says_both_colors()
    {
        BoardTableSheet sheet = SchematicsOf(
            new BoardData { Schematics = [Picture(" DeepSkyBlue ")] },
            new BoardData { Schematics = [Picture("#1E90FF")] });

        BoardTableRow row = sheet.Rows.Single();

        Assert.Equal(
            new BoardTableColourChange("DeepSkyBlue", "#1E90FF"),
            BoardTableColourCells.Of(Cell(row, BoardWorkbookSchema.ColSchematicHighlightColor)));
    }

    // A side with no color of its own is blank - CRT draws it in its default.
    [Fact]
    public void A_color_set_or_cleared_has_a_blank_side()
    {
        BoardTableSheet set = SchematicsOf(
            new BoardData { Schematics = [Picture(string.Empty)] },
            new BoardData { Schematics = [Picture("Red")] });

        Assert.Equal(new BoardTableColourChange(string.Empty, "Red"), BoardTableColourCells.Of(Cell(set.Rows.Single(), BoardWorkbookSchema.ColSchematicHighlightColor)));

        BoardTableSheet cleared = SchematicsOf(
            new BoardData { Schematics = [Picture("Red")] },
            new BoardData { Schematics = [Picture(string.Empty)] });

        Assert.Equal(new BoardTableColourChange("Red", string.Empty), BoardTableColourCells.Of(Cell(cleared.Rows.Single(), BoardWorkbookSchema.ColSchematicHighlightColor)));
    }

    // No card for a color that did not change, for an added row, or for any other changed cell.
    [Fact]
    public void Only_a_color_that_changed_has_two_sides()
    {
        BoardTableSheet same = SchematicsOf(
            new BoardData { Schematics = [Picture("Red")] },
            new BoardData { Schematics = [Picture("Red")] });

        Assert.Null(BoardTableColourCells.Of(Cell(same.Rows.Single(), BoardWorkbookSchema.ColSchematicHighlightColor)));

        BoardTableSheet added = SchematicsOf(new BoardData(), new BoardData { Schematics = [Picture("Red")] });

        Assert.Null(BoardTableColourCells.Of(Cell(added.Rows.Single(), BoardWorkbookSchema.ColSchematicHighlightColor)));

        BoardTableSheet renamed = SchematicsOf(
            new BoardData { Schematics = [Picture("Red", "Top")] },
            new BoardData { Schematics = [Picture("Red", "Front")] });

        Assert.Null(BoardTableColourCells.Of(Cell(renamed.Rows.Single(), BoardWorkbookSchema.ColSchematicName)));
        Assert.Null(BoardTableColourCells.Of(null));
    }
}
