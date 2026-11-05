using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// "COMPONENT IMAGES" SHOWN IN TWO (owner request, 2026-11-03: "should probably be divided into two now, as
// it now has 'Pictures' and 'Oscilloscope baseline' seperately"): one sheet, as the board file and the
// server keep it; a row is a baseline by its pin alone.
// ###########################################################################################
public sealed class BoardTableSheetPartsTests
{
    private static BoardData Board() => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/a.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/b.png" },
        ],
    };

    private static BoardTableSheet Images(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!;

    private static BoardTableDocument Document(BoardData? published = null) =>
        BoardTableDocument.Create(published ?? Board(), Board());

    private static BoardTableCell Cell(BoardTableRow row, string column) =>
        row.Cells[row.Sheet.Columns.ToList().IndexOf(column)];

    // What says whether a step really changed a row: its values and whether it is deleted, wherever it stands.
    [Fact]
    public void A_row_stands_in_a_sheet_holding_the_same_values_in_the_same_state()
    {
        BoardTableSheet before = Images(Document());

        var published = new BoardData { Components = Board().Components, ComponentImages = [Board().ComponentImages[1]] };
        BoardTableSheet moved = Images(BoardTableDocument.Create(published, new BoardData { Components = Board().Components, ComponentImages = [Board().ComponentImages[1]] }));

        Assert.True(BoardTableSheetParts.StandsIn(before, moved.Rows[0]));

        Cell(moved.Rows[0], BoardWorkbookSchema.ColName).Text = "Changed";
        Assert.False(BoardTableSheetParts.StandsIn(before, moved.Rows[0]));

        // The same values, one of them deleted: a step took it away or brought it back.
        BoardTableSheet gone = Images(BoardTableDocument.Create(Board(), new BoardData { Components = Board().Components, ComponentImages = [Board().ComponentImages[0]] }));
        BoardTableRow ghost = gone.Rows.Single(row => row.IsDeleted);

        Assert.False(BoardTableSheetParts.StandsIn(before, ghost));
    }

    [Fact]
    public void The_row_nearest_a_place_the_part_shows_is_looked_for_down_then_up()
    {
        BoardData board = new()
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "A", File = "a.png" },
                new ComponentImageEntry { BoardLabel = "U8", Name = "B", File = "b.png" },
                new ComponentImageEntry { BoardLabel = "U8", Name = "C", File = "c.png" },
                new ComponentImageEntry { BoardLabel = "U8", Pin = "2", Name = "D", File = "d.png" },
                new ComponentImageEntry { BoardLabel = "U8", Name = "E", File = "e.png" },
            ],
        };

        BoardTableSheet sheet = Images(BoardTableDocument.Create(board, board));

        string NameAt(BoardTableSheetPart part, int index) => Cell(BoardTableSheetParts.NearestShown(sheet, part, index)!, BoardWorkbookSchema.ColName).Text;

        Assert.Equal("D", NameAt(BoardTableSheetPart.Baselines, 1));
        Assert.Equal("D", NameAt(BoardTableSheetPart.Baselines, 3));
        Assert.Equal("D", NameAt(BoardTableSheetPart.Baselines, 4));
        Assert.Equal("B", NameAt(BoardTableSheetPart.Pictures, 0));
        Assert.Equal("E", NameAt(BoardTableSheetPart.Pictures, 3));

        // None of that part: none.
        var onlyPictures = new BoardData { Components = board.Components, ComponentImages = [board.ComponentImages[1]] };

        Assert.Null(BoardTableSheetParts.NearestShown(Images(BoardTableDocument.Create(onlyPictures, onlyPictures)), BoardTableSheetPart.Baselines, 0));
    }

    [Fact]
    public void A_row_is_a_baseline_by_its_pin_alone()
    {
        BoardTableSheet sheet = Images(Document());

        Assert.Equal(BoardTableSheetPart.Pictures, BoardTableSheetParts.PartOf(sheet.Rows[0]));
        Assert.Equal(BoardTableSheetPart.Baselines, BoardTableSheetParts.PartOf(sheet.Rows[1]));

        // The rule the rest of CRT goes by, read off the row.
        Assert.False(ComponentImageKinds.IsScopeBaseline(Board().ComponentImages[0]));
        Assert.True(ComponentImageKinds.IsScopeBaseline(Board().ComponentImages[1]));
    }

    [Fact]
    public void Each_part_shows_its_own_rows_and_the_whole_sheet_shows_them_all()
    {
        BoardTableSheet sheet = Images(Document());

        Assert.Equal(
            [sheet.Rows[0]],
            BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.None, BoardTableSearch.None, rowInHand: null, BoardTableSheetPart.Pictures));

        Assert.Equal(
            [sheet.Rows[1]],
            BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.None, BoardTableSearch.None, rowInHand: null, BoardTableSheetPart.Baselines));

        Assert.Equal(2, BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.None, BoardTableSearch.None, rowInHand: null).Count);
    }

    // A pin typed moves the row across - and the row the cursor is on stays until the cursor leaves it, or
    // the next text typed would land on another row.
    [Fact]
    public void A_pin_typed_moves_the_row_across_once_the_cursor_has_left_it()
    {
        BoardTableSheet sheet = Images(Document());
        BoardTableRow picture = sheet.Rows[0];

        Cell(picture, BoardWorkbookSchema.ColPin).Text = "2";
        sheet.Refresh();

        Assert.Equal(BoardTableSheetPart.Baselines, BoardTableSheetParts.PartOf(picture));
        Assert.DoesNotContain(picture, BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.None, BoardTableSearch.None, rowInHand: null, BoardTableSheetPart.Pictures));
        Assert.Contains(picture, BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.None, BoardTableSearch.None, picture, BoardTableSheetPart.Pictures));
    }

    // Somebody's new row: hidden the moment it appeared, the insert would seem to have failed.
    [Theory]
    [InlineData(BoardTableSheetPart.Pictures)]
    [InlineData(BoardTableSheetPart.Baselines)]
    public void A_new_row_with_nothing_in_it_shows_in_either_part(BoardTableSheetPart part)
    {
        BoardTableSheet sheet = Images(Document());

        BoardTableRow added = sheet.InsertRow(sheet.Rows[0]);

        Assert.True(BoardTableSheetParts.Shows(part, added));
    }

    // ###########################################################################################
    // *** A ROW THAT CAN STILL BE A BASELINE SHOWS WHERE ITS PIN CAN BE TYPED (code review, 2026-11-04). ***
    // "Pictures" has no Pin column, and the pane's page no pin field, so a row that was only a picture by
    // then could not be given one in the table: Ctrl+Z or deleting it was all that got out. It stays a
    // picture by its pin alone, and shows among the pictures too - so that table loses nothing.
    // ###########################################################################################
    private static BoardTableSheet ImagesOf(params ComponentImageEntry[] rows)
    {
        var board = new BoardData { Components = Board().Components, ComponentImages = [.. rows] };

        return Images(BoardTableDocument.Create(board, board));
    }

    [Theory]
    [InlineData(BoardWorkbookSchema.ColTimeDiv, "50us")]
    [InlineData(BoardWorkbookSchema.ColExpectedOscilloscopeReading, "DRAM FAIL")]
    public void A_row_with_no_pin_holding_what_only_a_baseline_has_shows_among_the_baselines_too(string column, string value)
    {
        BoardTableSheet sheet = ImagesOf(
            new ComponentImageEntry { BoardLabel = "U8", Name = "Odd", File = "a.png" },
            new ComponentImageEntry { BoardLabel = "U8", Name = "Plain", File = "b.png" });

        Cell(sheet.Rows[0], column).Text = value;
        sheet.Refresh();

        Assert.Equal(BoardTableSheetPart.Pictures, BoardTableSheetParts.PartOf(sheet.Rows[0]));
        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Pictures, sheet.Rows[0]));
        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Baselines, sheet.Rows[0]));

        Assert.False(BoardTableSheetParts.Shows(BoardTableSheetPart.Baselines, sheet.Rows[1]));
    }

    // A label says nothing of which it is, a name neither - a baseline has one too. A file or a note on a
    // row with no pin says picture.
    [Fact]
    public void A_row_saying_nothing_of_what_it_is_shows_in_either_part_until_it_does()
    {
        BoardTableSheet sheet = Images(Document());
        BoardTableRow row = sheet.InsertRow(sheet.Rows[0]);

        Cell(row, BoardWorkbookSchema.ColBoardLabel).Text = "U9";
        Cell(row, BoardWorkbookSchema.ColName).Text = "VCC";
        sheet.Refresh();

        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Baselines, row));
        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Pictures, row));

        Cell(row, BoardWorkbookSchema.ColFile).Text = "Component images/c.png";
        sheet.Refresh();

        Assert.False(BoardTableSheetParts.Shows(BoardTableSheetPart.Baselines, row));
        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Pictures, row));
    }

    // What the editor holds of a row worked on among the baselines, while the table stands.
    [Fact]
    public void A_row_kept_with_the_baselines_shows_there_with_no_pin()
    {
        BoardTableSheet sheet = ImagesOf(new ComponentImageEntry { BoardLabel = "U8", Name = "Picked", File = "a.png" });
        BoardTableRow row = sheet.Rows[0];

        Assert.False(BoardTableSheetParts.Shows(BoardTableSheetPart.Baselines, row));

        row.StaysWithBaselines = true;

        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Baselines, row));
        Assert.True(BoardTableSheetParts.Shows(BoardTableSheetPart.Pictures, row));
    }

    // ###########################################################################################
    // *** "PICTURES" JUDGES ONLY THE COLUMNS IT SHOWS (code review, 2026-11-04). *** A row listed for a word
    // or a problem in a column it hides shows nothing that says why - and the cursor, sent to a column the
    // grid does not have, went nowhere, so what was typed next had no cell to go to. Such a row shows among
    // the baselines too, every cell in sight, and the table goes there as a pill goes to its rows.
    // ###########################################################################################
    [Fact]
    public void A_search_among_the_pictures_finds_only_what_they_show()
    {
        var board = new BoardData
        {
            Components = Board().Components,
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Name = "Pinout", File = "a.png", TimeDiv = "50us" }],
        };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet sheet = Images(document);
        BoardTableRow row = sheet.Rows[0];

        List<BoardTableRow> Listed(string text, BoardTableSheetPart part) =>
            BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.None, BoardTableSearch.For(document, text), rowInHand: null, part);

        Assert.DoesNotContain(row, Listed("50us", BoardTableSheetPart.Pictures));
        Assert.Contains(row, Listed("50us", BoardTableSheetPart.Baselines));
        Assert.Equal(BoardTableSheetPart.Baselines, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Pictures, BoardTableRowKinds.None, BoardTableSearch.For(document, "50us")));

        // What it shows is found as ever.
        Assert.Contains(row, Listed("Pinout", BoardTableSheetPart.Pictures));
        Assert.Equal(BoardTableSheetPart.Pictures, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Pictures, BoardTableRowKinds.None, BoardTableSearch.For(document, "Pinout")));

        // A word left out, in a column it hides: nothing there to leave it out for.
        Assert.Contains(row, Listed("-50us", BoardTableSheetPart.Pictures));
        Assert.DoesNotContain(row, Listed("-50us", BoardTableSheetPart.Baselines));

        // The whole sheet, as the Maintainer tab shows it, is searched whole.
        Assert.Contains(row, Listed("50us", BoardTableSheetPart.Whole));
    }

    [Fact]
    public void A_pill_among_the_pictures_goes_only_by_problems_in_the_columns_they_show()
    {
        var board = new BoardData
        {
            Components = Board().Components,
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Name = "Pinout", File = "a.png", TimeDiv = "nonsense" }],
        };

        BoardTableSheet sheet = Images(BoardTableDocument.Create(board, board));
        BoardTableRow row = sheet.Rows[0];

        Assert.True(row.HasWarnings);

        Assert.Null(BoardTableRowFilter.FirstProblemCell(BoardTableRowKinds.Warnings, row, BoardTableSheetPart.Pictures));
        Assert.Same(Cell(row, BoardWorkbookSchema.ColTimeDiv), BoardTableRowFilter.FirstProblemCell(BoardTableRowKinds.Warnings, row, BoardTableSheetPart.Baselines));
        Assert.Same(Cell(row, BoardWorkbookSchema.ColTimeDiv), BoardTableRowFilter.FirstProblemCell(BoardTableRowKinds.Warnings, row));

        Assert.DoesNotContain(row, BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.Warnings, BoardTableSearch.None, rowInHand: null, BoardTableSheetPart.Pictures));
        Assert.Contains(row, BoardTableRowFilter.RowsShown(sheet.Rows, BoardTableRowKinds.Warnings, BoardTableSearch.None, rowInHand: null, BoardTableSheetPart.Baselines));
        Assert.Equal(BoardTableSheetPart.Baselines, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Pictures, BoardTableRowKinds.Warnings, BoardTableSearch.None));
    }

    [Fact]
    public void A_picture_has_no_column_of_what_only_a_baseline_has()
    {
        BoardTableSheet sheet = Images(Document());

        string[] picture = BoardTableSheetParts.ColumnsShown(BoardTableSheetPart.Pictures, sheet).Select(index => sheet.Columns[index]).ToArray();

        Assert.Equal(
            [BoardWorkbookSchema.ColBoardLabel, BoardWorkbookSchema.ColRegion, BoardWorkbookSchema.ColName, BoardWorkbookSchema.ColFile, BoardWorkbookSchema.ColNote],
            picture);

        Assert.Equal(sheet.Columns.Count, BoardTableSheetParts.ColumnsShown(BoardTableSheetPart.Baselines, sheet).Count);
        Assert.Equal(sheet.Columns.Count, BoardTableSheetParts.ColumnsShown(BoardTableSheetPart.Whole, sheet).Count);
    }

    // ###########################################################################################
    // *** A BLOCK PASTED INTO A PICTURE'S ROW SKIPS WHAT IS NOT SHOWN. *** Walking the sheet's own columns,
    // its third value landed in the hidden Pin - nobody could see it, and the picture became a baseline.
    // ###########################################################################################
    [Fact]
    public void A_block_pasted_into_a_part_lands_in_the_columns_it_shows()
    {
        BoardTableSheet sheet = Images(Document());
        BoardTableRow picture = sheet.Rows[0];
        List<int> columns = BoardTableSheetParts.ColumnsShown(BoardTableSheetPart.Pictures, sheet);

        BoardTablePasteResult result = sheet.PasteBlock(
            Cell(picture, BoardWorkbookSchema.ColBoardLabel),
            [["U9", "PAL", "Pasted name"]],
            [picture],
            mayAddRows: false,
            columns);

        Assert.True(result.Pasted);
        Assert.Equal("U9", Cell(picture, BoardWorkbookSchema.ColBoardLabel).Text);
        Assert.Equal("PAL", Cell(picture, BoardWorkbookSchema.ColRegion).Text);
        Assert.Equal("Pasted name", Cell(picture, BoardWorkbookSchema.ColName).Text);
        Assert.Equal(string.Empty, Cell(picture, BoardWorkbookSchema.ColPin).Text);
    }

    // Past the columns shown it is said to be left out, as past a sheet's last column.
    [Fact]
    public void What_a_part_has_no_column_for_is_left_out()
    {
        BoardTableSheet sheet = Images(Document());
        BoardTableRow picture = sheet.Rows[0];
        List<int> columns = BoardTableSheetParts.ColumnsShown(BoardTableSheetPart.Pictures, sheet);

        BoardTablePasteResult result = sheet.PasteBlock(
            Cell(picture, BoardWorkbookSchema.ColNote),
            [["A note", "One too many"]],
            [picture],
            mayAddRows: false,
            columns);

        Assert.Equal("A note", Cell(picture, BoardWorkbookSchema.ColNote).Text);
        Assert.Equal(1, result.ColumnsLeftOut);
        Assert.Equal(string.Empty, Cell(picture, BoardWorkbookSchema.ColTimeDiv).Text);
    }

    // ###########################################################################################
    // PICKING "Errors" GOES TO THE ERRORS INSIDE A SHEET IN TWO, as it does between sheets - with no answer
    // in either part it stays, and any other sheet is whole.
    // ###########################################################################################
    [Fact]
    public void A_pill_with_nothing_in_the_part_shown_goes_to_the_other_part()
    {
        BoardTableSheet sheet = Images(Document());

        Assert.Equal(BoardTableSheetPart.Pictures, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Whole, BoardTableRowKinds.None, BoardTableSearch.None));
        Assert.Equal(BoardTableSheetPart.Baselines, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Baselines, BoardTableRowKinds.None, BoardTableSearch.None));

        // Only the baseline has been changed.
        Cell(sheet.Rows[1], BoardWorkbookSchema.ColName).Text = "Changed";
        sheet.Refresh();

        Assert.Equal(BoardTableSheetPart.Baselines, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Pictures, BoardTableRowKinds.Modified, BoardTableSearch.None));

        // Nothing of what is asked for in either: where it was.
        Assert.Equal(BoardTableSheetPart.Pictures, BoardTableSheetParts.PartToShow(sheet, BoardTableSheetPart.Pictures, BoardTableRowKinds.Deleted, BoardTableSearch.None));
    }

    [Fact]
    public void A_sheet_not_in_two_is_always_whole()
    {
        BoardTableSheet components = Document().FindSheet(BoardWorkbookSchema.SheetComponents)!;

        Assert.Equal(BoardTableSheetPart.Whole, BoardTableSheetParts.PartToShow(components, BoardTableSheetPart.Baselines, BoardTableRowKinds.None, BoardTableSearch.None));
        Assert.False(BoardTableSheetParts.IsSplit(BoardWorkbookSchema.SheetComponents));
    }
}
