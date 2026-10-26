using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// One cell to and from the clipboard, the way Excel writes and reads it - the table editor's
// Ctrl+C / Ctrl+V on a selected cell.
//
// The trailing line break is the case that would actually bite: Excel puts one after EVERY
// copied cell, including a single one, so a naive paste carries a newline into the cell.
// ###########################################################################################
public sealed class BoardTableClipboardTests
{
    [Theory]
    [InlineData("U8\r\n", "U8")]
    [InlineData("U8\n", "U8")]
    [InlineData("U8", "U8")]
    [InlineData("6510 CPU", "6510 CPU")]
    public void A_single_cell_is_taken_without_the_line_break_Excel_adds(string clipboard, string expected)
    {
        Assert.True(BoardTableClipboard.TryReadSingleCell(clipboard, out string cell));
        Assert.Equal(expected, cell);
    }

    [Theory]
    [InlineData("U8\tCPU\r\n")]
    [InlineData("U8\r\nU9\r\n")]
    [InlineData("\"a\"\t\"b\"\r\n")]
    public void A_block_of_several_cells_is_not_read_as_one_cell(string clipboard)
    {
        // A block is ReadBlock's (since CRT 4.0); read as one cell it would put "U8<tab>CPU" into a
        // single Board label cell.
        Assert.False(BoardTableClipboard.TryReadSingleCell(clipboard, out string cell));
        Assert.Equal(string.Empty, cell);
    }

    // ------------------------------------------------------------------ Blocks (CRT 4.0, 2026-10-25)

    [Fact]
    public void A_block_is_read_row_by_row_and_cell_by_cell_without_the_last_line_break()
    {
        Assert.Equal(
            [["U8", "CPU"], ["U9", "VIC"]],
            BoardTableClipboard.ReadBlock("U8\tCPU\r\nU9\tVIC\r\n"));
    }

    // LibreOffice and Google Sheets end lines with "\n"; an old Mac program with "\r".
    [Theory]
    [InlineData("U8\nU9\n")]
    [InlineData("U8\rU9")]
    [InlineData("U8\r\nU9")]
    public void Every_kind_of_line_break_ends_a_row(string clipboard)
    {
        Assert.Equal([["U8"], ["U9"]], BoardTableClipboard.ReadBlock(clipboard));
    }

    [Fact]
    public void A_quoted_cell_keeps_its_tabs_line_breaks_and_quotes()
    {
        Assert.Equal(
            [["a\tb", "line one\r\nline two"], ["5.25\" drive", "x"]],
            BoardTableClipboard.ReadBlock("\"a\tb\"\t\"line one\r\nline two\"\r\n\"5.25\"\" drive\"\tx\r\n"));
    }

    // An empty cell is a cell: copied empty cells clear the cells they land on, as in a spreadsheet.
    [Fact]
    public void Empty_cells_are_read_as_empty_cells()
    {
        Assert.Equal([["a", "", "c"], ["", "", ""]], BoardTableClipboard.ReadBlock("a\t\tc\r\n\t\t\r\n"));
    }

    // Text from somewhere else than a spreadsheet: rows of different lengths, and a quote that does
    // not close a cell the way a spreadsheet writes one - kept as it is.
    [Fact]
    public void Ragged_rows_and_stray_quotes_are_taken_as_they_are()
    {
        Assert.Equal(
            [["\"5.25\" drive", "b"], ["c"]],
            BoardTableClipboard.ReadBlock("\"5.25\" drive\tb\nc"));
    }

    [Fact]
    public void One_cell_and_an_empty_clipboard_read_as_one_cell_and_as_nothing()
    {
        Assert.Equal([["U8"]], BoardTableClipboard.ReadBlock("U8\r\n"));
        Assert.Empty(BoardTableClipboard.ReadBlock(string.Empty));
        Assert.Empty(BoardTableClipboard.ReadBlock(null));
    }

    // The table's own copy pasted back is the same block - whatever the cells hold.
    [Fact]
    public void A_block_written_and_read_back_is_the_same_block()
    {
        string[][] block =
        [
            ["U8", "5.25\" drive", ""],
            ["a\tb", "line one\nline two", "\"quoted\""],
        ];

        Assert.Equal(block, BoardTableClipboard.ReadBlock(BoardTableClipboard.ForBlock(block)));
    }

    [Fact]
    public void Selected_cells_are_copied_in_the_order_the_rows_are_shown_with_gaps_left_empty()
    {
        BoardTableSheet sheet = BoardTableClipboardTests.Sheet("U1", "U2", "U3");
        BoardTableRow u1 = sheet.Rows[0];
        BoardTableRow u2 = sheet.Rows[1];
        BoardTableRow u3 = sheet.Rows[2];

        // U3 and U1 selected in label and name, but U1's name not - and U2 not at all: U2's line is
        // left out, U1's unselected cell is empty, and the rows come in the order shown.
        string copied = BoardTableClipboard.ForCells(
            [u3.Cells[0], u1.Cells[0], u3.Cells[1]],
            [u1, u2, u3]);

        Assert.Equal("U1\t\r\nU3\tname U3", copied);
    }

    [Fact]
    public void One_selected_cell_is_copied_as_one_cell()
    {
        BoardTableSheet sheet = BoardTableClipboardTests.Sheet("U1");

        Assert.Equal("U1", BoardTableClipboard.ForCells([sheet.Rows[0].Cells[0]], sheet.Rows));
        Assert.Equal(string.Empty, BoardTableClipboard.ForCells([], sheet.Rows));
    }

    // The Components sheet of a board with these labels, each named "name <label>".
    private static BoardTableSheet Sheet(params string[] labels)
    {
        var board = new BoardData();

        foreach (string label in labels)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = label, FriendlyName = $"name {label}", TechnicalNameOrValue = $"part {label}" });
        }

        BoardTableSheet sheet = BoardTableDocument.Create(null, board).FindSheet(BoardWorkbookSchema.SheetComponents)!;

        Assert.Equal(BoardWorkbookSchema.ColBoardLabel, sheet.Columns[0]);
        Assert.Equal(BoardWorkbookSchema.ColFriendlyName, sheet.Columns[1]);

        return sheet;
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void An_empty_clipboard_pastes_nothing(string? clipboard)
    {
        Assert.False(BoardTableClipboard.TryReadSingleCell(clipboard, out _));
    }

    [Fact]
    public void A_quoted_cell_holding_a_line_break_is_ONE_cell()
    {
        // Excel quotes a cell containing Alt+Enter line breaks; it is still one value.
        Assert.True(BoardTableClipboard.TryReadSingleCell("\"line one\nline two\"\r\n", out string cell));
        Assert.Equal("line one\nline two", cell);
    }

    [Fact]
    public void Doubled_quotes_inside_a_quoted_cell_become_single_quotes()
    {
        Assert.True(BoardTableClipboard.TryReadSingleCell("\"5.25\"\" drive\"\r\n", out string cell));
        Assert.Equal("5.25\" drive", cell);
    }

    [Theory]
    [InlineData("U8", "U8")]
    [InlineData("", "")]
    [InlineData("5.25\" drive", "\"5.25\"\" drive\"")]
    [InlineData("a\tb", "\"a\tb\"")]
    [InlineData("line one\nline two", "\"line one\nline two\"")]
    public void A_copied_cell_is_quoted_exactly_when_Excel_would_quote_it(string cell, string expected)
    {
        Assert.Equal(expected, BoardTableClipboard.ForSingleCell(cell));
    }

    [Theory]
    [InlineData("U8")]
    [InlineData("5.25\" drive")]
    [InlineData("a\tb")]
    [InlineData("line one\r\nline two")]
    public void Copying_a_cell_and_pasting_it_back_gives_the_same_text(string cell)
    {
        Assert.True(BoardTableClipboard.TryReadSingleCell(BoardTableClipboard.ForSingleCell(cell), out string pasted));
        Assert.Equal(cell, pasted);
    }
}
