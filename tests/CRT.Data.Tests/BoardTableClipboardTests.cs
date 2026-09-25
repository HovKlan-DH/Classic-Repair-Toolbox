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
    public void A_block_of_several_cells_is_refused_rather_than_squeezed_into_one(string clipboard)
    {
        // Pasting a block is a later feature; half-doing it here would put "U8<tab>CPU" into a
        // single Board label cell.
        Assert.False(BoardTableClipboard.TryReadSingleCell(clipboard, out string cell));
        Assert.Equal(string.Empty, cell);
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
