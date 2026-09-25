using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for HardwareBoardEntry.ShortHardwareBoardLabel - the short
// "hardware/board" label derived from ExcelDataFile's own folder structure, used by the worklog
// bar's cross-board workbook picker (Main.FormatBoardKeyForDisplay) in place of the full
// HardwareName/BoardName pair, which runs too long once several boards share one dropdown.
public class HardwareBoardEntryTests
{
    [Theory]
    [InlineData("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", "C64/250407")]
    [InlineData("Amstrad/CPC 464/Some Board/Data.xlsx", "CPC 464/Some Board")]
    // A deeper path still only takes the immediate parent (board) and its own parent (hardware) -
    // not the whole path.
    [InlineData("Commodore/Shared files/C64/250407/Data.xlsx", "C64/250407")]
    public void The_short_label_is_the_boards_immediate_parent_folders(string excelDataFile, string expected)
    {
        var entry = new HardwareBoardEntry { ExcelDataFile = excelDataFile };

        Assert.Equal(expected, entry.ShortHardwareBoardLabel);
    }

    [Fact]
    public void A_blank_excel_data_file_yields_a_blank_label_rather_than_throwing()
    {
        Assert.Equal(string.Empty, new HardwareBoardEntry { ExcelDataFile = "" }.ShortHardwareBoardLabel);
        Assert.Equal(string.Empty, new HardwareBoardEntry { ExcelDataFile = "   " }.ShortHardwareBoardLabel);
    }

    // Fewer than two folder segments ahead of the file itself means there is no board/hardware
    // pair to take - falls back to the raw path rather than guessing or throwing.
    [Theory]
    [InlineData("Data.xlsx")]
    [InlineData("C64/Data.xlsx")]
    public void A_path_with_too_few_folder_segments_falls_back_to_the_raw_path(string excelDataFile)
    {
        var entry = new HardwareBoardEntry { ExcelDataFile = excelDataFile };

        Assert.Equal(excelDataFile, entry.ShortHardwareBoardLabel);
    }

    [Fact]
    public void A_leading_slash_does_not_shift_which_segments_are_taken()
    {
        var entry = new HardwareBoardEntry { ExcelDataFile = "/Commodore/C64/250407/Data.xlsx" };

        Assert.Equal("C64/250407", entry.ShortHardwareBoardLabel);
    }
}
