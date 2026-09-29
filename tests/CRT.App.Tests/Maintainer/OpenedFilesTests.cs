using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers OpenedFiles' names - what a file tree and the table's hover card may hand to the
// operating system, and under which name.
// ###########################################################################################
public sealed class OpenedFilesTests
{
    // ###########################################################################################
    // What the tree opens: what a submission may carry, and a board's own files - workbook,
    // highlight file, KiCad data. A web page as text. Nothing else, and only the file's own name.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx", "Data C128 310378 v2.0.0.xlsx")]
    [InlineData("Commodore/C128/310378/Data C128 310378 v2.0.0.json", "Data C128 310378 v2.0.0.json")]
    [InlineData("Commodore/C128/310378/KiCad data/c128.kicad_pcb", "c128.kicad_pcb")]
    [InlineData("Commodore/C128/310378/Images/a.png", "a.png")]
    [InlineData("Commodore/C128/310378/notes.html", "notes.html.txt")]
    public void A_board_file_opens_under_its_own_name(string path, string expected)
    {
        Assert.True(OpenedFiles.TryGetOpenName(path, out string name));
        Assert.Equal(expected, name);
    }

    [Theory]
    [InlineData("Commodore/C128/310378/run.exe")]
    [InlineData("Commodore/C128/310378/macro.xlsm")]
    [InlineData("Commodore/C128/310378/")]
    public void Anything_else_is_not_opened(string path)
    {
        Assert.False(OpenedFiles.TryGetOpenName(path, out _));
    }

    // ###########################################################################################
    // Opened files go under CRT's own folder in the temp folder (2026-09-29) - not a
    // "CRT Maintainer" folder named after an application that no longer exists.
    // ###########################################################################################
    [Fact]
    public void Opened_files_are_written_under_CRTs_own_temp_folder()
    {
        string expected = Path.Combine(Path.GetTempPath(), CRT.AppConfig.AppFolderName, "Maintainer");

        Assert.Equal(expected, OpenedFiles.TempRoot);
        Assert.DoesNotContain("CRT Maintainer", OpenedFiles.TempRoot, StringComparison.Ordinal);
    }
}
