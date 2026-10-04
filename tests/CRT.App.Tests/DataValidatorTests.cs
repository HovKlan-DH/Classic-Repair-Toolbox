using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// Smoke tests for DataValidator - the contributor-mode pass that walks every board and warns
// about missing files, duplicate UUIDs, orphan components and bad oscilloscope values.
//
// LIMIT OF THIS FILE, on purpose: ValidateAllDataAsync returns a bare Task and reports
// everything it finds through Logger only, so there is no result to assert on. These tests can
// prove it walks real data without throwing - which matters, because it runs at startup for
// contributors and an exception there is a broken launch - but they cannot check that a given
// problem is actually detected.
//
// To test the findings themselves, ValidateAllDataAsync would need to return them (e.g. a list
// of validation messages) instead of only logging. That is a public API change and a separate
// decision; see the Tests section of .claude/CLAUDE.md. *** Since 2026-10-02 the RULES are
// BoardDataChecks' (CRT.Data), tested there - this logs what they find - so only the log line's
// wording is tested here, through Describe. ***
[Collection("DataManager")]
public sealed class DataValidatorTests : IDisposable
{
    // A problem in a sheet names the file, sheet, row (1-based, as Excel numbers the data rows) and
    // column; a problem in no sheet - a highlight - names the JSON beside the workbook.
    [Fact]
    public void A_logged_problem_says_where_it_is_and_what_it_is()
    {
        var inSheet = new BoardDataProblem(
            BoardProblemLevel.Warning, "image.time_div", "5 ms", "T/DIV [5 ms] is not a setting CRT knows.",
            BoardWorkbookSchema.SheetComponentImages, 2, BoardWorkbookSchema.ColTimeDiv);

        Assert.Equal(
            "Excel data file [Commodore/C64/250407/Data.xlsx] sheet [Component images] row [3] column [T/DIV] has warning [image.time_div]: " +
            "T/DIV [5 ms] is not a setting CRT knows. - please fix!",
            DataValidator.Describe("Commodore/C64/250407/Data.xlsx", inSheet));

        var highlight = new BoardDataProblem(
            BoardProblemLevel.Error, "highlight.unknown_schematic", "X / U1", "A highlight names schematic [X].", null, -1, null);

        Assert.Equal(
            "Excel data file [Commodore/C64/250407/Data.xlsx] JSON file [Commodore/C64/250407/Data.json] has error [highlight.unknown_schematic]: " +
            "A highlight names schematic [X]. - please fix!",
            DataValidator.Describe("Commodore/C64/250407/Data.xlsx", highlight));
    }

    private readonly TempWorkspace thisWorkspace = new();

    static DataValidatorTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        this.thisWorkspace.Dispose();
    }

    private static readonly string[] HardwareHeaders =
    {
        "Hardware name in drop-down", "Board name in drop-down",
        "Excel data file", "Hardware notes in \"Overview\" tab"
    };

    private void LoadMasterWorkbook(params string?[][] rows)
    {
        new BoardWorkbookBuilder()
            .Sheet("Hardware & Board", HardwareHeaders, rows)
            .SaveTo(Path.Combine(this.thisWorkspace.Root, "master.xlsx"));

        DataManager.LoadFrom(this.thisWorkspace.Root, "master.xlsx");
    }

    [Fact]
    public async Task Validation_of_an_empty_dataset_completes()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "nothing.xlsx");

        Exception? thrown = await Record.ExceptionAsync(DataValidator.ValidateAllDataAsync);

        Assert.True(thrown is null, thrown?.ToString());
    }

    [Fact]
    public async Task Validation_survives_a_board_whose_excel_file_is_missing()
    {
        // A board listed in the master workbook whose file was never contributed must produce
        // warnings, not an unhandled exception at startup.
        this.LoadMasterWorkbook(new[] { "C64", "250407", "Commodore/C64/250407/missing.xlsx", "" });

        Exception? thrown = await Record.ExceptionAsync(DataValidator.ValidateAllDataAsync);

        Assert.True(thrown is null, thrown?.ToString());
    }

    [Fact]
    public async Task Validation_survives_a_board_with_a_blank_excel_file_reference()
    {
        this.LoadMasterWorkbook(new[] { "C64", "250407", "", "" });

        Exception? thrown = await Record.ExceptionAsync(DataValidator.ValidateAllDataAsync);

        Assert.True(thrown is null, thrown?.ToString());
    }

    [Fact]
    public async Task Validation_walks_a_real_board_workbook_without_throwing()
    {
        string relative = Path.Combine("Commodore", "C64", "250407", "board.xlsx");
        BoardWorkbookBuilder.WriteCompleteBoard(this.thisWorkspace.Path_(relative));

        this.LoadMasterWorkbook(new[] { "C64", "250407", relative.Replace('\\', '/'), "" });

        Exception? thrown = await Record.ExceptionAsync(DataValidator.ValidateAllDataAsync);

        Assert.True(thrown is null, thrown?.ToString());
    }

    [Fact]
    public async Task Validation_survives_a_board_with_duplicate_uuids_across_sheets()
    {
        // The duplicate-UUID check walks every sheet; feed it an actual duplicate so that path
        // executes rather than short-circuiting.
        string relative = Path.Combine("Commodore", "C64", "250407", "dupes.xlsx");
        string full = this.thisWorkspace.Path_(relative);

        new BoardWorkbookBuilder()
            .Sheet("Board schematics", BoardWorkbookBuilder.SchematicsHeaders,
                new[] { "same-uuid", "Sheet 1", "", "s1.png", "", "", "", "", "" })
            .Sheet("Components", BoardWorkbookBuilder.ComponentsHeaders,
                new[] { "same-uuid", "U1", "PLA", "906114", null, "IC", null, null })
            .SaveTo(full);

        this.LoadMasterWorkbook(new[] { "C64", "250407", relative.Replace('\\', '/'), "" });

        Exception? thrown = await Record.ExceptionAsync(DataValidator.ValidateAllDataAsync);

        Assert.True(thrown is null, thrown?.ToString());
    }

    [Fact]
    public async Task Validation_survives_a_board_with_out_of_range_oscilloscope_values()
    {
        string relative = Path.Combine("Commodore", "C64", "250407", "badscope.xlsx");

        new BoardWorkbookBuilder()
            .Sheet("Components", BoardWorkbookBuilder.ComponentsHeaders,
                new[] { "u1", "U1", "PLA", "906114", null, "IC", null, null })
            .Sheet("Component images", BoardWorkbookBuilder.ComponentImagesHeaders,
                new[] { "u2", "U1", "", "1", "Clock", "", "u1.png", "", "not-a-time", "not-a-volt", "nonsense" })
            .SaveTo(this.thisWorkspace.Path_(relative));

        this.LoadMasterWorkbook(new[] { "C64", "250407", relative.Replace('\\', '/'), "" });

        Exception? thrown = await Record.ExceptionAsync(DataValidator.ValidateAllDataAsync);

        Assert.True(thrown is null, thrown?.ToString());
    }
}
