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
// decision; see the Tests section of .claude/CLAUDE.md.
[Collection("DataManager")]
public sealed class DataValidatorTests : IDisposable
{
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
