using System.Reflection;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for DataManager's local half - resolving the data root and reading the
// master workbook that lists every hardware/board and every oscilloscope command set.
//
// DataManager is a static singleton whose InitializeAsync also syncs over the network and seeds
// the data folder. These tests never call it: LoadFrom(dataRoot, workbookName) is the local
// load with no network and no seeding, pointed at a temporary folder.
//
// Static state again, so this file is its own sequential collection.
[Collection("DataManager")]
public sealed class DataManagerTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    static DataManagerTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public void Dispose()
    {
        // Leave the static lists empty so nothing later mistakes test data for real data.
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        this.thisWorkspace.Dispose();
    }

    // ------------------------------------------------------------------ data root

    [Fact]
    public void ResolveDataRoot_uses_the_data_root_command_line_argument()
    {
        Assert.Equal(
            @"D:\somewhere\Data",
            DataManager.ResolveDataRoot(new[] { @"--data-root=D:\somewhere\Data" }));
    }

    [Fact]
    public void ResolveDataRoot_strips_surrounding_quotes()
    {
        // Shells and shortcuts quote paths that contain spaces.
        Assert.Equal(
            @"D:\my data\Data",
            DataManager.ResolveDataRoot(new[] { "--data-root=\"D:\\my data\\Data\"" }));
    }

    [Fact]
    public void ResolveDataRoot_matches_the_argument_case_insensitively()
    {
        Assert.Equal("X", DataManager.ResolveDataRoot(new[] { "--DATA-ROOT=X" }));
    }

    [Fact]
    public void ResolveDataRoot_takes_the_first_matching_argument()
    {
        Assert.Equal(
            "first",
            DataManager.ResolveDataRoot(new[] { "--data-root=first", "--data-root=second" }));
    }

    [Fact]
    public void ResolveDataRoot_ignores_unrelated_arguments()
    {
        Assert.Equal(
            "X",
            DataManager.ResolveDataRoot(new[] { "--verbose", "--data-root=X", "somefile.txt" }));
    }

    [Fact]
    public void ResolveDataRoot_falls_back_to_a_Data_folder_under_local_appdata()
    {
        string resolved = DataManager.ResolveDataRoot(Array.Empty<string>());

        Assert.EndsWith(Path.Combine("Classic-Repair-Toolbox", "Data"), resolved);
        Assert.True(Path.IsPathRooted(resolved));
    }

    // ------------------------------------------------------------ master workbook

    private const string HardwareSheet = "Hardware & Board";
    private const string OscilloscopeSheet = "Oscilloscope";

    private static readonly string[] HardwareHeaders =
    {
        "Hardware name in drop-down", "Board name in drop-down",
        "Excel data file", "Hardware notes in \"Overview\" tab"
    };

    private static readonly string[] OscilloscopeHeaders =
    {
        "Brand", "Series or model", "Port", "Identify", "DrainErrorQueue", "Operation-Complete",
        "Clear-Statistics", "QueryActiveTrigger", "Stop", "Single", "Run", "QueryTriggerMode",
        "QueryTriggerLevel", "SetTriggerLevel", "QueryTimeDiv", "SetTimeDiv", "QueryVoltsDiv",
        "SetVoltsDiv", "DumpImage", "TIME/DIV", "VOLTS/DIV", "Debounce-Time"
    };

    /// <summary>Writes a master workbook into the temp data root and loads it.</summary>
    private void LoadWorkbook(BoardWorkbookBuilder builder, string fileName = "Classic-Repair-Toolbox.xlsx")
    {
        builder.SaveTo(Path.Combine(this.thisWorkspace.Root, fileName));
        DataManager.LoadFrom(this.thisWorkspace.Root, fileName);
    }

    private static BoardWorkbookBuilder HardwareWorkbook(params string?[][] rows) =>
        new BoardWorkbookBuilder().Sheet(HardwareSheet, HardwareHeaders, rows);

    [Fact]
    public void A_missing_workbook_leaves_the_lists_empty_instead_of_throwing()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "not-there.xlsx");

        Assert.Empty(DataManager.HardwareBoards);
        Assert.Empty(DataManager.Oscilloscopes);
    }

    [Fact]
    public void Hardware_and_board_rows_are_read()
    {
        this.LoadWorkbook(HardwareWorkbook(
            new[] { "C64", "250407", "Commodore/C64/250407/Data C64 250407.xlsx", "The breadbin" },
            new[] { "Plus/4", "310163", "Commodore/Plus4/310163/Data Plus4 310163.xlsx", "The Plus/4" }));

        Assert.Equal(2, DataManager.HardwareBoards.Count);

        HardwareBoardEntry c64 = DataManager.HardwareBoards[0];
        Assert.Equal("C64", c64.HardwareName);
        Assert.Equal("250407", c64.BoardName);
        Assert.Equal("Commodore/C64/250407/Data C64 250407.xlsx", c64.ExcelDataFile);
        Assert.Equal("The breadbin", c64.HardwareNotes);
    }

    [Fact]
    public void A_blank_hardware_name_is_carried_forward_from_the_row_above()
    {
        // The sheet uses a merged-cell style: one hardware name, several board rows under it.
        this.LoadWorkbook(HardwareWorkbook(
            new[] { "C64", "250407", "a.xlsx", "notes" },
            new[] { null, "250425", "b.xlsx", "notes" },
            new[] { null, "326298", "c.xlsx", "notes" }));

        Assert.Equal(3, DataManager.HardwareBoards.Count);
        Assert.All(DataManager.HardwareBoards, e => Assert.Equal("C64", e.HardwareName));
        Assert.Equal(new[] { "250407", "250425", "326298" },
            DataManager.HardwareBoards.Select(e => e.BoardName));
    }

    [Fact]
    public void A_new_hardware_name_starts_a_new_carry_forward_group()
    {
        this.LoadWorkbook(HardwareWorkbook(
            new[] { "C64", "250407", "a.xlsx", "" },
            new[] { null, "250425", "b.xlsx", "" },
            new[] { "Plus/4", "310163", "c.xlsx", "" },
            new[] { null, "310164", "d.xlsx", "" }));

        Assert.Equal(
            new[] { "C64", "C64", "Plus/4", "Plus/4" },
            DataManager.HardwareBoards.Select(e => e.HardwareName));
    }

    [Fact]
    public void Column_order_in_the_master_workbook_does_not_matter()
    {
        string[] reversed = HardwareHeaders.Reverse().ToArray();

        this.LoadWorkbook(new BoardWorkbookBuilder().Sheet(
            HardwareSheet, reversed,
            new[] { "The breadbin", "a.xlsx", "250407", "C64" }));

        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards);
        Assert.Equal("C64", entry.HardwareName);
        Assert.Equal("250407", entry.BoardName);
    }

    [Fact]
    public void Comment_rows_above_the_header_are_skipped()
    {
        this.LoadWorkbook(new BoardWorkbookBuilder().SheetWithLeadingRows(
            HardwareSheet,
            new[] { "# Master data file", "# Revision date: 2026-01-01" },
            HardwareHeaders,
            new[] { "C64", "250407", "a.xlsx", "" }));

        Assert.Single(DataManager.HardwareBoards);
    }

    [Fact]
    public void A_workbook_with_the_wrong_headers_yields_no_hardware()
    {
        this.LoadWorkbook(new BoardWorkbookBuilder().Sheet(
            HardwareSheet, new[] { "Nope", "Not", "These", "Either" },
            new[] { "C64", "250407", "a.xlsx", "" }));

        Assert.Empty(DataManager.HardwareBoards);
    }

    [Fact]
    public void Loading_a_second_workbook_replaces_the_previous_entries()
    {
        this.LoadWorkbook(HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" }), "first.xlsx");
        Assert.Single(DataManager.HardwareBoards);

        this.LoadWorkbook(
            HardwareWorkbook(
                new[] { "Plus/4", "310163", "b.xlsx", "" },
                new[] { "VIC-20", "250403", "c.xlsx", "" }),
            "second.xlsx");

        Assert.Equal(2, DataManager.HardwareBoards.Count);
        Assert.DoesNotContain(DataManager.HardwareBoards, e => e.HardwareName == "C64");
    }

    [Fact]
    public void The_resolved_workbook_name_is_reported()
    {
        this.LoadWorkbook(HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" }), "chosen.xlsx");

        Assert.Equal("chosen.xlsx", DataManager.ResolvedMainExcelFileName);
    }

    // --------------------------------------------------------------- oscilloscopes

    private static string?[] OscilloscopeRow(
        string brand, string series, string timeDivList = "1ms, 2ms", string voltsDivList = "1V, 2V") =>
        new string?[]
        {
            brand, series, "5025", "*IDN?", ":SYST:ERR?", "*OPC?", ":MEAS:CLE", ":TRIG:STAT?",
            ":STOP", ":SING", ":RUN", ":TRIG:MODE?", ":TRIG:EDGE:LEV?", ":TRIG:EDGE:LEV {0}",
            ":TIM:SCAL?", ":TIM:SCAL {0}", ":CHAN1:SCAL?", ":CHAN1:SCAL {0}", ":DISP:DATA?",
            timeDivList, voltsDivList, "250"
        };

    [Fact]
    public void Oscilloscope_definitions_are_read_from_their_sheet()
    {
        this.LoadWorkbook(
            HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" })
                .Sheet(OscilloscopeSheet, OscilloscopeHeaders,
                    OscilloscopeRow("Rigol", "DS1000Z")));

        OscilloscopeEntry scope = Assert.Single(DataManager.Oscilloscopes);

        Assert.Equal("Rigol", scope.Brand);
        Assert.Equal("DS1000Z", scope.SeriesOrModel);
        Assert.Equal("5025", scope.Port);
        Assert.Equal("*IDN?", scope.Identify);
        Assert.Equal(":TIM:SCAL {0}", scope.SetTimeDiv);
        Assert.Equal("1ms, 2ms", scope.TimeDivList);
        Assert.Equal("1V, 2V", scope.VoltsDivList);
    }

    [Fact]
    public void A_blank_oscilloscope_brand_is_carried_forward()
    {
        string?[] second = OscilloscopeRow("", "DS1000Z Plus");

        this.LoadWorkbook(
            HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" })
                .Sheet(OscilloscopeSheet, OscilloscopeHeaders,
                    OscilloscopeRow("Rigol", "DS1000Z"),
                    second));

        Assert.Equal(2, DataManager.Oscilloscopes.Count);
        Assert.All(DataManager.Oscilloscopes, s => Assert.Equal("Rigol", s.Brand));
    }

    [Fact]
    public void An_oscilloscope_row_with_no_series_is_skipped()
    {
        this.LoadWorkbook(
            HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" })
                .Sheet(OscilloscopeSheet, OscilloscopeHeaders,
                    OscilloscopeRow("Rigol", "DS1000Z"),
                    OscilloscopeRow("Siglent", "")));

        Assert.Single(DataManager.Oscilloscopes);
    }

    [Fact]
    public void A_workbook_with_no_oscilloscope_sheet_still_loads_the_hardware()
    {
        this.LoadWorkbook(HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" }));

        Assert.Single(DataManager.HardwareBoards);
        Assert.Empty(DataManager.Oscilloscopes);
    }

    [Fact]
    public void A_loaded_scope_definition_drives_the_value_mapper_end_to_end()
    {
        // Proves the two halves agree: what the workbook declares as supported is what the
        // mapper will actually snap a board's T/DIV value to.
        this.LoadWorkbook(
            HardwareWorkbook(new[] { "C64", "250407", "a.xlsx", "" })
                .Sheet(OscilloscopeSheet, OscilloscopeHeaders,
                    OscilloscopeRow("Rigol", "DS1000Z", timeDivList: "500us, 1ms, 2ms, 5ms")));

        OscilloscopeEntry scope = Assert.Single(DataManager.Oscilloscopes);

        Assert.True(Handlers.Oscilloscope.ScopeValueMapper.TryMapTimeDiv(
            new ComponentImageEntry { TimeDiv = "1000us" },
            scope,
            out Handlers.Oscilloscope.ScopeMappedValue mapped));

        Assert.Equal("1ms", mapped.MatchedDisplayValue);
        Assert.Equal("0.001", mapped.ScpiValue);
    }

    // ------------------------------------------------------------ data-root containment
    //
    // thisIsPathWithinDataRoot guards the two DELETING paths in this class - orphan-file cleanup
    // and the empty-directory sweep that follows it - so what it calls "inside the root" decides
    // what may be removed from a user's disk. It is private, and reached here by reflection for
    // the same reason ExternalTargetLauncherTests reaches its own containment predicate that way:
    // driving it through the public sync path would mean actually deleting files to observe the
    // answer.

    private static bool IsPathWithinDataRoot(string dataRootFullPath, string candidateFullPath)
    {
        MethodInfo? method = typeof(DataManager).GetMethod(
            "thisIsPathWithinDataRoot", BindingFlags.NonPublic | BindingFlags.Static);

        Assert.True(method is not null,
            "DataManager.thisIsPathWithinDataRoot not found - if it was renamed, update these tests.");

        return (bool)method!.Invoke(null, new object?[] { dataRootFullPath, candidateFullPath })!;
    }

    // Built with Path.Combine rather than a hardcoded @"C:\..." literal - thisIsPathWithinDataRoot
    // appends Path.DirectorySeparatorChar (a backslash on Windows, a forward slash on Linux/macOS),
    // and CI runs this suite on Linux. A literal backslash path compared against a forward-slash
    // trailing separator fails containment for the wrong reason - reported by a CI run that was
    // green locally on Windows and red on GitHub Actions for exactly this test.
    private static string DataRootPath => Path.Combine("CRT-root", "Data");

    [Fact]
    public void A_file_inside_the_data_root_is_within_it()
    {
        Assert.True(IsPathWithinDataRoot(
            DataRootPath, Path.Combine(DataRootPath, "Commodore", "C64", "board.xlsx")));
    }

    // The bug this pins: a bare StartsWith against the root treats a SIBLING whose name merely
    // BEGINS with the root's name as being inside it, because "CRT-root/DataBackup" does start
    // with "CRT-root/Data". Both callers delete what this admits, so that would let orphan cleanup
    // and the empty-directory sweep walk into a folder sitting next to the data root. Fails against
    // the version without the trailing-separator guard.
    [Fact]
    public void A_sibling_folder_whose_name_merely_starts_with_the_root_name_is_not_within_it()
    {
        Assert.False(IsPathWithinDataRoot(
            DataRootPath, Path.Combine("CRT-root", "DataBackup", "secret.xlsx")));
    }

    // The root is not inside itself - the empty-directory sweep would otherwise consider deleting
    // the data root itself once it happened to be empty.
    [Fact]
    public void The_data_root_itself_is_not_within_itself()
    {
        Assert.False(IsPathWithinDataRoot(DataRootPath, DataRootPath));
    }

    [Fact]
    public void A_path_entirely_outside_the_data_root_is_not_within_it()
    {
        Assert.False(IsPathWithinDataRoot(
            DataRootPath, Path.Combine("elsewhere", "System32", "config")));
    }
}
