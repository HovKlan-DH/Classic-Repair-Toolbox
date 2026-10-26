using System.Reflection;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for DataManager's local half - resolving the data root and reading the
// main file that lists every hardware/board and every oscilloscope command set.
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
        // windows-path-literal: command-line TEXT, parsed and compared as text - no path API sees it.
        Assert.Equal(
            @"D:\somewhere\Data",
            DataManager.ResolveDataRoot(new[] { @"--data-root=D:\somewhere\Data" }));
    }

    [Fact]
    public void ResolveDataRoot_strips_surrounding_quotes()
    {
        // Shells and shortcuts quote paths that contain spaces.
        // windows-path-literal: command-line TEXT, parsed and compared as text - no path API sees it.
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

    // CRT 4.0 keeps its tree in a folder of its own beside 3.x's "Data" (2026-10-24): a new folder
    // is seeded from the install package, and 3.x's is left as it was.
    [Fact]
    public void ResolveDataRoot_falls_back_to_the_4_0_data_folder_under_local_appdata()
    {
        string resolved = DataManager.ResolveDataRoot(Array.Empty<string>());

        Assert.EndsWith(Path.Combine("Classic-Repair-Toolbox", AppConfig.DataFolderName), resolved);
        Assert.NotEqual(AppConfig.LegacyDataFolderName, AppConfig.DataFolderName);
        Assert.True(Path.IsPathRooted(resolved));
    }

    // ------------------------------------------------------------ main file
    //
    // *** IN CRT 4.0 THE MAIN FILE IS JSON (2026-10-24). *** "Classic-Repair-Toolbox.v4.0.0.json"
    // lists every board by its id, in the drop-downs' order, and every oscilloscope with its brand
    // on each entry - the Excel sheet's merged-cell carrying forward and header searching are the
    // converter's business now (DataTreeConversion), not the app's.

    private void LoadMain(IEnumerable<MainFiles.Board> boards, params MainDocument.OscilloscopeRow[] oscilloscopes) =>
        DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.Write(this.thisWorkspace.Root, boards, oscilloscopes));

    private void LoadMain(params MainFiles.Board[] boards) => this.LoadMain(boards, []);

    [Fact]
    public void The_boards_are_listed_in_the_main_files_order_each_by_its_board_file()
    {
        this.LoadMain(
            new MainFiles.Board("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "The breadbin"),
            new MainFiles.Board("Commodore/C128/310378", "Commodore 128", "310378 (C128 & C128D)"),
            new MainFiles.Board("ZX Spectrum/Spectrum 16K-48K/Issue 4B", "ZX Spectrum 16K/48K", "Issue 4B"));

        Assert.Equal(
            ["250407 (long board)", "310378 (C128 & C128D)", "Issue 4B"],
            DataManager.HardwareBoards.Select(entry => entry.BoardName));

        HardwareBoardEntry c64 = DataManager.HardwareBoards[0];
        Assert.Equal("Commodore 64", c64.HardwareName);
        Assert.Equal("Commodore/C64/250407/board.json", c64.ExcelDataFile);
        Assert.Equal("The breadbin", c64.HardwareNotes);
        Assert.True(c64.IsPublished);
    }

    [Fact]
    public void A_missing_main_file_leaves_the_lists_empty_instead_of_throwing()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "not-there.json");

        Assert.Empty(DataManager.HardwareBoards);
        Assert.Empty(DataManager.Oscilloscopes);
    }

    // A main file that is not JSON, or one a later CRT wrote in a format this one cannot read,
    // lists nothing - never a half-read list.
    [Theory]
    [InlineData("not json at all")]
    [InlineData("{ \"formatVersion\": 999, \"boards\": [ { \"boardId\": \"Commodore/C64/250407\" } ] }")]
    public void A_main_file_that_cannot_be_read_leaves_the_lists_empty(string content)
    {
        File.WriteAllText(Path.Combine(this.thisWorkspace.Root, MainFiles.FileName), content);

        DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.FileName);

        Assert.Empty(DataManager.HardwareBoards);
    }

    [Fact]
    public void A_listed_board_with_no_id_is_skipped()
    {
        this.LoadMain(
            new MainFiles.Board("", "C64", "nameless"),
            new MainFiles.Board("Commodore/C64/250407", "C64", "250407"));

        Assert.Equal(["250407"], DataManager.HardwareBoards.Select(entry => entry.BoardName));
    }

    // ###########################################################################################
    // *** ONE WAY TO READ THE LIST (code review, 2026-10-26). *** A row with no hardware name of its
    // own is under the hardware of the row above - the rule the server reads the main file by for
    // its name checks, the Boards screen and placement (MasterListing.TryRead). CRT read the name as
    // it stood and listed such a board under a blank hardware, so what users saw and what the
    // server judged were two lists. Both now read the file the same, names trimmed.
    // ###########################################################################################
    [Fact]
    public void The_boards_are_listed_as_the_server_reads_the_main_file_a_blank_hardware_name_under_the_one_above()
    {
        this.LoadMain(
            new MainFiles.Board("Commodore/C64/250407", "Commodore 64", "250407"),
            new MainFiles.Board("Commodore/C64/250425", "", "250425"),
            new MainFiles.Board("Commodore/C128/310378", " Commodore 128 ", " 310378 "));

        Assert.Equal(["Commodore 64", "Commodore 64", "Commodore 128"], DataManager.HardwareBoards.Select(entry => entry.HardwareName));
        Assert.Equal("310378", DataManager.HardwareBoards[2].BoardName);

        Assert.True(MasterListing.TryRead(Path.Combine(this.thisWorkspace.Root, MainFiles.FileName), out IReadOnlyList<MasterListingRow> rows, out _));
        Assert.Equal(
            rows.Select(row => (row.HardwareName, row.BoardName, row.ExcelDataFile)),
            DataManager.HardwareBoards.Select(entry => (entry.HardwareName, entry.BoardName, entry.ExcelDataFile)));
    }

    [Fact]
    public void Loading_a_second_main_file_replaces_the_previous_entries()
    {
        this.LoadMain(new MainFiles.Board("Commodore/C64/250407", "C64", "250407"));
        Assert.Single(DataManager.HardwareBoards);

        this.LoadMain(
            new MainFiles.Board("Commodore/Plus4/310163", "Plus/4", "310163"),
            new MainFiles.Board("Commodore/VIC-20/250403", "VIC-20", "250403"));

        Assert.Equal(2, DataManager.HardwareBoards.Count);
        Assert.DoesNotContain(DataManager.HardwareBoards, e => e.HardwareName == "C64");
    }

    [Fact]
    public void The_resolved_main_file_name_is_reported()
    {
        this.LoadMain(new MainFiles.Board("Commodore/C64/250407", "C64", "250407"));

        Assert.Equal(MainFiles.FileName, DataManager.ResolvedMainExcelFileName);
    }

    // --------------------------------------------------------------- oscilloscopes

    private static MainDocument.OscilloscopeRow Oscilloscope(
        string brand, string series, string timeDivList = "1ms, 2ms", string voltsDivList = "1V, 2V") => new()
    {
        Brand = brand,
        SeriesOrModel = series,
        Port = "5025",
        Identify = "*IDN?",
        DrainErrorQueue = ":SYST:ERR?",
        OperationComplete = "*OPC?",
        ClearStatistics = ":MEAS:CLE",
        QueryActiveTrigger = ":TRIG:STAT?",
        Stop = ":STOP",
        Single = ":SING",
        Run = ":RUN",
        QueryTriggerMode = ":TRIG:MODE?",
        QueryTriggerLevel = ":TRIG:EDGE:LEV?",
        SetTriggerLevel = ":TRIG:EDGE:LEV {0}",
        QueryTimeDiv = ":TIM:SCAL?",
        SetTimeDiv = ":TIM:SCAL {0}",
        QueryVoltsDiv = ":CHAN1:SCAL?",
        SetVoltsDiv = ":CHAN1:SCAL {0}",
        DumpImage = ":DISP:DATA?",
        TimeDivList = timeDivList,
        VoltsDivList = voltsDivList,
        DebounceTime = "250",
    };

    private static readonly MainFiles.Board AnyBoard = new("Commodore/C64/250407", "C64", "250407");

    [Fact]
    public void Oscilloscope_definitions_are_read_from_the_main_file()
    {
        this.LoadMain([AnyBoard], Oscilloscope("Rigol", "DS1000Z"));

        OscilloscopeEntry scope = Assert.Single(DataManager.Oscilloscopes);

        Assert.Equal("Rigol", scope.Brand);
        Assert.Equal("DS1000Z", scope.SeriesOrModel);
        Assert.Equal("5025", scope.Port);
        Assert.Equal("*IDN?", scope.Identify);
        Assert.Equal(":TIM:SCAL {0}", scope.SetTimeDiv);
        Assert.Equal("1ms, 2ms", scope.TimeDivList);
        Assert.Equal("1V, 2V", scope.VoltsDivList);
        Assert.Equal("250", scope.DebounceTime);
    }

    [Fact]
    public void Each_oscilloscope_carries_its_own_brand()
    {
        this.LoadMain([AnyBoard], Oscilloscope("Rigol", "DS1000Z"), Oscilloscope("Siglent", "SDS1000X-E"));

        Assert.Equal(["Rigol", "Siglent"], DataManager.Oscilloscopes.Select(scope => scope.Brand));
    }

    [Fact]
    public void An_oscilloscope_with_no_series_is_skipped()
    {
        this.LoadMain([AnyBoard], Oscilloscope("Rigol", "DS1000Z"), Oscilloscope("Siglent", ""));

        Assert.Single(DataManager.Oscilloscopes);
    }

    [Fact]
    public void A_main_file_with_no_oscilloscopes_still_lists_the_boards()
    {
        this.LoadMain(AnyBoard);

        Assert.Single(DataManager.HardwareBoards);
        Assert.Empty(DataManager.Oscilloscopes);
    }

    [Fact]
    public void A_loaded_scope_definition_drives_the_value_mapper_end_to_end()
    {
        // Proves the two halves agree: what the main file declares as supported is what the
        // mapper will actually snap a board's T/DIV value to.
        this.LoadMain([AnyBoard], Oscilloscope("Rigol", "DS1000Z", timeDivList: "500us, 1ms, 2ms, 5ms"));

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
            DataRootPath, Path.Combine("elsewhere", "Board32", "config")));
    }
}
