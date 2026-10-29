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

    // ------------------------------------------------------------ the user's own scope models

    private string OwnModelsFile => this.thisWorkspace.Path_("own-oscilloscopes.json");

    private void LoadWithOwnModels(MainDocument.OscilloscopeRow[] published, params MainDocument.OscilloscopeRow[] own)
    {
        OwnOscilloscopeFile.Write(this.OwnModelsFile, own);
        DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], published), this.OwnModelsFile);
    }

    // The user's own models come after the published ones, marked as theirs - the Oscilloscope tab
    // edits only those.
    [Fact]
    public void The_user_s_own_models_are_listed_after_the_published_ones_and_marked_as_theirs()
    {
        this.LoadWithOwnModels([Oscilloscope("Rigol", "DS1000Z")], Oscilloscope("Voltcraft", "DOV702"));

        Assert.Equal(["DS1000Z", "DOV702"], DataManager.Oscilloscopes.Select(scope => scope.SeriesOrModel));
        Assert.Equal([false, true], DataManager.Oscilloscopes.Select(scope => scope.IsOwn));
        Assert.Equal("DS1000Z", Assert.Single(DataManager.PublishedScopeModels).SeriesOrModel);
        Assert.Equal("DOV702", Assert.Single(DataManager.OwnScopeModels).SeriesOrModel);
    }

    // The editor's change is listed at once and kept in the user's own file.
    [Fact]
    public async Task A_change_to_the_own_models_is_listed_at_once_and_written_to_the_user_s_file()
    {
        this.LoadWithOwnModels([Oscilloscope("Rigol", "DS1000Z")]);

        Task<string?> save = DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Voltcraft", "DOV702")]);

        Assert.Equal("DOV702", DataManager.Oscilloscopes[^1].SeriesOrModel);
        Assert.Null(await save);

        MainDocument written = DataTreeJson.ReadMain(File.ReadAllText(this.OwnModelsFile));
        Assert.Equal("DOV702", Assert.Single(written.Oscilloscopes).SeriesOrModel);
        Assert.Empty(written.Boards);
    }

    // Saves follow each other and the file ends with the latest list, whichever finishes last.
    [Fact]
    public async Task Saves_one_after_another_leave_the_latest_list_in_the_file()
    {
        this.LoadWithOwnModels([]);

        Task<string?> first = DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Voltcraft", "DOV702")]);
        Task<string?> second = DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Voltcraft", "DOV702"), Oscilloscope("Owon", "ADS800")]);

        await Task.WhenAll(first, second);

        Assert.Equal(2, DataTreeJson.ReadMain(File.ReadAllText(this.OwnModelsFile)).Oscilloscopes.Count);
    }

    // ###########################################################################################
    // *** AN OWN-MODELS FILE THAT COULD NOT BE READ IS NEVER WRITTEN OVER (code review, 2026-10-29). ***
    // It read as "no models", and the editor's first save replaced it with the one model just made.
    // The save says why instead, and nothing changes - in the list or on disk.
    // ###########################################################################################
    [Fact]
    public async Task An_own_models_file_that_could_not_be_read_is_never_written_over()
    {
        File.WriteAllText(this.OwnModelsFile, "{ a hand edit gone wrong");
        DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z")]), this.OwnModelsFile);

        Assert.NotNull(DataManager.OwnScopeModelsProblem);

        string? refused = await DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Voltcraft", "DOV702")]);

        Assert.Equal(DataManager.OwnScopeModelsProblem, refused);
        Assert.Equal("{ a hand edit gone wrong", File.ReadAllText(this.OwnModelsFile));
        Assert.Empty(DataManager.OwnScopeModels);
        Assert.Single(DataManager.Oscilloscopes);
    }

    // ###########################################################################################
    // A reload of the same main file, HELD just before the own models are taken up - the main file
    // read, the board list in place (DataManager.BeforeOwnScopeModelsForTests) - so a test can look at
    // what the lists say WHILE a manual data check runs on the pool. Disposing the hold lets the
    // reload finish; Reload is the reload itself. (Held at the write gate instead, as this was first,
    // a reload of a file read well never stopped - it takes the gate only to write - and the test
    // looked after the reload had as good as finished: code review, 2026-10-29.)
    // ###########################################################################################
    private sealed class HeldReload : IDisposable
    {
        private readonly ManualResetEventSlim thisLetGo = new(false);
        private readonly ManualResetEventSlim thisHeld = new(false);

        public HeldReload(Action reload)
        {
            DataManager.BeforeOwnScopeModelsForTests = () =>
            {
                this.thisHeld.Set();
                this.thisLetGo.Wait(TimeSpan.FromSeconds(30));
            };

            this.Reload = Task.Run(reload);
            Assert.True(this.thisHeld.Wait(TimeSpan.FromSeconds(10)));
        }

        public Task Reload { get; }

        public void Dispose()
        {
            this.thisLetGo.Set();
            DataManager.BeforeOwnScopeModelsForTests = null;
        }
    }

    private HeldReload ReloadHeld(string main) =>
        new(() => DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile));

    // ###########################################################################################
    // *** A RELOAD NEVER LIFTS "COULD NOT BE READ" WHILE IT RUNS (code review, 2026-10-29). *** The
    // start of the load cleared the problem - the file read last had failed, so it looked like another
    // file - and a model made while the data check ran was written over the unreadable file.
    // ###########################################################################################
    [Fact]
    public async Task A_save_while_a_reload_runs_is_still_refused_for_a_file_that_could_not_be_read()
    {
        File.WriteAllText(this.OwnModelsFile, "{ a hand edit gone wrong");
        string main = MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z")]);
        DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile);
        string? problem = DataManager.OwnScopeModelsProblem;
        Assert.NotNull(problem);

        Task reload;
        Task<string?> save;

        using (HeldReload held = this.ReloadHeld(main))
        {
            reload = held.Reload;
            Assert.Equal(problem, DataManager.OwnScopeModelsProblem);

            save = DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Voltcraft", "DOV702")]);
            Assert.True(save.IsCompleted);
        }

        await reload;
        Assert.Equal(problem, await save);
        Assert.Equal("{ a hand edit gone wrong", File.ReadAllText(this.OwnModelsFile));
    }

    // ###########################################################################################
    // *** THE PUBLISHED MODELS STAY LISTED WHILE A RELOAD RUNS (code review, 2026-10-29). *** They were
    // taken out at its start: a model the editor named meanwhile was checked against own models alone,
    // and the drop-downs it rebuilt showed no published model.
    // ###########################################################################################
    [Fact]
    public async Task The_published_models_stay_listed_while_a_reload_runs()
    {
        string main = MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z")]);
        OwnOscilloscopeFile.Write(this.OwnModelsFile, [Oscilloscope("Voltcraft", "DOV702")]);
        DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile);

        Task reload;

        using (HeldReload held = this.ReloadHeld(main))
        {
            reload = held.Reload;
            Assert.Equal(["DS1000Z", "DOV702"], DataManager.Oscilloscopes.Select(scope => scope.SeriesOrModel));
            Assert.Equal("DS1000Z", Assert.Single(DataManager.PublishedScopeModels).SeriesOrModel);
        }

        await reload;
    }

    // ###########################################################################################
    // *** THE TAB IS TOLD AS SOON AS THE MODELS ARE LISTED AGAIN (code review, 2026-10-29). *** It was
    // told only after the whole data check - the board files and the rest - and a change typed into a
    // model the reload had renamed meanwhile was dropped. Raised once per load, off the UI thread.
    // ###########################################################################################
    [Fact]
    public void Every_load_says_when_the_scope_models_are_listed_again()
    {
        int told = 0;
        int ownWhenTold = -1;

        void OnReloaded()
        {
            told++;
            ownWhenTold = DataManager.OwnScopeModels.Count;
        }

        OwnOscilloscopeFile.Write(this.OwnModelsFile, [Oscilloscope("Voltcraft", "DOV702")]);
        DataManager.ScopeModelsListed += OnReloaded;

        try
        {
            DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z")]), this.OwnModelsFile);
        }
        finally
        {
            DataManager.ScopeModelsListed -= OnReloaded;
        }

        Assert.Equal(1, told);
        Assert.Equal(1, ownWhenTold);
    }

    // ###########################################################################################
    // *** A FILE NOT READ YET IS NEVER WRITTEN EITHER (code review, 2026-10-29). *** With no main file
    // at launch, the own models were never read - and the editor's save would have replaced the
    // user's file with the one list it held. The editor changes the list only once its file was read.
    // ###########################################################################################
    [Fact]
    public async Task An_own_models_file_not_read_yet_is_never_written()
    {
        OwnOscilloscopeFile.Write(this.OwnModelsFile, [Oscilloscope("Voltcraft", "DOV702")]);
        string before = File.ReadAllText(this.OwnModelsFile);
        DataManager.LoadFrom(this.thisWorkspace.Root, "no-main-file.json", this.OwnModelsFile);

        string? refused = await DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Owon", "ADS800")]);

        Assert.Equal(DataManager.OwnScopeModelsNotReadYet, refused);
        Assert.Equal(before, File.ReadAllText(this.OwnModelsFile));
        Assert.Empty(DataManager.OwnScopeModels);
    }

    // A main file that cannot be read lists no published model - the own ones stay.
    [Fact]
    public void A_main_file_that_cannot_be_read_takes_the_published_models_out_of_the_lists()
    {
        string main = MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z")]);
        OwnOscilloscopeFile.Write(this.OwnModelsFile, [Oscilloscope("Voltcraft", "DOV702")]);
        DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile);

        File.WriteAllText(Path.Combine(this.thisWorkspace.Root, main), "{ not a main file");
        DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile);

        Assert.Empty(DataManager.PublishedScopeModels);
        Assert.Equal(["DOV702"], DataManager.Oscilloscopes.Select(scope => scope.SeriesOrModel));
    }

    // ###########################################################################################
    // *** A RELOAD KEEPS THE OWN MODELS CRT HOLDS (code review, 2026-10-29). *** A manual data check
    // reads the main file again on the pool while the editor may be saving; reading the own file again
    // then could put back a list older than the one just saved. CRT is the file's only writer, so what
    // it holds is the truth once read - a reload only reconciles it with the published models.
    // ###########################################################################################
    [Fact]
    public async Task A_reload_keeps_the_own_models_held_and_only_reconciles_them()
    {
        string main = MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z")]);
        OwnOscilloscopeFile.Write(this.OwnModelsFile, [Oscilloscope("Voltcraft", "DOV702")]);
        DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile);

        Assert.Null(await DataManager.SaveOwnScopeModelsAsync([Oscilloscope("Voltcraft", "DOV702"), Oscilloscope("Owon", "ADS800")]));

        // Something else wrote the file meanwhile - a reload does not read it back over what CRT holds.
        OwnOscilloscopeFile.Write(this.OwnModelsFile, []);
        DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.Write(this.thisWorkspace.Root, [AnyBoard], [Oscilloscope("Rigol", "DS1000Z"), Oscilloscope("Owon", "ADS800")]), this.OwnModelsFile);

        // ADS800 published as it is: taken over; DOV702 kept.
        Assert.Equal(["DOV702"], DataManager.OwnScopeModels.Select(model => model.SeriesOrModel));
        Assert.Equal(["DS1000Z", "ADS800", "DOV702"], DataManager.Oscilloscopes.Select(scope => scope.SeriesOrModel));
    }

    // The user's model sent in and since published as it is: the own copy goes, from the list and
    // from the file - one model, listed once.
    [Fact]
    public void An_own_model_published_since_is_taken_over_at_the_load()
    {
        this.LoadWithOwnModels([Oscilloscope("Voltcraft", "DOV702")], Oscilloscope("Voltcraft", "DOV702"), Oscilloscope("Owon", "ADS800"));

        Assert.Equal(["DOV702", "ADS800"], DataManager.Oscilloscopes.Select(scope => scope.SeriesOrModel));
        Assert.Equal("ADS800", Assert.Single(DataTreeJson.ReadMain(File.ReadAllText(this.OwnModelsFile)).Oscilloscopes).SeriesOrModel);
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
