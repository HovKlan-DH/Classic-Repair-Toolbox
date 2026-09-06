using System.Text.Json;
using System.Text.Json.Nodes;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for UserSettings - the JSON preferences file in the user's AppData.
//
// UserSettings is a static singleton whose setters save to disk immediately, which is exactly
// why it was untestable before: touching it from a test would rewrite the developer's real
// settings. Load() now resolves the AppData path and delegates to the internal LoadFrom(path),
// so these tests point it at a temporary file instead. NOTHING here calls Load().
//
// The class is global mutable state, so this whole file is one xUnit collection: the tests run
// sequentially and each one re-loads a fresh settings file first.
[Collection("UserSettings")]
public sealed class UserSettingsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        // Detach from the temp file so nothing written later can reach a real settings file.
        this.LoadSettings("{}");
        this.thisWorkspace.Dispose();
    }

    /// <summary>Writes a settings file with the given JSON and loads it.</summary>
    private string LoadSettings(string json)
    {
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(path, json);
        UserSettings.LoadFrom(path);
        return path;
    }

    private static JsonNode ReadJson(string path) =>
        JsonNode.Parse(File.ReadAllText(path))!;

    // ---------------------------------------------------------------------- defaults

    [Fact]
    public void An_empty_settings_file_yields_documented_defaults()
    {
        this.LoadSettings("{}");

        Assert.Equal("PAL", UserSettings.Region);
        Assert.Equal("Light", UserSettings.ThemeVariant);
        Assert.Equal(5025, UserSettings.OscilloscopePort);
        Assert.Equal("192.168.0.100", UserSettings.OscilloscopeHost);
        Assert.True(UserSettings.SchematicsShowZones);
        Assert.True(UserSettings.SchematicsShowOppositeSideTraces);
        Assert.False(UserSettings.BlinkSelected);
        Assert.True(UserSettings.EnableNetworkConnectedOscilloscopeTab);
        Assert.True(UserSettings.EnableMiniproExperimentalMode);
        Assert.False(UserSettings.EnableMiniproExperimentalDemoMode);
        Assert.True(UserSettings.EnableWorklog);
        Assert.Equal("CurrentBoard", UserSettings.WorkbooksScope);
    }

    [Fact]
    public void The_two_external_tooling_toggles_are_opt_out_but_the_demo_mode_is_opt_in()
    {
        // The two "External tooling availability" toggles default to TRUE, unlike most flags here,
        // because they gate whether a whole feature is offered at all: the "Oscilloscope" tab and
        // the MiniPro IC-test affordance in the component popup. An existing user upgrading into
        // these settings has no key in their file and must keep the features they already had.
        //
        // The MiniPro demo mode next to them is the opposite - it simulates a programmer that is
        // not attached, is only useful for CRT development, and so stays opt-in.
        this.LoadSettings("{}");

        Assert.True(UserSettings.EnableNetworkConnectedOscilloscopeTab);
        Assert.True(UserSettings.EnableMiniproExperimentalMode);
        Assert.False(UserSettings.EnableMiniproExperimentalDemoMode);
    }

    [Fact]
    public void The_worklog_toggle_is_opt_out_like_the_other_feature_gates()
    {
        // Same reasoning as the two "External tooling availability" toggles above: this gates
        // whether the worklog bar is offered at all, so an existing user upgrading into this
        // setting - no key in their file yet - must see the feature already turned on.
        this.LoadSettings("{}");

        Assert.True(UserSettings.EnableWorklog);
    }

    [Fact]
    public void An_explicit_false_beats_an_opt_out_default()
    {
        // The whole point of a default-true setting is that the stored false must win, otherwise
        // the feature switches itself back on at the next launch.
        this.LoadSettings("""
        {
          "enableNetworkConnectedOscilloscopeTab": false,
          "enableMiniproExperimentalMode": false,
          "enableWorklog": false
        }
        """);

        Assert.False(UserSettings.EnableNetworkConnectedOscilloscopeTab);
        Assert.False(UserSettings.EnableMiniproExperimentalMode);
        Assert.False(UserSettings.EnableWorklog);
    }

    [Fact]
    public void Turning_the_network_connected_oscilloscope_tab_off_persists_and_survives_a_reload()
    {
        // The false value must actually reach the file - a default-true setting that is not
        // written back would silently re-enable the tab on the next launch.
        string path = this.LoadSettings("{}");

        UserSettings.EnableNetworkConnectedOscilloscopeTab = false;

        Assert.False(ReadJson(path)["enableNetworkConnectedOscilloscopeTab"]!.GetValue<bool>());

        UserSettings.LoadFrom(path);
        Assert.False(UserSettings.EnableNetworkConnectedOscilloscopeTab);
    }

    [Fact]
    public void Turning_the_worklog_bar_off_persists_and_survives_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.EnableWorklog = false;

        Assert.False(ReadJson(path)["enableWorklog"]!.GetValue<bool>());

        UserSettings.LoadFrom(path);
        Assert.False(UserSettings.EnableWorklog);
    }

    [Fact]
    public void Worklog_comments_and_work_done_sort_order_default_to_newest_first()
    {
        // Matches the worklog entry editor's own in-memory default from before this was persisted,
        // so an existing user upgrading into this setting sees no change in behaviour.
        this.LoadSettings("{}");

        Assert.True(UserSettings.WorklogCommentsSortNewestFirst);
        Assert.True(UserSettings.WorklogWorkDoneSortNewestFirst);
    }

    [Fact]
    public void Switching_the_worklog_comments_sort_order_persists_and_survives_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.WorklogCommentsSortNewestFirst = false;

        Assert.False(ReadJson(path)["worklogCommentsSortNewestFirst"]!.GetValue<bool>());

        UserSettings.LoadFrom(path);
        Assert.False(UserSettings.WorklogCommentsSortNewestFirst);
    }

    [Fact]
    public void Switching_the_worklog_work_done_sort_order_persists_and_survives_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.WorklogWorkDoneSortNewestFirst = false;

        Assert.False(ReadJson(path)["worklogWorkDoneSortNewestFirst"]!.GetValue<bool>());

        UserSettings.LoadFrom(path);
        Assert.False(UserSettings.WorklogWorkDoneSortNewestFirst);
    }

    [Fact]
    public void The_two_worklog_sort_orders_are_independent()
    {
        string path = this.LoadSettings("{}");

        UserSettings.WorklogCommentsSortNewestFirst = false;

        Assert.True(UserSettings.WorklogWorkDoneSortNewestFirst);

        UserSettings.LoadFrom(path);
        Assert.False(UserSettings.WorklogCommentsSortNewestFirst);
        Assert.True(UserSettings.WorklogWorkDoneSortNewestFirst);
    }

    [Fact]
    public void Worklog_show_entries_checkbox_defaults_to_checked()
    {
        this.LoadSettings("{}");

        Assert.True(UserSettings.WorklogShowEntriesChecked);
    }

    [Fact]
    public void Unchecking_the_worklog_show_entries_checkbox_persists_and_survives_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.WorklogShowEntriesChecked = false;

        Assert.False(ReadJson(path)["worklogShowEntriesChecked"]!.GetValue<bool>());

        UserSettings.LoadFrom(path);
        Assert.False(UserSettings.WorklogShowEntriesChecked);
    }

    [Fact]
    public void Values_present_in_the_file_are_read()
    {
        this.LoadSettings("""
        {
          "region": "NTSC",
          "theme": "Dark",
          "oscilloscopePort": 1234,
          "blinkSelected": true
        }
        """);

        Assert.Equal("NTSC", UserSettings.Region);
        Assert.Equal("Dark", UserSettings.ThemeVariant);
        Assert.Equal(1234, UserSettings.OscilloscopePort);
        Assert.True(UserSettings.BlinkSelected);
    }

    [Fact]
    public void A_malformed_settings_file_falls_back_to_defaults_instead_of_throwing()
    {
        // A hand-edited or truncated settings file must not stop the app from starting.
        this.LoadSettings("{ this is not json");

        Assert.Equal("PAL", UserSettings.Region);
    }

    [Fact]
    public void A_missing_settings_file_is_not_an_error()
    {
        string path = this.thisWorkspace.Path_("does-not-exist.json");

        Exception? thrown = Record.Exception(() => UserSettings.LoadFrom(path));

        Assert.True(thrown is null);
    }

    // ----------------------------------------------------------------- persistence

    [Fact]
    public void Setting_a_value_writes_it_to_the_settings_file_immediately()
    {
        // There is no explicit Save button: every setter persists.
        string path = this.LoadSettings("{}");

        UserSettings.Region = "NTSC";

        Assert.Equal("NTSC", ReadJson(path)["region"]!.GetValue<string>());
    }

    [Fact]
    public void A_persisted_value_survives_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.OscilloscopePort = 4242;
        UserSettings.LoadFrom(path);

        Assert.Equal(4242, UserSettings.OscilloscopePort);
    }

    [Fact]
    public void A_short_circuiting_setter_does_not_rewrite_the_file_for_an_unchanged_value()
    {
        string path = this.LoadSettings("""{"interactiveCadTraceHoverMode": "Always"}""");
        DateTime before = MarkFileAsOld(path);

        UserSettings.InteractiveCadTraceHoverMode = "Always";

        Assert.Equal(before, File.GetLastWriteTimeUtc(path));
    }

    [Fact]
    public void Not_every_setter_short_circuits_an_unchanged_value()
    {
        // CURRENT BEHAVIOUR, and inconsistent: InteractiveCadTraceHoverMode, ContributorMode
        // and SetLastHardware all return early when the value has not changed, but
        // BlinkSelected (and several like it) write the whole file again regardless. Harmless,
        // but it means re-applying UI state does real disk I/O and logs a "Setting changed"
        // line that did not change anything.
        string path = this.LoadSettings("""{"blinkSelected": true}""");
        DateTime before = MarkFileAsOld(path);

        UserSettings.BlinkSelected = true;

        Assert.NotEqual(before, File.GetLastWriteTimeUtc(path));
    }

    /// <summary>Backdates the file so a rewrite is unambiguous rather than a timer-resolution race.</summary>
    private static DateTime MarkFileAsOld(string path)
    {
        File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(-5));
        return File.GetLastWriteTimeUtc(path);
    }

    [Fact]
    public void The_written_file_is_valid_indented_json()
    {
        string path = this.LoadSettings("{}");

        UserSettings.Region = "NTSC";

        string json = File.ReadAllText(path);
        Assert.Contains("\n", json);
        Exception? thrown = Record.Exception(() => JsonDocument.Parse(json));
        Assert.True(thrown is null, "the settings file must stay parseable");
    }

    [Fact]
    public void Unknown_properties_in_the_file_do_not_break_loading()
    {
        // A settings file written by a newer build must still load in an older one.
        this.LoadSettings("""{"region": "NTSC", "somethingFromTheFuture": 42}""");

        Assert.Equal("NTSC", UserSettings.Region);
    }

    // ------------------------------------------------------------- atomic saving
    //
    // Save must never write the settings file in place: an in-place write truncates first, so a
    // crash, power loss or full disk in that window destroys EVERY preference at once, and the
    // next launch silently falls back to defaults. Save therefore writes a sibling ".tmp" file
    // and swaps it over the real one (the same pattern OnlineServices.DownloadFileAsync uses).
    // A hard crash mid-write cannot be forced from a unit test, so these tests pin the two
    // observable halves of that mechanism instead: the swap (not an in-place write) is what
    // touches the real file, and a temp left behind by a crash is harmless and cleaned up.

    [Fact]
    public void Saving_succeeds_while_another_handle_reads_the_settings_file()
    {
        // Backup tools, sync clients and antivirus scanners read the settings file while the app
        // runs. On Windows an in-place rewrite needs write access and dies on a sharing
        // violation, silently losing the change; the atomic swap only needs to replace the name,
        // which a read-sharing handle permits. (Unix does not enforce sharing modes, so there
        // this documents the swap; on Windows it proves it.)
        string path = this.LoadSettings("""{"region": "PAL"}""");

        using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete))
        {
            UserSettings.Region = "NTSC";
        }

        UserSettings.LoadFrom(path);
        Assert.Equal("NTSC", UserSettings.Region);
    }

    [Fact]
    public void A_leftover_temp_file_from_a_crashed_save_is_harmless_and_cleaned_up()
    {
        // A crash between writing the temp and swapping it in leaves "<settings>.tmp" behind
        // while the real file stays intact. The next load must read the real file, and the next
        // save must clear the leftover away instead of accumulating it forever.
        string path = this.LoadSettings("""{"region": "NTSC"}""");
        string tempPath = path + ".tmp";
        File.WriteAllText(tempPath, "{ garbage from a crashed save");

        UserSettings.LoadFrom(path);
        Assert.Equal("NTSC", UserSettings.Region);

        UserSettings.OscilloscopePort = 4242;

        Assert.False(File.Exists(tempPath), "a successful save must leave no .tmp file behind");
        Assert.Equal(4242, ReadJson(path)["oscilloscopePort"]!.GetValue<int>());
    }

    [Fact]
    public void A_save_that_cannot_complete_leaves_the_previous_settings_file_intact()
    {
        // Fail closed: blocking the temp path (a directory squats on the name) makes the save
        // fail before the swap, so the real file must still hold the old, parseable content.
        string path = this.LoadSettings("""{"region": "NTSC"}""");
        Directory.CreateDirectory(path + ".tmp");

        UserSettings.Region = "PAL";   // save fails silently - logged, not thrown

        Assert.Equal("NTSC", ReadJson(path)["region"]!.GetValue<string>());
    }

    // ------------------------------------------------------------- per-board values

    [Fact]
    public void An_unset_board_splitter_ratio_is_reported_as_absent()
    {
        this.LoadSettings("{}");

        Assert.False(UserSettings.HasSchematicsSplitterRatio("C64/250407"));
    }

    [Fact]
    public void A_board_splitter_ratio_round_trips()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetSchematicsSplitterRatio("C64/250407", 0.375);
        UserSettings.LoadFrom(path);

        Assert.True(UserSettings.HasSchematicsSplitterRatio("C64/250407"));
        Assert.Equal(0.375, UserSettings.GetSchematicsSplitterRatio("C64/250407"));
    }

    [Fact]
    public void Board_settings_are_kept_apart_per_board()
    {
        this.LoadSettings("{}");

        UserSettings.SetSchematicsSplitterRatio("C64/250407", 0.25);
        UserSettings.SetSchematicsSplitterRatio("Plus4/310163", 0.75);

        Assert.Equal(0.25, UserSettings.GetSchematicsSplitterRatio("C64/250407"));
        Assert.Equal(0.75, UserSettings.GetSchematicsSplitterRatio("Plus4/310163"));
    }

    [Fact]
    public void The_last_schematic_round_trips_per_board()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetLastSchematicForBoard("C64/250407", "Sheet 3");
        UserSettings.LoadFrom(path);

        Assert.Equal("Sheet 3", UserSettings.GetLastSchematicForBoard("C64/250407"));
        Assert.Null(UserSettings.GetLastSchematicForBoard("Plus4/310163"));
    }

    [Fact]
    public void A_blank_board_key_is_ignored_by_the_per_board_setters()
    {
        this.LoadSettings("{}");

        Exception? thrown = Record.Exception(() =>
            UserSettings.SetSchematicsSplitterRatio("   ", 0.5));

        Assert.True(thrown is null);
        Assert.False(UserSettings.HasSchematicsSplitterRatio("   "));
    }

    [Fact]
    public void Selected_categories_round_trip_per_board()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetSelectedCategories("C64/250407", new List<string> { "IC", "Capacitor" });
        UserSettings.LoadFrom(path);

        Assert.Equal(new[] { "IC", "Capacitor" }, UserSettings.GetSelectedCategories("C64/250407"));
        Assert.Null(UserSettings.GetSelectedCategories("Plus4/310163"));
    }

    // The workbook the Workbooks tab "activated" (Main.ActivateWorkbook) for a board - what
    // Main.ResolveActiveWorkbookForBoard reads back to decide which workbook the worklog bar,
    // "Show worklogs" and "Add worklog" all act on, in place of the board's newest workbook.
    [Fact]
    public void Active_workbook_id_round_trips_per_board()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetActiveWorkbookId("C64/250407", 31);
        UserSettings.LoadFrom(path);

        Assert.Equal(31, UserSettings.GetActiveWorkbookId("C64/250407"));

        // A board with nothing activated has no saved id at all - that is what tells
        // ResolveActiveWorkbookForBoard to fall back to the newest workbook, rather than to
        // workbook id 0 or some other sentinel that could collide with a real id.
        Assert.Null(UserSettings.GetActiveWorkbookId("Plus4/310163"));
    }

    [Fact]
    public void Activating_a_different_workbook_for_the_same_board_overwrites_the_previous_one()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetActiveWorkbookId("C64/250407", 31);
        UserSettings.SetActiveWorkbookId("C64/250407", 42);
        UserSettings.LoadFrom(path);

        // Only the most recently activated workbook is kept per board - activating a second one
        // is a replacement, not an addition to some history of activations.
        Assert.Equal(42, UserSettings.GetActiveWorkbookId("C64/250407"));
    }

    // The Workbooks tab's two splitters (left workbook list / rest, board pane / entry list) -
    // plain app-wide pixel widths, not per-board like SchematicsSplitterRatios, since this tab's
    // layout does not depend on which board is selected. Defaults match the AXAML's own design-time
    // column widths (200 and 347) so a first run does not jump on its first save.
    [Fact]
    public void Workbooks_left_panel_width_round_trips()
    {
        string path = this.LoadSettings("{}");

        Assert.Equal(200.0, UserSettings.WorkbooksLeftPanelWidth);

        UserSettings.WorkbooksLeftPanelWidth = 260.0;
        UserSettings.LoadFrom(path);

        Assert.Equal(260.0, UserSettings.WorkbooksLeftPanelWidth);
    }

    [Fact]
    public void Workbooks_entry_list_width_round_trips()
    {
        string path = this.LoadSettings("{}");

        Assert.Equal(347.0, UserSettings.WorkbooksEntryListWidth);

        UserSettings.WorkbooksEntryListWidth = 410.0;
        UserSettings.LoadFrom(path);

        Assert.Equal(410.0, UserSettings.WorkbooksEntryListWidth);
    }

    [Fact]
    public void Schematic_order_round_trips_per_board()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetSchematicsOrder("C64/250407", new List<string> { "Sheet 2", "Sheet 1" });
        UserSettings.LoadFrom(path);

        Assert.Equal(new[] { "Sheet 2", "Sheet 1" }, UserSettings.GetSchematicsOrder("C64/250407"));
    }

    // ------------------------------------------------------- last-used selections

    [Fact]
    public void The_last_hardware_round_trips()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetLastHardware("C64");
        UserSettings.LoadFrom(path);

        Assert.Equal("C64", UserSettings.GetLastHardware());
    }

    [Fact]
    public void A_blank_last_hardware_is_ignored()
    {
        this.LoadSettings("""{"lastHardware": "C64"}""");

        UserSettings.SetLastHardware("   ");

        Assert.Equal("C64", UserSettings.GetLastHardware());
    }

    [Fact]
    public void The_last_board_is_remembered_per_hardware()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetLastBoardForHardware("C64", "250407");
        UserSettings.SetLastBoardForHardware("Plus4", "310163");
        UserSettings.LoadFrom(path);

        Assert.Equal("250407", UserSettings.GetLastBoardForHardware("C64"));
        Assert.Equal("310163", UserSettings.GetLastBoardForHardware("Plus4"));
        Assert.Null(UserSettings.GetLastBoardForHardware("VIC20"));
    }

    [Fact]
    public void Hardware_lookup_is_case_insensitive()
    {
        this.LoadSettings("{}");

        UserSettings.SetLastBoardForHardware("C64", "250407");

        Assert.Equal("250407", UserSettings.GetLastBoardForHardware("c64"));
    }

    [Fact]
    public void The_last_oscilloscope_series_is_remembered_per_vendor()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetLastOscilloscopeSeriesForVendor("Rigol", "DS1000Z");
        UserSettings.LoadFrom(path);

        Assert.Equal("DS1000Z", UserSettings.GetLastOscilloscopeSeriesForVendor("Rigol"));
        Assert.Null(UserSettings.GetLastOscilloscopeSeriesForVendor("Siglent"));
    }

    // ------------------------------------------------------------- theme colours

    [Fact]
    public void Missing_theme_colours_are_filled_in_with_defaults_on_load()
    {
        this.LoadSettings("{}");

        var colors = UserSettings.GetUserThemeColors();

        Assert.NotEmpty(colors);
        Assert.True(colors.ContainsKey("Bg"));
        Assert.True(colors.ContainsKey("Schematics_FirstPin"));
    }

    [Fact]
    public void A_customised_theme_colour_is_not_overwritten_by_the_default()
    {
        this.LoadSettings("""{"userThemeColors": {"Bg": "#123456"}}""");

        Assert.Equal("#123456", UserSettings.GetUserThemeColors()["Bg"]);
    }

    [Fact]
    public void A_customised_theme_still_gains_any_newly_added_default_keys()
    {
        // Upgrading to a build with a new themed element must not leave that element unstyled.
        this.LoadSettings("""{"userThemeColors": {"Bg": "#123456"}}""");

        var colors = UserSettings.GetUserThemeColors();

        Assert.Equal("#123456", colors["Bg"]);
        Assert.True(colors.ContainsKey("Thumbnail_Border"));
    }

    [Fact]
    public void IsUserThemeColorResourceKey_only_matches_the_four_trace_palette_colours()
    {
        // Despite the broad name, this predicate is NOT "is this a user theme colour" - it
        // recognises only the four Schematics_TracePalette_ColorN keys. "Bg" is a themed
        // colour and still returns false.
        this.LoadSettings("{}");

        Assert.True(UserSettings.IsUserThemeColorResourceKey("Schematics_TracePalette_Color1"));
        Assert.True(UserSettings.IsUserThemeColorResourceKey("Schematics_TracePalette_Color4"));

        Assert.False(UserSettings.IsUserThemeColorResourceKey("Bg"));
        Assert.False(UserSettings.IsUserThemeColorResourceKey("NotAThemeKey"));
    }

    [Fact]
    public void IsUserThemeColorResourceKey_is_case_sensitive()
    {
        this.LoadSettings("{}");

        Assert.False(UserSettings.IsUserThemeColorResourceKey("schematics_tracepalette_color1"));
    }

    // --------------------------------------------------------- contributor mode

    [Fact]
    public void Contributor_mode_defaults_to_off()
    {
        this.LoadSettings("{}");

        Assert.False(UserSettings.ContributorMode);
    }

    [Fact]
    public void Legacy_settings_are_migrated_into_contributor_mode()
    {
        // Before contributor mode existed, the pair of validate-on-launch + debug-logging was
        // what a contributor turned on. Those users keep their setup after upgrading.
        this.LoadSettings("""{"validateDataOnLaunch": true, "debugLogging": true}""");

        Assert.True(UserSettings.ContributorMode);
    }

    [Fact]
    public void Only_one_legacy_flag_does_not_enable_contributor_mode()
    {
        this.LoadSettings("""{"validateDataOnLaunch": true, "debugLogging": false}""");

        Assert.False(UserSettings.ContributorMode);
    }

    [Fact]
    public void An_explicit_contributor_mode_value_is_never_overridden_by_migration()
    {
        this.LoadSettings("""
        {"contributorMode": false, "validateDataOnLaunch": true, "debugLogging": true}
        """);

        Assert.False(UserSettings.ContributorMode);
    }

    [Fact]
    public void Contributor_mode_drives_debug_logging()
    {
        this.LoadSettings("""{"contributorMode": true}""");

        Assert.True(Logger.IsDebugEnabled);

        this.LoadSettings("""{"contributorMode": false}""");

        Assert.False(Logger.IsDebugEnabled);
    }

    [Fact]
    public void Contributor_mode_round_trips_per_board()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SetContributorModeForBoard("C64/250407", true);
        UserSettings.LoadFrom(path);

        Assert.True(UserSettings.GetContributorModeForBoard("C64/250407"));
    }

    // ------------------------------------------------------------------- events

    [Fact]
    public void Changing_the_trace_hover_mode_raises_its_change_event()
    {
        this.LoadSettings("""{"interactiveCadTraceHoverMode": "Always"}""");

        int raised = 0;
        void Handler() => raised++;

        UserSettings.InteractiveCadTraceHoverModeChanged += Handler;
        try
        {
            UserSettings.InteractiveCadTraceHoverMode = "HoldShift";
            Assert.Equal(1, raised);

            UserSettings.InteractiveCadTraceHoverMode = "HoldShift";   // unchanged
            Assert.Equal(1, raised);
        }
        finally
        {
            UserSettings.InteractiveCadTraceHoverModeChanged -= Handler;
        }
    }

    [Theory]
    [InlineData("Disabled", "Disabled")]
    [InlineData("HoldShift", "HoldShift")]
    [InlineData("holdshift", "HoldShift")]   // matched case-insensitively...
    [InlineData("Always", "Always")]
    [InlineData("Hold shift", "Always")]     // ...but a space is NOT the same token
    [InlineData("nonsense", "Always")]       // anything unrecognised falls back to Always
    public void The_trace_hover_mode_is_normalised_to_one_of_three_tokens(string input, string expected)
    {
        this.LoadSettings("{}");

        UserSettings.InteractiveCadTraceHoverMode = input;

        Assert.Equal(expected, UserSettings.InteractiveCadTraceHoverMode);
    }

    [Theory]
    [InlineData("AllBoards", "AllBoards")]
    [InlineData("allboards", "AllBoards")]     // matched case-insensitively...
    [InlineData("CurrentBoard", "CurrentBoard")]
    [InlineData("All Boards", "CurrentBoard")] // ...but a space is NOT the same token
    [InlineData("nonsense", "CurrentBoard")]   // anything unrecognised falls back to CurrentBoard
    public void The_workbooks_scope_is_normalised_to_one_of_two_tokens(string input, string expected)
    {
        this.LoadSettings("{}");

        UserSettings.WorkbooksScope = input;

        Assert.Equal(expected, UserSettings.WorkbooksScope);
    }

    [Fact]
    public void The_workbooks_scope_persists_across_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.WorkbooksScope = "AllBoards";

        UserSettings.LoadFrom(path);

        Assert.Equal("AllBoards", UserSettings.WorkbooksScope);
    }

    [Fact]
    public void Changing_the_workbooks_scope_raises_its_change_event_only_when_it_actually_changes()
    {
        this.LoadSettings("""{"workbooksScope": "CurrentBoard"}""");

        int raised = 0;
        void Handler() => raised++;

        UserSettings.WorkbooksScopeChanged += Handler;
        try
        {
            UserSettings.WorkbooksScope = "AllBoards";
            Assert.Equal(1, raised);

            UserSettings.WorkbooksScope = "AllBoards";   // unchanged
            Assert.Equal(1, raised);
        }
        finally
        {
            UserSettings.WorkbooksScopeChanged -= Handler;
        }
    }

    // ---------------------------------------------------------------------- worklog currency

    // An existing installation has no key in its file, and every cost it has already recorded was
    // typed as a bare number. USD is what those numbers become - the choice the Configuration tab
    // opens on until the user picks their own country.
    [Fact]
    public void The_worklog_currency_defaults_to_the_dollar()
    {
        this.LoadSettings("{}");

        Assert.Equal("USD", UserSettings.WorklogCurrencyCode);
    }

    // The getter normalises rather than trusting the stored string - settings.json is a plain file
    // a user can edit, and this code is printed beside every cost in a customer's exported PDF.
    [Theory]
    [InlineData("DKK", "DKK")]
    [InlineData("dkk", "DKK")]        // matched case-insensitively
    [InlineData("  NOK  ", "NOK")]    // and trimmed
    [InlineData("XYZ", "USD")]        // a code no country here uses falls back
    [InlineData("nonsense", "USD")]
    [InlineData("", "USD")]
    public void The_worklog_currency_is_normalised_to_a_known_code(string input, string expected)
    {
        this.LoadSettings("{}");

        UserSettings.WorklogCurrencyCode = input;

        Assert.Equal(expected, UserSettings.WorklogCurrencyCode);
    }

    // A code stored by hand in the file, not through the setter, is normalised on the way OUT too -
    // otherwise the setter's guard is the only defence and a hand-edited file bypasses it entirely.
    [Fact]
    public void A_hand_edited_currency_code_is_normalised_when_read_back()
    {
        this.LoadSettings("""{"worklogCurrencyCode": "eur"}""");

        Assert.Equal("EUR", UserSettings.WorklogCurrencyCode);
    }

    [Fact]
    public void The_worklog_currency_persists_across_a_reload()
    {
        string path = this.LoadSettings("{}");

        UserSettings.WorklogCurrencyCode = "DKK";

        UserSettings.LoadFrom(path);

        Assert.Equal("DKK", UserSettings.WorklogCurrencyCode);
    }

    // The Workbooks tab's summary strip and entry cards are built once per refresh, so they only
    // reprint their figures because this event fires. The unchanged-value guard matters for the
    // same reason WorkbooksScope's does: a rebuild is a full disk rescan and schematic re-decode,
    // and re-selecting the country already chosen must not cost one.
    [Fact]
    public void Changing_the_worklog_currency_raises_its_change_event_only_when_it_actually_changes()
    {
        this.LoadSettings("""{"worklogCurrencyCode": "USD"}""");

        int raised = 0;
        void Handler() => raised++;

        UserSettings.WorklogCurrencyChanged += Handler;
        try
        {
            UserSettings.WorklogCurrencyCode = "DKK";
            Assert.Equal(1, raised);

            UserSettings.WorklogCurrencyCode = "DKK";   // unchanged
            Assert.Equal(1, raised);

            // Same code, different case - still unchanged once normalised, so still no rebuild.
            UserSettings.WorklogCurrencyCode = "dkk";
            Assert.Equal(1, raised);
        }
        finally
        {
            UserSettings.WorklogCurrencyChanged -= Handler;
        }
    }

    [Fact]
    public void Changing_check_data_on_launch_raises_its_change_event_with_the_new_value()
    {
        this.LoadSettings("""{"checkDataOnLaunch": true}""");

        bool? observed = null;
        void Handler(bool value) => observed = value;

        UserSettings.CheckDataOnLaunchChanged += Handler;
        try
        {
            UserSettings.CheckDataOnLaunch = false;
            Assert.False(observed);
        }
        finally
        {
            UserSettings.CheckDataOnLaunchChanged -= Handler;
        }
    }

    // --------------------------------------------------------- window placement

    [Fact]
    public void Window_placement_round_trips()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SaveWindowPlacement(
            "Maximized", 1280, 800, 10, 20, 0, 0, 1920, 1080, 1.5);
        UserSettings.LoadFrom(path);

        JsonNode json = ReadJson(path);

        Assert.Equal("Maximized", json["windowState"]!.GetValue<string>());
        Assert.Equal(1280, json["windowWidth"]!.GetValue<double>());
        Assert.Equal(1.5, json["windowScreenScaling"]!.GetValue<double>());
        Assert.True(json["hasWindowPlacement"]!.GetValue<bool>());
    }

    [Fact]
    public void Component_info_window_layout_round_trips()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SaveComponentInfoWindowLayout("Normal", 900, 500, 0.4, 120, 150, 80);

        JsonNode json = ReadJson(path);

        Assert.Equal(900, json["componentInfoWindowWidth"]!.GetValue<double>());
        Assert.Equal(0.4, json["componentInfoWindowLeftColumnRatio"]!.GetValue<double>());
        Assert.Equal(150, json["componentInfoWindowX"]!.GetValue<int>());
        Assert.Equal(80, json["componentInfoWindowY"]!.GetValue<int>());
        Assert.True(json["hasComponentInfoWindowLayout"]!.GetValue<bool>());
    }

    [Fact]
    public void Worklog_entry_window_layout_round_trips()
    {
        string path = this.LoadSettings("{}");

        UserSettings.SaveWorklogEntryWindowLayout("Maximized", 1000, 720, 240, 130, 1920, 0, 0.42);
        UserSettings.LoadFrom(path);

        JsonNode json = ReadJson(path);

        Assert.Equal("Maximized", json["worklogEntryWindowState"]!.GetValue<string>());
        Assert.Equal(1000, json["worklogEntryWindowWidth"]!.GetValue<double>());
        Assert.Equal(720, json["worklogEntryWindowHeight"]!.GetValue<double>());
        Assert.Equal(240, json["worklogEntryWindowX"]!.GetValue<int>());
        Assert.Equal(130, json["worklogEntryWindowY"]!.GetValue<int>());
        Assert.Equal(1920, json["worklogEntryWindowScreenX"]!.GetValue<int>());
        Assert.Equal(0, json["worklogEntryWindowScreenY"]!.GetValue<int>());
        Assert.Equal(0.42, json["worklogEntryWindowLeftColumnRatio"]!.GetValue<double>());
        Assert.True(json["hasWorklogEntryWindowLayout"]!.GetValue<bool>());

        // And it reads back through the accessors the window restores from.
        Assert.True(UserSettings.HasWorklogEntryWindowLayout);
        Assert.Equal("Maximized", UserSettings.WorklogEntryWindowState);
        Assert.Equal(1000, UserSettings.WorklogEntryWindowWidth);
        Assert.Equal(720, UserSettings.WorklogEntryWindowHeight);
        Assert.Equal(240, UserSettings.WorklogEntryWindowX);
        Assert.Equal(130, UserSettings.WorklogEntryWindowY);
        Assert.Equal(1920, UserSettings.WorklogEntryWindowScreenX);
        Assert.Equal(0, UserSettings.WorklogEntryWindowScreenY);
        Assert.Equal(0.42, UserSettings.WorklogEntryWindowLeftColumnRatio);
    }

    // THE BUG THIS EXISTS TO CLOSE: a window maximized on a second monitor, without ever being
    // un-maximized there, must still be told to restore onto that monitor - not onto whichever one
    // its last WINDOWED position named, which can be a different one entirely (see
    // WorklogEntryWindowScreenX's own comment). Asserted separately from the round-trip test above
    // so a future change to one cannot silently stop covering the other.
    [Fact]
    public void Worklog_entry_window_screen_position_is_independent_of_the_normal_state_position()
    {
        string path = this.LoadSettings("{}");

        // Last windowed on the primary monitor (0,0), but maximized on a second one to the right.
        UserSettings.SaveWorklogEntryWindowLayout("Maximized", 1000, 720, 50, 50, 1920, 0, 0.5);
        UserSettings.LoadFrom(path);

        Assert.Equal(50, UserSettings.WorklogEntryWindowX);
        Assert.Equal(50, UserSettings.WorklogEntryWindowY);
        Assert.Equal(1920, UserSettings.WorklogEntryWindowScreenX);
        Assert.Equal(0, UserSettings.WorklogEntryWindowScreenY);
    }

    // The flag is what tells the window whether to apply a saved position at all. Without a saved
    // layout the window must keep WindowStartupLocation="CenterOwner" rather than jumping to (0,0),
    // so a fresh settings file has to report false.
    [Fact]
    public void A_fresh_settings_file_has_no_worklog_entry_window_layout()
    {
        this.LoadSettings("{}");

        Assert.False(UserSettings.HasWorklogEntryWindowLayout);
    }

    // Defaults match the window's own declared size, so a first run and a settings file that
    // predates this feature both open at the size the XAML asks for.
    [Fact]
    public void The_default_worklog_entry_window_size_matches_the_windows_declared_size()
    {
        this.LoadSettings("{}");

        Assert.Equal(1200, UserSettings.WorklogEntryWindowWidth);
        Assert.Equal(800, UserSettings.WorklogEntryWindowHeight);
        Assert.Equal("Normal", UserSettings.WorklogEntryWindowState);

        // 0.6 is the 3*:2* split the XAML declares, so an unsaved window opens exactly as designed.
        Assert.Equal(0.6, UserSettings.WorklogEntryWindowLeftColumnRatio);
    }

    // A settings file written by an older build has none of these keys; reading it must not throw
    // and must fall back to the defaults rather than a zero-sized window.
    [Fact]
    public void A_settings_file_without_the_worklog_keys_falls_back_to_defaults()
    {
        this.LoadSettings("""{ "region": "NTSC" }""");

        Assert.False(UserSettings.HasWorklogEntryWindowLayout);
        Assert.Equal(1200, UserSettings.WorklogEntryWindowWidth);
        Assert.Equal(800, UserSettings.WorklogEntryWindowHeight);
        Assert.Equal(0.6, UserSettings.WorklogEntryWindowLeftColumnRatio);
    }
}
