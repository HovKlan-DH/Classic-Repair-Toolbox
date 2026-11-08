using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using System.IO;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The PER-BOARD half of "Detach thumbnails into their own window", driven through a real Main so
// the tier-switching is exercised end to end: which tier a toggle WRITES, which tier the
// Configuration checkbox READS back, and what happens to an already-open window when the tier
// itself changes underneath it.
//
// Everything here was reachable only through Main. TabConfigurationDetachThumbnailsTests
// constructs a bare TabConfiguration with no Main above it, so every one of its tests takes the
// global-flag fallback - which is exactly why none of them caught the checkbox reading the global
// flag while the toggle wrote the per-board one.
//
// Board selection is faked with Main.BoardKeyOverrideForTests: selecting a board for real
// populates the dropdowns, raising OnBoardSelectionChanged and loading that board's Excel data
// off disk.
//
// COLLECTION NOTE: "HeadlessUi" - same reason as MainWindowTests' own header: these construct a
// Window and so need the shared dispatcher thread. They drive UserSettings' static state, which
// is safe only because xunit.runner.json sets "parallelizeTestCollections": false.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ThumbnailsDetachPerBoardTests : IDisposable
{
    private const string BoardA = "Commodore C64|250407";
    private const string BoardB = "Commodore Plus4|310163";

    private readonly TempWorkspace thisWorkspace = new();

    public ThumbnailsDetachPerBoardTests()
    {
        this.LoadEmptySettings();
    }

    public void Dispose()
    {
        this.LoadEmptySettings();
        this.thisWorkspace.Dispose();
    }

    // Writes an actual (empty) settings FILE before loading it rather than pointing UserSettings at
    // a path that does not exist - LoadFrom only resets its in-memory _data when it can deserialize
    // something, so a missing file leaves whatever the previous test left behind. Same reason and
    // same shape as SchematicsThumbnailsWindowPerBoardTests' own helper.
    private void LoadEmptySettings()
    {
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(path, "{}");
        UserSettings.LoadFrom(path);
    }

    // Main.ApplyWorklogBarVisibility's enable path reads the workbook folder, so point it at a temp
    // one rather than the developer's real Workbook directory.
    private void RedirectWorklogToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
    }

    // Builds a shown Main on a fake board selection, and hands it to the body. Shown because the
    // detached window is opened with Show(owner), which Avalonia refuses against a non-visible
    // owner. Never calls StartAsync - see MainWindowTests' header for why that rule matters.
    private void WithMainOnBoard(string? boardKey, Action<CRT.Main> body)
    {
        this.RedirectWorklogToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            main.BoardKeyOverrideForTests = boardKey;
            main.Show();

            try
            {
                body(main);
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    private static CheckBox DetachCheckBox(CRT.Main main)
        => main.GetControl<TabConfiguration>("TabConfiguration")
               .GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");

    private static CheckBox RememberPerBoardCheckBox(CRT.Main main)
        => main.GetControl<TabConfiguration>("TabConfiguration")
               .GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");

    // Selects the Configuration tab and lets the layout run, so its content is actually attached to
    // the visual tree. Needed by any test that drives a checkbox whose handler reaches Main through
    // TopLevel.GetTopLevel (the idiom throughout TabConfiguration): a TabControl does not realise an
    // unselected tab's content, so GetTopLevel returns null for it and the handler silently takes
    // its no-Main branch. Selecting the tab is also what a user does before clicking the box.
    private static void ShowConfigurationTab(CRT.Main main)
    {
        main.GetControl<TabControl>("MainTabControl").SelectedItem =
            main.GetControl<TabItem>("ConfigurationTabItem");

        main.UpdateLayout();
    }

    // -----------------------------------------------------------------------------------------
    // The checkbox must read back the SAME tier the toggle writes
    // -----------------------------------------------------------------------------------------

    // Reported shape: with per-board mode on, SetThumbnailsDetached wrote the PER-BOARD flag while
    // the checkbox resynced itself from the GLOBAL one - so the box visibly re-ticked itself to the
    // wrong state immediately after every toggle. Fails against that version, where the box lands
    // back on the global false.
    [Fact]
    public void Ticking_detach_in_per_board_mode_leaves_the_checkbox_ticked()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.SetThumbnailsDetached(true);

            // Written to this board's own record, and the box agrees with it.
            Assert.True(UserSettings.GetThumbnailWindowSettingsForBoard(BoardA)?.IsDetached);
            Assert.True(DetachCheckBox(main).IsChecked);

            // The global flag is NOT what was written, and must be left alone entirely.
            Assert.False(UserSettings.DetachSchematicsThumbnails);
        });
    }

    // The mirror: unticking must land on false rather than bouncing back to whatever the global
    // flag happens to hold. A global true is set here precisely so a checkbox reading the wrong
    // tier would visibly re-tick itself.
    [Fact]
    public void Unticking_detach_in_per_board_mode_leaves_the_checkbox_unticked()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.SetThumbnailsDetached(true);
            Assert.True(DetachCheckBox(main).IsChecked);

            main.SetThumbnailsDetached(false);

            Assert.False(UserSettings.GetThumbnailWindowSettingsForBoard(BoardA)?.IsDetached);
            Assert.False(DetachCheckBox(main).IsChecked);
        });
    }

    // Each board keeps its own answer, and the checkbox follows the board being shown - the whole
    // point of the feature.
    [Fact]
    public void Each_board_keeps_its_own_detach_state_and_the_checkbox_follows_the_shown_board()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.SetThumbnailsDetached(true);
            Assert.True(DetachCheckBox(main).IsChecked);

            // Switch board (as OnBoardSelectionChanged would) and resync: board B has never had a
            // choice made for it, so it starts embedded.
            main.BoardKeyOverrideForTests = BoardB;
            main.RefreshDetachThumbnailsCheckBox();

            Assert.False(DetachCheckBox(main).IsChecked);

            // Board A's own choice is untouched by any of that.
            Assert.True(UserSettings.GetThumbnailWindowSettingsForBoard(BoardA)?.IsDetached);
            Assert.Null(UserSettings.GetThumbnailWindowSettingsForBoard(BoardB));
        });
    }

    // With per-board mode OFF the checkbox must still read the global flag, exactly as before -
    // the fix must not have moved the non-per-board case onto the per-board tier.
    [Fact]
    public void With_per_board_mode_off_the_checkbox_still_follows_the_global_flag()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = false;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.SetThumbnailsDetached(true);

            Assert.True(UserSettings.DetachSchematicsThumbnails);
            Assert.True(DetachCheckBox(main).IsChecked);

            // Nothing was written into the per-board tier.
            Assert.Null(UserSettings.GetThumbnailWindowSettingsForBoard(BoardA));
        });
    }

    // -----------------------------------------------------------------------------------------
    // No board selected
    // -----------------------------------------------------------------------------------------

    // GetCurrentBoardKey documents an empty string for "nothing selected". Storing a detach state
    // under that key gave every no-selection moment one shared "" record, which a real board could
    // then inherit a detach state from. The global flag is the honest place for the choice, and is
    // what the resolver falls back to for the same key.
    [Fact]
    public void With_no_board_selected_the_choice_goes_to_the_global_flag_not_a_shared_empty_key()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;

        this.WithMainOnBoard(string.Empty, main =>
        {
            main.SetThumbnailsDetached(true);

            Assert.Null(UserSettings.GetThumbnailWindowSettingsForBoard(string.Empty));
            Assert.True(UserSettings.DetachSchematicsThumbnails);

            // And the checkbox still agrees with where it actually went.
            Assert.True(DetachCheckBox(main).IsChecked);
        });
    }

    // The leak the empty key would cause, asserted from the other end: a board selected afterwards
    // must not come up detached because of something stored while nothing was selected.
    [Fact]
    public void A_real_board_does_not_inherit_a_detach_state_stored_with_no_selection()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;

        this.WithMainOnBoard(string.Empty, main =>
        {
            main.SetThumbnailsDetached(true);

            main.BoardKeyOverrideForTests = BoardA;

            Assert.False(main.ResolveThumbnailsDetachedForCurrentBoard());
        });
    }

    // -----------------------------------------------------------------------------------------
    // Toggling "Remember thumbnail window settings per board" itself
    // -----------------------------------------------------------------------------------------

    // Reported shape: the per-board checkbox persisted its setting but never re-applied the layout,
    // so the open window and the collapsed thumbnail column stayed out of line with the newly
    // selected tier until the next board change or restart. Here the global flag says detached
    // while this board has no per-board record, so ticking the option must EMBED the thumbnails
    // immediately. Fails against the persist-only version, where the window stays open.
    [Fact]
    public void Ticking_per_board_mode_applies_the_new_tier_to_the_open_window_at_once()
    {
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            ShowConfigurationTab(main);

            main.ApplyThumbnailsDetachedState();
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            // The user ticks "Remember thumbnail window settings per board". This board has no
            // record of its own, so its answer becomes "embedded".
            RememberPerBoardCheckBox(main).IsChecked = true;

            Assert.True(UserSettings.RememberThumbnailWindowSettingsPerBoard);
            Assert.False(main.IsThumbnailsWindowOpenForTests);
            Assert.False(DetachCheckBox(main).IsChecked);
        });
    }

    // The other direction: unticking it hands the answer back to the global flag, which here says
    // detached - so the window must open, even though this board's own record says embedded.
    [Fact]
    public void Unticking_per_board_mode_hands_the_answer_back_to_the_global_flag_at_once()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            ShowConfigurationTab(main);

            // This board is explicitly embedded in the per-board tier.
            UserSettings.SetThumbnailWindowDetachedForBoard(BoardA, false);
            main.ApplyThumbnailsDetachedState();
            Assert.False(main.IsThumbnailsWindowOpenForTests);

            RememberPerBoardCheckBox(main).IsChecked = false;

            Assert.False(UserSettings.RememberThumbnailWindowSettingsPerBoard);
            Assert.True(main.IsThumbnailsWindowOpenForTests);
            Assert.True(DetachCheckBox(main).IsChecked);
        });
    }

    // Switching tier must not be read by the window's Closed handler as the user dismissing it -
    // that would turn the detach setting off as a side effect of a bookkeeping close. Both tiers
    // say detached here, so the window must simply be replaced and the settings left alone.
    [Fact]
    public void Switching_tier_with_a_window_open_reopens_it_without_turning_detach_off()
    {
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            ShowConfigurationTab(main);

            UserSettings.SetThumbnailWindowDetachedForBoard(BoardA, true);
            main.ApplyThumbnailsDetachedState();
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            RememberPerBoardCheckBox(main).IsChecked = true;

            // Still open (reopened against the per-board tier), and NEITHER tier was turned off by
            // the close that did it.
            Assert.True(main.IsThumbnailsWindowOpenForTests);
            Assert.True(UserSettings.DetachSchematicsThumbnails);
            Assert.True(UserSettings.GetThumbnailWindowSettingsForBoard(BoardA)?.IsDetached);
        });
    }

    // -----------------------------------------------------------------------------------------
    // A board change with a window open
    // -----------------------------------------------------------------------------------------

    // Unticking per-board mode while a window was open used to make the board-change path return
    // early forever, so the window kept the thisPerBoardKey it was opened with and wrote the NEW
    // board's geometry into the OLD board's saved layout. The window resolves its key once, in
    // Initialize, so the only fix is a fresh window - asserted here by the old board's layout
    // staying unwritten while a window opened under the per-board tier is replaced.
    [Fact]
    public void Turning_per_board_mode_off_with_a_window_open_replaces_it_rather_than_leaving_a_stale_key()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            ShowConfigurationTab(main);

            UserSettings.SetThumbnailWindowDetachedForBoard(BoardA, true);
            main.ApplyThumbnailsDetachedState();
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            // Per-board mode goes off; the window is replaced by one on the global tier.
            RememberPerBoardCheckBox(main).IsChecked = false;
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            // Board B is now selected. The window in front of the user must be saving into the
            // GLOBAL tier, not still into board A's record - so closing it stamps the global one.
            main.BoardKeyOverrideForTests = BoardB;
            main.CloseThumbnailsWindowAsUserForTests();

            Assert.True(UserSettings.HasSchematicsThumbnailsWindowLayout);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Fullscreen owning the detached thumbnails window
    // -----------------------------------------------------------------------------------------

    // While fullscreen is open the thumbnails window is OWNED by it (so clicking into the
    // thumbnails raises the fullscreen window rather than Main - see ReownSchematicsThumbnailsWindow).
    // Closing an owner closes what it owns, so leaving fullscreen destroyed the thumbnails window,
    // and its Closed handler read that as the user dismissing it and silently turned detach off.
    // The compensating ReownSchematicsThumbnailsWindow() ran from the fullscreen Closed handler,
    // i.e. after the thumbnails field was already null, so it never prevented anything.
    //
    // Fails against that version: the setting comes back false and the window stays gone.
    [Fact]
    public void Leaving_fullscreen_keeps_detach_on_and_brings_the_thumbnails_window_back()
    {
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.ApplyThumbnailsDetachedState();
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            main.ToggleSchematicsFullscreenWindow();
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            // F11 again - fullscreen closes, taking the window it owns with it.
            main.ToggleSchematicsFullscreenWindow();

            Assert.True(UserSettings.DetachSchematicsThumbnails);
            Assert.True(main.IsThumbnailsWindowOpenForTests);
            Assert.True(DetachCheckBox(main).IsChecked);
        });
    }

    // The per-board tier must survive the same cascade, and the reopened window must come back on
    // the board that is actually selected.
    [Fact]
    public void Leaving_fullscreen_keeps_the_per_board_detach_state_on()
    {
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.SetThumbnailsDetached(true);
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            main.ToggleSchematicsFullscreenWindow();
            main.ToggleSchematicsFullscreenWindow();

            Assert.True(UserSettings.GetThumbnailWindowSettingsForBoard(BoardA)?.IsDetached);
            Assert.True(main.IsThumbnailsWindowOpenForTests);
        });
    }

    // The distinction the fix rests on: the cascade must be ignored, but a genuine close by the
    // user WHILE fullscreen is open must still turn the feature off, exactly as it does otherwise.
    [Fact]
    public void Closing_the_thumbnails_window_by_hand_during_fullscreen_still_turns_detach_off()
    {
        UserSettings.DetachSchematicsThumbnails = true;

        this.WithMainOnBoard(BoardA, main =>
        {
            main.ApplyThumbnailsDetachedState();
            main.ToggleSchematicsFullscreenWindow();
            Assert.True(main.IsThumbnailsWindowOpenForTests);

            main.CloseThumbnailsWindowAsUserForTests();

            Assert.False(UserSettings.DetachSchematicsThumbnails);
            Assert.False(main.IsThumbnailsWindowOpenForTests);

            // Leaving fullscreen afterwards must not resurrect it - the user turned it off.
            main.ToggleSchematicsFullscreenWindow();
            Assert.False(main.IsThumbnailsWindowOpenForTests);
        });
    }

    // Leaving fullscreen with detach OFF must not open a window at all - the reopen is a restore of
    // what was there, not an unconditional open.
    [Fact]
    public void Leaving_fullscreen_with_detach_off_opens_no_thumbnails_window()
    {
        this.WithMainOnBoard(BoardA, main =>
        {
            Assert.False(main.IsThumbnailsWindowOpenForTests);

            main.ToggleSchematicsFullscreenWindow();
            main.ToggleSchematicsFullscreenWindow();

            Assert.False(main.IsThumbnailsWindowOpenForTests);
        });
    }
}
