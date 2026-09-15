using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The "Detach thumbnails into their own window" checkbox on the Configuration tab: that it sits
// directly under "Open multiple component info windows" as asked for, that it loads its checked
// state from UserSettings and persists a toggle back, and that RefreshDetachSchematicsThumbnailsCheckBoxFromSettings
// (used when the detached window is closed via its own OS close button rather than the checkbox)
// updates the checkbox WITHOUT re-triggering the handler that would otherwise call back into Main
// and redundantly reapply the state a second time.
//
// COLLECTION NOTE: "HeadlessUi" because TabConfiguration's constructor reads UserSettings for
// every checkbox's initial state - see UserSettingsTests's own header for why every test here
// points UserSettings at a temp file first.
[Collection("HeadlessUi")]
public sealed class TabConfigurationDetachThumbnailsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public TabConfigurationDetachThumbnailsTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
        this.thisWorkspace.Dispose();
    }

    [Fact]
    public void The_detach_checkbox_sits_directly_after_the_multiple_popups_checkbox()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var popupsRow = tab.GetControl<CheckBox>("MultipleInstancesForComponentPopupCheckBox");
            var detachRow = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");

            var parent = popupsRow.GetVisualParent();
            Assert.NotNull(parent);
            Assert.Same(parent, detachRow.GetVisualParent());

            var siblings = ((StackPanel)parent!).Children;
            int popupsIndex = siblings.IndexOf(popupsRow);
            int detachIndex = siblings.IndexOf(detachRow);

            Assert.Equal(popupsIndex + 1, detachIndex);
        });
    }

    [Fact]
    public void Constructing_the_tab_loads_the_checkbox_from_settings()
    {
        UserSettings.DetachSchematicsThumbnails = true;

        try
        {
            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var checkBox = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");

                Assert.True(checkBox.IsChecked);
            });
        }
        finally
        {
            // Leaving this true and relying on Dispose()'s own LoadFrom to reset it left a window
            // where a test from ANOTHER class in this same HeadlessUi collection could observe it
            // still true - reported as SchematicsThumbnailsContextMenuTests failing only when run
            // alongside this class, never in isolation.
            UserSettings.DetachSchematicsThumbnails = false;
        }
    }

    [Fact]
    public void Toggling_the_checkbox_persists_the_setting()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();
            var checkBox = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");

            checkBox.IsChecked = true;

            Assert.True(UserSettings.DetachSchematicsThumbnails);

            checkBox.IsChecked = false;

            Assert.False(UserSettings.DetachSchematicsThumbnails);
        });
    }

    // The resync seam Main calls when the detached window is closed via its own close button: it
    // must update the checkbox from UserSettings without re-entering the handler and writing the
    // value straight back out again (which would be harmless here, but in the running app calls
    // back into Main.ApplyThumbnailsDetachedState a second time for nothing).
    [Fact]
    public void RefreshFromSettings_updates_the_checkbox_without_writing_the_setting_back()
    {
        UiTest.Run(() =>
        {
            // Start checked, as if the tab was constructed while the setting was already on.
            UserSettings.DetachSchematicsThumbnails = true;
            var tab = new TabConfiguration();
            var checkBox = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");
            Assert.True(checkBox.IsChecked);

            // Simulate the detached window closing itself (its own OS close button): UserSettings
            // changes WITHOUT going through the checkbox, so the box is now stale relative to it.
            UserSettings.DetachSchematicsThumbnails = false;
            Assert.True(checkBox.IsChecked); // still stale - nothing has told the box yet

            tab.RefreshDetachSchematicsThumbnailsCheckBoxFromSettings();

            Assert.False(checkBox.IsChecked);
            Assert.False(UserSettings.DetachSchematicsThumbnails);
        });
    }

    // ---------------------------------------------------------------------------------------
    // "Remember thumbnail window settings per board"
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void The_remember_per_board_checkbox_is_disabled_by_default_since_detach_starts_off()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var detachCheckBox = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");
            var rememberCheckBox = tab.GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");

            Assert.False(detachCheckBox.IsChecked);
            Assert.False(rememberCheckBox.IsEnabled);
        });
    }

    // Constructing the tab while "Detach thumbnails" is already on (a prior session's setting)
    // must enable the dependent checkbox immediately, not only after the user re-toggles the
    // parent - otherwise it would look permanently greyed out for anyone who already has detach on.
    [Fact]
    public void The_remember_per_board_checkbox_is_enabled_on_construction_when_detach_is_already_on()
    {
        UserSettings.DetachSchematicsThumbnails = true;

        try
        {
            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var rememberCheckBox = tab.GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");

                Assert.True(rememberCheckBox.IsEnabled);
            });
        }
        finally
        {
            UserSettings.DetachSchematicsThumbnails = false;
        }
    }

    [Fact]
    public void Toggling_detach_enables_and_disables_the_remember_per_board_checkbox()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();
            var detachCheckBox = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");
            var rememberCheckBox = tab.GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");

            Assert.False(rememberCheckBox.IsEnabled);

            detachCheckBox.IsChecked = true;
            Assert.True(rememberCheckBox.IsEnabled);

            detachCheckBox.IsChecked = false;
            Assert.False(rememberCheckBox.IsEnabled);
        });
    }

    [Fact]
    public void Toggling_the_remember_per_board_checkbox_persists_the_setting()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();
            var detachCheckBox = tab.GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");
            var rememberCheckBox = tab.GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");

            // Only reachable while enabled, as it would be in the running app.
            detachCheckBox.IsChecked = true;

            rememberCheckBox.IsChecked = true;
            Assert.True(UserSettings.RememberThumbnailWindowSettingsPerBoard);

            rememberCheckBox.IsChecked = false;
            Assert.False(UserSettings.RememberThumbnailWindowSettingsPerBoard);
        });
    }

    [Fact]
    public void Constructing_the_tab_loads_the_remember_per_board_checkbox_from_settings()
    {
        UserSettings.DetachSchematicsThumbnails = true;
        UserSettings.RememberThumbnailWindowSettingsPerBoard = true;

        try
        {
            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var rememberCheckBox = tab.GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");

                Assert.True(rememberCheckBox.IsChecked);
            });
        }
        finally
        {
            UserSettings.RememberThumbnailWindowSettingsPerBoard = false;
            UserSettings.DetachSchematicsThumbnails = false;
        }
    }

    // RefreshDetachSchematicsThumbnailsCheckBoxFromSettings is also the seam that fires when the
    // detached window closes itself - it must re-evaluate the dependent checkbox's enabled state
    // too, or turning detach off that way would leave "remember per board" visibly enabled while
    // meaning nothing.
    [Fact]
    public void RefreshFromSettings_also_updates_the_remember_per_board_checkbox_enabled_state()
    {
        UiTest.Run(() =>
        {
            UserSettings.DetachSchematicsThumbnails = true;
            var tab = new TabConfiguration();
            var rememberCheckBox = tab.GetControl<CheckBox>("RememberThumbnailWindowSettingsPerBoardCheckBox");
            Assert.True(rememberCheckBox.IsEnabled);

            UserSettings.DetachSchematicsThumbnails = false;
            tab.RefreshDetachSchematicsThumbnailsCheckBoxFromSettings();

            Assert.False(rememberCheckBox.IsEnabled);
        });
    }
}
