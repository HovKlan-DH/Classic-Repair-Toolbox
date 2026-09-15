using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The "Detach thumbnails to its own window" checkbox on the Configuration tab: that it sits
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
}
