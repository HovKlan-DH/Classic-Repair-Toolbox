using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using System.IO;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// SchematicsThumbnailsWindow.Initialize's optional boardKey parameter - the "Remember thumbnail
// window settings per board" half of the detach feature (SchematicsThumbnailsDetachTests covers
// the rest: mode toggling, F11, the gallery wiring). With boardKey null (every call site when the
// setting is off, and every pre-existing test's direct construction) the window must behave exactly
// as before, reading and later saving the single global SchematicsThumbnailsWindow* tier. With a
// boardKey given, it must read and save UserSettings' per-board ThumbnailWindowSettingsByBoard
// tier instead, and never touch the global tier at all.
//
// COLLECTION NOTE: "HeadlessUi" - same reason as TabConfigurationDetachThumbnailsTests's own
// header: constructing a Window and a UserControl needs Avalonia's dispatcher, and every test here
// points UserSettings at a temp file first so nothing touches the developer's real settings.
[Collection("HeadlessUi")]
public sealed class SchematicsThumbnailsWindowPerBoardTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SchematicsThumbnailsWindowPerBoardTests()
    {
        this.LoadEmptySettings();
    }

    public void Dispose()
    {
        this.LoadEmptySettings();
        this.thisWorkspace.Dispose();
    }

    // Writes an actual (empty) settings FILE before loading it, rather than pointing UserSettings
    // at a path that does not exist - LoadFrom only resets its in-memory _data when it can
    // deserialize something, so a missing file leaves whatever the PREVIOUS test in this class left
    // behind untouched. UserSettingsTests' own LoadSettings helper does the same for the same
    // reason. Fails intermittently without this: HasSchematicsThumbnailsWindowLayout stays true
    // from an earlier test in this class once one has called SaveSchematicsThumbnailsWindowLayout.
    private void LoadEmptySettings()
    {
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(path, "{}");
        UserSettings.LoadFrom(path);
    }

    [Fact]
    public void With_no_boardKey_a_saved_global_layout_is_still_applied()
    {
        UiTest.Run(() =>
        {
            UserSettings.SaveSchematicsThumbnailsWindowLayout("Maximized", 500, 700, 60, 40);

            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab);

            Assert.Equal(WindowState.Maximized, window.WindowState);
        });
    }

    [Fact]
    public void With_a_boardKey_the_per_board_layout_is_applied_instead_of_the_global_one()
    {
        UiTest.Run(() =>
        {
            // A global layout that must be IGNORED once a boardKey is given.
            UserSettings.SaveSchematicsThumbnailsWindowLayout("Maximized", 999, 999, 999, 999);

            UserSettings.SetThumbnailWindowDetachedForBoard("C64/250407", true);
            UserSettings.SaveThumbnailWindowLayoutForBoard("C64/250407", "Normal", 350, 450, 15, 25);

            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab, "C64/250407");

            Assert.Equal(WindowState.Normal, window.WindowState);
            Assert.Equal(350, window.Width);
            Assert.Equal(450, window.Height);
        });
    }

    // A board that has never had a layout saved must start at the plain 420x600 default, not at
    // whatever the global tier happens to hold - the two tiers are meant to be fully independent.
    [Fact]
    public void A_board_with_no_saved_layout_starts_at_the_default_size_ignoring_the_global_tier()
    {
        UiTest.Run(() =>
        {
            UserSettings.SaveSchematicsThumbnailsWindowLayout("Maximized", 999, 999, 999, 999);

            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab, "Plus4/310163");

            Assert.Equal(WindowState.Normal, window.WindowState);
            Assert.Equal(420.0, window.Width);
            Assert.Equal(600.0, window.Height);
        });
    }

    // Closing must save into the PER-BOARD tier when a boardKey was given, and must leave the
    // global tier completely untouched - the two must never cross-contaminate.
    [Fact]
    public void Closing_with_a_boardKey_saves_to_the_per_board_tier_and_not_the_global_one()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab, "C64/250407");
            window.Show();

            window.Close();

            var perBoard = UserSettings.GetThumbnailWindowSettingsForBoard("C64/250407");
            Assert.NotNull(perBoard);
            Assert.True(perBoard!.HasWindowLayout);

            Assert.False(UserSettings.HasSchematicsThumbnailsWindowLayout);
        });
    }

    // The mirror of the above: closing with no boardKey must save to the global tier and must not
    // create a per-board entry for whatever board happens to be selected elsewhere.
    [Fact]
    public void Closing_with_no_boardKey_saves_to_the_global_tier_and_not_any_per_board_entry()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab);
            window.Show();

            window.Close();

            Assert.True(UserSettings.HasSchematicsThumbnailsWindowLayout);
            Assert.Null(UserSettings.GetThumbnailWindowSettingsForBoard("C64/250407"));
        });
    }
}
