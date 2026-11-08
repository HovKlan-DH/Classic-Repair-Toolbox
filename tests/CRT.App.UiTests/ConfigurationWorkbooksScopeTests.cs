using Avalonia.Controls;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's "which workbooks does the Workbooks tab list" radio group, and the
// change notification it drives.
//
// The thing worth pinning here is that ONE click writes the setting ONCE. A radio group raises
// IsCheckedChanged twice per click - once for the button being unchecked, once for the one being
// checked - and both handlers see the same post-transition state, so a handler that reads the
// GROUP rather than its own sender fires the same write twice. That was harmless only while
// UserSettings.WorkbooksScope's unchanged-value guard absorbed the second one; now that the setter
// raises WorkbooksScopeChanged and Main rebuilds the entire Workbooks tab off it, a doubled write
// is a doubled full disk rescan and schematic re-decode on every click.
//
// COLLECTION NOTE: "HeadlessUi" rather than "UserSettings", because these construct a control and
// so need the shared dispatcher thread - a class can only join one collection. They nonetheless
// drive UserSettings' static state, which is safe because xunit.runner.json turns collection
// parallelism off; see WorkbooksListTests' own note. Every test restores WorkbooksScope in a
// finally block.
[Collection("HeadlessUi")]
public sealed class ConfigurationWorkbooksScopeTests
{
    // Counts WorkbooksScopeChanged while body runs, and unsubscribes however it ends - a leaked
    // handler would keep counting into every later test in this collection.
    private static int CountScopeChanges(Action body)
    {
        int changes = 0;
        Action handler = () => changes++;

        UserSettings.WorkbooksScopeChanged += handler;
        try
        {
            body();
        }
        finally
        {
            UserSettings.WorkbooksScopeChanged -= handler;
        }

        return changes;
    }

    // The setting follows the button the user actually picked, and is written exactly once.
    [Fact]
    public void Choosing_all_boards_writes_the_setting_once()
    {
        string saved = UserSettings.WorkbooksScope;
        try
        {
            UserSettings.WorkbooksScope = "CurrentBoard";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();

                int changes = CountScopeChanges(() =>
                    tab.GetControl<RadioButton>("WorkbooksScopeAllBoardsRadioButton").IsChecked = true);

                Assert.Equal("AllBoards", UserSettings.WorkbooksScope);

                // ONE, not two. The uncheck of the sibling raises the same handler and must not
                // write - see this file's header for why that now costs a full tab rebuild.
                Assert.Equal(1, changes);
            });
        }
        finally
        {
            UserSettings.WorkbooksScope = saved;
        }
    }

    // The other direction, so neither button is correct only by being the one the default happens
    // to start on.
    [Fact]
    public void Choosing_current_board_writes_the_setting_once()
    {
        string saved = UserSettings.WorkbooksScope;
        try
        {
            UserSettings.WorkbooksScope = "AllBoards";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();

                int changes = CountScopeChanges(() =>
                    tab.GetControl<RadioButton>("WorkbooksScopeCurrentBoardRadioButton").IsChecked = true);

                Assert.Equal("CurrentBoard", UserSettings.WorkbooksScope);
                Assert.Equal(1, changes);
            });
        }
        finally
        {
            UserSettings.WorkbooksScope = saved;
        }
    }

    // ###########################################################################################
    // The third choice, in the middle (owner request, 2026-11-06: "add a new option (in the middle) for 'Show only
    // devices for selected hardware'") - written once too, and the group worded by the devices in the owner's order.
    // ###########################################################################################
    [Fact]
    public void Choosing_the_selected_hardware_writes_the_setting_once()
    {
        string saved = UserSettings.WorkbooksScope;
        try
        {
            UserSettings.WorkbooksScope = "CurrentBoard";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var hardware = tab.GetControl<RadioButton>("WorkbooksScopeCurrentHardwareRadioButton");

                Assert.Equal(
                    ["Show all devices", "Show only devices for selected hardware", "Show only devices for selected board"],
                    Assert.IsAssignableFrom<Panel>(hardware.Parent).Children.OfType<RadioButton>().Select(button => button.Content as string));

                int changes = CountScopeChanges(() => hardware.IsChecked = true);

                Assert.Equal("CurrentHardware", UserSettings.WorkbooksScope);
                Assert.Equal(1, changes);
            });

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();

                Assert.True(tab.GetControl<RadioButton>("WorkbooksScopeCurrentHardwareRadioButton").IsChecked);
                Assert.False(tab.GetControl<RadioButton>("WorkbooksScopeAllBoardsRadioButton").IsChecked);
                Assert.False(tab.GetControl<RadioButton>("WorkbooksScopeCurrentBoardRadioButton").IsChecked);
            });
        }
        finally
        {
            UserSettings.WorkbooksScope = saved;
        }
    }

    // Re-picking the already-selected scope changes nothing and notifies nobody: subscribers rebuild
    // the whole Workbooks tab, so a no-op click must not cost one.
    [Fact]
    public void Re_choosing_the_current_scope_notifies_nobody()
    {
        string saved = UserSettings.WorkbooksScope;
        try
        {
            UserSettings.WorkbooksScope = "AllBoards";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var allBoards = tab.GetControl<RadioButton>("WorkbooksScopeAllBoardsRadioButton");

                // The tab's constructor already checked it from the setting; re-asserting it is the
                // "clicked the one that was already on" case.
                int changes = CountScopeChanges(() => allBoards.IsChecked = true);

                Assert.Equal("AllBoards", UserSettings.WorkbooksScope);
                Assert.Equal(0, changes);
            });
        }
        finally
        {
            UserSettings.WorkbooksScope = saved;
        }
    }

    // The tab shows the persisted scope when it is built, rather than whatever the markup declares -
    // otherwise reopening Configuration reports a setting the app is not using.
    [Fact]
    public void The_radio_group_reflects_the_saved_scope_on_construction()
    {
        string saved = UserSettings.WorkbooksScope;
        try
        {
            UserSettings.WorkbooksScope = "AllBoards";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();

                Assert.True(tab.GetControl<RadioButton>("WorkbooksScopeAllBoardsRadioButton").IsChecked);
                Assert.False(tab.GetControl<RadioButton>("WorkbooksScopeCurrentBoardRadioButton").IsChecked);
            });

            UserSettings.WorkbooksScope = "CurrentBoard";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();

                Assert.False(tab.GetControl<RadioButton>("WorkbooksScopeAllBoardsRadioButton").IsChecked);
                Assert.True(tab.GetControl<RadioButton>("WorkbooksScopeCurrentBoardRadioButton").IsChecked);
            });
        }
        finally
        {
            UserSettings.WorkbooksScope = saved;
        }
    }

    // ###########################################################################################
    // The radio buttons' labels WRAP (2026-10-13). The theme's check box wraps its text and its
    // radio button does not, so "Show only workbooks for selected board" was cut off at the
    // window's minimum width once the tab's scroll bar showed. The tab's own style sets the wrap on
    // the template's presenter. Read off the style: the wrap itself needs real fonts, and showing
    // the tab starts its catalogue build against whatever data another test left behind.
    // ###########################################################################################
    // ###########################################################################################
    // "SORTING OF DEVICES" (owner request, 2026-11-07), below the scope: "Sort automatically by name" - the default -,
    // "... by hardware/board/name" and "Sort manually in list (drag up/down)". The button picked writes the setting
    // once, as the scope's do, and raises DeviceSortingChanged, which Main draws the lists again by.
    // ###########################################################################################
    [Fact]
    public void Sorting_of_devices_is_by_name_unless_picked_otherwise_and_one_click_writes_it_once()
    {
        string saved = UserSettings.DeviceSorting;

        try
        {
            UserSettings.DeviceSorting = OwnedDeviceSorting.ByName;

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();

                Assert.True(tab.GetControl<RadioButton>("DeviceSortingByNameRadioButton").IsChecked);
                Assert.Equal("Sort automatically by name", tab.GetControl<RadioButton>("DeviceSortingByNameRadioButton").Content);
                Assert.Equal("Sort automatically by hardware/board/name", tab.GetControl<RadioButton>("DeviceSortingByHardwareBoardNameRadioButton").Content);
                Assert.Equal("Sort manually in list (drag up/down)", tab.GetControl<RadioButton>("DeviceSortingManualRadioButton").Content);

                int changes = 0;
                Action counted = () => changes++;
                UserSettings.DeviceSortingChanged += counted;

                try
                {
                    tab.GetControl<RadioButton>("DeviceSortingManualRadioButton").IsChecked = true;
                    Assert.Equal(OwnedDeviceSorting.Manual, UserSettings.DeviceSorting);
                    Assert.Equal(1, changes);

                    tab.GetControl<RadioButton>("DeviceSortingByHardwareBoardNameRadioButton").IsChecked = true;
                    Assert.Equal(OwnedDeviceSorting.ByHardwareBoardName, UserSettings.DeviceSorting);
                    Assert.Equal(2, changes);
                }
                finally
                {
                    UserSettings.DeviceSortingChanged -= counted;
                }
            });
        }
        finally
        {
            UserSettings.DeviceSorting = saved;
        }
    }

    [Fact]
    public void The_scope_labels_are_styled_to_wrap()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var style = Assert.Single(
                tab.Styles.OfType<Avalonia.Styling.Style>(),
                style => style.Selector?.ToString() == "RadioButton /template/ ContentPresenter");

            var setter = Assert.Single(style.Setters.OfType<Avalonia.Styling.Setter>());

            Assert.Equal(Avalonia.Controls.Presenters.ContentPresenter.TextWrappingProperty, setter.Property);
            Assert.Equal(Avalonia.Media.TextWrapping.Wrap, setter.Value);
        });
    }
}
