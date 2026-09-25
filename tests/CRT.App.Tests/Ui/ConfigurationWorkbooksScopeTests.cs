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
}
