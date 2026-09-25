using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// TabWorkbooks.FocusSearchBox - puts cursor focus in "Find a previous repair" the moment the
// Workbooks tab becomes selected, so a user landing on it can start typing straight away.
//
// Main is what actually calls this (from OnMainTabControlSelectionChanged, wired to
// MainTabControl.SelectionChanged), and Main is never constructed by any test - see CLAUDE.md.
// So this covers what IS testable in isolation: that calling FocusSearchBox moves real focus onto
// the real FindRepairTextBox control, which is the contract Main's handler relies on.
//
// Needs a real shown Window, unlike most of this tab's other tests: FocusSearchBox's whole job is
// actual keyboard focus, and a control never attached to a visual tree cannot receive it.
[Collection("HeadlessUi")]
public sealed class WorkbooksSearchFocusTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private readonly string thisBoardKey = "Commodore 64|250469 " + Guid.NewGuid().ToString("N");

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    private void LoadWorklog()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
    }

    // A sibling button OUTSIDE the tab, so a test can move focus away from the search box without
    // needing to know the internal layout of the tab's own control tree.
    private static void WithShownTab(string boardKey, Action<TabWorkbooks, TextBox, Button> body)
    {
        UiTest.Run(() =>
        {
            var tab = new TabWorkbooks { BoardKeyOverrideForTests = boardKey };
            tab.RefreshWorkbooks();

            var elsewhere = new Button { Content = "Elsewhere" };
            var root = new DockPanel();
            DockPanel.SetDock(elsewhere, Dock.Top);
            root.Children.Add(elsewhere);
            root.Children.Add(tab);

            var window = new Window { Width = 900, Height = 600, Content = root };

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                var searchBox = tab.GetControl<TextBox>("FindRepairTextBox");
                body(tab, searchBox, elsewhere);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The direct contract: calling it puts focus on the box. Posted at Background priority (see
    // FocusSearchBox's own comment), so the dispatcher has to actually run the posted job before
    // this can observe it.
    [Fact]
    public void FocusSearchBox_moves_focus_onto_the_search_box()
    {
        this.LoadWorklog();

        WithShownTab(this.thisBoardKey, (tab, searchBox, _) =>
        {
            Assert.False(searchBox.IsFocused);

            tab.FocusSearchBox();
            Dispatcher.UIThread.RunJobs();

            Assert.True(searchBox.IsFocused);
        });
    }

    // Calling it again (a second tab visit) must not throw or misbehave - the box already has
    // focus, and re-focusing an already-focused control is a normal no-op in Avalonia.
    [Fact]
    public void Calling_FocusSearchBox_twice_leaves_the_box_focused()
    {
        this.LoadWorklog();

        WithShownTab(this.thisBoardKey, (tab, searchBox, _) =>
        {
            tab.FocusSearchBox();
            Dispatcher.UIThread.RunJobs();

            tab.FocusSearchBox();
            Dispatcher.UIThread.RunJobs();

            Assert.True(searchBox.IsFocused);
        });
    }

    // Focus moves elsewhere afterwards exactly as normal user interaction would - FocusSearchBox
    // is a one-shot "focus on tab entry", not something that keeps re-stealing focus back. This is
    // the other half of the requested behaviour: "when user clicks on something, then again the
    // Filter components box should take over again, as normal" - which this test shows nothing on
    // the Workbooks tab's own side is fighting.
    [Fact]
    public void Focus_can_move_away_from_the_search_box_afterwards()
    {
        this.LoadWorklog();

        WithShownTab(this.thisBoardKey, (tab, searchBox, elsewhere) =>
        {
            tab.FocusSearchBox();
            Dispatcher.UIThread.RunJobs();
            Assert.True(searchBox.IsFocused);

            elsewhere.Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.False(searchBox.IsFocused);
        });
    }
}
