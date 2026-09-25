using Avalonia.Controls;
using Avalonia.Layout;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's three "Open ... folder" buttons, which replaced a single
// "Open data/workbooks/log/settings folder" button.
//
// The single button was wrong in a way that is worth recording: it named four things and opened
// ONE folder - the AppData parent - so "Open data folder" landed the user a level above the data,
// and it ignored the data-root and workbooks-root command-line switches entirely, meaning a user
// who had relocated either one was shown a folder the app was not reading from.
//
// The CLICKS are deliberately not exercised, for the same reason ConfigurationHelpIconTests does
// not exercise its own: the handlers end in a Process.Start through the platform file manager, and
// rule 6 keeps that out of the suite. What is pinned here is everything a reader of the tab can
// see - that all three buttons exist, that they are in the asked-for order, that they are the same
// width, and that each still has its Click wired (a mis-typed handler name fails the XAML parse,
// but a Click attribute deleted outright leaves a button that silently does nothing).
[Collection("HeadlessUi")]
public sealed class ConfigurationFolderButtonsTests
{
    private const string DataButton = "OpenDataFolderButton";
    private const string WorkbooksButton = "OpenWorkbooksFolderButton";
    private const string LogsButton = "OpenLogsFolderButton";

    [Fact]
    public void All_three_folder_buttons_exist_with_their_own_labels()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            Assert.Equal("Open data folder", tab.GetControl<Button>(DataButton).Content);
            Assert.Equal("Open workbooks folder", tab.GetControl<Button>(WorkbooksButton).Content);
            Assert.Equal("Open logs and settings folder", tab.GetControl<Button>(LogsButton).Content);
        });
    }

    // The replaced button is gone rather than merely hidden - leaving it in the markup would give
    // the tab a fourth button opening the parent folder, which is the ambiguity being removed.
    [Fact]
    public void The_old_combined_button_is_gone()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            Assert.Null(tab.FindControl<Button>("OpenAppDataFolderButton"));
        });
    }

    // Asked for explicitly: data, then workbooks, then logs, each on its own line. Asserted by
    // position within the shared parent rather than by reading the markup, so a re-ordering that
    // looks harmless in a diff still fails.
    [Fact]
    public void The_buttons_are_stacked_in_the_asked_for_order()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var data = tab.GetControl<Button>(DataButton);
            var workbooks = tab.GetControl<Button>(WorkbooksButton);
            var logs = tab.GetControl<Button>(LogsButton);

            var parent = Assert.IsType<StackPanel>(data.Parent);

            Assert.Same(parent, workbooks.Parent);
            Assert.Same(parent, logs.Parent);

            int dataIndex = parent.Children.IndexOf(data);
            int workbooksIndex = parent.Children.IndexOf(workbooks);
            int logsIndex = parent.Children.IndexOf(logs);

            Assert.True(dataIndex < workbooksIndex, "data must come before workbooks");
            Assert.True(workbooksIndex < logsIndex, "workbooks must come before logs");
        });
    }

    // Asked for explicitly. A Left-aligned Button sizes to its own text, so without an explicit
    // width these three would each be a different length and read as a ragged stack rather than as
    // one group of related actions.
    [Fact]
    public void The_buttons_are_all_the_same_width()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            double data = tab.GetControl<Button>(DataButton).Width;
            double workbooks = tab.GetControl<Button>(WorkbooksButton).Width;
            double logs = tab.GetControl<Button>(LogsButton).Width;

            Assert.False(double.IsNaN(data), "the buttons need an explicit Width or they size to their own text");
            Assert.Equal(data, workbooks);
            Assert.Equal(data, logs);

            // The longest label has to fit, or the widest button clips the text it is sized for.
            Assert.True(data >= 200, $"width {data} is too narrow for \"Open logs and settings folder\"");
        });
    }

    // Left-aligned so they line up with the controls above them rather than stretching the full
    // width of the tab, which is what a Button in a StackPanel does by default.
    [Fact]
    public void The_buttons_are_left_aligned_like_the_rest_of_the_tab()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            foreach (string name in new[] { DataButton, WorkbooksButton, LogsButton })
            {
                Assert.Equal(HorizontalAlignment.Left, tab.GetControl<Button>(name).HorizontalAlignment);
            }
        });
    }

    // Each button says what it opens. These are three folders whose names alone do not say what is
    // in them - "data" in particular means nothing to someone who has never opened it.
    [Fact]
    public void Each_button_explains_what_is_in_the_folder()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            foreach (string name in new[] { DataButton, WorkbooksButton, LogsButton })
            {
                var tip = ToolTip.GetTip(tab.GetControl<Button>(name)) as string;

                Assert.False(string.IsNullOrWhiteSpace(tip), $"{name} has no tooltip");
            }
        });
    }

    // ###############################################################################################
    // The workbooks button opens the root the app is REALLY using, and reports rather than guesses
    // when there isn't one.
    //
    // The first version of this handler fell back to rebuilding the AppData default whenever
    // WorklogManager.WorkbookRoot was empty. That is the one case where the default is certainly
    // WRONG: the root is empty because loading it failed, and the likeliest reason is a
    // "--workbooks-root=" pointing at something unreachable. The fallback then called
    // Directory.CreateDirectory on the default and opened it, so a user whose external drive was not
    // connected was shown a newly-created empty folder presented as their workbooks - the "my
    // workbooks are gone" report, caused by the button meant to help diagnose it.
    //
    // These two tests are about the RESOLUTION, which is the half that was wrong. The opening itself
    // still ends in Process.Start and stays out of the suite under rule 6.
    // ###############################################################################################
    [Fact]
    public void The_workbooks_button_targets_the_root_the_app_is_actually_using()
    {
        using var workspace = new TempWorkspace();
        string root = workspace.Path_("Relocated-Workbooks");

        WorklogManager.LoadFrom(root);

        // What the handler reads. Non-empty means it opens this, not a rebuilt default.
        Assert.False(string.IsNullOrEmpty(WorklogManager.WorkbookRoot));
        Assert.Equal(root, WorklogManager.WorkbookRoot);
    }

    // And the case the fix is for: an unresolved root must stay empty, so the handler takes its
    // "report it" branch. A root that silently reported the AppData default here would put the
    // handler back on the path that creates and opens the wrong folder.
    [Fact]
    public void An_unresolved_workbook_root_stays_empty_rather_than_reporting_a_default()
    {
        // LoadFrom fails to a deliberately unusable path, which is what an unreachable
        // "--workbooks-root=" produces in practice.
        WorklogManager.LoadFrom(InvalidRootPath);

        Assert.True(
            string.IsNullOrEmpty(WorklogManager.WorkbookRoot),
            "an unresolved root must be empty so the button reports it instead of opening the AppData default");
    }

    // A path that cannot be created on any platform this app ships for: an empty path throws on
    // Directory.CreateDirectory, which is the failure LoadFrom catches by clearing the root.
    private const string InvalidRootPath = "";
}
