using Avalonia.Controls;
using Avalonia.Layout;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's "Open ... folder" buttons, which replaced a single
// "Open data/workbooks/log/settings folder" button. A fourth, for the oscilloscope captures, joined
// them on 2026-10-13 (owner request), between workbooks and logs.
//
// The single button was wrong in a way that is worth recording: it named four things and opened
// ONE folder - the AppData parent - so "Open data folder" landed the user a level above the data,
// and it ignored the data-root and workbooks-root command-line switches entirely, meaning a user
// who had relocated either one was shown a folder the app was not reading from.
//
// The CLICKS are deliberately not exercised, for the same reason ConfigurationHelpIconTests does
// not exercise its own: the handlers end in a Process.Start through the platform file manager, and
// rule 6 keeps that out of the suite. What is pinned here is everything a reader of the tab can
// see - that all four buttons exist, that they are in the asked-for order, that they are the same
// width, and that each still has its Click wired (a mis-typed handler name fails the XAML parse,
// but a Click attribute deleted outright leaves a button that silently does nothing).
[Collection("HeadlessUi")]
public sealed class ConfigurationFolderButtonsTests
{
    private const string DataButton = "OpenDataFolderButton";
    private const string WorkbooksButton = "OpenWorkbooksFolderButton";
    private const string OscilloscopeButton = "OpenOscilloscopeImagesFolderButton";
    private const string LogsButton = "OpenLogsFolderButton";

    private static readonly string[] AllButtons = [DataButton, WorkbooksButton, OscilloscopeButton, LogsButton];

    [Fact]
    public void All_four_folder_buttons_exist_with_their_own_labels()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            Assert.Equal("Open data folder", tab.GetControl<Button>(DataButton).Content);
            Assert.Equal("Open workbooks folder", tab.GetControl<Button>(WorkbooksButton).Content);
            Assert.Equal("Open oscilloscope image folder", tab.GetControl<Button>(OscilloscopeButton).Content);
            Assert.Equal("Open logs and settings folder", tab.GetControl<Button>(LogsButton).Content);
        });
    }

    // The replaced button is gone rather than merely hidden - leaving it in the markup would give
    // the tab another button opening the parent folder, which is the ambiguity being removed.
    [Fact]
    public void The_old_combined_button_is_gone()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            Assert.Null(tab.FindControl<Button>("OpenAppDataFolderButton"));
        });
    }

    // Asked for explicitly: data, then workbooks, then oscilloscope images ("after the workbooks
    // button", 2026-10-13), then logs, each on its own line. Asserted by position within the shared
    // parent rather than by reading the markup, so a re-ordering that looks harmless in a diff
    // still fails.
    [Fact]
    public void The_buttons_are_stacked_in_the_asked_for_order()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var data = tab.GetControl<Button>(DataButton);
            var workbooks = tab.GetControl<Button>(WorkbooksButton);
            var oscilloscope = tab.GetControl<Button>(OscilloscopeButton);
            var logs = tab.GetControl<Button>(LogsButton);

            var parent = Assert.IsType<StackPanel>(data.Parent);

            Assert.Same(parent, workbooks.Parent);
            Assert.Same(parent, oscilloscope.Parent);
            Assert.Same(parent, logs.Parent);

            int dataIndex = parent.Children.IndexOf(data);
            int workbooksIndex = parent.Children.IndexOf(workbooks);
            int oscilloscopeIndex = parent.Children.IndexOf(oscilloscope);
            int logsIndex = parent.Children.IndexOf(logs);

            Assert.True(dataIndex < workbooksIndex, "data must come before workbooks");
            Assert.Equal(workbooksIndex + 1, oscilloscopeIndex);
            Assert.True(oscilloscopeIndex < logsIndex, "oscilloscope images must come before logs");
        });
    }

    // Asked for explicitly. A Left-aligned Button sizes to its own text, so without an explicit
    // width these four would each be a different length and read as a ragged stack rather than as
    // one group of related actions.
    [Fact]
    public void The_buttons_are_all_the_same_width()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            double data = tab.GetControl<Button>(DataButton).Width;

            Assert.False(double.IsNaN(data), "the buttons need an explicit Width or they size to their own text");

            foreach (string name in AllButtons)
            {
                Assert.Equal(data, tab.GetControl<Button>(name).Width);
            }

            // The longest label has to fit, or the widest button clips the text it is sized for.
            // Measured with CRT's real font on 2026-10-13 (the suite's headless fonts cannot):
            // "Open oscilloscope image folder" needs 226 - a longer label tried that day needed 286
            // and was cut off at 260.
            Assert.True(data >= 230, $"width {data} is too narrow for \"Open oscilloscope image folder\"");
        });
    }

    // ###############################################################################################
    // The oscilloscope images button is there only while the oscilloscope is enabled (owner
    // request, 2026-10-13), and follows the check box as it is ticked - not only at launch.
    //
    // Ticking the box writes UserSettings, a shared static this "HeadlessUi" class cannot isolate by
    // collection, so the value is put back however the test ends.
    // ###############################################################################################
    [Fact]
    public void The_oscilloscope_images_button_shows_only_while_the_oscilloscope_is_enabled()
    {
        bool before = UserSettings.EnableNetworkConnectedOscilloscopeTab;

        try
        {
            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var enabled = tab.GetControl<CheckBox>("EnableNetworkConnectedOscilloscopeTabCheckBox");
                var button = tab.GetControl<Button>(OscilloscopeButton);

                enabled.IsChecked = false;
                Assert.False(button.IsVisible);

                enabled.IsChecked = true;
                Assert.True(button.IsVisible);

                enabled.IsChecked = false;
                Assert.False(button.IsVisible);
            });
        }
        finally
        {
            UserSettings.EnableNetworkConnectedOscilloscopeTab = before;
        }
    }

    // And at construction, from the saved setting - the tab is built at launch, before anyone ticks.
    [Fact]
    public void The_oscilloscope_images_button_starts_as_the_saved_setting_says()
    {
        bool before = UserSettings.EnableNetworkConnectedOscilloscopeTab;

        try
        {
            foreach (bool saved in new[] { true, false })
            {
                UserSettings.EnableNetworkConnectedOscilloscopeTab = saved;

                UiTest.Run(() =>
                {
                    var tab = new TabConfiguration();

                    Assert.Equal(saved, tab.GetControl<Button>(OscilloscopeButton).IsVisible);
                });
            }
        }
        finally
        {
            UserSettings.EnableNetworkConnectedOscilloscopeTab = before;
        }
    }

    // ###############################################################################################
    // The left column's narrowest must still hold the buttons, plus the 16 of its vertical scroll
    // bar (AllowAutoHide off) - nothing on the tab scrolls sideways, so a button wider than that is
    // cut off at the window's minimum width. Widening the buttons without the column is the mistake
    // this catches.
    // ###############################################################################################
    [Fact]
    public void The_left_column_never_gets_narrower_than_the_buttons()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            double buttons = tab.GetControl<Button>(DataButton).Width;
            double leftColumnMin = tab.GetControl<Grid>("RootSplitGrid").ColumnDefinitions[0].MinWidth;

            Assert.True(
                leftColumnMin >= buttons + 16,
                $"the left column's MinWidth {leftColumnMin} cannot hold {buttons}-wide buttons and the scroll bar");
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

            foreach (string name in AllButtons)
            {
                Assert.Equal(HorizontalAlignment.Left, tab.GetControl<Button>(name).HorizontalAlignment);
            }
        });
    }

    // Each button says what it opens. These are folders whose names alone do not say what is
    // in them - "data" in particular means nothing to someone who has never opened it.
    [Fact]
    public void Each_button_explains_what_is_in_the_folder()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            foreach (string name in AllButtons)
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
