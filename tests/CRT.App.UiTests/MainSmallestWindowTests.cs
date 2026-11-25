using Avalonia;
using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// CRT'S WINDOW AT ITS SMALLEST: EVERYTHING ON EVERY TAB CAN STILL BE REACHED (owner report, 2026-11-25: "when not logged
// in and I have forgot my password, and I then need to type in the "Reset code", then on my Linux computer I cannot see
// this input field, as it is outside viewport and there is no possibility for me to scroll down ... we need to
// investigate is happening elsewhere also, as it makes the UI useless in these situation").
//
// The window's own minimum (800 x 500) is the least room any tab ever gets, and a Linux desktop's larger font only
// makes it less. Measured that day on the real C64 250407 with real fonts: the sign-in needed 590 high, 910 after "I
// forgot my password"; "Feedback" 623, its text box squeezed to nothing and "Send feedback" below the window;
// "Oscilloscope" 588; the component list on the left 9 px high; and "Contribute" at 800 wide had its top line in a
// column so narrow it pushed the whole workspace out of the window. Each was a stack of controls with nothing to
// scroll it - so a tab's content that can be taller than the window SCROLLS, and these tests read every tab at the
// smallest window through ControlsInReach.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainSmallestWindowTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MainSmallestWindowTests()
    {
        // Every tab that can be switched on, switched on; the settings and workbooks in a folder of the test's own.
        string settings = this.thisWorkspace.Path_("settings-" + Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(settings, "{\"enableWorklog\":true,\"enableMaintainerTab\":true,\"enableNetworkConnectedOscilloscopeTab\":true}");
        UserSettings.LoadFrom(settings);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N")));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_("after-" + Guid.NewGuid().ToString("N") + ".json"));
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-after-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    // CRT's window as low as it goes. `widthTimes`: how many times its narrowest width - a whole tab is read twice as wide
    // (ControlsInReach.OutOfReach: the headless font is twice as wide as a real one, so a window twice as wide wraps a
    // sentence over as many lines as CRT's narrowest does on a desktop - and how many lines a sentence takes is how high
    // it is).
    private static CRT.Main Smallest(int widthTimes = 1)
    {
        var window = new CRT.Main();
        window.Width = window.MinWidth * widthTimes;
        window.Height = window.MinHeight;
        window.ApplyWorklogBarVisibility();
        window.ApplyOscilloscopeTabVisibility();
        window.ApplyMaintainerTabVisibility();
        window.Show();
        ControlsInReach.Settle(window);
        return window;
    }

    private static void Close(CRT.Main window)
    {
        window.IsApplicationShuttingDownForTests = true;
        window.Close();
    }

    [Fact]
    public void At_the_smallest_window_every_tab_keeps_its_controls_in_reach()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = MainSmallestWindowTests.Smallest(widthTimes: 2);

            try
            {
                var problems = new List<string>();
                List<TabItem> tabs = window.MainTabControl.Items.OfType<TabItem>().Where(tab => tab.IsVisible).ToList();

                // Not vacuous: the optional tabs are there too.
                Assert.Contains(window.OscilloscopeTabItem, tabs);
                Assert.Contains(window.MaintainerTabItem, tabs);

                // How HIGH each tab's content is, not how wide: the headless font is twice as wide as a real one
                // (ControlsInReach.OutOfReach), so the window is twice CRT's narrowest - the narrow tab's top line has a
                // test of its own (TabDraftsWorkspaceTests).
                foreach (TabItem tab in tabs)
                {
                    window.MainTabControl.SelectedItem = tab;
                    ControlsInReach.Settle(window);
                    problems.AddRange(ControlsInReach.OutOfReach(window, sideways: false).Select(problem => tab.Name + ": " + problem));
                }

                Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
            }
            finally
            {
                MainSmallestWindowTests.Close(window);
            }
        });
    }

    // ###########################################################################################
    // THE REPORTED CASE: "I forgot my password" put the reset code, the new password and "Set password" under the "Maintainer"
    // tab's sign-in form - and on CRT's smallest window that was below its edge, with nothing to scroll. Since the same day
    // (2026-11-25, accounts for everybody) those are CRT's account window's, whose every view is held to a window as low as
    // this (AccountWindowTests.Every_view_is_in_reach_in_a_window_as_low_as_CRTs_smallest); the tab keeps what it is and
    // "Sign in", and an account that maintains no board said so - each in reach here, its message line filled.
    // ###########################################################################################
    [Fact]
    public void The_Maintainer_tabs_sign_in_panel_is_in_reach_on_the_smallest_window()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = MainSmallestWindowTests.Smallest();

            try
            {
                window.MainTabControl.SelectedItem = window.MaintainerTabItem;
                ControlsInReach.Settle(window);

                TabMaintainer tab = window.TabMaintainer;
                TextBlock message = tab.GetControl<TextBlock>("SignInMessageText");
                message.Text = "You are signed out.";
                message.IsVisible = true;
                ControlsInReach.Settle(window);

                Assert.True(tab.GetControl<Button>("SignInButton").IsEffectivelyVisible);
                Assert.Empty(ControlsInReach.OutOfReach(window, sideways: false));
            }
            finally
            {
                MainSmallestWindowTests.Close(window);
            }
        });
    }

    // ###########################################################################################
    // THE COMPONENT LIST KEEPS ITS HALF OF THE ROOM ON THE FAR LEFT (Main.SideLists.cs). Measured that day on the C64
    // 250407 at CRT's smallest window: the category list took the room its seven rows wanted, and the component list
    // under it was 9 high - not one row to pick. With more categories than the room holds they scroll in half of it.
    // ###########################################################################################
    [Fact]
    public void The_categories_leave_the_component_list_half_the_room_on_the_smallest_window()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = MainSmallestWindowTests.Smallest();

            try
            {
                window.CategoryFilterListBox.ItemsSource = Enumerable.Range(1, 20).Select(index => "Category " + index).ToList();
                window.ComponentFilterListBox.ItemsSource = Enumerable.Range(1, 60)
                    .Select(index => new ComponentListItem { DisplayText = "U" + index, BoardLabel = "U" + index, SelectionKey = "U" + index })
                    .ToList();
                ControlsInReach.Settle(window);

                Assert.Empty(ControlsInReach.OutOfReach(window, sideways: false));

                double categories = window.CategoryFilterListBox.Bounds.Height;
                double components = window.ComponentFilterListBox.Bounds.Height;

                Assert.True(components >= categories - 1, $"the categories took {categories} and left the components {components}");

                // Not vacuous: the twenty categories do not fit, so their list is the one that gave way.
                ScrollViewer categoryScroll = window.CategoryFilterListBox.GetVisualDescendants().OfType<ScrollViewer>().First();
                Assert.True(categoryScroll.Extent.Height > categoryScroll.Viewport.Height + 1, "every category fits, so nothing here was squeezed");
            }
            finally
            {
                MainSmallestWindowTests.Close(window);
            }
        });
    }

}
