using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// HOME, END, PAGE UP AND PAGE DOWN SCROLL THE "ABOUT" TAB (owner request, 2026-11-06: "When I am placed anywhere
// in the 'About' tab and I press 'Home', then I want it to navigate to the top of the page. Likewise with 'End'
// should go to the end. Also, 'Page/Up' and down should work") - Main.KeyIsForAboutPage hands them to
// TabAbout.TakeScrollKey from wherever the keyboard is. Through the real window, as a key's route needs it: the
// focus is in the component filter, where every click on the tab sends it.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainAboutTabKeysTests
{
    [Fact]
    public void Home_End_and_the_page_keys_scroll_the_about_tab_with_the_focus_in_the_component_filter()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { Width = 1300, Height = 800 };
            window.Show();

            try
            {
                window.MainTabControl.SelectedItem = window.AboutTabItem;
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                ItemsControl items = window.TabAbout.AboutItems;
                ScrollViewer viewer = items.GetVisualDescendants().OfType<ScrollViewer>().First();
                Assert.True(viewer.Extent.Height > viewer.Viewport.Height * 3, "the page is too short to scroll");

                TextBox filter = window.ComponentSearchTextBox;
                filter.Text = "U1";
                filter.Focus();
                filter.CaretIndex = 1;
                Assert.True(filter.IsFocused);

                void Press(Key key, PhysicalKey physical)
                {
                    window.KeyPress(key, RawInputModifiers.None, physical, keySymbol: null);
                    Dispatcher.UIThread.RunJobs();
                    window.UpdateLayout();
                }

                // End: the very end, the last block of the changelog built - not where a guess at the heights of
                // blocks never built put it.
                Press(Key.End, PhysicalKey.End);
                Assert.True(viewer.Offset.Y + viewer.Viewport.Height >= viewer.Extent.Height - 1, $"End stopped at {viewer.Offset.Y} of {viewer.Extent.Height - viewer.Viewport.Height}");
                Assert.NotNull(items.ContainerFromIndex(items.ItemCount - 1));

                Press(Key.Home, PhysicalKey.Home);
                Assert.Equal(0, viewer.Offset.Y);

                Press(Key.PageDown, PhysicalKey.PageDown);
                double onePage = viewer.Offset.Y;
                Assert.True(onePage > viewer.Viewport.Height / 2, $"Page Down moved {onePage} of a {viewer.Viewport.Height} page");

                Press(Key.PageDown, PhysicalKey.PageDown);
                Assert.True(viewer.Offset.Y > onePage);

                Press(Key.PageUp, PhysicalKey.PageUp);
                Assert.Equal(onePage, viewer.Offset.Y, 1.0);

                // The filter kept its words, its caret where it was - the keys were the page's.
                Assert.Equal("U1", filter.Text);
                Assert.Equal(1, filter.CaretIndex);
            }
            finally
            {
                window.IsApplicationShuttingDownForTests = true;
                window.Close();
            }
        });
    }

    // On any other tab the keys are where the keyboard is: the filter's Home moves its caret, the page stays.
    [Fact]
    public void On_another_tab_the_keys_stay_where_the_keyboard_is()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { Width = 1300, Height = 800 };
            window.Show();

            try
            {
                window.MainTabControl.SelectedItem = window.SchematicsTabItem;
                Dispatcher.UIThread.RunJobs();

                TextBox filter = window.ComponentSearchTextBox;
                filter.Text = "U1";
                filter.Focus();
                filter.CaretIndex = 2;

                window.KeyPress(Key.Home, RawInputModifiers.None, PhysicalKey.Home, keySymbol: null);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(0, filter.CaretIndex);
            }
            finally
            {
                window.IsApplicationShuttingDownForTests = true;
                window.Close();
            }
        });
    }
}
