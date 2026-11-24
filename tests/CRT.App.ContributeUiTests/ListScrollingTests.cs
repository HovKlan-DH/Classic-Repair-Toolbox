using System.Collections.ObjectModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ListScrolling: a list scrolled to a row after its rows changed, and never a hidden one. Its user is the "Contribute"
// tab's component list (TabDrafts.ReadComponentLists, ShowNavSelection) - so its tests stand with that tab's.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ListScrollingTests
{
    // Avalonia's own ScrollIntoView builds the row and lets go of it again in its next layout pass, the row lying outside
    // the old view - the list stayed at its top (measured on 12.1.6 and 12.2.0, 2026-11-24).
    [Fact]
    public async Task A_row_put_in_since_the_last_layout_pass_is_scrolled_to()
    {
        await UiTest.RunAsync(() =>
        {
            var items = new ObservableCollection<string> { "U400" };
            (Window window, ListBox list) = Shown(items);
            list.SelectedIndex = 0;
            Dispatcher.UIThread.RunJobs();

            for (int number = 1; number < 400; number++)
            {
                items.Insert(number - 1, "U" + number);
            }

            ListScrolling.ToRow(list, list.SelectedIndex);
            Tick();

            Assert.True(InView(list, 399));

            window.Close();
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task A_list_laid_out_already_is_scrolled_to_the_row()
    {
        await UiTest.RunAsync(() =>
        {
            (Window window, ListBox list) = Shown(new ObservableCollection<string>(Enumerable.Range(1, 400).Select(number => "U" + number)));

            ListScrolling.ToRow(list, 399);
            Tick();

            Assert.True(InView(list, 399));

            window.Close();
            return Task.CompletedTask;
        });
    }

    // Avalonia 12.2.0 keeps a scroll asked of a hidden list and does it as the list is shown again; the lists on the far
    // left are shown and hidden as the "Contribute" tab is entered and left, each put where the other stood.
    [Fact]
    public async Task A_hidden_list_is_not_scrolled_not_even_once_it_is_shown_again()
    {
        await UiTest.RunAsync(() =>
        {
            (Window window, ListBox list) = Shown(new ObservableCollection<string>(Enumerable.Range(1, 400).Select(number => "U" + number)));
            ScrollViewer scroller = list.GetVisualDescendants().OfType<ScrollViewer>().First();

            list.IsVisible = false;
            Tick();

            ListScrolling.ToRow(list, 399);
            Tick();

            list.IsVisible = true;
            Tick();
            Tick();

            Assert.Equal(0, scroller.Offset.Y);
            Assert.False(InView(list, 399));

            window.Close();
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task No_row_is_no_scroll()
    {
        await UiTest.RunAsync(() =>
        {
            (Window window, ListBox list) = Shown(new ObservableCollection<string>(Enumerable.Range(1, 400).Select(number => "U" + number)));
            ScrollViewer scroller = list.GetVisualDescendants().OfType<ScrollViewer>().First();

            ListScrolling.ToRow(list, -1);
            Tick();

            Assert.Equal(0, scroller.Offset.Y);

            window.Close();
            return Task.CompletedTask;
        });
    }

    // As the tab's list in Main.axaml: Avalonia's own scrolling to the selection off.
    private static (Window Window, ListBox List) Shown(ObservableCollection<string> items)
    {
        var list = new ListBox { ItemsSource = items, AutoScrollToSelectedItem = false };
        var window = new Window { Content = list, Width = 300, Height = 300 };

        window.Show();
        Tick();

        return (window, list);
    }

    private static void Tick()
    {
        Dispatcher.UIThread.RunJobs();
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        Dispatcher.UIThread.RunJobs();
    }

    private static bool InView(ListBox list, int index)
    {
        ScrollViewer scroller = list.GetVisualDescendants().OfType<ScrollViewer>().First();

        return list.ContainerFromIndex(index) is Control row &&
               row.IsEffectivelyVisible &&
               row.TranslatePoint(default, scroller) is { } at &&
               at.Y >= -1 &&
               at.Y < scroller.Viewport.Height;
    }
}
