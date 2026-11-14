using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// RowsInViewPanel - a list that builds only the rows in view, its rows editors (2026-11-13, measured). Avalonia's
// VirtualizingStackPanel let go of the row being typed in as it scrolled away, and Tab went by the order its rows were
// BUILT in, not their order in the list (code review, 2026-11-14 - Avalonia 12.1.6's FocusHelpers.GetInputElementChildren
// walks the panel's children as they stand, and a row built again is added at their end).
// ###########################################################################################
[Collection("HeadlessUi")]
public class RowsInViewPanelTests
{
    public sealed class Row(int index)
    {
        public string A { get; } = $"A {index}";

        public string B { get; } = $"B {index}";
    }

    private const int RowCount = 60;

    private sealed record Shown(Window Window, ScrollViewer Viewer, ItemsControl List, Button Before, Button After);

    // A page as the component pane's: its own scroll viewer, a button above the list and one under it, the list's rows two
    // boxes each - bound, as a row built again shows another item.
    private static Shown Show()
    {
        var list = new ItemsControl
        {
            ItemsSource = Enumerable.Range(0, RowCount).Select(index => new Row(index)).ToList(),
            ItemsPanel = new FuncTemplate<Panel?>(() => new RowsInViewPanel()),
            ItemTemplate = new FuncDataTemplate<Row>((_, _) => new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Height = 30,
                Children =
                {
                    new TextBox { Width = 150, [!TextBox.TextProperty] = new Binding(nameof(Row.A)) },
                    new TextBox { Width = 150, [!TextBox.TextProperty] = new Binding(nameof(Row.B)) },
                },
            }),
        };

        var before = new Button { Content = "Before", Height = 30 };
        var after = new Button { Content = "After", Height = 30 };
        var viewer = new ScrollViewer { Content = new StackPanel { Children = { before, list, after } }, Height = 200 };
        var window = new Window { Content = viewer, Width = 400, Height = 300 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return new Shown(window, viewer, list, before, after);
    }

    private static void ScrollTo(Shown shown, int index)
    {
        // The "Before" button's 30, then 30 a row.
        shown.Viewer.Offset = new Vector(0, 30 + (index * 30));
        Dispatcher.UIThread.RunJobs();
    }

    private static TextBox Box(Shown shown, int index, string which)
    {
        Control container = shown.List.ContainerFromIndex(index) ?? throw new InvalidOperationException($"Row {index} is not built");

        return container.GetVisualDescendants().OfType<TextBox>().Single(box => box.Text == $"{which} {index}");
    }

    private static object? Focused(Shown shown) => TopLevel.GetTopLevel(shown.Window)!.FocusManager!.GetFocusedElement();

    private static string? FocusedText(Shown shown) => Focused(shown) switch
    {
        TextBox box => box.Text,
        Button button => button.Content as string,
        _ => null,
    };

    private static void Press(Shown shown, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        shown.Window.KeyPress(Key.Tab, modifiers, PhysicalKey.Tab, keySymbol: null);
        shown.Window.KeyRelease(Key.Tab, modifiers, PhysicalKey.Tab, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    // Scrolled down and back up a little: the rows above are built after the ones that stayed, so they stand last among
    // the panel's children - Tab from row 19 went out of the list, Shift+Tab from row 20 above it. Here by row number.
    [Fact]
    public void Tab_and_shift_tab_go_row_by_row_after_the_list_was_scrolled_up()
    {
        UiTest.Run(() =>
        {
            Shown shown = Show();
            ScrollTo(shown, 20);
            ScrollTo(shown, 15);

            Box(shown, 19, "B").Focus();
            Dispatcher.UIThread.RunJobs();
            Press(shown);
            Assert.Equal("A 20", FocusedText(shown));

            Press(shown, RawInputModifiers.Shift);
            Assert.Equal("B 19", FocusedText(shown));

            shown.Window.Close();
        });
    }

    // From the last row built, Tab goes to the next row - built and brought into view for it - never past rows not built.
    [Fact]
    public void Tab_from_the_last_row_built_goes_to_the_next_row_and_brings_it_into_view()
    {
        UiTest.Run(() =>
        {
            Shown shown = Show();

            int last = shown.List.GetRealizedContainers().Max(container => shown.List.IndexFromContainer(container));
            Box(shown, last, "B").Focus();
            Dispatcher.UIThread.RunJobs();

            Press(shown);

            Assert.Equal($"A {last + 1}", FocusedText(shown));
            Assert.True(((Visual)Focused(shown)!).IsEffectivelyVisible);

            // And back up, row by row, from a row built above the ones already there.
            ScrollTo(shown, 40);
            Box(shown, 40, "A").Focus();
            Dispatcher.UIThread.RunJobs();
            ScrollTo(shown, 40);

            Press(shown, RawInputModifiers.Shift);
            Assert.Equal("B 39", FocusedText(shown));

            shown.Window.Close();
        });
    }

    // The first row's first box and the last row's last box lead out of the list, to what stands before and after it -
    // as they did when every row was built.
    [Fact]
    public void Tab_leaves_the_list_after_its_last_row_and_shift_tab_before_its_first()
    {
        UiTest.Run(() =>
        {
            Shown shown = Show();

            // Scrolled to the end and back a row: the row above is built after the last, so it stands after it.
            ScrollTo(shown, RowCount - 1);
            ScrollTo(shown, RowCount - 7);
            Box(shown, RowCount - 1, "B").Focus();
            Dispatcher.UIThread.RunJobs();
            Press(shown);
            Assert.Same(shown.After, Focused(shown));

            ScrollTo(shown, 3);
            ScrollTo(shown, 0);
            Box(shown, 0, "A").Focus();
            Dispatcher.UIThread.RunJobs();
            Press(shown, RawInputModifiers.Shift);
            Assert.Same(shown.Before, Focused(shown));

            // And into the list from either side: its first row, its last.
            Press(shown);
            Assert.Equal("A 0", FocusedText(shown));

            ScrollTo(shown, RowCount - 1);
            shown.After.Focus();
            Dispatcher.UIThread.RunJobs();
            Press(shown, RawInputModifiers.Shift);
            Assert.Equal($"B {RowCount - 1}", FocusedText(shown));

            shown.Window.Close();
        });
    }

    // The row being typed in stays built as it scrolls away, the caret in its box.
    [Fact]
    public void The_row_being_typed_in_stays_in_the_window_as_it_scrolls_away()
    {
        UiTest.Run(() =>
        {
            Shown shown = Show();

            TextBox box = Box(shown, 1, "A");
            box.Focus();
            Dispatcher.UIThread.RunJobs();

            ScrollTo(shown, RowCount - 5);

            Assert.NotNull(TopLevel.GetTopLevel(box));
            Assert.Equal("A 1", box.Text);
            Assert.Same(box, Focused(shown));

            shown.Window.Close();
        });
    }

    // ###########################################################################################
    // The focus moving INSIDE that row while it is out of view: Avalonia's ItemsControl moves its TabOnceActiveElement from
    // the row to the box that took the focus, the panel let go of the row on that very change, and the box left the window
    // with the keyboard (code review, 2026-11-14). By Tab and by code alike; the row comes back into view first.
    // ###########################################################################################
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void The_focus_moving_inside_the_row_scrolled_away_stays_in_its_box(bool byTab)
    {
        UiTest.Run(() =>
        {
            Shown shown = Show();

            TextBox a = Box(shown, 1, "A");
            TextBox b = Box(shown, 1, "B");
            a.Focus();
            Dispatcher.UIThread.RunJobs();

            ScrollTo(shown, RowCount - 5);

            if (byTab)
            {
                Press(shown);
            }
            else
            {
                b.Focus();
                Dispatcher.UIThread.RunJobs();
            }

            Assert.Same(b, Focused(shown));
            Assert.NotNull(TopLevel.GetTopLevel(b));
            Assert.Equal("B 1", b.Text);
            Assert.True(b.IsEffectivelyVisible);

            // And it stays built as it scrolls away again.
            ScrollTo(shown, RowCount - 5);
            Assert.Same(b, Focused(shown));
            Assert.NotNull(TopLevel.GetTopLevel(b));

            shown.Window.Close();
        });
    }
}
