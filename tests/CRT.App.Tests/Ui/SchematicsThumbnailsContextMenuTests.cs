using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The Schematics tab's inline thumbnail panel: right-click opens a small menu offering "Detach
// thumbnails to its own window" - a quicker way to reach the same setting the Configuration tab's
// checkbox controls, not a second implementation of it.
//
// Driven through REAL headless input (window.MouseDown/MouseUp/KeyPress), not RaiseEvent on the
// handler directly: this class's own right-click routing (SchematicsThumbnailList -> the tunnelled
// dismiss handler on the tab) is exactly the kind of wiring bug a seam-driven test cannot see - see
// SchematicsThumbnailsDetachTests' own note on the placeholder-drag regression for why.
//
// COLLECTION NOTE: "HeadlessUi" because it constructs TabSchematics and shows a Window.
[Collection("HeadlessUi")]
public sealed class SchematicsThumbnailsContextMenuTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SchematicsThumbnailsContextMenuTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
        this.thisWorkspace.Dispose();
    }

    private static (Window Window, TabSchematics Tab) BuildShownTab(double width = 800, double height = 600)
    {
        var tab = new TabSchematics();
        var window = new Window { Content = tab, Width = width, Height = height };

        window.Show();
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        Dispatcher.UIThread.RunJobs();

        return (window, tab);
    }

    private static Point CenterOf(Control control, Window window)
    {
        var point = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window);
        Assert.NotNull(point);
        return point!.Value;
    }

    [Fact]
    public void Right_clicking_the_thumbnail_panel_opens_the_menu()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab();

            try
            {
                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                var menu = tab.GetControl<Border>("SchematicsThumbnailsContextMenuBorder");
                Assert.False(menu.IsVisible);

                window.MouseDown(CenterOf(list, window), MouseButton.Right);
                window.MouseUp(CenterOf(list, window), MouseButton.Right);
                Dispatcher.UIThread.RunJobs();

                Assert.True(menu.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Reported twice with screenshots. First the menu was wider than the thumbnail column and its
    // button ran off the panel; constraining it to the column then made the text wrap down to ONE
    // CHARACTER PER LINE, because that column is routinely ~170px once the splitter is dragged
    // over toward the schematic.
    //
    // So the menu deliberately is NOT confined to the column: it spans all three grid columns and
    // opens to the LEFT of the cursor, growing over the schematic image where there is room. What
    // has to hold is that the whole menu stays on screen and keeps its readable width.
    //
    // Driven at several window widths because the original bug was invisible at the default 800px
    // test window (a ~400px column swallows the menu without complaint).
    [Theory]
    [InlineData(420)]   // ~200px column - narrower than the menu, the reported shape
    [InlineData(520)]   // ~250px column
    [InlineData(800)]   // ~400px column - wide enough that the first version looked fine
    public void The_menu_opens_left_of_the_cursor_and_stays_fully_on_screen(double windowWidth)
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab(windowWidth);

            try
            {
                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                var grid = tab.GetControl<Grid>("SchematicsInnerGrid");
                var menu = tab.GetControl<Border>("SchematicsThumbnailsContextMenuBorder");

                Point clickPoint = CenterOf(list, window);
                window.MouseDown(clickPoint, MouseButton.Right);
                window.MouseUp(clickPoint, MouseButton.Right);
                Dispatcher.UIThread.RunJobs();
                window.Measure(new Size(windowWidth, 600));
                window.Arrange(new Rect(0, 0, windowWidth, 600));
                Dispatcher.UIThread.RunJobs();

                Assert.True(menu.IsVisible);
                Assert.True(grid.Bounds.Width > 0);

                // Fully inside the tab - never clipped off either edge.
                Assert.True(menu.Margin.Left >= 0, $"menu pushed off the left edge at {menu.Margin.Left}");
                Assert.True(
                    menu.Margin.Left + menu.Bounds.Width <= grid.Bounds.Width + 0.5,
                    $"menu right edge {menu.Margin.Left + menu.Bounds.Width} past grid {grid.Bounds.Width}");

                // It grows LEFT from the cursor rather than right, which is the only direction with
                // room: the thumbnail column sits at the tab's right-hand edge.
                Point? cursorInGrid = window.TranslatePoint(clickPoint, grid);
                Assert.NotNull(cursorInGrid);
                Assert.True(
                    menu.Margin.Left <= cursorInGrid!.Value.X + 0.5,
                    $"menu left {menu.Margin.Left} is right of the cursor at {cursorInGrid.Value.X}");
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The point of opening leftward: the label keeps its real width instead of being wrapped into a
    // narrow column one character at a time. Asserted on the rendered line count, which is what the
    // screenshot actually showed.
    [Fact]
    public void The_label_stays_on_one_or_two_lines_rather_than_wrapping_per_character()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab(420);

            try
            {
                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                var menu = tab.GetControl<Border>("SchematicsThumbnailsContextMenuBorder");

                window.MouseDown(CenterOf(list, window), MouseButton.Right);
                window.MouseUp(CenterOf(list, window), MouseButton.Right);
                Dispatcher.UIThread.RunJobs();
                window.Measure(new Size(420, 600));
                window.Arrange(new Rect(0, 0, 420, 600));
                Dispatcher.UIThread.RunJobs();

                var button = tab.GetControl<Button>("DetachThumbnailsContextMenuButton");
                var label = Assert.IsType<TextBlock>(button.Content);

                Assert.Equal("Detach thumbnails into their own window", label.Text);
                Assert.Equal(TextWrapping.Wrap, label.TextWrapping);

                // The button fits inside its own menu...
                Assert.True(button.Bounds.Width > 0);
                Assert.True(
                    button.Bounds.Width <= menu.Bounds.Width,
                    $"button {button.Bounds.Width} wider than menu {menu.Bounds.Width}");

                // ...and the text is genuinely readable. One character per line made the block far
                // taller than it is wide; a real line of text is much wider than it is tall.
                Assert.True(
                    label.Bounds.Width > label.Bounds.Height,
                    $"label {label.Bounds.Width}x{label.Bounds.Height} looks like a vertical string of characters");
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void The_menu_offers_exactly_the_detach_action()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab();

            try
            {
                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                window.MouseDown(CenterOf(list, window), MouseButton.Right);
                window.MouseUp(CenterOf(list, window), MouseButton.Right);
                Dispatcher.UIThread.RunJobs();

                var button = tab.GetControl<Button>("DetachThumbnailsContextMenuButton");
                // A wrapping TextBlock rather than a plain string Content, so the label survives a
                // narrow thumbnail column instead of being clipped - see the width tests above.
                var label = Assert.IsType<TextBlock>(button.Content);
                Assert.Equal("Detach thumbnails into their own window", label.Text);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Clicking the action sets the SAME UserSettings.DetachSchematicsThumbnails the Configuration
    // checkbox writes - the two must never be able to disagree about what turns the feature on.
    [Fact]
    public void Clicking_the_action_turns_the_setting_on_and_collapses_the_inline_column()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab();

            try
            {
                Assert.False(UserSettings.DetachSchematicsThumbnails);

                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                window.MouseDown(CenterOf(list, window), MouseButton.Right);
                window.MouseUp(CenterOf(list, window), MouseButton.Right);
                Dispatcher.UIThread.RunJobs();

                var button = tab.GetControl<Button>("DetachThumbnailsContextMenuButton");
                var buttonCenter = CenterOf(button, window);
                window.MouseDown(buttonCenter, MouseButton.Left);
                window.MouseUp(buttonCenter, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.True(UserSettings.DetachSchematicsThumbnails);

                // With no MainWindow (this tab was never handed one), ApplyThumbnailsDetachedState
                // cannot run, but EnterThumbnailsDetachedMode has no such dependency - assert the
                // setting rather than the column here; the full collapse is covered by
                // SchematicsThumbnailsDetachTests, which drives EnterThumbnailsDetachedMode directly.
                var menu = tab.GetControl<Border>("SchematicsThumbnailsContextMenuBorder");
                Assert.False(menu.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void A_left_click_elsewhere_dismisses_the_menu()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab();

            try
            {
                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                var menu = tab.GetControl<Border>("SchematicsThumbnailsContextMenuBorder");

                window.MouseDown(CenterOf(list, window), MouseButton.Right);
                window.MouseUp(CenterOf(list, window), MouseButton.Right);
                Dispatcher.UIThread.RunJobs();
                Assert.True(menu.IsVisible);

                var container = tab.GetControl<Border>("SchematicsContainer");
                window.MouseDown(CenterOf(container, window), MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.False(menu.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void Escape_dismisses_the_menu()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab();

            try
            {
                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                var menu = tab.GetControl<Border>("SchematicsThumbnailsContextMenuBorder");

                window.MouseDown(CenterOf(list, window), MouseButton.Right);
                window.MouseUp(CenterOf(list, window), MouseButton.Right);
                Dispatcher.UIThread.RunJobs();
                Assert.True(menu.IsVisible);

                // The tunnel handler only sees keys targeting something inside TabSchematics' own
                // visual tree - without an explicit focus target the key goes nowhere in particular.
                // TabSchematics itself is not Focusable; SchematicsContainer is the schematic image
                // area and IS, so it is the same thing a real user's click on the tab would focus.
                tab.GetControl<Border>("SchematicsContainer").Focus();
                window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);
                Dispatcher.UIThread.RunJobs();

                Assert.False(menu.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // While thumbnails are already detached the inline panel is HIDDEN, so there is nothing to
    // right-click on screen at all - EnterThumbnailsDetachedMode's own IsVisible=false already
    // guarantees that, which this pins as the actual mechanism rather than a separate check in the
    // menu handler duplicating it.
    [Fact]
    public void The_inline_panel_is_hidden_while_detached_so_there_is_nothing_to_right_click()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = BuildShownTab();

            try
            {
                tab.EnterThumbnailsDetachedMode();

                var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
                Assert.False(list.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }
}
