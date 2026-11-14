using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The color dropper beside a color picker (owner request, 2026-11-03: "I really miss the 'color picker'
// option, as it then would be easy to select the same color as something existing"). A headless display
// draws nothing to read, so the picture of the window is stood in for (ScreenPixelsTests holds how it is
// read); the render harness proved a real one on 2026-11-03.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ColorDropperButtonTests
{
    private static readonly Color Red = Color.FromRgb(255, 0, 0);

    // A window with the dropper in it, every pixel of the stand-in picture `color`.
    private static (ColorDropperButton Button, Window Window) Shown(Color color)
    {
        byte[] pixels = new byte[400 * 300 * 4];

        for (int at = 0; at < pixels.Length; at += 4)
        {
            pixels[at] = color.B;
            pixels[at + 1] = color.G;
            pixels[at + 2] = color.R;
            pixels[at + 3] = 255;
        }

        var button = new ColorDropperButton
        {
            PixelsOverrideForTests = () => new ScreenPixels(pixels, new PixelSize(400, 300), isBgra: true, 1.0),
        };

        var window = new Window { Content = new StackPanel { Children = { button } }, Width = 400, Height = 300 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (button, window);
    }

    // ###########################################################################################
    // The icon font is the icon's alone (owner report, 2026-11-03: the tooltip read as Font Awesome's
    // symbols). A tooltip counts as a child of its control, so a style on "this button's TextBlocks"
    // reached its words too.
    // ###########################################################################################
    [Fact]
    public void The_tooltip_is_not_in_the_icon_font()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);

            TextBlock icon = button.GetVisualDescendants().OfType<TextBlock>().Single();

            ToolTip.SetIsOpen(button, true);

            for (int i = 0; i < 3; i++)
            {
                Dispatcher.UIThread.RunJobs();
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            }

            // The ToolTip control Avalonia made for the button - held in an internal attached property.
            var tipProperty = (AvaloniaProperty)typeof(ToolTip)
                .GetField("ToolTipProperty", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public)!
                .GetValue(null)!;

            var tip = (ToolTip)button.GetValue(tipProperty)!;
            TextBlock words = tip.GetVisualDescendants().OfType<TextBlock>().First();

            Assert.Equal(ColorDropperButton.Tip, words.Text);
            Assert.Contains("Font Awesome", icon.FontFamily.ToString(), StringComparison.Ordinal);
            Assert.DoesNotContain("Font Awesome", words.FontFamily.ToString(), StringComparison.Ordinal);

            ToolTip.SetIsOpen(button, false);
            window.Close();
        });
    }

    // Pressed, and its cover laid out - a pointer finds nothing on it until then, as for a hand on a mouse.
    private static void Press(ColorDropperButton button)
    {
        button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        Dispatcher.UIThread.RunJobs();
    }

    // ###########################################################################################
    // The cover and the picture are of the window as it was pressed (code review, 2026-11-04). Maximized
    // from its title bar, which does not deactivate it, the part past the old cover took clicks while every
    // key was still the dropper's, and the part still covered read colors from the old layout.
    // ###########################################################################################
    [Fact]
    public void A_window_resized_lets_go_and_a_layout_at_the_same_size_does_not()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);

            Press(button);
            Assert.True(button.IsPicking);

            window.InvalidateMeasure();
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();
            Assert.True(button.IsPicking);

            window.Width = 700;
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            Assert.False(button.IsPicking);
            Assert.False(ColorDropperButton.IsAnyPicking);
            Assert.Null(button.CoverForTests);

            window.Close();
        });
    }

    [Fact]
    public void Pressed_it_covers_the_window_and_a_click_takes_the_color_under_the_pointer()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);
            var told = new List<Color>();
            button.Picked += (_, color) => told.Add(color);

            Press(button);

            Assert.True(button.IsPicking);
            Assert.True(ColorDropperButton.IsAnyPicking);
            Assert.Contains(button.CoverForTests!, OverlayLayer.GetOverlayLayer(window)!.Children);

            window.MouseMove(new Point(100, 100));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("#FF0000", button.LensTextForTests);

            window.MouseDown(new Point(100, 100), MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(Red, button.Color);
            Assert.Equal([Red], told);
            Assert.False(button.IsPicking);
            Assert.False(ColorDropperButton.IsAnyPicking);
            Assert.Empty(OverlayLayer.GetOverlayLayer(window)!.Children.OfType<Canvas>());

            window.MouseUp(new Point(100, 100), MouseButton.Left);
            window.Close();
        });
    }

    // Bound two-way as a picker is, the color reaches what it is bound to - and the binding stays.
    [Fact]
    public void The_color_taken_reaches_what_it_is_bound_to()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);
            var picker = new ColorView { Color = Colors.White };
            button.Bind(ColorDropperButton.ColorProperty, new Avalonia.Data.Binding(nameof(ColorView.Color)) { Source = picker, Mode = Avalonia.Data.BindingMode.TwoWay });

            Press(button);
            window.MouseDown(new Point(50, 50), MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(Red, picker.Color);

            picker.Color = Colors.Blue;
            Assert.Equal(Colors.Blue, button.Color);

            window.MouseUp(new Point(50, 50), MouseButton.Left);
            window.Close();
        });
    }

    // Esc lets go, and the color stays - nor does the key go anywhere else meanwhile.
    [Fact]
    public void Escape_lets_go_with_nothing_taken()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);
            button.Color = Colors.White;

            Press(button);

            window.KeyPressQwerty(PhysicalKey.Escape, RawInputModifiers.None);
            Dispatcher.UIThread.RunJobs();

            Assert.False(button.IsPicking);
            Assert.False(ColorDropperButton.IsAnyPicking);
            Assert.Equal(Colors.White, button.Color);

            window.Close();
        });
    }

    // A right click lets go too - the way out for a hand on the mouse.
    [Fact]
    public void Another_button_lets_go_with_nothing_taken()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);
            button.Color = Colors.White;

            Press(button);
            window.MouseDown(new Point(100, 100), MouseButton.Right);
            Dispatcher.UIThread.RunJobs();

            Assert.False(button.IsPicking);
            Assert.Equal(Colors.White, button.Color);

            window.MouseUp(new Point(100, 100), MouseButton.Right);
            window.Close();
        });
    }

    // The window closing, or the button leaving it, ends it - left on, CRT's Esc stays swallowed for ever.
    [Fact]
    public void Leaving_the_visual_tree_lets_go()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);

            Press(button);
            Assert.True(ColorDropperButton.IsAnyPicking);

            window.Content = null;
            Dispatcher.UIThread.RunJobs();

            Assert.False(button.IsPicking);
            Assert.False(ColorDropperButton.IsAnyPicking);

            window.Close();
        });
    }

    // A picture that cannot be taken is a button that does nothing.
    [Fact]
    public void With_no_picture_of_the_window_nothing_starts()
    {
        UiTest.Run(() =>
        {
            (ColorDropperButton button, Window window) = Shown(Red);
            button.PixelsOverrideForTests = () => null;

            Press(button);

            Assert.False(button.IsPicking);
            Assert.False(ColorDropperButton.IsAnyPicking);

            window.Close();
        });
    }

    // "A general thing": the Schematics tab's trace colors have one too, and it sets what the flyout sets.
    [Fact]
    public void The_trace_palette_has_a_dropper_that_sets_its_color()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var dropper = tab.GetControl<ColorDropperButton>("TraceColorDropper");
            var view = (ColorView)((Flyout)tab.GetControl<Button>("CustomColorButton").Flyout!).Content!;

            dropper.PixelsOverrideForTests = () => new ScreenPixels([0, 0, 255, 255], new PixelSize(1, 1), isBgra: true, 1.0);

            var window = new Window { Content = tab, Width = 1000, Height = 700 };
            window.Show();
            tab.GetControl<Border>("TraceFloatingPalette").IsVisible = true;
            Dispatcher.UIThread.RunJobs();

            Press(dropper);
            window.MouseDown(new Point(0, 0), MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(Red, view.Color);

            window.MouseUp(new Point(0, 0), MouseButton.Left);
            window.Close();
        });
    }
}
