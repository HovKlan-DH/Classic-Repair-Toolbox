using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Threading;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Covers ToolTipPlacement - every tooltip opens at its control's edge, never under the pointer
// (owner report, 2026-09-30: "I need to click the "Approve and publish to BETA" multiple times for
// it to react").
//
// Avalonia's own default opened a tooltip 20 px below the POINTER, flipped it above with no room
// below, kept the offset - and so put it over the pointer at the bottom of the window, where the
// next click landed on the tooltip. The first test here reproduces exactly that on a plain button
// with a long tooltip, and fails against Avalonia's default (0 clicks reach the button).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ToolTipPlacementTests
{
    private const string LongTip =
        "Publishes this submission into the BETA data - once every approval it needs is given. " +
        "Everyone gets it once it is published from Production. This cannot be undone.";

    // A window as tall as a screen with one button, placed at its top or its bottom.
    private static (Window Window, Button Button) Show(VerticalAlignment where)
    {
        var button = new Button
        {
            Content = "Approve and publish to BETA",
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = where,
            Margin = new Thickness(0, 8, 16, 8)
        };

        ToolTip.SetTip(button, ToolTipPlacementTests.LongTip);

        var window = new Window { Content = button, Width = 1400, Height = 1080, Position = new PixelPoint(0, 0) };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, button);
    }

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    // Rests on the button and opens its tooltip, as the tooltip service does after its delay, and
    // answers where the tooltip landed, in the window's coordinates.
    private static Rect OpenTip(Window window, Button button, Point pointer)
    {
        window.MouseMove(pointer);
        Dispatcher.UIThread.RunJobs();
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
        Point origin = tip.TranslatePoint(new Point(0, 0), window)!.Value;

        return new Rect(origin, tip.Bounds.Size);
    }

    [Fact]
    public void At_the_bottom_of_the_window_the_tooltip_is_above_the_button_and_one_click_reaches_it()
    {
        UiTest.Run(() =>
        {
            (Window window, Button button) = Show(VerticalAlignment.Bottom);
            Point pointer = CentreOf(window, button);

            int clicks = 0;
            button.Click += (_, _) => clicks++;

            Rect tip = OpenTip(window, button, pointer);
            Point buttonTop = button.TranslatePoint(new Point(0, 0), window)!.Value;

            Assert.False(tip.Contains(pointer), $"the tooltip {tip} lies over the pointer {pointer}");
            Assert.True(tip.Bottom <= buttonTop.Y + 0.5, $"the tooltip {tip} reaches down into the button at {buttonTop.Y}");

            window.MouseDown(pointer, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            window.MouseUp(pointer, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(1, clicks);

            window.Close();
        });
    }

    // With room below, it is below the button - under it, not over it.
    [Fact]
    public void With_room_below_the_tooltip_opens_just_under_the_button()
    {
        UiTest.Run(() =>
        {
            (Window window, Button button) = Show(VerticalAlignment.Top);
            Point pointer = CentreOf(window, button);

            Rect tip = OpenTip(window, button, pointer);
            double buttonBottom = button.TranslatePoint(new Point(0, button.Bounds.Height), window)!.Value.Y;

            Assert.False(tip.Contains(pointer));
            Assert.True(tip.Top >= buttonBottom - 0.5, $"the tooltip {tip} starts above the button's bottom edge {buttonBottom}");

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE DEFAULT IS THE WEAKEST OPINION (code review, 2026-10-01). *** It is written onto the
    // control when its tooltip first opens; written as a local value it was indistinguishable from
    // a choice by the control's author and beat every style and trigger set afterwards. A placement
    // chosen AFTER the first open must win - locally, and through a style.
    // ###########################################################################################
    [Fact]
    public void A_placement_chosen_after_the_first_open_beats_the_default_the_handler_wrote()
    {
        UiTest.Run(() =>
        {
            var local = new Button { Content = "Local" };
            ToolTip.SetTip(local, "A tip");

            var window = new Window { Content = local, Width = 600, Height = 400 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ToolTip.SetIsOpen(local, true);
            Dispatcher.UIThread.RunJobs();
            ToolTip.SetIsOpen(local, false);

            // The first open placed it at its edge; its author then chose otherwise, and a later
            // open must not put the default back.
            Assert.Equal(PlacementMode.Bottom, ToolTip.GetPlacement(local));

            ToolTip.SetPlacement(local, PlacementMode.Right);
            Assert.Equal(PlacementMode.Right, ToolTip.GetPlacement(local));

            ToolTip.SetIsOpen(local, true);
            Dispatcher.UIThread.RunJobs();
            ToolTip.SetIsOpen(local, false);
            Assert.Equal(PlacementMode.Right, ToolTip.GetPlacement(local));

            // And the handler's own value is not a local one: clearing the author's choice returns
            // to the weakest opinion rather than to nothing.
            local.ClearValue(ToolTip.PlacementProperty);
            Assert.NotEqual(PlacementMode.Right, ToolTip.GetPlacement(local));

            window.Close();
        });
    }

    // ###########################################################################################
    // Any control's tooltip is placed at its edge as it opens - a TextBlock and a Border as much as
    // a button (the table's cells, the file tree's rows) - and a control that chose a placement of
    // its own keeps it.
    // ###########################################################################################
    [Fact]
    public void Every_kind_of_control_is_placed_at_its_edge_as_its_tooltip_opens_unless_it_chose()
    {
        UiTest.Run(() =>
        {
            var own = new Button { Content = "Own" };
            ToolTip.SetPlacement(own, PlacementMode.Right);
            ToolTip.SetVerticalOffset(own, 7);

            Control[] controls = [new TextBlock { Text = "Cell" }, new Border { Width = 40, Height = 20 }, new CheckBox(), own];
            var panel = new StackPanel();

            foreach (Control control in controls)
            {
                ToolTip.SetTip(control, "A tip");
                panel.Children.Add(control);
            }

            var window = new Window { Content = panel, Width = 600, Height = 400 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            foreach (Control control in controls)
            {
                ToolTip.SetIsOpen(control, true);
                Dispatcher.UIThread.RunJobs();
                ToolTip.SetIsOpen(control, false);
            }

            foreach (Control control in controls.Take(3))
            {
                Assert.Equal(PlacementMode.Bottom, ToolTip.GetPlacement(control));
                Assert.Equal(0, ToolTip.GetVerticalOffset(control));
            }

            Assert.Equal(PlacementMode.Right, ToolTip.GetPlacement(own));
            Assert.Equal(7, ToolTip.GetVerticalOffset(own));

            window.Close();
        });
    }
}
