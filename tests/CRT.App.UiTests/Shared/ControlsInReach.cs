using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// WHAT A PERSON CANNOT GET TO IN A WINDOW (owner report, 2026-11-25: on Linux the "Maintainer" tab's "Reset code" box
// lay below the window, "outside viewport and there is no possibility for me to scroll down ... it makes the UI useless
// in these situation").
//
// Every box, button, list, drop-down or tick drawn in the window of which less than HALF lies inside it, unless a
// scroll viewer around it that scrolls that way has its viewport inside the window: scrolling it brings the control
// in. A control mostly in the window can still be clicked and typed in - one wholly below its edge, as the reset code
// was, cannot. And a box lower than a line of text, or a list lower than two rows: drawn, but squeezed until nobody
// can use it (the "Feedback" tab's text box at 500 high).
//
// Parts of a control's own template (a scroll bar's buttons, a drop-down's toggle) are the control's, so they are
// left to it. Read off a SHOWN window after a layout pass.
// ###########################################################################################
public static class ControlsInReach
{
    private const double Slack = 1;

    // A box lower than a line of text, a list lower than two of its rows.
    private const double LowestBox = 18;
    private const double LowestList = 40;

    // One line per control out of reach or squeezed - its type, name, words and where it is - or none.
    //
    // `sideways`: whether a control past the window's LEFT OR RIGHT edge counts too. Off for a whole tab of CRT: the
    // headless platform's font is about twice as wide as any real one ("I forgot my password" 280 wide against 140 in
    // Segoe UI) while as high, so a line of buttons that fits on every desktop runs past the edge here - a test of how
    // HIGH content is stays true to life, one of how wide it is does not.
    public static List<string> OutOfReach(Window window, bool sideways = true)
    {
        var client = new Rect(window.ClientSize);
        var found = new List<string>();

        foreach (Control control in window.GetVisualDescendants().OfType<Control>())
        {
            if (!ControlsInReach.IsInteractive(control)
                || control.TemplatedParent is not null
                || !control.IsEffectivelyVisible
                || control.Bounds.Width < Slack
                || (control.Bounds.Height < Slack && control is not (TextBox or ListBox)))
            {
                continue;
            }

            if (ControlsInReach.PlaceIn(control, window) is not { } rect)
            {
                continue;
            }

            // An empty list may be as low as it likes - there is nothing in it to pick.
            if ((control is TextBox && rect.Height < ControlsInReach.LowestBox) || (control is ListBox { ItemCount: > 0 } && rect.Height < ControlsInReach.LowestList))
            {
                found.Add($"squeezed to {rect.Height:0} high: {ControlsInReach.Describe(control, rect, client)}");
                continue;
            }

            if (control is ListBox && rect.Height < Slack)
            {
                continue;
            }

            // Less than half of it inside the window, that way.
            double seenHigh = Math.Min(rect.Bottom, client.Bottom) - Math.Max(rect.Top, client.Top);
            double seenWide = Math.Min(rect.Right, client.Right) - Math.Max(rect.Left, client.Left);
            bool outVertically = seenHigh < rect.Height / 2;
            bool outHorizontally = sideways && seenWide < rect.Width / 2;

            if ((outVertically && !ControlsInReach.ScrollsIn(control, window, client, vertical: true))
                || (outHorizontally && !ControlsInReach.ScrollsIn(control, window, client, vertical: false)))
            {
                found.Add("out of reach: " + ControlsInReach.Describe(control, rect, client));
            }
        }

        return found;
    }

    // Layout and whatever it posts, run until the window has settled.
    public static void Settle(Window window)
    {
        for (int i = 0; i < 3; i++)
        {
            Dispatcher.UIThread.RunJobs();
            window.UpdateLayout();
        }
    }

    // Where a control is in the window, after every scroll viewer's offset.
    public static Rect? PlaceIn(Visual visual, Visual window)
    {
        Matrix? matrix = visual.TransformToVisual(window);
        return matrix is null ? null : new Rect(visual.Bounds.Size).TransformToAABB(matrix.Value);
    }

    private static bool IsInteractive(Control control) =>
        control is TextBox or Button or ToggleButton or ComboBox or ListBox or Slider or AutoCompleteBox or NumericUpDown or TreeView;

    private static bool ScrollsIn(Control control, Window window, Rect client, bool vertical)
    {
        foreach (ScrollViewer viewer in control.GetVisualAncestors().OfType<ScrollViewer>())
        {
            ScrollBarVisibility visibility = vertical ? viewer.VerticalScrollBarVisibility : viewer.HorizontalScrollBarVisibility;

            if (visibility == ScrollBarVisibility.Disabled)
            {
                continue;
            }

            Visual viewport = viewer.GetVisualDescendants().OfType<ScrollContentPresenter>().FirstOrDefault() ?? (Visual)viewer;

            if (ControlsInReach.PlaceIn(viewport, window) is not { } rect || rect.Width < Slack || rect.Height < Slack)
            {
                continue;
            }

            bool inside = vertical
                ? rect.Top >= client.Top - Slack && rect.Bottom <= client.Bottom + Slack
                : rect.Left >= client.Left - Slack && rect.Right <= client.Right + Slack;

            if (inside)
            {
                return true;
            }
        }

        return false;
    }

    private static string Describe(Control control, Rect rect, Rect client)
    {
        string words = control switch
        {
            TextBox box => box.PlaceholderText ?? box.Text ?? string.Empty,
            ContentControl { Content: string text } => text,
            _ => string.Empty,
        };

        string within = string.Join(" < ", control.GetVisualAncestors().OfType<StyledElement>()
            .Where(element => !string.IsNullOrEmpty(element.Name))
            .Take(3)
            .Select(element => element.Name));

        return $"{control.GetType().Name} '{control.Name}' \"{words}\" at {rect.X:0},{rect.Y:0} {rect.Width:0}x{rect.Height:0} in a window of {client.Width:0}x{client.Height:0}, in {within}";
    }
}
