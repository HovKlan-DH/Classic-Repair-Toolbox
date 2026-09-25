using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Headless;
using Avalonia.Threading;

namespace ClassicRepairToolbox.Tests.Ui;

// Pins down the Avalonia hit-testing rule that a control with no Background does not receive
// pointer events, while the same control with Background="Transparent" does.
//
// This is not an abstract curiosity: the worklog Files list attached PointerPressed directly to a
// TextBlock, and clicking the file link silently did nothing because the press never reached the
// handler. The fix was to wrap it in a Border carrying a transparent Background - the same shape
// the Links rows already used. Without a test the next person to write "put the handler on the
// TextBlock" gets a control that looks wired up and is inert.
[Collection("HeadlessUi")]
public class PointerHitTestingTests
{
    // Arranges a control inside a window and reports whether a press at its centre reaches it.
    private static bool ReceivesPointerPress(Control target)
    {
        bool pressed = false;
        target.PointerPressed += (_, _) => pressed = true;

        var window = new Window
        {
            Width = 200,
            Height = 200,
            Content = new Panel { Children = { target } }
        };

        // Closed in a finally: every UI test shares one headless session and dispatcher, so a window
        // left open by a failing assertion would persist into subsequent tests in this collection
        // and could steal input routing from them - an order-dependent failure that is miserable to
        // track down from the symptom.
        try
        {
            window.Show();
            window.Measure(new Size(200, 200));
            window.Arrange(new Rect(0, 0, 200, 200));
            Dispatcher.UIThread.RunJobs();

            var centre = target.TranslatePoint(new Point(target.Bounds.Width / 2, target.Bounds.Height / 2), window);
            Assert.NotNull(centre);

            window.MouseDown(centre!.Value, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            return pressed;
        }
        finally
        {
            window.Close();
        }
    }

    // The bug: a bare TextBlock has a null Background, so it is skipped by hit-testing entirely
    // and its PointerPressed never fires however precisely it is clicked.
    [Fact]
    public void A_text_block_without_a_background_never_receives_a_pointer_press()
    {
        bool pressed = false;

        UiTest.Run(() =>
        {
            var textBlock = new TextBlock { Text = "file link", Width = 120, Height = 20 };
            pressed = ReceivesPointerPress(textBlock);
        });

        Assert.False(pressed);
    }

    // The fix: Transparent is a real brush for hit-testing purposes even though it paints nothing,
    // so the same click now lands. This is why the link is wrapped in a Border.
    [Fact]
    public void A_border_with_a_transparent_background_receives_a_pointer_press()
    {
        bool pressed = false;

        UiTest.Run(() =>
        {
            var border = new Border
            {
                Background = Brushes.Transparent,
                Width = 120,
                Height = 20,
                Child = new TextBlock { Text = "file link" }
            };

            pressed = ReceivesPointerPress(border);
        });

        Assert.True(pressed);
    }
}
