using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// *** ONE CLICK ON APPROVE IS ONE CLICK (owner report, 2026-09-30: "Sometimes I do feel that I need
// to click the "Approve and publish to BETA" multiple times for it to react ... I once needed to
// click it 3 times"). ***
//
// The cause was the button's TOOLTIP. Avalonia places a tooltip 20 px below the pointer; at the
// bottom of the window there is no room, so it flips above the pointer - and keeps the 20 px
// offset, which moves it back down OVER the pointer. The click then lands on the tooltip, not on
// the button. With the pointer on Approve and its tooltip open, as it is after a moment's hover,
// the first test's click did not reach the button at all against the version that had one.
//
// A real pointer in a shown window, the tab at the bottom of a window as tall as a real screen.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerDecisionClickTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row() =>
        new(42, "Commodore/C64/250407", "pending", "Corrected U8.", "c@example.com", DateTimeOffset.UtcNow.AddHours(-1), false, false, true);

    private static ReviewSubmissionDetail Detail(ReviewQueueRow row, bool canPublish = true) =>
        new(row, canPublish, new ReviewChangeSummaryView(false, []), [], new ReviewSubmissionAssets([]), [],
            new Dictionary<string, string>(), new Dictionary<string, string>(), [], null);

    // The tab filling a window as tall as a screen, with a submission and its decision bar shown.
    private static (Window Window, TabMaintainer Tab) Show(bool canPublish = true)
    {
        var tab = new TabMaintainer();
        var window = new Window { Content = tab, Width = 1400, Height = 1080, Position = new PixelPoint(0, 0) };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(tab, null);
        ReviewQueueRow row = Row();
        tab.ApplyQueueResponse(new ReviewQueueResponse(true, [row], true));
        typeof(TabMaintainer).GetMethod("ShowSubmission", Any)!.Invoke(tab, [row]);
        tab.ShowDetail(Detail(row, canPublish));
        Dispatcher.UIThread.RunJobs();

        return (window, tab);
    }

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    [Fact]
    public void One_click_on_Approve_at_the_bottom_of_the_window_reaches_it_after_hovering()
    {
        UiTest.Run(() =>
        {
            (Window window, TabMaintainer tab) = Show();
            Button approve = tab.FindControl<Button>("ApproveButton")!;
            Point centre = CentreOf(window, approve);

            int clicks = 0;
            approve.Click += (_, _) => clicks++;

            // Resting on the button: what the tooltip service does after its delay.
            window.MouseMove(centre);
            Dispatcher.UIThread.RunJobs();
            ToolTip.SetIsOpen(approve, true);
            Dispatcher.UIThread.RunJobs();

            window.MouseDown(centre, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            window.MouseUp(centre, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(1, clicks);

            window.Close();
        });
    }

    // ###########################################################################################
    // And none when Approve is off: neither for a system gated elsewhere (the reason is the first
    // line above the table) nor for an account that may not publish - which is said in the line
    // above the buttons instead, where it can be read without hovering at all.
    // ###########################################################################################
    [Fact]
    public void Approve_carries_no_tooltip_and_an_account_that_cannot_publish_is_told_beside_it()
    {
        UiTest.Run(() =>
        {
            (Window window, TabMaintainer tab) = Show(canPublish: false);
            Button approve = tab.FindControl<Button>("ApproveButton")!;
            TextBlock status = tab.FindControl<TextBlock>("ApprovalStatusText")!;

            Assert.False(approve.IsEnabled);
            Assert.Null(ToolTip.GetTip(approve));
            Assert.True(status.IsVisible);
            Assert.Equal(ApprovalWording.NotAMaintainer, status.Text);

            typeof(TabMaintainer).GetMethod("BlockApproval", Any)!.Invoke(tab, null);
            Assert.Null(ToolTip.GetTip(approve));

            window.Close();
        });
    }
}
