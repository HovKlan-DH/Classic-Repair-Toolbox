using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// *** THE FOUR SCREENS ARE A TAB STRIP, A SUBMISSION'S VIEWS ONE JOINED SWITCH (owner request,
// 2026-10-01: "I do not know how to handle the 4 button in top-left corner ... It looks like bad UX
// design - do you have any better ideas"). ***
//
// They were grey buttons wrapped onto two rows of the 370-wide list column, looking exactly like
// the Board data / Files / Contributor buttons beside them. Now: the screens are tabs across the
// whole tab, under CRT's own and in its colours (the red underline on the chosen one), and the
// views are a joined switch with the chosen part filled. A real window, read as drawn - the
// template's ContentPresenter is what paints a button, so that is what is asserted.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerScreenTabsTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static readonly string[] ScreenTabs = ["BoardsModeButton", "ReviewModeButton", "BetaModeButton", "AccountModeButton"];

    private static ReviewQueueRow Row() =>
        new(42, "Commodore/C64/250407", "pending", "Corrected U8.", "c@example.com", DateTimeOffset.UtcNow.AddHours(-1), false, false, true);

    private static ReviewSubmissionDetail Detail(ReviewQueueRow row) =>
        new(row, true, new ReviewChangeSummaryView(false, []), [], new ReviewSubmissionAssets([]), [],
            new Dictionary<string, string>(), new Dictionary<string, string>(), [], null);

    // The tab in a window, signed in as an administrator, a submission open. 1700
    // wide so the four fit on one line in the suite's stub font, which is far wider than the real one
    // (TabMaintainerModesTests.The_screen_tabs_stay_on_one_line_inside_the_tab_with_two_digit_badges).
    private static (Window Window, TabMaintainer Tab) Show()
    {
        var tab = new TabMaintainer();
        var window = new Window { Content = tab, Width = 1700, Height = 900, Position = new PixelPoint(0, 0) };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(tab, null);
        typeof(TabMaintainer).GetMethod("SetAdministrator", Any)!.Invoke(tab, [true]);

        ReviewQueueRow row = Row();
        tab.ApplyQueueResponse(new ReviewQueueResponse(true, [row], true));
        typeof(TabMaintainer).GetMethod("ShowSubmission", Any)!.Invoke(tab, [row]);
        tab.ShowDetail(Detail(row));

        // The submission's table, opened without the server - which is what shows its three views.
        tab.OpenTableForTests(row, new ReviewTableData(
            Version: 1,
            Published: new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] },
            Submitted: new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] }));
        Dispatcher.UIThread.RunJobs();

        return (window, tab);
    }

    private static Rect BoundsIn(Window window, Control control) =>
        new(control.TranslatePoint(default, window)!.Value, control.Bounds.Size);

    private static ContentPresenter Painted(Button button) =>
        button.GetVisualDescendants().OfType<ContentPresenter>().First(presenter => presenter.Name == "PART_ContentPresenter");

    private static Color ColourOf(IBrush? brush) => Assert.IsAssignableFrom<ISolidColorBrush>(brush).Color;

    private static Color Resource(string key)
    {
        Assert.True(Application.Current!.TryGetResource(key, ThemeVariant.Light, out object? value));
        return ColourOf((IBrush?)value);
    }

    // ###########################################################################################
    // One line across the WHOLE tab - not wrapped inside the list column - and above both panels,
    // so it reads as choosing the screen rather than as a button of the list.
    // ###########################################################################################
    [Fact]
    public void The_four_screens_are_one_row_of_tabs_across_the_top_above_both_panels()
    {
        UiTest.Run(() =>
        {
            (Window window, TabMaintainer tab) = Show();

            List<Rect> tabs = ScreenTabs.Select(name => BoundsIn(window, tab.FindControl<Button>(name)!)).ToList();
            Rect strip = BoundsIn(window, tab.FindControl<WrapPanel>("ModeBar")!);
            Rect queue = BoundsIn(window, tab.FindControl<Grid>("QueuePanel")!);

            // One line: every tab at the same height, left to right.
            Assert.All(tabs, bounds => Assert.Equal(tabs[0].Y, bounds.Y, 0.5));
            Assert.Equal(tabs.OrderBy(bounds => bounds.X), tabs);

            // Wider than the 370-wide list column, which is where they used to wrap.
            Assert.True(tabs[^1].Right > 370, $"The Account tab ends at {tabs[^1].Right}, inside the list column.");
            Assert.Equal(queue.X, strip.X, 0.5);

            // Above the list on the left and the submission on the right.
            Assert.True(BoundsIn(window, tab.FindControl<Control>("ReviewList")!).Top >= strip.Bottom);
            Assert.True(BoundsIn(window, tab.FindControl<Control>("SubmissionViewBar")!).Top >= strip.Bottom);
        });
    }

    // ###########################################################################################
    // CRT's tab look: no fill, the chosen screen underlined in the main tabs' red, the others not -
    // and greyed by their LABEL only, so a red badge on a screen not chosen keeps its colour.
    // ###########################################################################################
    [Fact]
    public void The_chosen_screen_is_underlined_in_CRTs_tab_red_and_the_others_are_greyed()
    {
        UiTest.Run(() =>
        {
            (_, TabMaintainer tab) = Show();
            Color red = Resource("Main_TabUnderline_Selected");

            ContentPresenter chosen = Painted(tab.FindControl<Button>("ReviewModeButton")!);
            ContentPresenter other = Painted(tab.FindControl<Button>("BoardsModeButton")!);

            Assert.Equal(red, ColourOf(chosen.BorderBrush));
            Assert.Equal(new Thickness(0, 0, 0, 3), chosen.BorderThickness);
            Assert.NotEqual(red, ColourOf(other.BorderBrush));
            Assert.Equal(Colors.Transparent, ColourOf(chosen.Background));

            TextBlock Label(string button) =>
                tab.FindControl<Button>(button)!.GetVisualDescendants().OfType<TextBlock>().First(text => text.Classes.Contains("ScreenTabLabel"));

            Assert.Equal(1, Label("ReviewModeButton").Opacity);
            Assert.True(Label("BoardsModeButton").Opacity < 1);

            // The badge on a screen not chosen is not greyed with its label.
            Assert.Equal(1, tab.FindControl<TextBlock>("BetaBadgeText")!.Opacity);
        });
    }

    // ###########################################################################################
    // A submission's views are ONE joined control - neighbours sharing a border line - with the
    // chosen part filled: nothing like the tabs above it, so the two levels cannot be confused.
    // ###########################################################################################
    [Fact]
    public void The_submissions_views_are_one_joined_switch_with_the_chosen_part_filled()
    {
        UiTest.Run(() =>
        {
            (Window window, TabMaintainer tab) = Show();

            Rect boardData = BoundsIn(window, tab.FindControl<Button>("BoardDataViewButton")!);
            Rect files = BoundsIn(window, tab.FindControl<Button>("FilesViewButton")!);
            Rect contributor = BoundsIn(window, tab.FindControl<Button>("ContributorViewButton")!);

            // Joined: each starts on the border line the one before it ends with.
            Assert.True(boardData.Width > 0, "The views are not laid out - the submission's table is not open.");
            Assert.Equal(boardData.Right - 1, files.X, 0.5);
            Assert.Equal(files.Right - 1, contributor.X, 0.5);

            Assert.Equal(Resource("Mode_Selected_Bg"), ColourOf(Painted(tab.FindControl<Button>("BoardDataViewButton")!).Background));
            Assert.NotEqual(Resource("Mode_Selected_Bg"), ColourOf(Painted(tab.FindControl<Button>("FilesViewButton")!).Background));
        });
    }
}
