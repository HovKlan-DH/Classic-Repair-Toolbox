using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ThemeResources' two helpers of the theme color review (2026-11-12): Paint, a color set in code that follows the theme as
// {DynamicResource} does, and TextOn, the white or black that reads on a fill.
// ###########################################################################################
[Collection("HeadlessUi")]
public class ThemeResourcesTests
{
    private static Color ThemeColour(string key, ThemeVariant variant)
    {
        Assert.True(Application.Current!.TryGetResource(key, variant, out object? resource));
        return Assert.IsAssignableFrom<ISolidColorBrush>(resource).Color;
    }

    // The point of Paint over a fixed brush: the same line is the light theme's red in a light window and the dark theme's
    // in a dark one - the Maintainer tab's messages were IndianRed in both until that day.
    [Fact]
    public void A_painted_text_takes_the_colour_of_the_theme_it_is_shown_in()
    {
        UiTest.Run(() =>
        {
            foreach (ThemeVariant variant in new[] { ThemeVariant.Light, ThemeVariant.Dark })
            {
                var text = new TextBlock { Text = "Refused" };
                ThemeResources.Paint(text, TextBlock.ForegroundProperty, "Text_Fail_Fg");

                var window = new Window { Content = text, RequestedThemeVariant = variant };
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(ThemeColour("Text_Fail_Fg", variant), Assert.IsAssignableFrom<ISolidColorBrush>(text.Foreground).Color);

                window.Close();
            }

            Assert.NotEqual(ThemeColour("Text_Fail_Fg", ThemeVariant.Light), ThemeColour("Text_Fail_Fg", ThemeVariant.Dark));
        });
    }

    // And follows a theme CHANGED while it is shown - a fixed brush kept the theme it was set in.
    [Fact]
    public void A_painted_text_follows_its_window_into_another_theme()
    {
        UiTest.Run(() =>
        {
            var text = new TextBlock { Text = "Saved" };
            ThemeResources.Paint(text, TextBlock.ForegroundProperty, "Text_Success_Fg");

            var window = new Window { Content = text, RequestedThemeVariant = ThemeVariant.Light };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(ThemeColour("Text_Success_Fg", ThemeVariant.Light), ((ISolidColorBrush)text.Foreground!).Color);

            window.RequestedThemeVariant = ThemeVariant.Dark;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(ThemeColour("Text_Success_Fg", ThemeVariant.Dark), ((ISolidColorBrush)text.Foreground!).Color);

            window.Close();
        });
    }

    // ###########################################################################################
    // A painted color let go of stays let go when the theme changes (code review, 2026-11-13): Avalonia 12.1.4's
    // ClearValue does not end a binding, so the Configuration tab's board ticked back on turned red again at the next theme
    // switch - and a value assigned over a paint lost to it at the next theme change too (measured). Unpaint ends it, so
    // a value assigned after it stays.
    // ###########################################################################################
    [Fact]
    public void A_painted_colour_let_go_of_stays_let_go_when_the_theme_changes()
    {
        UiTest.Run(() =>
        {
            var unpainted = new TextBlock { Text = "250407" };
            var assigned = new TextBlock { Text = "250425" };
            ThemeResources.Paint(unpainted, TextBlock.ForegroundProperty, "Text_Fail_Fg");
            ThemeResources.Paint(assigned, TextBlock.ForegroundProperty, "Text_Fail_Fg");

            var window = new Window { Content = new StackPanel { Children = { unpainted, assigned } }, RequestedThemeVariant = ThemeVariant.Light };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ThemeResources.Unpaint(unpainted, TextBlock.ForegroundProperty);
            ThemeResources.Unpaint(assigned, TextBlock.ForegroundProperty);
            assigned.Foreground = Brushes.Blue;
            Dispatcher.UIThread.RunJobs();

            window.RequestedThemeVariant = ThemeVariant.Dark;
            Dispatcher.UIThread.RunJobs();

            Assert.NotEqual(ThemeColour("Text_Fail_Fg", ThemeVariant.Dark), (unpainted.Foreground as ISolidColorBrush)?.Color);
            Assert.Equal(Colors.Blue, ((ISolidColorBrush)assigned.Foreground!).Color);

            window.Close();
        });
    }

    // A wash of the theme's color follows the theme as the words over it do (code review, 2026-11-13: the MiniPro test's
    // result kept the light theme's green wash under the dark theme's green PASS).
    [Fact]
    public void A_painted_wash_is_a_faint_copy_of_the_theme_colour_and_follows_its_window()
    {
        UiTest.Run(() =>
        {
            var border = new Border();
            ThemeResources.PaintWash(border, Border.BackgroundProperty, "Text_Success_Fg", 0x22);

            var window = new Window { Content = border, RequestedThemeVariant = ThemeVariant.Light };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Color light = ThemeColour("Text_Success_Fg", ThemeVariant.Light);
            Assert.Equal(Color.FromArgb(0x22, light.R, light.G, light.B), ((ISolidColorBrush)border.Background!).Color);

            window.RequestedThemeVariant = ThemeVariant.Dark;
            Dispatcher.UIThread.RunJobs();

            Color dark = ThemeColour("Text_Success_Fg", ThemeVariant.Dark);
            Assert.Equal(Color.FromArgb(0x22, dark.R, dark.G, dark.B), ((ISolidColorBrush)border.Background!).Color);

            window.Close();
        });
    }

    [Theory]
    [InlineData("#000000", "#FFFFFF")]
    [InlineData("#FFFFFF", "#000000")]
    [InlineData("#B53535", "#FFFFFF")]   // the light theme's red: white 6.0:1, black 3.5:1
    [InlineData("#F26C6C", "#000000")]   // the dark theme's: white 2.9:1, black 7.1:1
    [InlineData("#C8880E", "#000000")]   // an amber in between: black 7.0:1, white 3.0:1
    public void The_text_on_a_fill_is_the_one_of_white_and_black_that_reads_better(string fill, string expected)
    {
        Assert.Equal(Color.Parse(expected), ((ISolidColorBrush)ThemeResources.TextOn(Color.Parse(fill))).Color);
    }

    // What the change is about: a worklog's "#N" stays white on the light theme's categories, and turns black on the dark
    // theme's lighter ones, where white read at 2:1 to 3:1.
    [Fact]
    public void A_worklog_badge_writes_white_in_the_light_theme_and_black_in_the_dark()
    {
        UiTest.Run(() =>
        {
            foreach (string key in new[] { "Worklog_Category_Note", "Worklog_Category_Issue", "Worklog_Status_Open", "Worklog_Status_Closed" })
            {
                Assert.Equal(Colors.White, ((ISolidColorBrush)ThemeResources.TextOn(ThemeColour(key, ThemeVariant.Light))).Color);
                Assert.Equal(Colors.Black, ((ISolidColorBrush)ThemeResources.TextOn(ThemeColour(key, ThemeVariant.Dark))).Color);
            }
        });
    }
}
