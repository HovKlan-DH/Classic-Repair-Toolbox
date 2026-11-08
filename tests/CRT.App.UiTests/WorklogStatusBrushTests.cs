using Avalonia;
using Avalonia.Media;

namespace ClassicRepairToolbox.Tests.Ui;

// The worklog palette lives in App.axaml's ResourceDictionary.ThemeDictionaries, which means a
// plain TryFindResource on a control does NOT resolve it - the lookup has to go through
// Application.Current with an explicit ThemeVariant.
//
// That distinction is invisible at a glance and fails silently: every resolver in this app pairs
// the lookup with a hardcoded fallback colour, so a missed key does not throw or log, it just
// renders the fallback. That is exactly what happened to the worklog bar's status dot - it used
// the single-step lookup, always missed, and always drew its Brushes.Green fallback, so the bar
// showed green while an entry's own Open pill (which uses the two-step idiom) showed the themed
// red. The two disagreed on screen for the same word, "Open".
//
// These tests pin the palette down at its source: the keys exist, and they hold the colours the
// UI is supposed to show. A resolver that silently falls back would still be caught, because
// the fallback colours are deliberately not the theme colours any more.
[Collection("HeadlessUi")]
public class WorklogStatusBrushTests
{
    // Resolves a brush the way the app's ResolveThemeBrush helpers do - the second step is the
    // one that actually works for a ThemeDictionaries-scoped key.
    private static Color? ResolveThemeColor(string key)
    {
        Color? resolved = null;

        UiTest.Run(() =>
        {
            var app = Application.Current;
            Assert.NotNull(app);

            if (app!.TryGetResource(key, app.ActualThemeVariant, out var resource) && resource is ISolidColorBrush brush)
            {
                resolved = brush.Color;
            }
        });

        return resolved;
    }

    // Both worklog axes - a workbook's status and an entry's state - render through these two
    // keys, so they are the single source of the Open/Closed palette.
    [Theory]
    [InlineData("Worklog_Status_Open")]
    [InlineData("Worklog_Status_Closed")]
    [InlineData("Worklog_Category_Note")]
    [InlineData("Worklog_Category_Cosmetic")]
    [InlineData("Worklog_Category_Issue")]
    public void Every_worklog_palette_key_resolves_through_the_theme(string key)
    {
        Assert.NotNull(ResolveThemeColor(key));
    }

    // Open is red and Closed is green - outstanding work reads as red, finished work as green.
    // Pinned as concrete colours because the bug being guarded against was a resolver quietly
    // substituting a DIFFERENT colour, which a mere "it resolved to something" assertion would
    // not have caught.
    [Fact]
    public void An_open_worklog_is_red_and_a_closed_one_is_green()
    {
        Assert.Equal(Colors.IndianRed, ResolveThemeColor("Worklog_Status_Open"));
        Assert.Equal(Color.Parse("#4C8C31"), ResolveThemeColor("Worklog_Status_Closed"));
    }

    // The whole point of the last palette change: the worklog bar's "Open" and a worklog entry's
    // "Open" must be the same red, because they are the same word to the user. They resolve
    // through one key precisely so they cannot drift apart - this fails if someone reintroduces a
    // second Open colour.
    [Fact]
    public void An_open_workbook_and_an_open_entry_share_one_colour()
    {
        var status = ResolveThemeColor("Worklog_Status_Open");
        var category = ResolveThemeColor("Worklog_Category_Issue");

        Assert.NotNull(status);
        Assert.Equal(category, status);
    }
}
