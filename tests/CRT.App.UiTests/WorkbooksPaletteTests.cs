using Avalonia;
using Avalonia.Media;
using Avalonia.Styling;

namespace ClassicRepairToolbox.Tests.Ui;

// The Workbooks tab's palette, pinned the same way WorklogStatusBrushTests pins the worklog bar's.
//
// The reason is the same one that file explains at length: these keys live in App.axaml's
// ResourceDictionary.ThemeDictionaries, and a missed key does not throw or log - the app's
// resolvers all pair a lookup with a hardcoded fallback colour, so a typo renders the fallback
// and nothing says so. The tab itself binds these through DynamicResource, where a missing key
// is even quieter: Avalonia silently leaves the property unset, so a mistyped brush shows up as
// an invisible pin or an uncoloured chip that only a human running the app would notice.
//
// WHAT IS COVERED HERE IS ONLY WHAT THE TAB ACTUALLY RENDERS. Sixteen further Workbooks_* keys
// (Workbooks_Category_Note/_Cosmetic/_Suspected/_Confirmed, Workbooks_State_Fixed/_RuledOut,
// _FixedBadge_*, _ChipSelected_*, _Placeholder_*, _ScopeCapture_*, _PinRing) used to be asserted
// here, and were referenced by nothing but these tests - so the suite certified a palette no
// surface drew. Several of them encoded the MOCKUP's four-category vocabulary
// (Note/Cosmetic/Suspected/Confirmed, Pending/Fixed/Ruled out), which is not the shipped
// WorklogManager model (Note/Cosmetic/Issue, Open/Closed) that every real part of the tab uses -
// and the tab resolves its categories through Worklog_Category_* instead. Someone adding a category
// would have wired it to the Workbooks_* family and got a wrong-coloured chip that a passing
// palette test could not catch. Both the keys and their tests are gone.
[Collection("HeadlessUi")]
public class WorkbooksPaletteTests
{
    private static Color? ResolveThemeColor(string key, ThemeVariant? variant = null)
    {
        Color? resolved = null;

        UiTest.Run(() =>
        {
            var app = Application.Current;
            Assert.NotNull(app);

            if (app!.TryGetResource(key, variant ?? app.ActualThemeVariant, out var resource) && resource is ISolidColorBrush brush)
            {
                resolved = brush.Color;
            }
        });

        return resolved;
    }

    // Every key the tab's markup names. A rename in App.axaml that misses the .axaml - or the
    // reverse - fails here rather than on someone's screen.
    [Theory]
    [InlineData("Workbooks_SelectedCard_Bg")]
    [InlineData("Workbooks_Panel_Bg")]
    [InlineData("Workbooks_Separator")]
    [InlineData("Workbooks_RowSeparator")]
    [InlineData("Workbooks_Muted_Fg")]
    [InlineData("Workbooks_Faint_Fg")]
    [InlineData("Workbooks_ZeroCount_Fg")]
    [InlineData("Workbooks_SearchHit_Bg")]
    [InlineData("Workbooks_SearchHit_Fg")]
    public void Every_workbooks_palette_key_resolves_through_the_theme(string key)
    {
        Assert.NotNull(ResolveThemeColor(key));
    }

    // Both themes must define the whole set. A key added to only one of the two dictionaries is
    // the easy mistake here - App.axaml lists them twice, a few hundred lines apart - and it
    // would leave the tab looking correct in whichever theme the author happened to be using.
    [Theory]
    [InlineData("Workbooks_SelectedCard_Bg")]
    [InlineData("Workbooks_Panel_Bg")]
    [InlineData("Workbooks_Separator")]
    [InlineData("Workbooks_RowSeparator")]
    [InlineData("Workbooks_Muted_Fg")]
    [InlineData("Workbooks_Faint_Fg")]
    [InlineData("Workbooks_ZeroCount_Fg")]
    [InlineData("Workbooks_SearchHit_Bg")]
    [InlineData("Workbooks_SearchHit_Fg")]
    public void Every_workbooks_palette_key_is_defined_in_both_themes(string key)
    {
        Assert.NotNull(ResolveThemeColor(key, ThemeVariant.Light));
        Assert.NotNull(ResolveThemeColor(key, ThemeVariant.Dark));
    }
}
