using Avalonia;
using Avalonia.Media;
using Avalonia.Styling;

namespace ClassicRepairToolbox.Tests.Ui;

// The Draft_* theme keys, pinned the same way WorkbooksPaletteTests pins the Workbooks tab's -
// see that file's own header comment for why: a missed key here does not throw, it silently
// renders a fallback colour or (for a DynamicResource binding) leaves the property unset, so
// nothing short of a human looking at both themes catches a typo.
//
// Only keys something actually renders are covered: Draft_Chip_Bg/Fg/Border by
// ComponentFilterListBox's ItemTemplate in Main.axaml, Draft_Highlight_Tint by
// TabSchematics.Highlights.cs's ApplyHighlightVisuals (NewContributeStrategy.md Phase 2,
// session 2b), and Draft_ReferenceBadge_Bg/Border by BoardFilesWindow's Border.ReferenceBadge
// style (one badge per component in the KiCad match report).
[Collection("HeadlessUi")]
public class DraftPaletteTests
{
    private static object? ResolveThemeResource(string key, ThemeVariant? variant = null)
    {
        object? resolved = null;

        UiTest.Run(() =>
        {
            var app = Application.Current;
            Assert.NotNull(app);

            if (app!.TryGetResource(key, variant ?? app.ActualThemeVariant, out var resource))
            {
                resolved = resource;
            }
        });

        return resolved;
    }

    [Theory]
    [InlineData("Draft_Chip_Bg")]
    [InlineData("Draft_Chip_Fg")]
    [InlineData("Draft_Chip_Border")]
    [InlineData("Draft_ReferenceBadge_Bg")]
    [InlineData("Draft_ReferenceBadge_Border")]
    public void Every_draft_chip_brush_resolves_as_a_solid_colour_brush(string key)
    {
        Assert.IsAssignableFrom<ISolidColorBrush>(ResolveThemeResource(key));
    }

    [Fact]
    public void Draft_Highlight_Tint_resolves_as_a_colour()
    {
        Assert.IsType<Color>(ResolveThemeResource("Draft_Highlight_Tint"));
    }

    // Both themes must define the whole set - App.axaml lists Light and Dark a few hundred lines
    // apart, and a key added to only one leaves the other theme showing a hardcoded fallback.
    [Theory]
    [InlineData("Draft_Chip_Bg")]
    [InlineData("Draft_Chip_Fg")]
    [InlineData("Draft_Chip_Border")]
    [InlineData("Draft_Highlight_Tint")]
    [InlineData("Draft_ReferenceBadge_Bg")]
    [InlineData("Draft_ReferenceBadge_Border")]
    public void Every_draft_key_is_defined_in_both_themes(string key)
    {
        Assert.NotNull(ResolveThemeResource(key, ThemeVariant.Light));
        Assert.NotNull(ResolveThemeResource(key, ThemeVariant.Dark));
    }

    // The two themes must actually differ - a key present in both dictionaries with the same
    // literal value would pass every check above while still failing to adapt to dark mode.
    [Theory]
    [InlineData("Draft_Chip_Bg")]
    [InlineData("Draft_Chip_Fg")]
    [InlineData("Draft_ReferenceBadge_Bg")]
    [InlineData("Draft_ReferenceBadge_Border")]
    public void Draft_chip_colours_differ_between_light_and_dark(string key)
    {
        var light = ResolveThemeResource(key, ThemeVariant.Light) as ISolidColorBrush;
        var dark = ResolveThemeResource(key, ThemeVariant.Dark) as ISolidColorBrush;

        Assert.NotNull(light);
        Assert.NotNull(dark);
        Assert.NotEqual(light!.Color, dark!.Color);
    }
}
