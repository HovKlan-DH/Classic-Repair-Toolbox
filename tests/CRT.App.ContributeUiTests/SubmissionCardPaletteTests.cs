using Avalonia;
using Avalonia.Media;
using Avalonia.Styling;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The two theme keys the "My submissions" cards added (owner request, 2026-09-22):
// Card_Bg and Text_Waiting_Fg.
//
// *** PINNED FOR THE REASON WorkbooksPaletteTests SPELLS OUT: A MISSING KEY IS SILENT. *** The
// row model pairs every lookup with a hardcoded fallback, so a typo renders the fallback and
// nothing says so; where the markup binds a key through DynamicResource, Avalonia simply leaves
// the property unset. Either way the only thing that notices is a human running the app in the
// theme that happens to be broken.
//
// *** BOTH THEMES, ALWAYS. *** These were added because the card had no fill and was invisible
// against the window. A hardcoded light grey would have fixed the light theme and glared in the
// dark one, so the keys step in OPPOSITE directions - darker than the background in the light
// theme, lighter in the dark one - and that is exactly the kind of thing that gets "tidied" into
// one shared value later.
// ###########################################################################################
[Collection("HeadlessUi")]
public class SubmissionCardPaletteTests
{
    private static Color? ResolveThemeColor(string key, ThemeVariant variant)
    {
        Color? resolved = null;

        UiTest.Run(() =>
        {
            var app = Application.Current;
            Assert.NotNull(app);

            if (app!.TryGetResource(key, variant, out var resource) && resource is ISolidColorBrush brush)
            {
                resolved = brush.Color;
            }
        });

        return resolved;
    }

    [Theory]
    [InlineData("Card_Bg")]
    [InlineData("Text_Waiting_Fg")]
    [InlineData("Text_Warning_Fg")]
    public void Each_new_key_resolves_in_BOTH_themes(string key)
    {
        Assert.NotNull(SubmissionCardPaletteTests.ResolveThemeColor(key, ThemeVariant.Light));
        Assert.NotNull(SubmissionCardPaletteTests.ResolveThemeColor(key, ThemeVariant.Dark));
    }

    [Fact]
    public void The_card_fill_DIFFERS_from_the_window_background_in_both_themes()
    {
        // *** THE WHOLE POINT OF THE KEY. *** The cards were reported as not noticeable because
        // they used Table_Bg, which is plain White in the light theme - a card filled with the
        // window's own colour is not a card. A future edit that points Card_Bg back at the
        // background colour would silently restore exactly that.
        foreach (ThemeVariant variant in new[] { ThemeVariant.Light, ThemeVariant.Dark })
        {
            Color? card = SubmissionCardPaletteTests.ResolveThemeColor("Card_Bg", variant);
            Color? background = SubmissionCardPaletteTests.ResolveThemeColor("Bg", variant);

            Assert.NotNull(card);
            Assert.NotNull(background);
            Assert.NotEqual(background!.Value, card!.Value);
        }
    }

    [Fact]
    public void The_status_colours_are_DISTINGUISHABLE_from_each_other_in_both_themes()
    {
        // ALL FOUR outcome buckets have to be different colours, or the accent edge says nothing.
        // Asserted per theme because the dark theme defines its own values.
        foreach (ThemeVariant variant in new[] { ThemeVariant.Light, ThemeVariant.Dark })
        {
            var colours = new Dictionary<string, Color>();

            foreach (string key in new[]
            {
                "Text_Success_Fg", "Text_Fail_Fg", "Text_Waiting_Fg", "Text_Warning_Fg"
            })
            {
                Color? resolved = SubmissionCardPaletteTests.ResolveThemeColor(key, variant);
                Assert.NotNull(resolved);

                colours[key] = resolved!.Value;
            }

            Assert.Equal(colours.Count, colours.Values.Distinct().Count());
        }
    }

    // ###########################################################################################
    // *** THE REPORTED BUG - AND TWO WRONG TESTS FOR IT BEFORE THIS ONE (owner request,
    // 2026-09-23, reported twice). ***
    //
    // The bug: "Changes requested" and "Not accepted" both painted themselves Text_Fail_Fg, so a
    // submission the contributor could still rescue looked exactly like one that had been turned
    // down. The replacement orange was then ALSO reported as reading red, and both of the tests
    // written to catch that passed it. They are worth recording, because each was wrong in a way
    // that looks right:
    //
    //   1. GREEN CHANNEL, absolute. "Orange has more green than red does." False against this
    //      palette: the light theme's Text_Fail_Fg is IndianRed, a DESATURATED red whose green
    //      channel is already as high as a dark readable orange's.
    //
    //   2. HUE. "25 degrees from red is orange." Also false, and this is the interesting one -
    //      hue reported the two colours 25 degrees apart while they still looked identical.
    //
    // *** WHAT ACTUALLY MATCHED WAS THE GREEN-TO-RED RATIO. *** IndianRed is 205,92,92; the
    // rejected orange was 194,87,10. The R and G channels are practically the same and only blue
    // differs, so the eye reads "the same colour, less washed out" rather than "a different hue".
    // Both sat at G/R = 0.45.
    //
    // So that ratio is what this test asserts, bracketed on both sides: clear of the failure red
    // below and of the waiting amber above. It is a measurement of the thing that was actually
    // reported, which is why it rejects the colour the other two admitted.
    // ###########################################################################################
    [Fact]
    public void The_WARNING_colour_does_not_share_the_FAILURE_colours_red_to_green_balance()
    {
        foreach (ThemeVariant variant in new[] { ThemeVariant.Light, ThemeVariant.Dark })
        {
            double warning = SubmissionCardPaletteTests.GreenOverRed(
                SubmissionCardPaletteTests.ResolveThemeColor("Text_Warning_Fg", variant)!.Value);

            double fail = SubmissionCardPaletteTests.GreenOverRed(
                SubmissionCardPaletteTests.ResolveThemeColor("Text_Fail_Fg", variant)!.Value);

            Assert.True(
                warning - fail >= 0.10,
                $"Text_Warning_Fg must not share Text_Fail_Fg's red-to-green balance in the " +
                $"{variant} theme - that is what made the rejected #C2570A read as red beside " +
                $"IndianRed. Its green-to-red ratio leads by only {warning - fail:0.00}.");
        }
    }

    // ###########################################################################################
    // The other side of the bracket: warning must not drift into the amber that already means
    // "just waiting" either. One is a call to ACT and the other is passive, and this is the pair
    // most at risk of being "tidied" into a single colour later, since both are orange-ish.
    //
    // Measured on the same axis as the test above, deliberately - a colour squeezed between two
    // others has to be checked against both on the SAME ruler, or satisfying one end can quietly
    // violate the other.
    // ###########################################################################################
    [Fact]
    public void The_WARNING_colour_stays_clear_of_the_WAITING_amber_too()
    {
        foreach (ThemeVariant variant in new[] { ThemeVariant.Light, ThemeVariant.Dark })
        {
            double warning = SubmissionCardPaletteTests.GreenOverRed(
                SubmissionCardPaletteTests.ResolveThemeColor("Text_Warning_Fg", variant)!.Value);

            double waiting = SubmissionCardPaletteTests.GreenOverRed(
                SubmissionCardPaletteTests.ResolveThemeColor("Text_Waiting_Fg", variant)!.Value);

            Assert.True(
                waiting - warning >= 0.05,
                $"Text_Waiting_Fg must stay yellower than Text_Warning_Fg in the {variant} theme, " +
                $"but its green-to-red ratio leads by only {waiting - warning:0.00}.");
        }
    }

    [Fact]
    public void The_card_fill_steps_in_OPPOSITE_directions_in_the_two_themes()
    {
        // Light theme: a card is DARKER than its background. Dark theme: LIGHTER. Anyone
        // "simplifying" these to one shared grey breaks one of the two, and which one depends on
        // the theme they happen to be running - so it would very likely ship.
        Color light = SubmissionCardPaletteTests.ResolveThemeColor("Card_Bg", ThemeVariant.Light)!.Value;
        Color lightBackground = SubmissionCardPaletteTests.ResolveThemeColor("Bg", ThemeVariant.Light)!.Value;

        Color dark = SubmissionCardPaletteTests.ResolveThemeColor("Card_Bg", ThemeVariant.Dark)!.Value;
        Color darkBackground = SubmissionCardPaletteTests.ResolveThemeColor("Bg", ThemeVariant.Dark)!.Value;

        Assert.True(
            SubmissionCardPaletteTests.Brightness(light) < SubmissionCardPaletteTests.Brightness(lightBackground),
            "Card_Bg must be darker than the background in the LIGHT theme.");

        Assert.True(
            SubmissionCardPaletteTests.Brightness(dark) > SubmissionCardPaletteTests.Brightness(darkBackground),
            "Card_Bg must be lighter than the background in the DARK theme.");
    }

    // Plain average rather than a perceptual luminance: this only has to answer "lighter or
    // darker", and these are near-neutral greys where the two agree.
    private static double Brightness(Color color) => (color.R + color.G + color.B) / 3.0;

    // ###########################################################################################
    // Hue in degrees, 0 = red, 60 = yellow - the standard HSL hue, written out because neither
    // Avalonia's Color nor System.Drawing is available here as a one-liner that returns it.
    //
    // Every colour these tests compare sits in the red-to-yellow arc, so the wrap-around at 360 is
    // not a case that can arise; a grey (max == min) answers 0, which no caller passes.
    // ###########################################################################################
    // ###########################################################################################
    // THE GREEN-TO-RED RATIO - the axis that actually separates an orange from a washed-out red.
    //
    // Both colours compared here have red as their dominant channel, so this ratio says how far
    // the colour has moved toward yellow WITHOUT being fooled by how light or dark it is. Hue and
    // saturation both were: see the tests above for what each of them let through.
    //
    // Red can never be zero for any colour in this palette (they are all reds, oranges and
    // ambers), but it is guarded anyway rather than trusting that to stay true.
    // ###########################################################################################
    private static double GreenOverRed(Color color)
    {
        return color.R == 0 ? 0.0 : color.G / (double)color.R;
    }
}
