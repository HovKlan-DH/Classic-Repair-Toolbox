using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// UserThemeShades - a button's colors under the pointer and pressed in the "User preference" theme (code review,
// 2026-11-13): the light theme's fixed shades were used whatever the user's own button colors were, so a dark button with
// white words turned light grey under the pointer, 1.7:1.
// ###########################################################################################
public sealed class UserThemeShadesTests
{
    private static Dictionary<string, string> ColorsWith(params (string Key, string Value)[] changed)
    {
        var colors = new Dictionary<string, string>(UserSettings.UserThemeDefaultColors, StringComparer.Ordinal);

        foreach ((string key, string value) in changed)
        {
            colors[key] = value;
        }

        return colors;
    }

    private static double Contrast(Color one, Color other)
    {
        double a = UserThemeShades.Luminance(one);
        double b = UserThemeShades.Luminance(other);
        return (Math.Max(a, b) + 0.05) / (Math.Min(a, b) + 0.05);
    }

    // Colors left at their defaults keep the light theme's own shades - the ones ThemeContrastTests holds.
    [Fact]
    public void Button_colors_left_at_their_defaults_keep_the_light_themes_shades()
    {
        Assert.Empty(UserThemeShades.Of(ColorsWith(), UserSettings.UserThemeDefaultColors));
    }

    // ###########################################################################################
    // The reported case: a dark button with white words. Too dark for a deeper shade to show, its shades go LIGHTER - as
    // the dark theme's do - its white words still reading well on them; and the edge goes the way its ground goes.
    // ###########################################################################################
    [Fact]
    public void A_dark_button_with_white_words_gets_lighter_shades_its_words_still_read_on()
    {
        Dictionary<string, Color> shades = UserThemeShades.Of(
            ColorsWith(("Button_Bg", "#333333"), ("Button_Fg", "White")), UserSettings.UserThemeDefaultColors);

        Color resting = Color.Parse("#333333");
        Assert.Equal(["Button_Hover_Bg", "Button_Hover_Border", "Button_Pressed_Bg"], shades.Keys.Order(StringComparer.Ordinal).ToArray());
        Assert.True(UserThemeShades.Luminance(shades["Button_Hover_Bg"]) > UserThemeShades.Luminance(resting));
        Assert.True(UserThemeShades.Luminance(shades["Button_Pressed_Bg"]) > UserThemeShades.Luminance(shades["Button_Hover_Bg"]));
        Assert.True(UserThemeShades.Luminance(shades["Button_Hover_Border"]) > UserThemeShades.Luminance(Color.Parse("DarkGray")));
        Assert.All(
            new[] { shades["Button_Hover_Bg"], shades["Button_Pressed_Bg"] },
            shade => Assert.True(Contrast(shade, Colors.White) >= 4.5, $"{shade} against white"));
    }

    // A white button with black words: no lighter shade exists, so they go DEEPER - something shows under the pointer and
    // pressed (code review, 2026-11-13: lighter, it stayed white).
    [Fact]
    public void A_white_button_with_black_words_gets_deeper_shades_that_show()
    {
        Dictionary<string, Color> shades = UserThemeShades.Of(ColorsWith(("Button_Bg", "#FFFFFF")), UserSettings.UserThemeDefaultColors);

        Assert.NotEqual(Colors.White, shades["Button_Hover_Bg"]);
        Assert.True(UserThemeShades.Luminance(shades["Button_Pressed_Bg"]) < UserThemeShades.Luminance(shades["Button_Hover_Bg"]));
        Assert.True(Contrast(shades["Button_Pressed_Bg"], Colors.Black) >= 4.5);
    }

    // A light green button the user gave dark words: its shades go LIGHTER, away from its words - the light theme's own
    // fixed dark green under the pointer read at under 3:1 with them.
    [Fact]
    public void A_light_button_with_dark_words_gets_lighter_shades()
    {
        Dictionary<string, Color> shades = UserThemeShades.Of(
            ColorsWith(("Button_Ok_Bg", "#A0E0A0"), ("Button_Ok_Fg", "#202020")), UserSettings.UserThemeDefaultColors);

        Color resting = Color.Parse("#A0E0A0");
        Color words = Color.Parse("#202020");
        Assert.True(UserThemeShades.Luminance(shades["Button_Ok_Hover_Bg"]) > UserThemeShades.Luminance(resting));
        Assert.True(UserThemeShades.Luminance(shades["Button_Ok_Pressed_Bg"]) > UserThemeShades.Luminance(shades["Button_Ok_Hover_Bg"]));
        Assert.True(Contrast(shades["Button_Ok_Pressed_Bg"], words) >= Contrast(resting, words));
        Assert.False(shades.ContainsKey("Button_Hover_Bg"));
        Assert.False(shades.ContainsKey("Button_Cancel_Hover_Bg"));
    }

    // White words on a light red the user picked: deeper, as the light theme's own red goes.
    [Fact]
    public void A_light_button_with_white_words_gets_deeper_shades()
    {
        Dictionary<string, Color> shades = UserThemeShades.Of(ColorsWith(("Button_Cancel_Bg", "#F0B0B0")), UserSettings.UserThemeDefaultColors);

        Assert.True(UserThemeShades.Luminance(shades["Button_Cancel_Hover_Bg"]) < UserThemeShades.Luminance(Color.Parse("#F0B0B0")));
        Assert.True(UserThemeShades.Luminance(shades["Button_Cancel_Pressed_Bg"]) < UserThemeShades.Luminance(shades["Button_Cancel_Hover_Bg"]));
    }

    // ###########################################################################################
    // The steps are the light theme's own (App.axaml): from its own button colors, towards black as it goes, they land
    // on its own shades - within a few levels per channel, as those were picked by eye. Every step is here.
    // ###########################################################################################
    [Fact]
    public void Each_step_from_the_light_themes_color_lands_near_its_own_shade()
    {
        var light = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Button_Bg"] = "LightGray", ["Button_Border"] = "DarkGray", ["Button_Ok_Bg"] = "#497F54", ["Button_Cancel_Bg"] = "#BC5454",
            ["Button_Hover_Bg"] = "#C4C4C4", ["Button_Pressed_Bg"] = "#B3B3B3", ["Button_Hover_Border"] = "#8A8A8A",
            ["Button_Ok_Hover_Bg"] = "#3F7049", ["Button_Ok_Pressed_Bg"] = "#355F3D",
            ["Button_Cancel_Hover_Bg"] = "#A84848", ["Button_Cancel_Pressed_Bg"] = "#943E3E",
        };

        foreach ((string shade, string from, _, _, double amount) in UserThemeShades.Steps)
        {
            Color worked = UserThemeShades.Moved(Color.Parse(light[from]), amount, towardsWhite: false);
            Color own = Color.Parse(light[shade]);

            Assert.True(
                Math.Abs(worked.R - own.R) <= 4 && Math.Abs(worked.G - own.G) <= 4 && Math.Abs(worked.B - own.B) <= 4,
                $"{shade}: {worked} worked out, the light theme's {own}");
        }
    }

    // A shade typed into the settings file by hand is the user's, and a color CRT cannot read gives nothing.
    [Fact]
    public void A_shade_typed_by_hand_stays_and_an_unreadable_color_gives_none()
    {
        Dictionary<string, Color> shades = UserThemeShades.Of(
            ColorsWith(("Button_Bg", "#333333"), ("Button_Fg", "White"), ("Button_Hover_Bg", "#123456"), ("Button_Ok_Bg", "not a color")),
            UserSettings.UserThemeDefaultColors);

        Assert.Equal(["Button_Hover_Border", "Button_Pressed_Bg"], shades.Keys.Order(StringComparer.Ordinal).ToArray());
    }
}
