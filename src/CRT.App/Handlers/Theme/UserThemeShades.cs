using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Theming
{
    // ###########################################################################################
    // A BUTTON'S SHADES UNDER THE POINTER AND PRESSED, IN THE "User preference" THEME (code review, 2026-11-13). The
    // shades are no colors the theme seeds (UserSettings.UserThemeDefaultColors has none of them), so the theme took the
    // light theme's fixed greys and greens whatever the button's own colors were: a Button_Bg of #333333 with white words
    // turned #C4C4C4 under the pointer, 1.7:1. So for a button whose color or words' color was CHANGED from its default,
    // each shade is worked out from its own colors - a shade further from its words' color where there is room for one to
    // show (GoesTowardsWhite) - and a button left at its defaults keeps the light theme's own shades, which
    // ThemeContrastTests holds. A shade typed into the settings file by hand is the user's and stays.
    // ###########################################################################################
    internal static class UserThemeShades
    {
        // Each shade: the color it is worked out from, the button's ground and words (which way it goes) and how far - the
        // light theme's own steps (LightGray #D3D3D3 to #C4C4C4 and #B3B3B3, DarkGray #A9A9A9 to #8A8A8A, #497F54 to
        // #3F7049 and #355F3D, #BC5454 to #A84848 and #943E3E), rounded - those were picked by eye, a few levels apart from
        // one rule. The edge goes the way its button's ground goes.
        internal static readonly (string Shade, string From, string Ground, string Words, double Amount)[] Steps =
        [
            ("Button_Hover_Bg", "Button_Bg", "Button_Bg", "Button_Fg", 0.07),
            ("Button_Pressed_Bg", "Button_Bg", "Button_Bg", "Button_Fg", 0.15),
            ("Button_Hover_Border", "Button_Border", "Button_Bg", "Button_Fg", 0.185),
            ("Button_Ok_Hover_Bg", "Button_Ok_Bg", "Button_Ok_Bg", "Button_Ok_Fg", 0.13),
            ("Button_Ok_Pressed_Bg", "Button_Ok_Bg", "Button_Ok_Bg", "Button_Ok_Fg", 0.26),
            ("Button_Cancel_Hover_Bg", "Button_Cancel_Bg", "Button_Cancel_Bg", "Button_Cancel_Fg", 0.11),
            ("Button_Cancel_Pressed_Bg", "Button_Cancel_Bg", "Button_Cancel_Bg", "Button_Cancel_Fg", 0.21),
        ];

        // Every key this may set - what applying another theme takes away again.
        public static IEnumerable<string> Keys => UserThemeShades.Steps.Select(step => step.Shade);

        // ###########################################################################################
        // The shades to set over the light theme's, for the user's `colors` against the theme's `defaults`.
        // ###########################################################################################
        public static Dictionary<string, Color> Of(IReadOnlyDictionary<string, string> colors, IReadOnlyDictionary<string, string> defaults)
        {
            ArgumentNullException.ThrowIfNull(colors);
            ArgumentNullException.ThrowIfNull(defaults);

            Color? Read(string key) => colors.TryGetValue(key, out string? value) && Color.TryParse(value?.Trim(), out Color color) ? color : null;

            bool AtDefault(string key) =>
                Read(key) is { } color && defaults.TryGetValue(key, out string? byDefault) && Color.TryParse(byDefault, out Color defaultColor) && defaultColor == color;

            var shades = new Dictionary<string, Color>(StringComparer.Ordinal);

            foreach ((string shade, string from, string ground, string words, double amount) in UserThemeShades.Steps)
            {
                if (colors.ContainsKey(shade) ||
                    Read(from) is not { } color || Read(ground) is not { } groundColor || Read(words) is not { } wordsColor ||
                    (AtDefault(from) && AtDefault(ground) && AtDefault(words)))
                {
                    continue;
                }

                shades[shade] = UserThemeShades.Moved(color, amount, UserThemeShades.GoesTowardsWhite(groundColor, wordsColor));
            }

            return shades;
        }

        // ###########################################################################################
        // Which way a button's shades go: away from its words, so they read at least as well - unless its ground is already
        // within a quarter of the range of that end, where a step shows nothing (white words on #333333, black on white: no
        // change under the pointer at all - code review, 2026-11-13); then towards them, which leaves them reading well, as
        // the ground is that far from them. One way for every shade of a button, hover and pressed alike.
        // ###########################################################################################
        internal static bool GoesTowardsWhite(Color ground, Color words)
        {
            bool awayIsWhite = UserThemeShades.Luminance(words) < UserThemeShades.Luminance(ground);
            int roomThatWay = awayIsWhite
                ? 255 - Math.Min(ground.R, Math.Min(ground.G, ground.B))
                : Math.Max(ground.R, Math.Max(ground.G, ground.B));

            return roomThatWay >= 64 ? awayIsWhite : !awayIsWhite;
        }

        // ###########################################################################################
        // `color` moved `amount` of the way towards white or black. Its alpha stays.
        // ###########################################################################################
        internal static Color Moved(Color color, double amount, bool towardsWhite)
        {
            double target = towardsWhite ? 255 : 0;

            byte Move(byte channel) => (byte)Math.Round(channel + ((target - channel) * amount), MidpointRounding.AwayFromZero);

            return Color.FromArgb(color.A, Move(color.R), Move(color.G), Move(color.B));
        }

        // WCAG's relative luminance - CRT.Data's one copy of the formula.
        internal static double Luminance(Color color) => Handlers.DataHandling.RegionColors.Luminance(color.R, color.G, color.B);
    }
}
