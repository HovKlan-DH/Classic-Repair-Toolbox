using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // THE THEME'S ICON IN THE WINDOW'S TOP RIGHT CORNER (owner request, 2026-11-17: "I would like to get a half-moon and a sun
    // icon somewhere, so the user easily can toggle this from a permanent, always-visible location. Maybe that could be in the
    // top-right corner, just to the right of the "update data" icon?"). ONE ICON, as the data icon beside it (owner requests,
    // the same day: "same size as the "update data" icon", then of a joined sun | moon switch, "I am not satisfied how it
    // looks. Can you try something else, so it looks better?"): the moon while the light theme is on, the sun while the dark
    // is - what a click switches to, its tooltip saying so. THE ONLY PLACE THE THEME IS PICKED: the "Configuration" tab's
    // drop-down went the same day (owner request: "you can entirely remove the "Appearance" with the theme").
    // ###########################################################################################
    public partial class Main
    {
        // Font Awesome's moon and sun.
        internal const string MoonGlyph = "";
        internal const string SunGlyph = "";

        internal const string ToDarkTip = "Switch to the dark theme";
        internal const string ToLightTip = "Switch to the light theme";

        // Called instead of App.ApplyConfiguredTheme when set - a test must never change the Application the whole headless
        // run shares (CLAUDE.md, "Headless UI tests"): a test turning it dark made a later test comparing a drawn color with
        // the light theme's fail on CI at random (2026-10-04).
        internal Action? ApplyThemeOverrideForTests { get; set; }

        private void WireThemeToggle() => this.ShowThemeToggle();

        private void OnThemeToggleClick(object? sender, RoutedEventArgs e) =>
            this.PickTheme(UserSettings.ThemeVariant == "Dark" ? "Light" : "Dark");

        // The theme picked: stored, put on CRT, and the icon showing the other.
        internal void PickTheme(string theme)
        {
            if (string.Equals(UserSettings.ThemeVariant, theme, StringComparison.Ordinal))
            {
                return;
            }

            UserSettings.ThemeVariant = theme;

            if (this.ApplyThemeOverrideForTests is Action apply)
            {
                apply();
            }
            else if (Application.Current is App app)
            {
                app.ApplyConfiguredTheme();
            }

            this.ShowThemeToggle();
        }

        // The moon in the light theme - a stored theme CRT does not know reads as light, as App.ApplyConfiguredTheme takes
        // it - and the sun in the dark.
        private void ShowThemeToggle()
        {
            bool dark = UserSettings.ThemeVariant == "Dark";

            this.ThemeToggleGlyph.Text = dark ? SunGlyph : MoonGlyph;
            ToolTip.SetTip(this.ThemeToggle, dark ? ToLightTip : ToDarkTip);
        }
    }
}
