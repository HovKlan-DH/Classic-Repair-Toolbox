using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A BUTTON KEEPS ITS COLOR UNDER THE POINTER AND WHILE PRESSED, A SHADE DEEPER (App.axaml's "Button:pointerover" styles;
// 2026-11-12, the theme color review). Fluent's own looks replaced the fill: a green or red button went grey under the pointer
// with black words, and a plain one turned lighter in the light theme and darker than at rest in the dark one. Read off the
// template's presenter of a SHOWN window, with the state forced, in each theme - a Fluent update that wins over these styles
// fails here.
// ###########################################################################################
[Collection("HeadlessUi")]
public class ButtonStatesTests
{
    private static readonly ThemeVariant[] Themes = [ThemeVariant.Light, ThemeVariant.Dark];

    private static Color ThemeColour(string key, ThemeVariant variant)
    {
        Assert.True(Application.Current!.TryGetResource(key, variant, out object? resource));
        return Assert.IsAssignableFrom<ISolidColorBrush>(resource).Color;
    }

    // The button shown in a window of `variant`, `state` (":pointerover", ":pressed") on.
    private static (ContentPresenter Presenter, Window Window) Shown(Button button, ThemeVariant variant, string state)
    {
        var window = new Window { Content = button, RequestedThemeVariant = variant };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        ((IPseudoClasses)button.Classes).Set(state, true);
        Dispatcher.UIThread.RunJobs();

        ContentPresenter presenter = button.GetVisualDescendants().OfType<ContentPresenter>().First(p => p.Name == "PART_ContentPresenter");
        return (presenter, window);
    }

    private static Color ColourOf(IBrush? brush) => Assert.IsAssignableFrom<ISolidColorBrush>(brush).Color;

    [Theory]
    [InlineData("Accept", ":pointerover", "Button_Ok_Hover_Bg", "Button_Ok_Fg")]
    [InlineData("Accept", ":pressed", "Button_Ok_Pressed_Bg", "Button_Ok_Fg")]
    [InlineData("Danger", ":pointerover", "Button_Cancel_Hover_Bg", "Button_Cancel_Fg")]
    [InlineData("Danger", ":pressed", "Button_Cancel_Pressed_Bg", "Button_Cancel_Fg")]
    public void A_green_or_red_button_keeps_its_colour_and_its_words_colour_under_the_pointer(string kind, string state, string fill, string words)
    {
        UiTest.Run(() =>
        {
            foreach (ThemeVariant variant in ButtonStatesTests.Themes)
            {
                // Its words as plain content, and as a TextBlock of their own (App.axaml's "Button.Accept > TextBlock").
                var plain = new Button { Content = "Save", Classes = { kind } };
                (ContentPresenter presenter, Window window) = ButtonStatesTests.Shown(plain, variant, state);

                Assert.Equal(ButtonStatesTests.ThemeColour(fill, variant), ButtonStatesTests.ColourOf(presenter.Background));
                Assert.Equal(ButtonStatesTests.ThemeColour(words, variant), ButtonStatesTests.ColourOf(presenter.Foreground));
                window.Close();

                var label = new TextBlock { Text = "Replace all" };
                var labelled = new Button { Content = label, Classes = { kind } };
                (ContentPresenter _, Window second) = ButtonStatesTests.Shown(labelled, variant, state);

                Assert.Equal(ButtonStatesTests.ThemeColour(words, variant), ButtonStatesTests.ColourOf(label.Foreground));
                second.Close();
            }
        });
    }

    [Theory]
    [InlineData(":pointerover", "Button_Hover_Bg")]
    [InlineData(":pressed", "Button_Pressed_Bg")]
    public void A_plain_button_turns_a_shade_deeper_under_the_pointer_in_both_themes(string state, string fill)
    {
        UiTest.Run(() =>
        {
            foreach (ThemeVariant variant in ButtonStatesTests.Themes)
            {
                (ContentPresenter presenter, Window window) = ButtonStatesTests.Shown(new Button { Content = "Browse" }, variant, state);

                Assert.Equal(ButtonStatesTests.ThemeColour(fill, variant), ButtonStatesTests.ColourOf(presenter.Background));
                Assert.Equal(ButtonStatesTests.ThemeColour("Button_Fg", variant), ButtonStatesTests.ColourOf(presenter.Foreground));
                Assert.NotEqual(ButtonStatesTests.ThemeColour("Button_Bg", variant), ButtonStatesTests.ThemeColour(fill, variant));

                window.Close();
            }
        });
    }

    // A toggle button and a drop-down button look as a button does, at rest and under the pointer - a style names an exact
    // type, so "Button" never reached them (Fluent's half-transparent grey until 2026-11-12).
    [Fact]
    public void A_toggle_or_drop_down_button_looks_as_a_button_does()
    {
        UiTest.Run(() =>
        {
            foreach (ThemeVariant variant in ButtonStatesTests.Themes)
            {
                var toggle = new ToggleButton { Content = "U8" };
                var dropDown = new DropDownButton { Content = "Columns: All" };
                var window = new Window { Content = new StackPanel { Children = { toggle, dropDown } }, RequestedThemeVariant = variant };
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(ButtonStatesTests.ThemeColour("Button_Bg", variant), ButtonStatesTests.ColourOf(toggle.Background));
                Assert.Equal(ButtonStatesTests.ThemeColour("Button_Bg", variant), ButtonStatesTests.ColourOf(dropDown.Background));

                ((IPseudoClasses)toggle.Classes).Set(":pointerover", true);
                ((IPseudoClasses)dropDown.Classes).Set(":pointerover", true);
                Dispatcher.UIThread.RunJobs();

                ContentPresenter togglePresenter = toggle.GetVisualDescendants().OfType<ContentPresenter>().First(p => p.Name == "PART_ContentPresenter");
                Border dropDownRoot = dropDown.GetVisualDescendants().OfType<Border>().First(b => b.Name == "RootBorder");

                Assert.Equal(ButtonStatesTests.ThemeColour("Button_Hover_Bg", variant), ButtonStatesTests.ColourOf(togglePresenter.Background));
                Assert.Equal(ButtonStatesTests.ThemeColour("Button_Hover_Bg", variant), ButtonStatesTests.ColourOf(dropDownRoot.Background));

                window.Close();
            }
        });
    }

    // ###########################################################################################
    // A note box's format button that is "On" (RichNoteTheme.axaml's ".On" - a press would turn its style off) keeps that
    // look under the pointer and pressed (code review, 2026-11-13): App.axaml's styles outrank a control theme's, so the
    // plain button's grey under the pointer took it away just as it was looked at - what 2026-11-08's review had fixed.
    // ###########################################################################################
    [Theory]
    [InlineData(":pointerover")]
    [InlineData(":pressed")]
    public void A_note_boxs_format_button_turned_on_keeps_that_look_under_the_pointer(string state)
    {
        UiTest.Run(() =>
        {
            Assert.True(Application.Current!.TryGetResource("RichNoteFormatButton", ThemeVariant.Default, out object? theme));

            foreach (ThemeVariant variant in ButtonStatesTests.Themes)
            {
                var bold = new Button { Content = "B", Theme = Assert.IsType<ControlTheme>(theme), Classes = { "FormatTool", "On" } };
                (ContentPresenter presenter, Window window) = ButtonStatesTests.Shown(bold, variant, state);

                Assert.Equal(ButtonStatesTests.ThemeColour("Mode_Selected_Bg", variant), ButtonStatesTests.ColourOf(presenter.Background));
                Assert.Equal(ButtonStatesTests.ThemeColour("Mode_Selected_Border", variant), ButtonStatesTests.ColourOf(presenter.BorderBrush));

                window.Close();
            }
        });
    }

    // ###########################################################################################
    // The region button lit keeps its region's words under the pointer (code review, 2026-11-13): the plain button's hover
    // style set its words to Button_Fg, which outranked the lit button's own - the dark theme's NTSC read 1.2:1 while the
    // pointer was still on it after the click.
    // ###########################################################################################
    [Theory]
    [InlineData("PalRegionButton", "Schematics_Region_PAL_Fg")]
    [InlineData("NtscRegionButton", "Schematics_Region_NTSC_Fg")]
    public void A_lit_region_button_keeps_its_regions_words_under_the_pointer(string name, string words)
    {
        UiTest.Run(() =>
        {
            foreach (ThemeVariant variant in ButtonStatesTests.Themes)
            {
                var region = new Button { Name = name, Content = "PAL", Classes = { "RegionToggle", "active" } };
                (ContentPresenter presenter, Window window) = ButtonStatesTests.Shown(region, variant, ":pointerover");

                Assert.Equal(ButtonStatesTests.ThemeColour(words, variant), ButtonStatesTests.ColourOf(presenter.Foreground));

                window.Close();
            }
        });
    }

    // An icon or a link - a button with no fill of its own - gets no box of CRT's under the pointer, set in markup or code.
    [Fact]
    public void A_transparent_button_gets_no_fill_of_its_own_under_the_pointer()
    {
        UiTest.Run(() =>
        {
            foreach (ThemeVariant variant in ButtonStatesTests.Themes)
            {
                (ContentPresenter presenter, Window window) = ButtonStatesTests.Shown(new Button { Content = "x", Background = Brushes.Transparent }, variant, ":pointerover");

                Assert.NotEqual(ButtonStatesTests.ThemeColour("Button_Hover_Bg", variant), ButtonStatesTests.ColourOf(presenter.Background));

                window.Close();
            }
        });
    }
}
