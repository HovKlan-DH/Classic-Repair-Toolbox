using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ColorPickerDropper - every color picker's list, built as it first opens, with the dropper in it (owner requests,
// 2026-11-03; 2026-11-13, measured). Moved here from ColorDropperButtonTests on 2026-11-14 (code review: one test file
// per class).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ColorPickerDropperTests
{
    private static readonly Color Red = Color.FromRgb(255, 0, 0);

    // ###########################################################################################
    // THE DROPPER IS IN EVERY COLOR PICKER'S LIST (owner request, 2026-11-03: "merge it with the other color
    // panel next to it? I am not sure if the color picker could be an 'internal' part when clicking the
    // color?"): Avalonia's picker has no place for one, so a change of its template shows here first. It
    // closes the list and picks from the window the picker is in - a flyout has nothing to pick from.
    // ###########################################################################################
    [Fact]
    public void A_color_pickers_list_has_the_dropper_which_takes_a_color_from_its_window()
    {
        UiTest.Run(() =>
        {
            byte[] pixels = new byte[400 * 300 * 4];

            for (int at = 0; at < pixels.Length; at += 4)
            {
                pixels[at + 2] = Red.R;
                pixels[at + 3] = 255;
            }

            var picker = new ColorPicker { Color = Color.FromRgb(0, 0, 255) };
            var window = new Window { Content = picker, Width = 400, Height = 300 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ColorDropperButton dropper = Assert.IsType<ColorDropperButton>(ColorPickerDropper.Of(picker));
            dropper.PixelsOverrideForTests = () => new ScreenPixels(pixels, new PixelSize(400, 300), isBgra: true, 1.0);

            Assert.Same(picker, dropper.Anchor);

            dropper.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            // The cover is over the picker's window, its pointer taking the color.
            Assert.True(dropper.IsPicking);
            Assert.Same(window, TopLevel.GetTopLevel(dropper.CoverForTests));

            window.MouseDown(new Point(100, 100), MouseButton.Left);
            window.MouseUp(new Point(100, 100), MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.False(dropper.IsPicking);
            Assert.Equal(Red, picker.Color);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A PICKER'S LIST IS BUILT WHEN IT FIRST OPENS (2026-11-13, measured: Avalonia's template built a whole copy of
    // the ColorView in every picker's flyout, opened or not - the largest single cost of "Board: Schematic images"). ***
    // Shown, the picker holds its swatch and nothing more; opened by a press on it, its list is a ColorView with the
    // dropper under it, drawn as Avalonia's copy was - no ground of its own, the tabs on the theme's base-low ground with
    // rounded tops - and the color is the picker's both ways.
    // ###########################################################################################
    [Fact]
    public void A_color_pickers_list_is_built_when_it_first_opens_with_the_pickers_color_both_ways()
    {
        UiTest.Run(() =>
        {
            var picker = new ColorPicker { Color = Color.FromRgb(0, 0, 255) };
            var window = new Window { Content = picker, Width = 600, Height = 500 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Assert.False(ColorPickerDropper.IsListBuilt(picker));

            DropDownButton swatch = picker.GetVisualDescendants().OfType<DropDownButton>().Single();
            swatch.Flyout!.ShowAt(swatch);
            Dispatcher.UIThread.RunJobs();

            Assert.True(ColorPickerDropper.IsListBuilt(picker));

            var list = Assert.IsType<StackPanel>(((Flyout)swatch.Flyout).Content);
            ColorView colors = Assert.IsType<ColorView>(list.Children[0]);
            Assert.IsType<ColorDropperButton>(list.Children[1]);
            Assert.Same(list.Children[1], ColorPickerDropper.Of(picker));

            Assert.Equal(Color.FromRgb(0, 0, 255), colors.Color);
            colors.Color = Red;
            Assert.Equal(Red, picker.Color);
            picker.Color = Color.FromRgb(0, 128, 0);
            Assert.Equal(Color.FromRgb(0, 128, 0), colors.Color);

            Assert.Equal(Brushes.Transparent, colors.Resources["ColorViewContentBackgroundBrush"]);
            Assert.Equal(new CornerRadius(5, 5, 0, 0), colors.Resources["ColorViewTabBackgroundCornerRadius"]);

            var tabGround = Assert.IsAssignableFrom<ISolidColorBrush>(colors.Resources["ColorViewTabBackgroundBrush"]);
            var baseLow = (ISolidColorBrush)picker.FindResource(picker.ActualThemeVariant, "SystemControlBackgroundBaseLowBrush")!;
            Assert.Equal(baseLow.Color, tabGround.Color);

            // Opened again: the same list.
            swatch.Flyout.Hide();
            swatch.Flyout.ShowAt(swatch);
            Dispatcher.UIThread.RunJobs();
            Assert.Same(list, ((Flyout)swatch.Flyout).Content);

            swatch.Flyout.Hide();
            window.Close();
        });
    }

    // The list follows how its picker is set up, as Avalonia's copy did - then and later (code review, 2026-11-13: three
    // settings copied once, so a picker without its palette or its alpha, or changed later, showed them anyway).
    [Fact]
    public void A_color_pickers_list_follows_how_the_picker_is_set_up()
    {
        UiTest.Run(() =>
        {
            var picker = new ColorPicker { IsColorPaletteVisible = false, IsAlphaVisible = false };
            ColorView colors = ColorPickerDropper.NewColorView(picker);

            Assert.False(colors.IsColorPaletteVisible);
            Assert.False(colors.IsAlphaVisible);

            picker.IsAlphaVisible = true;
            picker.IsHexInputVisible = false;
            picker.ColorModel = ColorModel.Hsva;

            Assert.True(colors.IsAlphaVisible);
            Assert.False(colors.IsHexInputVisible);
            Assert.Equal(ColorModel.Hsva, colors.ColorModel);

            colors.ColorModel = ColorModel.Rgba;
            Assert.Equal(ColorModel.Rgba, picker.ColorModel);
        });
    }

    // ###########################################################################################
    // Every setting a ColorView declares is followed, or left out here on purpose (code review, 2026-11-13: a hand-kept
    // list - a setting Avalonia adds would be ignored by the list without a word). The color and its model are bound both
    // ways of their own; the HSV color follows the color; which tab is picked is the list's own.
    // ###########################################################################################
    [Fact]
    public void Every_setting_a_color_view_declares_is_followed_or_left_out_on_purpose()
    {
        string[] boundOrLeftOut =
        [
            nameof(ColorView.Color), nameof(ColorView.ColorModel), nameof(ColorView.HsvColor), nameof(ColorView.SelectedIndex),
        ];

        string[] declared = AvaloniaPropertyRegistry.Instance.GetRegistered(typeof(ColorView))
            .Where(property => property.OwnerType == typeof(ColorView))
            .Select(property => property.Name)
            .Except(boundOrLeftOut)
            .Order(StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(declared, ColorPickerDropper.FollowedSettings.Select(property => property.Name).Order(StringComparer.Ordinal).ToArray());
    }
}
