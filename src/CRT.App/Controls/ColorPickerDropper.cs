using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.VisualTree;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE COLOR DROPPER IS PART OF EVERY COLOR PICKER'S LIST (owner request, 2026-11-03: "merge it with the
    // other color panel next to it? I am not sure if the color picker could be an 'internal' part when
    // clicking the color?" - it stood beside each picker until then, as large as the picker, and took the
    // room of the slider beside it). Avalonia's ColorPicker has no place for one, so it goes under what its
    // flyout holds: one rule for every picker, no markup at any of them.
    //
    // A flyout is a window of its own, so the dropper closes it first and picks from the window the picker
    // is in (ColorDropperButton.FlyoutToClose, Anchor).
    //
    // *** THE LIST IS BUILT WHEN IT FIRST OPENS (2026-11-13, measured). *** Avalonia's ColorPicker template holds a
    // whole copy of the ColorView inline in its flyout - three tabs, a spectrum, the palette's colors, sliders and
    // number boxes - built the moment the picker is, opened or not. Each picture card of "Board: Schematic images"
    // has four, and building them was the largest single cost of showing the page (~2 s for three cards on the
    // Open128-2 board). So App.axaml gives the picker Avalonia's own template less that copy - its flyout empty -
    // and the list is a ColorView made here as the flyout opens, looking as the copy did (NewColorView), the dropper
    // under it. A picker under Avalonia's own template (its copy in the flyout) gets the dropper under the copy.
    //
    // Registered in App's static constructor, as ToolTipPlacement is - a test's App has it too.
    // ###########################################################################################
    public static class ColorPickerDropper
    {
        public const string Label = "Pick a color";

        public static void UseInEveryColorPicker() =>
            TemplatedControl.TemplateAppliedEvent.AddClassHandler<ColorPicker>((picker, _) => ColorPickerDropper.Watch(picker));

        // The dropper in `picker`'s list - the list built as opening it builds it. Null when Avalonia's template no
        // longer has the place for it.
        internal static ColorDropperButton? Of(ColorPicker picker) =>
            ColorPickerDropper.FlyoutOf(picker) is { } flyout
                ? ColorPickerDropper.ListOf(picker, flyout).Children.OfType<ColorDropperButton>().FirstOrDefault()
                : null;

        // Whether `picker`'s list has been built - not before its flyout first opens.
        internal static bool IsListBuilt(ColorPicker picker) => ColorPickerDropper.FlyoutOf(picker)?.Content is not null;

        private static Flyout? FlyoutOf(ColorPicker picker) =>
            picker.GetVisualChildren().OfType<DropDownButton>().FirstOrDefault()?.Flyout as Flyout;

        // A template applied: its flyout builds its list as it opens. Each template has a flyout of its own.
        private static void Watch(ColorPicker picker)
        {
            if (ColorPickerDropper.FlyoutOf(picker) is { } flyout)
            {
                flyout.Opening += (_, _) => ColorPickerDropper.ListOf(picker, flyout);
            }
        }

        // What `flyout` lists: the colors - a ColorView made now, or the copy Avalonia's own template holds - and the
        // dropper under them. Made once; the same list every time after.
        private static StackPanel ListOf(ColorPicker picker, Flyout flyout)
        {
            if (flyout.Content is StackPanel built)
            {
                return built;
            }

            Control colors = flyout.Content as Control ?? ColorPickerDropper.NewColorView(picker);

            var dropper = new ColorDropperButton
            {
                Anchor = picker,
                FlyoutToClose = flyout,
                Content = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children =
                    {
                        new TextBlock { Text = ColorDropperButton.Icon, Classes = { "ColorDropperIcon" }, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = ColorPickerDropper.Label, VerticalAlignment = VerticalAlignment.Center },
                    },
                },
            };

            dropper.Classes.Add("InColorList");
            dropper.Picked += (_, color) => picker.SetCurrentValue(ColorView.ColorProperty, color);

            // A control has one parent: let go of as the flyout's content before it joins the panel.
            flyout.Content = null;

            var list = new StackPanel { Children = { colors, dropper } };
            flyout.Content = list;

            return list;
        }

        // ###########################################################################################
        // The picker's list as Avalonia's ColorPicker template draws it in its flyout - its copy of the ColorView, with
        // the flyout's own resources (Avalonia.Controls.ColorPicker's Themes/Fluent/ColorPicker.xaml): no ground and no
        // edge of its own, the tabs on the theme's base-low ground with rounded tops, following a theme switch as the
        // template's {StaticResource} in each theme dictionary did. The color is the picker's, both ways.
        // ###########################################################################################
        internal static ColorView NewColorView(ColorPicker picker)
        {
            var colors = new ColorView();

            // Every setting a picker can be given, FOLLOWED as the template's copy followed it - not copied once (code review,
            // 2026-11-13: a picker set up without its palette or alpha, or changed after its list was first made, was not
            // what its list showed). A ColorPicker IS a ColorView, so each is the same property on both.
            foreach (AvaloniaProperty setting in ColorPickerDropper.FollowedSettings)
            {
                colors.Bind(setting, picker.GetObservable(setting));
            }

            colors.Bind(ColorView.ColorModelProperty, new Binding(nameof(ColorView.ColorModel)) { Source = picker, Mode = BindingMode.TwoWay });

            var tabGround = new SolidColorBrush();
            tabGround.Bind(SolidColorBrush.ColorProperty, colors.GetResourceObservable("SystemControlBackgroundBaseLowBrush", brush => (brush as ISolidColorBrush)?.Color ?? Colors.Transparent));

            colors.Resources["ColorViewContentBackgroundBrush"] = Brushes.Transparent;
            colors.Resources["ColorViewContentBorderBrush"] = Brushes.Transparent;
            colors.Resources["ColorViewTabBorderBrush"] = Brushes.Transparent;
            colors.Resources["ColorViewTabBackgroundBrush"] = tabGround;
            colors.Resources["ColorViewTabBackgroundCornerRadius"] = new CornerRadius(5, 5, 0, 0);

            colors.Bind(ColorView.ColorProperty, new Binding(nameof(ColorView.Color)) { Source = picker, Mode = BindingMode.TwoWay });

            return colors;
        }

        internal static readonly AvaloniaProperty[] FollowedSettings =
        [
            ColorView.ColorSpectrumComponentsProperty,
            ColorView.ColorSpectrumShapeProperty,
            ColorView.HexInputAlphaPositionProperty,
            ColorView.IsAccentColorsVisibleProperty,
            ColorView.IsAlphaEnabledProperty,
            ColorView.IsAlphaVisibleProperty,
            ColorView.IsColorComponentsVisibleProperty,
            ColorView.IsColorModelVisibleProperty,
            ColorView.IsColorPaletteVisibleProperty,
            ColorView.IsColorPreviewVisibleProperty,
            ColorView.IsColorSpectrumVisibleProperty,
            ColorView.IsColorSpectrumSliderVisibleProperty,
            ColorView.IsComponentSliderVisibleProperty,
            ColorView.IsComponentTextInputVisibleProperty,
            ColorView.IsHexInputVisibleProperty,
            ColorView.MaxHueProperty,
            ColorView.MaxSaturationProperty,
            ColorView.MaxValueProperty,
            ColorView.MinHueProperty,
            ColorView.MinSaturationProperty,
            ColorView.MinValueProperty,
            ColorView.PaletteColorsProperty,
            ColorView.PaletteColumnCountProperty,
            ColorView.PaletteProperty,
        ];
    }
}
