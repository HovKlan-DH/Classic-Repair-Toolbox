using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.VisualTree;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE COLOR DROPPER IS PART OF EVERY COLOR PICKER'S LIST (owner request, 2026-11-03: "merge it with the
    // other color panel next to it? I am not sure if the color picker could be an 'internal' part when
    // clicking the color?" - it stood beside each picker until then, as large as the picker, and took the
    // room of the slider beside it). Avalonia's ColorPicker has no place for one, so it goes under what its
    // flyout holds, once its template is applied: one rule for every picker, no markup at any of them.
    //
    // A flyout is a window of its own, so the dropper closes it first and picks from the window the picker
    // is in (ColorDropperButton.FlyoutToClose, Anchor).
    //
    // Registered in App's static constructor, as ToolTipPlacement is - a test's App has it too.
    // ###########################################################################################
    public static class ColorPickerDropper
    {
        public const string Label = "Pick a color";

        public static void UseInEveryColorPicker() =>
            TemplatedControl.TemplateAppliedEvent.AddClassHandler<ColorPicker>((picker, _) => ColorPickerDropper.Add(picker));

        // The dropper in `picker`'s list - null when Avalonia's template no longer has the place for it.
        internal static ColorDropperButton? Of(ColorPicker picker) =>
            ColorPickerDropper.FlyoutOf(picker)?.Content is StackPanel panel ? panel.Children.OfType<ColorDropperButton>().FirstOrDefault() : null;

        private static Flyout? FlyoutOf(ColorPicker picker) =>
            picker.GetVisualChildren().OfType<DropDownButton>().FirstOrDefault()?.Flyout as Flyout;

        private static void Add(ColorPicker picker)
        {
            if (ColorPickerDropper.FlyoutOf(picker) is not { Content: Control colors } flyout || colors is StackPanel)
            {
                return;
            }

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
            flyout.Content = new StackPanel { Children = { colors, dropper } };
        }
    }
}
