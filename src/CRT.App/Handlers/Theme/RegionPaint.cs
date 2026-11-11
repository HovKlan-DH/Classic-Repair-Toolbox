using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;

namespace Handlers.Theming
{
    // ###########################################################################################
    // A REGION DRAWN IN THE BOARD'S OWN COLORS, WHERE IT HAS ANY (owner request, 2026-11-11; CRT.Data's RegionColors) - the
    // buttons under the component list and on the popup, the region's name over the Schematics tab's picture and the popup's,
    // and "Board: Regions" itself. A region with none keeps the colors CRT gives it: PAL's and NTSC's resources (which the
    // "Configuration" tab can set), a board's own region one of four (RegionChoice).
    //
    // A BUTTON is lit by App.axaml's styles: class "BoardColors" reads its colors from the button's own resources
    // (BackgroundKey, TextKey), declared after the per-region styles so it wins over them.
    // ###########################################################################################
    internal static class RegionPaint
    {
        public const string BoardColorsClass = "BoardColors";

        public const string BackgroundKey = "Region_Board_Bg";

        public const string TextKey = "Region_Board_Fg";

        // `button` lit in `colors` when picked - or, with none, in the colors CRT gives its region.
        public static void UseBoardColors(Button button, (string Background, string Text)? colors)
        {
            ArgumentNullException.ThrowIfNull(button);

            // A button that has none and had none is left untouched: asking for its Resources makes it a dictionary of its own,
            // and the change of resources that follows reaches every control in it.
            if (colors is null && !button.Classes.Contains(RegionPaint.BoardColorsClass))
            {
                return;
            }

            button.Classes.Set(RegionPaint.BoardColorsClass, colors is not null);

            if (colors is { } own)
            {
                // Set only where it changes: the buttons are laid out again at every read of the lists, and each resource set
                // tells every control in the button that its resources changed.
                RegionPaint.SetBrush(button, RegionPaint.BackgroundKey, own.Background);
                RegionPaint.SetBrush(button, RegionPaint.TextKey, own.Text);
            }
            else
            {
                button.Resources.Remove(RegionPaint.BackgroundKey);
                button.Resources.Remove(RegionPaint.TextKey);
            }
        }

        // ###########################################################################################
        // A region's name over a picture (`border` and its `text`): in `colors` when the board gives the region its own, else
        // in the resources `prefix` names (RegionChoice.ColorsOf - "Schematics_Region_PAL" with "_Bg", "_Border", "_Fg"),
        // followed as the theme changes. The frame stays the region's usual one either way.
        // ###########################################################################################
        public static void PaintLabel(Control host, Border border, TextBlock text, string prefix, (string Background, string Text)? colors)
        {
            ArgumentNullException.ThrowIfNull(host);
            ArgumentNullException.ThrowIfNull(border);
            ArgumentNullException.ThrowIfNull(text);

            border.Bind(Border.BorderBrushProperty, host.GetResourceObservable($"{prefix}_Border"));

            if (colors is { } own)
            {
                border.Bind(Border.BackgroundProperty, new Fixed(RegionPaint.Brush(own.Background)));
                text.Bind(TextBlock.ForegroundProperty, new Fixed(RegionPaint.Brush(own.Text)));
                return;
            }

            border.Bind(Border.BackgroundProperty, host.GetResourceObservable($"{prefix}_Bg"));
            text.Bind(TextBlock.ForegroundProperty, host.GetResourceObservable($"{prefix}_Fg"));
        }

        private static void SetBrush(Button button, string key, string? color)
        {
            IBrush wanted = RegionPaint.Brush(color);

            if (button.Resources.TryGetValue(key, out object? held) &&
                held is ISolidColorBrush shown &&
                wanted is ISolidColorBrush solid &&
                shown.Color == solid.Color)
            {
                return;
            }

            button.Resources[key] = wanted;
        }

        // A "#RRGGBB" the board gives (RegionColors keeps nothing else) - black for anything Avalonia cannot read anyway.
        public static IBrush Brush(string? color) =>
            new SolidColorBrush(Color.TryParse(color, out Color parsed) ? parsed : Colors.Black);

        // One value, at once, for a Bind beside the resource bindings it replaces.
        private sealed class Fixed : IObservable<object?>
        {
            private readonly object? thisValue;

            public Fixed(object? value) => this.thisValue = value;

            public IDisposable Subscribe(IObserver<object?> observer)
            {
                observer.OnNext(this.thisValue);
                return Nothing.Instance;
            }

            private sealed class Nothing : IDisposable
            {
                public static readonly Nothing Instance = new();

                public void Dispose()
                {
                }
            }
        }
    }
}
