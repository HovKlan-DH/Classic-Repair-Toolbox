using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A COMPONENT'S CASE DRAWN FROM ABOVE (CRT 4.0 step 2, 2026-10-25 - the plan's Pinout: "CRT draws
    // the top view, with the notch and pin 1"). The component pane hands it the package and the pins;
    // every position is PinoutGeometry's, so this only paints. A case CRT cannot draw takes no room.
    // ###########################################################################################
    public sealed class PinoutDrawing : Control
    {
        private const double TextSize = 11;

        private PinoutLayout? thisLayout;
        private Dictionary<int, string> thisNames = new();

        // The case and its pins, drawn again.
        public void Show(string? package, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            this.thisLayout = PinoutGeometry.For(package, pins.Select(pin => pin.Pin));
            this.thisNames = new Dictionary<int, string>();

            foreach (ComponentPinEntry pin in pins)
            {
                if (int.TryParse(pin.Pin?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int number) &&
                    !this.thisNames.ContainsKey(number))
                {
                    this.thisNames[number] = pin.Name?.Trim() ?? string.Empty;
                }
            }

            this.IsVisible = this.thisLayout is not null;
            this.InvalidateMeasure();
            this.InvalidateVisual();
        }

        // Whether there is anything to draw - for the tests.
        internal PinoutLayout? LayoutForTests => this.thisLayout;

        protected override Size MeasureOverride(Size availableSize) => this.thisLayout?.Size ?? default;

        public override void Render(DrawingContext context)
        {
            if (this.thisLayout is not { } layout)
            {
                return;
            }

            IBrush foreground = this.Brush("Fg", Brushes.Black);
            IBrush body = this.Brush("Card_Bg", Brushes.White);
            IBrush firstPin = this.Brush("Schematics_FirstPin", Brushes.IndianRed);
            var outline = new Pen(foreground, 1);

            context.DrawRectangle(body, outline, layout.Body, 3, 3);

            if (layout.IsDualInLine)
            {
                // The notch - half a circle cut into the top edge.
                var notch = new EllipseGeometry(new Rect(layout.Notch.X - 7, layout.Notch.Y - 7, 14, 14));
                using (context.PushClip(layout.Body))
                {
                    context.DrawGeometry(this.Brush("Bg", Brushes.White), outline, notch);
                }
            }

            context.DrawEllipse(firstPin, null, layout.FirstPinDot, 3, 3);

            foreach (PinoutPin pin in layout.Pins)
            {
                context.DrawRectangle(foreground, null, pin.Lead);

                string number = pin.Number.ToString(CultureInfo.InvariantCulture);
                FormattedText numberText = PinoutDrawing.Text(number, foreground, maxWidth: double.PositiveInfinity, bold: pin.Number == 1);

                Point numberAt = pin.Side switch
                {
                    PinSide.Left => new Point(pin.NumberAnchor.X, pin.NumberAnchor.Y - (numberText.Height / 2)),
                    PinSide.Right => new Point(pin.NumberAnchor.X - numberText.Width, pin.NumberAnchor.Y - (numberText.Height / 2)),
                    _ => new Point(pin.NumberAnchor.X - (numberText.Width / 2), pin.NumberAnchor.Y),
                };

                context.DrawText(numberText, numberAt);

                if (pin.Side == PinSide.Bottom ||
                    !this.thisNames.TryGetValue(pin.Number, out string? name) ||
                    name.Length == 0)
                {
                    continue;
                }

                FormattedText nameText = PinoutDrawing.Text(name, foreground, PinoutGeometry.NameWidth - 8, bold: false);

                Point nameAt = pin.Side == PinSide.Left
                    ? new Point(pin.NameAnchor.X - nameText.Width, pin.NameAnchor.Y - (nameText.Height / 2))
                    : new Point(pin.NameAnchor.X, pin.NameAnchor.Y - (nameText.Height / 2));

                context.DrawText(nameText, nameAt);
            }
        }

        private IBrush Brush(string key, IBrush fallback) =>
            this.TryFindResource(key, this.ActualThemeVariant, out object? value) && value is IBrush brush ? brush : fallback;

        private static FormattedText Text(string text, IBrush brush, double maxWidth, bool bold) =>
            new(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(FontFamily.Default, weight: bold ? FontWeight.Bold : FontWeight.Normal), TextSize, brush)
            {
                MaxTextWidth = maxWidth,
                MaxLineCount = 1,
                Trimming = TextTrimming.CharacterEllipsis,
            };
    }
}
