using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
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
    // the top view, with the notch and pin 1"). The component pane and the component popup hand it the
    // package and the pins - so it is in Controls/ since slice 3c - and every position is
    // PinoutGeometry's, so this only paints. A case CRT cannot draw takes no room.
    //
    // WITH PinsCanBePicked (the popup's - owner request, 2026-11-01) a pin lights up under the pointer, and
    // a click on it raises PinPicked. The host can light one up too (HoverPin - the list beside it), and
    // MARKS THE PIN OF THE IMAGE SHOWN, which stays lit whatever the pointer does (MarkPin - owner request,
    // 2026-11-02: "if I am viewing a specific pin, then do highlight it").
    // ###########################################################################################
    public sealed class PinoutDrawing : Control
    {
        private const double TextSize = 11;

        public static readonly RoutedEvent<PinPickedEventArgs> PinPickedEvent =
            RoutedEvent.Register<PinoutDrawing, PinPickedEventArgs>(nameof(PinPicked), RoutingStrategies.Bubble);

        private PinoutLayout? thisLayout;
        private Dictionary<int, string> thisNames = new();
        private double thisNameWidth = PinoutGeometry.NameWidth;

        // The pin under the pointer - here or in the host's list - and the pin the host says is shown.
        private int? thisHoveredPin;
        private int? thisMarkedPin;

        public event EventHandler<PinPickedEventArgs>? PinPicked
        {
            add => this.AddHandler(PinoutDrawing.PinPickedEvent, value);
            remove => this.RemoveHandler(PinoutDrawing.PinPickedEvent, value);
        }

        public bool PinsCanBePicked { get; set; }

        // The pin lit up by the pointer, else the one marked - for the tests.
        internal int? LitPinForTests => this.thisHoveredPin ?? this.thisMarkedPin;

        internal int? MarkedPinForTests => this.thisMarkedPin;

        // The case and its pins, drawn again.
        public void Show(string? package, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            this.thisNames = new Dictionary<int, string>();
            this.thisHoveredPin = null;
            this.thisMarkedPin = null;

            foreach (ComponentPinEntry pin in pins)
            {
                if (int.TryParse(pin.Pin?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int number) &&
                    !this.thisNames.ContainsKey(number))
                {
                    this.thisNames[number] = pin.Name?.Trim() ?? string.Empty;
                }
            }

            // As wide as its names need - the names not on the case's pins take no room either.
            this.thisNameWidth = PinoutGeometry.NameWidthFor(
                this.thisNames.Values
                    .Where(name => name.Length > 0)
                    .Select(name => PinoutDrawing.Text(name, Brushes.Black, double.PositiveInfinity, bold: false).Width)
                    .DefaultIfEmpty(0)
                    .Max());

            this.thisLayout = PinoutGeometry.For(package, pins.Select(pin => pin.Pin), this.thisNameWidth);
            this.IsVisible = this.thisLayout is not null;
            this.InvalidateMeasure();
            this.InvalidateVisual();
        }

        // Whether there is anything to draw - for the tests, and how wide for the popup's list beside it.
        internal PinoutLayout? LayoutForTests => this.thisLayout;

        // The pin of what the host shows - null for none. It stays lit until the host says otherwise.
        public void MarkPin(int? number)
        {
            if (this.thisMarkedPin != number)
            {
                this.thisMarkedPin = number;
                this.InvalidateVisual();
            }
        }

        // The pin under the pointer in the host's list - null for none.
        public void HoverPin(int? number)
        {
            if (this.thisHoveredPin != number)
            {
                this.thisHoveredPin = number;
                this.InvalidateVisual();
            }
        }

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);

            if (this.PinsCanBePicked && this.thisLayout is { } layout)
            {
                this.HoverOver(PinoutGeometry.PinAt(layout, e.GetPosition(this))?.Number);
            }
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);
            this.HoverOver(null);
        }

        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            if (this.PinsCanBePicked &&
                this.thisLayout is { } layout &&
                e.GetCurrentPoint(this).Properties.IsLeftButtonPressed &&
                PinoutGeometry.PinAt(layout, e.GetPosition(this)) is { } pin)
            {
                this.RaiseEvent(new PinPickedEventArgs(PinoutDrawing.PinPickedEvent, pin.Number));
                e.Handled = true;
            }
        }

        private void HoverOver(int? number)
        {
            if (this.thisHoveredPin != number)
            {
                this.thisHoveredPin = number;
                this.Cursor = number is null ? null : new Cursor(StandardCursorType.Hand);
                this.InvalidateVisual();
            }
        }

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

            // Clear, not nothing: only what is drawn takes the pointer, and the gaps between the pins' names
            // would otherwise let it through.
            if (this.PinsCanBePicked)
            {
                context.FillRectangle(Brushes.Transparent, new Rect(layout.Size));
            }

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

            // The pin shown and the pin under the pointer, each in its own - as a list's selected row and
            // the one under the pointer are.
            if (this.thisMarkedPin is int marked && layout.Pins.FirstOrDefault(pin => pin.Number == marked) is { } markedPin)
            {
                context.DrawRectangle(this.Brush("Form_Selected_Bg", Brushes.LightBlue), null, PinoutGeometry.AreaOf(layout, markedPin), 2, 2);
            }

            if (this.thisHoveredPin is int lit && lit != this.thisMarkedPin && layout.Pins.FirstOrDefault(pin => pin.Number == lit) is { } litPin)
            {
                context.DrawRectangle(this.Brush("Form_Selected_Hover_Bg", Brushes.SkyBlue), null, PinoutGeometry.AreaOf(layout, litPin), 2, 2);
            }

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

                FormattedText nameText = PinoutDrawing.Text(name, foreground, this.thisNameWidth - PinoutGeometry.NameGaps, bold: false);

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

    // A pin of a drawn case clicked - its number.
    public sealed class PinPickedEventArgs(RoutedEvent routedEvent, int pin) : RoutedEventArgs(routedEvent)
    {
        public int Pin { get; } = pin;
    }
}
