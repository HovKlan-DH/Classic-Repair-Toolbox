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
    // 2026-11-02: "if I am viewing a specific pin, then do highlight it"). A pin the POINTER moves onto or off
    // raises PinHovered, so the host lights its row in the list too (owner request, 2026-11-09: "hovering the IC
    // image pin will not highlight the table pin ... do highlight both places") - never for HoverPin, which is
    // the host's own say and would only come back to it.
    //
    // A CHIP CRT KNOWS IS DRAWN WITH ITS INSIDES (owner request, 2026-11-16; LogicChips decides, LogicChipGeometry places):
    // its gates, blocks, switches and wires in a case as a real chip's, each pin a square in the color of its direction, and
    // under it what the colors mean - painted by PinoutDrawing.Inside.cs. AND EVERY OTHER DIP IN THE SAME CASE, EMPTY (owner
    // request, 2026-11-17: "unify ALL pinout IC's so they look the same"; LogicChipGeometry.EmptyCaseFor). It is drawn
    // smaller, never cut, where its place is narrower than it, as a part's picture is (PartPictureDrawing). Only a single
    // row of pins - a SIP, a TO-92 or TO-220 - keeps the plain case (RenderCase).
    //
    // The parts: PinoutDrawing.cs (this - showing, the pointer, the single row's case), PinoutDrawing.Inside.cs (a DIP's
    // case, pins, wires and badges, and the legend under them), PinoutDrawing.Symbols.cs (the gates, blocks and switches inside it).
    // ###########################################################################################
    public sealed partial class PinoutDrawing : Control
    {
        private const double TextSize = 11;

        public static readonly RoutedEvent<PinPickedEventArgs> PinPickedEvent =
            RoutedEvent.Register<PinoutDrawing, PinPickedEventArgs>(nameof(PinPicked), RoutingStrategies.Bubble);

        public static readonly RoutedEvent<PinHoveredEventArgs> PinHoveredEvent =
            RoutedEvent.Register<PinoutDrawing, PinHoveredEventArgs>(nameof(PinHovered), RoutingStrategies.Bubble);

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

        public event EventHandler<PinHoveredEventArgs>? PinHovered
        {
            add => this.AddHandler(PinoutDrawing.PinHoveredEvent, value);
            remove => this.RemoveHandler(PinoutDrawing.PinHoveredEvent, value);
        }

        public bool PinsCanBePicked { get; set; }

        // The pin lit up by the pointer, else the one marked - for the tests.
        internal int? LitPinForTests => this.thisHoveredPin ?? this.thisMarkedPin;

        internal int? MarkedPinForTests => this.thisMarkedPin;

        // How much smaller than its own size it is drawn, for the place it has.
        private double thisScale = 1;

        // ###########################################################################################
        // The case and its pins, drawn again - with the insides of `chip` when CRT knows the chip (LogicChips.For), any other
        // DIP in the same case with nothing in it, each pin in the color of its direction (PinDirectionKinds.ByPin: the one
        // `pins` gives it, else the datasheet's).
        // ###########################################################################################
        public void Show(string? package, IReadOnlyList<ComponentPinEntry> pins, LogicChip? chip = null)
        {
            ArgumentNullException.ThrowIfNull(pins);

            this.thisNames = new Dictionary<int, string>();
            this.thisHoveredPin = null;
            this.thisMarkedPin = null;

            // A pin by its number as every other reading of it goes (ComponentVocabulary.PinNumberOf - "2 [64K]" is pin 2), as
            // its color is (code review, 2026-11-17: by int.TryParse here, such a pin was colored and left unnamed).
            foreach (ComponentPinEntry pin in pins)
            {
                if (ComponentVocabulary.PinNumberOf(pin.Pin) is int number && !this.thisNames.ContainsKey(number))
                {
                    this.thisNames[number] = pin.Name?.Trim() ?? string.Empty;
                }
            }

            // A chip drawn with its insides names every pin outside - the datasheet's name where the board has none - as nothing
            // is written at the pins inside it (owner request, 2026-11-16: "there will always be a text for the pin itself").
            if (chip is not null)
            {
                for (int number = 1; number <= chip.PinCount; number++)
                {
                    if (!this.thisNames.TryGetValue(number, out string? name) || name.Length == 0)
                    {
                        this.thisNames[number] = chip.NameOf(number);
                    }
                }
            }

            // As wide as its names need - the names not on the case's pins take no room either.
            this.thisNameWidth = PinoutGeometry.NameWidthFor(
                this.thisNames.Values
                    .Where(name => name.Length > 0)
                    .Select(name => PinoutDrawing.Text(name, Brushes.Black, double.PositiveInfinity, bold: false).Width)
                    .DefaultIfEmpty(0)
                    .Max());

            this.ShowInside(chip, package, pins);
            this.thisLayout = this.thisInside?.Case ?? PinoutGeometry.For(package, pins.Select(pin => pin.Pin));
            this.thisNaturalSize = this.thisLayout is not { } layout ? default : this.WithCaption(layout.Size);
            this.IsVisible = this.thisLayout is not null;
            this.InvalidateMeasure();
            this.InvalidateVisual();
        }

        // Whether there is anything to draw - for the tests.
        internal PinoutLayout? LayoutForTests => this.thisLayout;

        // Each pin's name as written beside it.
        internal IReadOnlyDictionary<int, string> NamesForTests => this.thisNames;

        // Its own size: the case, and under it what the colors mean - the popup's list stands beside it when there is room for
        // both. Worked out by Show, as nothing else changes it.
        private Size thisNaturalSize;

        internal Size NaturalSize => this.thisNaturalSize;

        protected override Size MeasureOverride(Size availableSize)
        {
            Size natural = this.NaturalSize;

            this.thisScale = natural.Width > 0 && availableSize.Width < natural.Width ? Math.Max(0.4, availableSize.Width / natural.Width) : 1;

            return new Size(natural.Width * this.thisScale, natural.Height * this.thisScale);
        }

        // Where the pointer is on the case as laid out, however small it is drawn.
        private Point Unscaled(PointerEventArgs e) => e.GetPosition(this) / this.thisScale;

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
                this.HoverOver(PinoutGeometry.PinAt(layout, this.Unscaled(e))?.Number);
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
                PinoutGeometry.PinAt(layout, this.Unscaled(e)) is { } pin)
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
                this.RaiseEvent(new PinHoveredEventArgs(PinoutDrawing.PinHoveredEvent, number));
            }
        }

        public override void Render(DrawingContext context)
        {
            if (this.thisLayout is not { } layout)
            {
                return;
            }

            using (context.PushTransform(Matrix.CreateScale(this.thisScale, this.thisScale)))
            {
                if (this.thisInside is { } inside)
                {
                    this.RenderInside(context, inside);
                }
                else
                {
                    this.RenderCase(context, layout);
                }
            }
        }

        // A single row of pins (a DIP is drawn by RenderInside): a box with its leads under it, each lead's number under it.
        private void RenderCase(DrawingContext context, PinoutLayout layout)
        {
            IBrush foreground = this.Brush("Fg", Brushes.Black);
            IBrush body = this.Brush("Card_Bg", Brushes.White);
            IBrush firstPin = this.Brush("Schematics_FirstPin", Brushes.IndianRed);
            var outline = new Pen(foreground, 1);

            // Clear, not nothing: only what is drawn takes the pointer, and the gaps between the leads would
            // otherwise let it through.
            if (this.PinsCanBePicked)
            {
                context.FillRectangle(Brushes.Transparent, new Rect(layout.Size));
            }

            context.DrawRectangle(body, outline, layout.Body, 3, 3);
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
                context.DrawText(numberText, new Point(pin.NumberAnchor.X - (numberText.Width / 2), pin.NumberAnchor.Y));
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

    // The pointer moved onto a pin of a drawn case - its number - or off every pin (null).
    public sealed class PinHoveredEventArgs(RoutedEvent routedEvent, int? pin) : RoutedEventArgs(routedEvent)
    {
        public int? Pin { get; } = pin;
    }
}
