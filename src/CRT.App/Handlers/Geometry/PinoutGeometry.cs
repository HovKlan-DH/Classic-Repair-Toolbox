using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    public enum PinSide
    {
        Left,
        Right,
        Bottom,
    }

    // One pin of a drawn case: its number, which side it leaves the case on, the pin itself, and
    // where its number and its signal name are written.
    public sealed record PinoutPin(int Number, PinSide Side, Rect Lead, Point NumberAnchor, Point NameAnchor);

    // A drawn case: its body, its pins, and - for a DIP - the notch at the top and the dot by pin 1.
    public sealed record PinoutLayout(Rect Body, IReadOnlyList<PinoutPin> Pins, bool IsDualInLine, Size Size)
    {
        // The notch's centre - the middle of the top edge.
        public Point Notch => new(this.Body.Center.X, this.Body.Top);

        // The dot that marks pin 1, inside the body beside it - clear of its number, which is written
        // against the edge.
        public Point FirstPinDot => this.IsDualInLine
            ? new Point(this.Body.Left + PinoutGeometry.DotInset, this.Pins[0].Lead.Center.Y)
            : new Point(this.Body.Left + PinoutGeometry.EndPadding + (PinoutGeometry.Pitch / 2), this.Body.Bottom - 8);
    }

    // ###########################################################################################
    // THE PINOUT DRAWING'S GEOMETRY (CRT 4.0 step 2, 2026-10-25 - the plan's Pinout: "Pick a package
    // ... CRT draws the top view, with the notch and pin 1").
    //
    // A DIP seen from above, notch at the top: pins 1 to N/2 down the left side, N/2+1 to N UP the
    // right side - so pin N sits opposite pin 1, as on every datasheet. Each pin's signal name is
    // written outside the case beside it, its number inside. A single row of pins (SIP, TO-92,
    // TO-220) leaves the case at the bottom, pin 1 on the left, its number under it; the names are
    // the pin list's to show.
    //
    // A case whose pin count CRT cannot tell has no drawing (null) - the pin list alone says it all.
    // Coordinates start at (0, 0); the drawing's size is the layout's Size.
    // ###########################################################################################
    public static class PinoutGeometry
    {
        public const double Pitch = 18;
        public const double LeadLength = 12;
        public const double LeadThickness = 6;
        public const double NameWidth = 120;
        public const double NarrowBodyWidth = 76;
        public const double WideBodyWidth = 116;
        public const double EndPadding = 6;
        public const double DotInset = 24;
        public const double SingleRowBodyHeight = 34;
        public const double NumberHeight = 14;

        // How much of a pin's number stands inside the case, which a pin's area takes in.
        private const double NumberInsideCase = 24;

        // What the pin list beside a drawing needs at least: its "Pin", "Signal" and "Direction" columns.
        public const double SmallestListWidth = 200;

        // The drawing's right margin (the markup's), between it and the list.
        public const double ListGap = 12;

        // ###########################################################################################
        // How wide a DIP's name columns are, for its longest name as measured: that and the gap on each
        // side of it, at most NameWidth, where a longer one is cut (2026-11-01: 120 pixels on each side,
        // mostly empty, left the popup's pin list no room beside the drawing).
        // ###########################################################################################
        public static double NameWidthFor(double longestName) =>
            Math.Min(PinoutGeometry.NameWidth, Math.Ceiling(Math.Max(0, longestName)) + PinoutGeometry.NameGaps);

        // The gap between a name and its lead, and on its other side.
        public const double NameGaps = 8;

        // ###########################################################################################
        // The part of a drawing that is one pin - to point at and click (owner request, 2026-11-01): the
        // strip of its row, from its name to its number inside the case; under a single row, the column
        // of its lead and number.
        // ###########################################################################################
        public static Rect AreaOf(PinoutLayout layout, PinoutPin pin)
        {
            ArgumentNullException.ThrowIfNull(layout);
            ArgumentNullException.ThrowIfNull(pin);

            double half = PinoutGeometry.Pitch / 2;

            return pin.Side switch
            {
                PinSide.Left => new Rect(0, pin.Lead.Center.Y - half, layout.Body.Left + PinoutGeometry.NumberInsideCase, PinoutGeometry.Pitch),
                PinSide.Right => new Rect(
                    layout.Body.Right - PinoutGeometry.NumberInsideCase,
                    pin.Lead.Center.Y - half,
                    layout.Size.Width - layout.Body.Right + PinoutGeometry.NumberInsideCase,
                    PinoutGeometry.Pitch),
                _ => new Rect(pin.Lead.Center.X - half, layout.Body.Bottom, PinoutGeometry.Pitch, layout.Size.Height - layout.Body.Bottom),
            };
        }

        // The pin at `point` of a drawing - null for none.
        public static PinoutPin? PinAt(PinoutLayout layout, Point point)
        {
            ArgumentNullException.ThrowIfNull(layout);

            return layout.Pins.FirstOrDefault(pin => PinoutGeometry.AreaOf(layout, pin).Contains(point));
        }

        // Whether the pin list fits beside a drawing `drawing` wide in `width` - with no drawing it has
        // the whole width anyway. A width not known yet (0: not laid out) says yes.
        public static bool ListFitsBeside(double width, double drawing) =>
            drawing <= 0 || width <= 0 || width - drawing - PinoutGeometry.ListGap >= PinoutGeometry.SmallestListWidth;

        // ###########################################################################################
        // The drawing of `package` - or null when CRT cannot tell how many pins it has. A SIP's count
        // is the highest pin number the component's pins name (`pinNumbers`); every other case's is
        // in its name (ComponentVocabulary.PinCountOf). `nameWidth`: how wide its name columns are.
        // ###########################################################################################
        public static PinoutLayout? For(string? package, IEnumerable<string?> pinNumbers, double nameWidth = PinoutGeometry.NameWidth)
        {
            ArgumentNullException.ThrowIfNull(pinNumbers);

            string text = (package ?? string.Empty).Trim();

            if (ComponentVocabulary.IsDualInLine(text) && ComponentVocabulary.PinCountOf(text) is int dipCount)
            {
                return PinoutGeometry.DualInLine(dipCount, wide: text.Contains("wide", StringComparison.OrdinalIgnoreCase), nameWidth);
            }

            if (ComponentVocabulary.PinCountOf(text) is int count)
            {
                return PinoutGeometry.SingleRow(count);
            }

            if (string.Equals(text, "SIP", StringComparison.OrdinalIgnoreCase))
            {
                int highest = pinNumbers
                    .Select(pin => int.TryParse(pin?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int number) ? number : 0)
                    .DefaultIfEmpty(0)
                    .Max();

                return highest is >= 2 and <= 64 ? PinoutGeometry.SingleRow(highest) : null;
            }

            return null;
        }

        // A DIP of `count` pins (even, at least 2), seen from above.
        public static PinoutLayout DualInLine(int count, bool wide, double nameWidth = PinoutGeometry.NameWidth)
        {
            if (count < 2 || count % 2 != 0)
            {
                throw new ArgumentOutOfRangeException(nameof(count), count, "A DIP has an even number of pins, at least two.");
            }

            int perSide = count / 2;
            double bodyWidth = wide ? PinoutGeometry.WideBodyWidth : PinoutGeometry.NarrowBodyWidth;
            double bodyLeft = nameWidth + PinoutGeometry.LeadLength;
            var body = new Rect(bodyLeft, 0, bodyWidth, (perSide * PinoutGeometry.Pitch) + (2 * PinoutGeometry.EndPadding));

            var pins = new List<PinoutPin>(count);

            for (int index = 0; index < count; index++)
            {
                bool left = index < perSide;
                int row = left ? index : count - 1 - index;
                double centreY = body.Top + PinoutGeometry.EndPadding + ((row + 0.5) * PinoutGeometry.Pitch);
                double top = centreY - (PinoutGeometry.LeadThickness / 2);

                pins.Add(left
                    ? new PinoutPin(
                        index + 1,
                        PinSide.Left,
                        new Rect(body.Left - PinoutGeometry.LeadLength, top, PinoutGeometry.LeadLength, PinoutGeometry.LeadThickness),
                        new Point(body.Left + 4, centreY),
                        new Point(body.Left - PinoutGeometry.LeadLength - 4, centreY))
                    : new PinoutPin(
                        index + 1,
                        PinSide.Right,
                        new Rect(body.Right, top, PinoutGeometry.LeadLength, PinoutGeometry.LeadThickness),
                        new Point(body.Right - 4, centreY),
                        new Point(body.Right + PinoutGeometry.LeadLength + 4, centreY)));
            }

            var size = new Size(body.Right + PinoutGeometry.LeadLength + nameWidth, body.Bottom);

            return new PinoutLayout(body, pins, IsDualInLine: true, size);
        }

        // A single row of `count` pins (at least 1), leaving the case at the bottom.
        public static PinoutLayout SingleRow(int count)
        {
            if (count < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(count), count, "A case has at least one pin.");
            }

            var body = new Rect(0, 0, (count * PinoutGeometry.Pitch) + (2 * PinoutGeometry.EndPadding), PinoutGeometry.SingleRowBodyHeight);
            var pins = new List<PinoutPin>(count);

            for (int index = 0; index < count; index++)
            {
                double centreX = body.Left + PinoutGeometry.EndPadding + ((index + 0.5) * PinoutGeometry.Pitch);
                var lead = new Rect(centreX - (PinoutGeometry.LeadThickness / 2), body.Bottom, PinoutGeometry.LeadThickness, PinoutGeometry.LeadLength);

                pins.Add(new PinoutPin(
                    index + 1,
                    PinSide.Bottom,
                    lead,
                    new Point(centreX, lead.Bottom + 2),
                    new Point(centreX, lead.Bottom + 2)));
            }

            var size = new Size(body.Right, body.Bottom + PinoutGeometry.LeadLength + PinoutGeometry.NumberHeight + 2);

            return new PinoutLayout(body, pins, IsDualInLine: false, size);
        }
    }
}
