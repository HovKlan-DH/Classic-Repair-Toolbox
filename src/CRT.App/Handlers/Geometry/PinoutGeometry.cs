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

    // A drawn case: its body and its pins. Pitch: how far apart its pins are (a DIP has them further apart -
    // LogicChipGeometry). A single row's dot by pin 1 is FirstPinDot; a DIP's notch and dot are LogicChipGeometry's.
    public sealed record PinoutLayout(Rect Body, IReadOnlyList<PinoutPin> Pins, Size Size, double Pitch = PinoutGeometry.Pitch)
    {
        // The dot that marks pin 1 of a single row, inside the body over it.
        public Point FirstPinDot => new(this.Body.Left + PinoutGeometry.EndPadding + (PinoutGeometry.Pitch / 2), this.Body.Bottom - 8);
    }

    // ###########################################################################################
    // THE PINOUT DRAWING'S GEOMETRY (CRT 4.0 step 2, 2026-10-25 - the plan's Pinout: "Pick a package
    // ... CRT draws the top view, with the notch and pin 1").
    //
    // A single row of pins (SIP, TO-92, TO-220) leaves the case at the bottom, pin 1 on the left, its
    // number under it; the names are the pin list's to show. A DIP is LogicChipGeometry's case since
    // 2026-11-17 (owner request: "unify ALL pinout IC's so they look the same" - EmptyCaseFor), the one
    // a chip drawn with its insides stands in; this keeps what every case shares - a pin's area under the
    // pointer, how wide the names are.
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
        public const double EndPadding = 6;
        public const double SingleRowBodyHeight = 34;
        public const double NumberHeight = 14;

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
        // strip of its row, from its name to the case's edge - its number is on its square, outside the
        // case (code review, 2026-11-17: the strip reached 24 into the case, where the old case wrote the
        // numbers, and took a badge drawn there for the pin of its row); under a single row, the column of
        // its lead and number.
        // ###########################################################################################
        public static Rect AreaOf(PinoutLayout layout, PinoutPin pin)
        {
            ArgumentNullException.ThrowIfNull(layout);
            ArgumentNullException.ThrowIfNull(pin);

            double half = layout.Pitch / 2;

            return pin.Side switch
            {
                PinSide.Left => new Rect(0, pin.Lead.Center.Y - half, layout.Body.Left, layout.Pitch),
                PinSide.Right => new Rect(layout.Body.Right, pin.Lead.Center.Y - half, layout.Size.Width - layout.Body.Right, layout.Pitch),
                _ => new Rect(pin.Lead.Center.X - half, layout.Body.Bottom, layout.Pitch, layout.Size.Height - layout.Body.Bottom),
            };
        }

        // The pin at `point` of a drawing - null for none.
        public static PinoutPin? PinAt(PinoutLayout layout, Point point)
        {
            ArgumentNullException.ThrowIfNull(layout);

            return layout.Pins.FirstOrDefault(pin => PinoutGeometry.AreaOf(layout, pin).Contains(point));
        }

        // Whether the pin list, `list` wide with every text whole (PinTableGeometry.WidthOf), fits beside a
        // drawing `drawing` wide in `width` - so no column of it wraps there (owner request, 2026-11-09: each
        // column as wide as its longest text). With no drawing it has the whole width anyway. A width not
        // known yet (0: not laid out) says yes.
        public static bool ListFitsBeside(double width, double drawing, double list) =>
            drawing <= 0 ||
            width <= 0 ||
            width - drawing - PinoutGeometry.ListGap >= list;

        // ###########################################################################################
        // The single row of pins `package` is - or null for a DIP (LogicChipGeometry.EmptyCaseFor) and when
        // CRT cannot tell how many pins it has. A SIP's count is the highest pin number the component's pins
        // name (`pinNumbers`); every other case's is in its name (ComponentVocabulary.PinCountOf).
        // ###########################################################################################
        public static PinoutLayout? For(string? package, IEnumerable<string?> pinNumbers)
        {
            ArgumentNullException.ThrowIfNull(pinNumbers);

            string text = (package ?? string.Empty).Trim();

            if (ComponentVocabulary.IsDualInLine(text))
            {
                return null;
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

            return new PinoutLayout(body, pins, size);
        }
    }
}
