using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A PART PICTURE'S PIECES STAND (PartPictureDrawing, the component popup and the Contribute tab's "Pinout" page -
    // owner request, 2026-11-09; what is drawn is PartPicture's). Pure, so a test checks it; in device-independent pixels
    // from (0, 0), the picture's size being the layout's Size.
    //
    // TWO FAMILIES. An AXIAL part lies across the picture, a lead out of each end and the label of each under that lead's
    // outer end - an axial capacitor, a diode, a ferrite bead, a coil. Every other part STANDS, its legs down out of its
    // body at an even pitch and each leg's label under its foot - a radial capacitor and a disc, a TO-92 or TO-220 case
    // seen from its printed side, a crystal can, an EMI filter, a resistor pack, a pin header.
    //
    // A lead's Area is what is pointed at and clicked for it (as PinoutGeometry.AreaOf is the DIP's): an axial lead's half
    // of the picture beyond the body, a leg's column from the body down through its label.
    // ###########################################################################################
    internal static class PartPictureGeometry
    {
        // A label is two lines - the pin number, the name under it - under its lead's end.
        public const double LabelLineHeight = 14;
        public const double LabelGap = 4;

        // ###########################################################################################
        // How far an axial part's lead stands out of each end of its body, and its label's width under that end. JUST ENOUGH
        // FOR THE LABEL (owner request, 2026-11-09: "there is no reason to have such long legs, as it does take space from the
        // right-side table/info. Have shorter legs, but same text") - the picture as wide as its body and that, no wider; a
        // fixed 250 until then. "Cathode" fits the label.
        // ###########################################################################################
        public const double AxialLabelWidth = 48;
        public const double AxialLeadOut = AxialLabelWidth + 6;

        // The room kept around a body for its soft shadow.
        public const double Margin = 6;

        // A TO-220's metal tab over its body, and the hole in it.
        public const double TabHeight = 44;
        public const double HoleSize = 20;

        // A pin header's posts over its plastic base.
        public const double PostHeight = 22;

        // How wide a name's letter is taken to be, to tell whether a name fits under its leg (NamesStagger).
        public const double NameLetterWidth = 6.5;

        // How long a radial electrolytic's minus leg is beside its plus leg.
        public const double ShorterMinusLeg = 0.7;

        // ###########################################################################################
        // A TO-92 AS IT LOOKS (owner request, 2026-11-09: "see if you can get this more "real" ... it does not look like a
        // transistor"): seen from its flat side and a little from above, so the top of the case - the D its round back makes -
        // shows over the flat face (Tab: the ellipse that D is half of, its middle on the face's top edge); its three thin
        // legs leave the case close together, 1.27 mm apart as they are, then bend out to their labels' pitch and run down
        // long (LegPath).
        // ###########################################################################################
        public const double CapDepth = 8;

        // How close the legs stand as they leave a TO-92, beside where they end; how far down they bend out, and over how much.
        public const double LegsTogether = 0.45;
        public const double BendStart = 7;
        public const double BendLength = 11;

        // ###########################################################################################
        // One lead: its pin number, the wire (or leg), the box its number is written in and the box its name is written in
        // (centred in each), what is clicked for it, and - a pin header's - the post standing up out of the base.
        // ###########################################################################################
        public sealed record LeadPlace(int Number, Rect Wire, Rect Label, Rect NameBox, Rect Area, Rect? Post);

        // ###########################################################################################
        // A drawn part. Band: a radial electrolytic's minus stripe, an axial electrolytic's groove round its plus end, a
        // diode's cathode band. Tab: a TO-220's metal tab, a crystal's base, a disc's coating run down round its legs, an
        // axial electrolytic's rubber end. Hole: the TO-220's. Dot: a resistor pack's pin 1. Marking: where what is printed
        // on the body stands (empty for nothing).
        // ###########################################################################################
        public sealed record Layout(
            PartPicture.Shape Shape,
            bool IsAxial,
            Rect Body,
            Rect Marking,
            Rect? Band,
            Rect? Tab,
            Rect? Hole,
            Point? Dot,
            IReadOnlyList<LeadPlace> Leads,
            Size Size);

        public static bool LiesAcross(PartPicture.Shape shape) => shape is
            PartPicture.Shape.AxialElectrolytic or
            PartPicture.Shape.AxialCapacitor or
            PartPicture.Shape.GlassDiode or
            PartPicture.Shape.EpoxyDiode or
            PartPicture.Shape.FerriteBead or
            PartPicture.Shape.Coil;

        public static Layout For(PartPicture.Picture picture)
        {
            ArgumentNullException.ThrowIfNull(picture);

            return PartPictureGeometry.LiesAcross(picture.Shape) ? PartPictureGeometry.Axial(picture) : PartPictureGeometry.Upright(picture);
        }

        // The lead at `point` - null for none.
        public static LeadPlace? LeadAt(Layout layout, Point point)
        {
            ArgumentNullException.ThrowIfNull(layout);

            return layout.Leads.FirstOrDefault(lead => lead.Area.Contains(point));
        }

        // ###########################################################################################
        // A part lying across: the body in the middle, a lead from each end to the picture's edge, the labels under the
        // leads' outer ends. A body's length and thickness are each part's own, as it looks beside the others.
        // ###########################################################################################
        private static Layout Axial(PartPicture.Picture picture)
        {
            (double length, double thickness, double lead) = picture.Shape switch
            {
                PartPicture.Shape.AxialElectrolytic => (120.0, 46.0, 5.0),
                PartPicture.Shape.AxialCapacitor => (62.0, 24.0, 4.0),
                PartPicture.Shape.GlassDiode => (52.0, 20.0, 3.5),
                PartPicture.Shape.EpoxyDiode => (60.0, 28.0, 4.5),
                PartPicture.Shape.FerriteBead => (42.0, 32.0, 4.0),
                _ => (92.0, 34.0, 4.5),
            };

            double width = length + (2 * (PartPictureGeometry.Margin + PartPictureGeometry.AxialLeadOut));
            double centreY = PartPictureGeometry.Margin + (thickness / 2);
            var body = new Rect((width - length) / 2, PartPictureGeometry.Margin, length, thickness);

            double labelTop = centreY + (lead / 2) + PartPictureGeometry.LabelGap;
            double labelHeight = 2 * PartPictureGeometry.LabelLineHeight;
            double height = Math.Max(body.Bottom + (thickness * 0.4), labelTop + labelHeight) + 2;
            var leftLabel = new Rect(PartPictureGeometry.Margin, labelTop, PartPictureGeometry.AxialLabelWidth, PartPictureGeometry.LabelLineHeight);
            var rightLabel = new Rect(width - PartPictureGeometry.Margin - PartPictureGeometry.AxialLabelWidth, labelTop, PartPictureGeometry.AxialLabelWidth, PartPictureGeometry.LabelLineHeight);

            // Each wire runs a little under the body, so no gap shows at its end.
            var leftWire = new Rect(PartPictureGeometry.Margin, centreY - (lead / 2), body.Left + 2 - PartPictureGeometry.Margin, lead);
            var rightWire = new Rect(body.Right - 2, centreY - (lead / 2), width - PartPictureGeometry.Margin - (body.Right - 2), lead);

            var leads = new List<LeadPlace>
            {
                new(
                    picture.Leads[0].Number,
                    leftWire,
                    leftLabel,
                    leftLabel.Translate(new Vector(0, PartPictureGeometry.LabelLineHeight)),
                    new Rect(0, 0, body.Left, height),
                    null),
                new(
                    picture.Leads[1].Number,
                    rightWire,
                    rightLabel,
                    rightLabel.Translate(new Vector(0, PartPictureGeometry.LabelLineHeight)),
                    new Rect(body.Right, 0, width - body.Right, height),
                    null),
            };

            bool markedLeft = picture.MarkedLead == 0;
            Rect? band = null, tab = null;
            Rect marking = default;

            switch (picture.Shape)
            {
                case PartPicture.Shape.GlassDiode:
                case PartPicture.Shape.EpoxyDiode:
                    band = PartPictureGeometry.EndBand(body, markedLeft, share: 0.15, inset: 0.1);
                    break;

                case PartPicture.Shape.AxialElectrolytic:
                    // NO MINUS BAND (owner request, 2026-11-09: "I have only seen capacitors with an "indention ring" at the
                    // plus-side"): the groove pressed round the can near its plus end, where the rubber seal stands out -
                    // the end away from the marked minus lead.
                    band = PartPictureGeometry.EndBand(body, atLeft: !markedLeft, share: 0.08, inset: 0.13);

                    double sealWidth = length * 0.05, sealHeight = thickness * 0.66;
                    tab = new Rect(markedLeft ? body.Right - 2 : body.Left - sealWidth + 2, centreY - (sealHeight / 2), sealWidth, sealHeight);
                    marking = new Rect(body.Left + (length * 0.24), body.Top + (thickness * 0.18), length * 0.52, thickness * 0.64);
                    break;

                case PartPicture.Shape.AxialCapacitor:
                    marking = new Rect(body.Left + (length * 0.12), body.Top + (thickness * 0.12), length * 0.76, thickness * 0.76);
                    break;
            }

            return new Layout(picture.Shape, IsAxial: true, body, marking, band, tab, Hole: null, Dot: null, leads, new Size(width, height));
        }

        // A band across an axial body, `share` of its length wide, `inset` of it in from the marked end.
        private static Rect EndBand(Rect body, bool atLeft, double share, double inset)
        {
            double width = body.Width * share;
            double left = atLeft ? body.Left + (body.Width * inset) : body.Right - (body.Width * inset) - width;

            return new Rect(left, body.Top, width, body.Height);
        }

        // ###########################################################################################
        // A part standing: its body at the top, centred; its legs down from under it at an even pitch; each label under
        // its leg's foot. The body's size, the pitch, the legs' length and thickness are each part's own; a resistor pack
        // and a pin header grow with their legs.
        // ###########################################################################################
        private static Layout Upright(PartPicture.Picture picture)
        {
            int count = picture.Leads.Count;

            (double bodyWidth, double bodyHeight, double pitch, double legLength, double leg, double legInset) = picture.Shape switch
            {
                PartPicture.Shape.RadialElectrolytic => (66.0, 100.0, 30.0, 44.0, 4.0, 3.0),
                PartPicture.Shape.RadialCapacitor => (48.0, 46.0, 26.0, 34.0, 3.5, 6.0),
                PartPicture.Shape.DiscCapacitor => (66.0, 58.0, 28.0, 30.0, 3.5, 4.0),
                PartPicture.Shape.TO92 => (44.0, 42.0, 22.0, 62.0, 3.0, 1.0),
                PartPicture.Shape.TO220 => (72.0, 60.0, 28.0, 44.0, 6.0, 1.0),
                PartPicture.Shape.Crystal => (74.0, 54.0, count == 3 ? 22.0 : 40.0, 32.0, 3.0, 1.0),
                PartPicture.Shape.EmiFilter => (64.0, 38.0, 24.0, 34.0, 3.5, 4.0),
                PartPicture.Shape.ResistorNetwork => ((count * 24.0) + 16, 34.0, 24.0, 30.0, 4.0, 1.0),
                _ => (count * 22.0, 16.0, 22.0, 26.0, 6.0, 1.0),
            };

            double labelWidth = Math.Max(pitch, 20);
            bool stagger = PartPictureGeometry.NamesStagger(picture, labelWidth);
            double nameWidth = stagger ? (2 * pitch) - 2 : labelWidth;
            double legSpan = (count - 1) * pitch;
            double width = Math.Max(bodyWidth + (2 * PartPictureGeometry.Margin) + 4, legSpan + Math.Max(labelWidth, nameWidth));
            double centreX = width / 2;
            double top = PartPictureGeometry.Margin / 2;

            Rect? tab = null, hole = null;
            Point? dot = null;

            double bodyTop = picture.Shape switch
            {
                PartPicture.Shape.TO220 => top + PartPictureGeometry.TabHeight - 4,
                PartPicture.Shape.PinHeader => top + PartPictureGeometry.PostHeight,
                PartPicture.Shape.TO92 => top + PartPictureGeometry.CapDepth,
                _ => top,
            };

            var body = new Rect(centreX - (bodyWidth / 2), bodyTop, bodyWidth, bodyHeight);

            switch (picture.Shape)
            {
                case PartPicture.Shape.TO220:
                    tab = new Rect(body.Left, top, bodyWidth, PartPictureGeometry.TabHeight);
                    hole = new Rect(centreX - (PartPictureGeometry.HoleSize / 2), top + 19 - (PartPictureGeometry.HoleSize / 2), PartPictureGeometry.HoleSize, PartPictureGeometry.HoleSize);
                    break;

                case PartPicture.Shape.Crystal:
                    tab = new Rect(body.Left - 3, body.Bottom - 1, bodyWidth + 6, 7);
                    break;

                case PartPicture.Shape.TO92:
                    tab = new Rect(body.Left, body.Top - PartPictureGeometry.CapDepth, bodyWidth, 2 * PartPictureGeometry.CapDepth);
                    break;

                case PartPicture.Shape.DiscCapacitor:
                    // The coating run down from the disc round each leg, joined between them.
                    tab = new Rect(centreX - ((pitch + 16) / 2), body.Bottom - 9, pitch + 16, 19);
                    break;

                case PartPicture.Shape.ResistorNetwork:
                    dot = new Point(body.Left + 9, body.Top + 9);
                    break;
            }

            // The legs come out of the lowest piece of the part.
            double legTop = (picture.Shape is PartPicture.Shape.Crystal or PartPicture.Shape.DiscCapacitor ? tab!.Value.Bottom : body.Bottom) - legInset;
            double legBottom = legTop + legLength;
            double labelTop = legBottom + PartPictureGeometry.LabelGap;
            double labelHeight = (stagger ? 3 : 2) * PartPictureGeometry.LabelLineHeight;
            double areaTop = Math.Min(body.Bottom, legTop);

            var leads = new List<LeadPlace>(count);

            for (int i = 0; i < count; i++)
            {
                double x = centreX - (legSpan / 2) + (i * pitch);

                // Staggered, every other name a line lower, so each has two pitches' room.
                double nameTop = labelTop + PartPictureGeometry.LabelLineHeight + (stagger && i % 2 == 1 ? PartPictureGeometry.LabelLineHeight : 0);

                // A radial electrolytic's plus leg is the longer, as it comes (owner request, 2026-11-09): the minus one
                // ends higher, every label still on one line.
                double length = picture.Shape == PartPicture.Shape.RadialElectrolytic && picture.MarkedLead == i
                    ? legLength * PartPictureGeometry.ShorterMinusLeg
                    : legLength;

                leads.Add(new LeadPlace(
                    picture.Leads[i].Number,
                    new Rect(x - (leg / 2), legTop, leg, length),
                    new Rect(x - (labelWidth / 2), labelTop, labelWidth, PartPictureGeometry.LabelLineHeight),
                    new Rect(x - (nameWidth / 2), nameTop, nameWidth, PartPictureGeometry.LabelLineHeight),
                    new Rect(x - (pitch / 2), areaTop, pitch, labelTop + labelHeight - areaTop),
                    picture.Shape == PartPicture.Shape.PinHeader ? new Rect(x - (leg / 2), top, leg, body.Top - top + 2) : null));
            }

            Rect? band = null;
            Rect marking = default;

            switch (picture.Shape)
            {
                case PartPicture.Shape.RadialElectrolytic:
                    // The minus stripe down the side of the can the minus leg is on; the print beside it.
                    double stripe = bodyWidth * 0.26;
                    bool minusLeft = picture.MarkedLead == 0;
                    band = new Rect(minusLeft ? body.Left + (bodyWidth * 0.1) : body.Right - (bodyWidth * 0.1) - stripe, body.Top, stripe, bodyHeight * 0.86);
                    double printLeft = minusLeft ? band.Value.Right + 2 : body.Left + (bodyWidth * 0.06);
                    double printRight = minusLeft ? body.Right - (bodyWidth * 0.06) : band.Value.Left - 2;
                    marking = new Rect(printLeft, body.Top + (bodyHeight * 0.2), printRight - printLeft, bodyHeight * 0.42);
                    break;

                case PartPicture.Shape.RadialCapacitor:
                    marking = new Rect(body.Left + (bodyWidth * 0.15), body.Top + (bodyHeight * 0.18), bodyWidth * 0.7, bodyHeight * 0.56);
                    break;

                case PartPicture.Shape.DiscCapacitor:
                    marking = new Rect(body.Left + (bodyWidth * 0.18), body.Top + (bodyHeight * 0.22), bodyWidth * 0.64, bodyHeight * 0.56);
                    break;

                case PartPicture.Shape.TO92:
                    marking = new Rect(body.Left + 3, body.Top + (bodyHeight * 0.28), bodyWidth - 6, bodyHeight * 0.44);
                    break;

                case PartPicture.Shape.TO220:
                    marking = new Rect(body.Left + 5, body.Top + (bodyHeight * 0.22), bodyWidth - 10, bodyHeight * 0.42);
                    break;

                case PartPicture.Shape.Crystal:
                    marking = new Rect(body.Left + 6, body.Top + (bodyHeight * 0.18), bodyWidth - 12, bodyHeight * 0.64);
                    break;

                case PartPicture.Shape.EmiFilter:
                    marking = new Rect(body.Left + 4, body.Top + (bodyHeight * 0.15), bodyWidth - 8, bodyHeight * 0.7);
                    break;

                case PartPicture.Shape.ResistorNetwork:
                    marking = new Rect(body.Left + 18, body.Top + 4, bodyWidth - 22, bodyHeight - 8);
                    break;
            }

            return new Layout(picture.Shape, IsAxial: false, body, marking, band, tab, hole, dot, leads, new Size(width, labelTop + labelHeight + 2));
        }

        // ###########################################################################################
        // The line a leg is drawn along, top to bottom, through the middle of its wire: straight down for every part but a
        // TO-92, whose legs leave the case LegsTogether as far apart as they end, run down BendStart, bend out over
        // BendLength and go on straight to the foot.
        // ###########################################################################################
        public static IReadOnlyList<Point> LegPath(Layout layout, LeadPlace lead)
        {
            ArgumentNullException.ThrowIfNull(layout);
            ArgumentNullException.ThrowIfNull(lead);

            double x = lead.Wire.Center.X;

            if (layout.Shape != PartPicture.Shape.TO92)
            {
                return [new Point(x, lead.Wire.Top), new Point(x, lead.Wire.Bottom)];
            }

            double centre = layout.Body.Center.X;
            double close = centre + ((x - centre) * PartPictureGeometry.LegsTogether);
            double bend = layout.Body.Bottom + PartPictureGeometry.BendStart;

            return
            [
                new Point(close, lead.Wire.Top),
                new Point(close, bend),
                new Point(x, bend + PartPictureGeometry.BendLength),
                new Point(x, lead.Wire.Bottom),
            ];
        }

        // ###########################################################################################
        // Whether the names under three or more legs stand on two lines, every other one lower: when any is too long for
        // its leg's pitch (a resistor pack's "common"), as cut short it would say nothing. A sign or letter never is.
        // ###########################################################################################
        public static bool NamesStagger(PartPicture.Picture picture, double labelWidth)
        {
            ArgumentNullException.ThrowIfNull(picture);

            return picture.Leads.Count > 2 &&
                   picture.Leads.Any(lead => lead.Name.Length * PartPictureGeometry.NameLetterWidth > labelWidth);
        }
    }
}
