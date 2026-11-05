using Avalonia.Media;
using Handlers.DataHandling;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // HOW A PICTURE'S KiCad TRACES ARE DRAWN - THIS SIDE'S AND THE OTHER SIDE'S (owner request,
    // 2026-10-30: "be able to set the 'this side' trace color + the opacity for 'this side' and
    // 'other side' traces").
    //
    // The ONE rule the Schematics tab's copper (TabSchematics.KiCad.Render.cs) and the Contribute tab's
    // previews of a schematic image (ComponentPane.BoardPictures.cs) draw by, so a colour picked on a
    // card is the colour the board shows. Each value the picture's row leaves blank is what CRT drew
    // before the columns existed:
    //
    //   - this side's colour: the picture's highlight colour (DeepSkyBlue when that is blank too, or
    //     `blankHighlight` - a schematic sheet's wires have always been Orange);
    //   - the other side's colour: DodgerBlue;
    //   - both opacities: the highlight opacity (20 % when blank) plus a quarter, at most 100 %.
    //
    // *** AN OPACITY GIVEN HOLDS FOR EVERY TRACE, THE ONE UNDER THE POINTER TOO (owner report,
    // 2026-11-03: "even if I change the opacity down to 5% ... it uses a solid/100% opacity"). *** A net
    // under the pointer was always drawn solid, whatever the row said, and that is all most of a
    // board's traces ever are. It still stands out while the opacity is blank, as it always did.
    //
    // Pure: a row in, colours out - every text read by RectGeometry's forgiving parsers, since the
    // text is contributed data.
    // ###########################################################################################
    public readonly record struct TraceColours(Color ThisSide, double ThisSideOpacity, Color OtherSide, double OtherSideOpacity)
    {
        public static readonly Color DefaultHighlight = Colors.DeepSkyBlue;
        public static readonly Color DefaultOtherSide = Colors.DodgerBlue;

        // A schematic sheet's `blankHighlight` - the Schematics tab draws by it, and a sheet's card says so.
        public static readonly Color SheetBlankHighlight = Colors.Orange;

        public const double DefaultHighlightOpacity = 0.20;

        // What the opacity of the traces adds to the highlight's when it is not given - the copper
        // reads stronger than a component's fill.
        public const double OpacityAboveHighlight = 0.25;

        // A zone's fill is lighter than its copper - at most this strong, however strong the copper.
        private const double ZoneFillShare = 0.65;
        private const double ZoneFillMost = 0.38;
        private const double ZoneFillLeast = 0.10;

        // Whether the picture's row says each side's opacity - else it is worked out from the highlight's.
        public bool ThisSideOpacityGiven { get; init; }

        public bool OtherSideOpacityGiven { get; init; }

        public static TraceColours Resolve(BoardSchematicEntry? row, Color? blankHighlight = null)
        {
            Color highlight = RectGeometry.ParseColorOrDefault(row?.SchematicHighlightColor ?? string.Empty, blankHighlight ?? TraceColours.DefaultHighlight);
            double highlightOpacity = RectGeometry.ParseOpacityOrDefault(row?.SchematicHighlightOpacity ?? string.Empty, TraceColours.DefaultHighlightOpacity);
            double derived = Math.Clamp(highlightOpacity + TraceColours.OpacityAboveHighlight, 0.0, 1.0);

            double thisSide = RectGeometry.ParseOpacityOrDefault(row?.TraceHighlightOpacity ?? string.Empty, double.NaN);
            double otherSide = RectGeometry.ParseOpacityOrDefault(row?.OppositeTraceHighlightOpacity ?? string.Empty, double.NaN);

            return new TraceColours(
                RectGeometry.ParseColorOrDefault(row?.TraceHighlightColor ?? string.Empty, highlight),
                double.IsNaN(thisSide) ? derived : thisSide,
                RectGeometry.ParseColorOrDefault(row?.OppositeTraceHighlightColor ?? string.Empty, TraceColours.DefaultOtherSide),
                double.IsNaN(otherSide) ? derived : otherSide)
            {
                ThisSideOpacityGiven = !double.IsNaN(thisSide),
                OtherSideOpacityGiven = !double.IsNaN(otherSide),
            };
        }

        // How strongly a trace of each side is stroked - `hovered`: the net under the pointer, not blinking.
        public double ThisSideStroke(bool hovered) => hovered && !this.ThisSideOpacityGiven ? 1.0 : this.ThisSideOpacity;

        public double OtherSideStroke(bool hovered) => hovered && !this.OtherSideOpacityGiven ? 1.0 : this.OtherSideOpacity;

        // How strongly a zone of each side is filled, by the same rule.
        public double ThisSideZoneFill(bool hovered) =>
            hovered && !this.ThisSideOpacityGiven ? 0.32 : TraceColours.ZoneFill(this.ThisSideOpacity);

        public double OtherSideZoneFill(bool hovered) =>
            hovered && !this.OtherSideOpacityGiven
                ? Math.Min(1.0, TraceColours.ZoneFill(this.OtherSideOpacity) + 0.12)
                : TraceColours.ZoneFill(this.OtherSideOpacity);

        // Never stronger than the copper it belongs to: a floor of 10 % made a zone at 5 % twice as
        // strong as its traces.
        public static double ZoneFill(double opacity) =>
            Math.Clamp(opacity * TraceColours.ZoneFillShare, Math.Min(TraceColours.ZoneFillLeast, opacity), TraceColours.ZoneFillMost);
    }
}
