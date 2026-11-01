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
    //   - this side's colour: the picture's highlight colour (DeepSkyBlue when that is blank too);
    //   - the other side's colour: DodgerBlue;
    //   - both opacities: the highlight opacity (20 % when blank) plus a quarter, at most 100 %.
    //
    // Pure: a row in, colours out - every text read by RectGeometry's forgiving parsers, since the
    // text is contributed data.
    // ###########################################################################################
    public readonly record struct TraceColours(Color ThisSide, double ThisSideOpacity, Color OtherSide, double OtherSideOpacity)
    {
        public static readonly Color DefaultHighlight = Colors.DeepSkyBlue;
        public static readonly Color DefaultOtherSide = Colors.DodgerBlue;
        public const double DefaultHighlightOpacity = 0.20;

        // What the opacity of the traces adds to the highlight's when it is not given - the copper
        // reads stronger than a component's fill.
        public const double OpacityAboveHighlight = 0.25;

        public static TraceColours Resolve(BoardSchematicEntry? row)
        {
            Color highlight = RectGeometry.ParseColorOrDefault(row?.SchematicHighlightColor ?? string.Empty, TraceColours.DefaultHighlight);
            double highlightOpacity = RectGeometry.ParseOpacityOrDefault(row?.SchematicHighlightOpacity ?? string.Empty, TraceColours.DefaultHighlightOpacity);
            double derived = Math.Clamp(highlightOpacity + TraceColours.OpacityAboveHighlight, 0.0, 1.0);

            return new TraceColours(
                RectGeometry.ParseColorOrDefault(row?.TraceHighlightColor ?? string.Empty, highlight),
                RectGeometry.ParseOpacityOrDefault(row?.TraceHighlightOpacity ?? string.Empty, derived),
                RectGeometry.ParseColorOrDefault(row?.OppositeTraceHighlightColor ?? string.Empty, TraceColours.DefaultOtherSide),
                RectGeometry.ParseOpacityOrDefault(row?.OppositeTraceHighlightOpacity ?? string.Empty, derived));
        }
    }
}
