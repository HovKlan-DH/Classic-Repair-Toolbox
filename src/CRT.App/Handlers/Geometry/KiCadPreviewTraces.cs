using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // One trace of the preview, in the picture's own pixels - from, to, and how wide it is drawn.
    public readonly record struct PreviewTrace(Point From, Point To, double Width);

    // The traces of a schematic image's two sides - its own (the side it shows) and the other.
    public sealed record PreviewTraces(IReadOnlyList<PreviewTrace> ThisSide, IReadOnlyList<PreviewTrace> OtherSide)
    {
        public static PreviewTraces None { get; } = new([], []);
    }

    // ###########################################################################################
    // THE TRACES OF A SCHEMATIC IMAGE'S TWO SIDES, FOR ITS CARD'S PREVIEW (owner requests, 2026-10-29:
    // "I would like to see the same for ... the 'Trace on the other side'", then "we also need an option
    // to see the traces on the 'normal' side, so we can have both in preview"). The Schematics tab draws
    // a lit net's copper on the image's own side in its highlight colour and on the other side in the
    // trace colour; the preview draws ALL of each side's copper so, to show how the colours read
    // against the photo.
    //
    //   - ON A PICTURE ALIGNED WITH A SIDE OF THE BOARD only - what "Mark all automatically" works on
    //     (KiCadFootprintMarks.Plan says why not, in the contributor's words).
    //   - THIS SIDE is the copper layer of the side shown - a top picture's F.Cu, a bottom picture's
    //     B.Cu - and THE OTHER SIDE the other one, as the Schematics tab picks them.
    //   - PLACED AS THE SCHEMATICS TAB PLACES COPPER: the board's world bounds and the mapping onto the
    //     picture are KiCadFootprintMarks' (the tab's own), and a trace's width is its world width
    //     scaled as the tab scales a length (the mean of the two scales). An arc is drawn as the two
    //     straight lines through its middle point - enough for a colour sample.
    // ###########################################################################################
    internal static class KiCadPreviewTraces
    {
        // A trace with no width in the data - KiCad's default track.
        public const double DefaultWidthMm = 0.25;

        public static PreviewTraces For(KiCadMarkPlan plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            return plan.Pcb is null ? PreviewTraces.None : KiCadPreviewTraces.For(plan.Pcb, plan.Bottom, plan.Calibration, plan.Width, plan.Height);
        }

        // Both sides of a picture showing `pcb` from the bottom or not.
        public static PreviewTraces For(KiCadPcb pcb, bool bottom, KiCadCalibrationValues calibration, int pixelWidth, int pixelHeight) => new(
            KiCadPreviewTraces.OnLayer(pcb, bottom ? "B.Cu" : "F.Cu", calibration, pixelWidth, pixelHeight),
            KiCadPreviewTraces.OnLayer(pcb, bottom ? "F.Cu" : "B.Cu", calibration, pixelWidth, pixelHeight));

        // Every trace and arc on the copper layer `layer`, placed on the picture.
        public static List<PreviewTrace> OnLayer(KiCadPcb pcb, string layer, KiCadCalibrationValues calibration, int pixelWidth, int pixelHeight)
        {
            ArgumentNullException.ThrowIfNull(pcb);

            var traces = new List<PreviewTrace>();
            Rect world = KiCadFootprintMarks.PcbWorldBounds(pcb);

            if (pixelWidth <= 0 || pixelHeight <= 0 || world.Width <= 0 || world.Height <= 0)
            {
                return traces;
            }

            Point OnPicture(KiCadPoint2D point) =>
                KiCadFootprintMarks.WorldToImagePixel(point.X, point.Y, world, calibration, pixelWidth, pixelHeight);

            // A world length on the picture - the Schematics tab's own rule.
            double WidthOf(double? width) => KiCadFootprintMarks.WorldLengthOnPicture(
                width is > 0 ? width.Value : DefaultWidthMm, world, calibration.ScaleX, calibration.ScaleY, pixelWidth, pixelHeight);

            bool OnIt(string? onLayer) => string.Equals(onLayer?.Trim(), layer, StringComparison.OrdinalIgnoreCase);

            foreach (KiCadPcbSegment segment in pcb.Routing.Segments)
            {
                if (segment.Start is null || segment.End is null || !OnIt(segment.Layer))
                {
                    continue;
                }

                traces.Add(new PreviewTrace(OnPicture(segment.Start), OnPicture(segment.End), WidthOf(segment.Width)));
            }

            foreach (KiCadPcbArc arc in pcb.Routing.Arcs)
            {
                if (arc.Start is null || arc.End is null || !OnIt(arc.Layer))
                {
                    continue;
                }

                double width = WidthOf(arc.Width);

                if (arc.Mid is null)
                {
                    traces.Add(new PreviewTrace(OnPicture(arc.Start), OnPicture(arc.End), width));
                    continue;
                }

                traces.Add(new PreviewTrace(OnPicture(arc.Start), OnPicture(arc.Mid), width));
                traces.Add(new PreviewTrace(OnPicture(arc.Mid), OnPicture(arc.End), width));
            }

            return traces;
        }
    }
}
