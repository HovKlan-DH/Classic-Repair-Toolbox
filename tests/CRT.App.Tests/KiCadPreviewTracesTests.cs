using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadPreviewTraces - the copper on both sides of a schematic image, for its card's preview (owner
// requests, 2026-10-29: "I would like to see the same for ... the 'Trace on the other side'", then
// "we also need an option to see the traces on the 'normal' side, so we can have both in preview").
// The board's copper spans 0..100 by 0..50 mm and the picture is 1000 x 500 pixels, so an unmoved
// calibration is 10 pixels a millimetre - every number below is worked out by hand from that.
// ###########################################################################################
public sealed class KiCadPreviewTracesTests
{
    private static readonly KiCadCalibrationValues Unmoved = new(1, 1, 0, 0, false, false);

    private static KiCadPoint2D At(double x, double y) => new() { X = x, Y = y };

    // A top trace, a bottom trace, a bottom arc, and one bottom trace with no width in the data.
    private static KiCadPcb Pcb() => new()
    {
        Routing = new KiCadPcbRouting
        {
            Segments =
            [
                new KiCadPcbSegment { Start = At(0, 0), End = At(100, 50), Width = 0.5, Layer = "F.Cu" },
                new KiCadPcbSegment { Start = At(10, 10), End = At(20, 10), Width = 0.3, Layer = "B.Cu" },
                new KiCadPcbSegment { Start = At(30, 40), End = At(40, 40), Layer = "b.cu" },
            ],
            Arcs =
            [
                new KiCadPcbArc { Start = At(50, 10), Mid = At(55, 15), End = At(60, 10), Width = 0.2, Layer = "B.Cu" },
            ],
        },
    };

    // Arithmetic on the way may leave a hair's difference - the trace is the same.
    private static void Near(PreviewTrace expected, PreviewTrace actual)
    {
        Assert.Equal(expected.From.X, actual.From.X, 6);
        Assert.Equal(expected.From.Y, actual.From.Y, 6);
        Assert.Equal(expected.To.X, actual.To.X, 6);
        Assert.Equal(expected.To.Y, actual.To.Y, 6);
        Assert.Equal(expected.Width, actual.Width, 6);
    }

    // A top picture's own side is the top's copper, its other side the bottom's - placed as the
    // Schematics tab places them.
    [Fact]
    public void A_top_image_shows_the_tops_copper_and_its_other_side_is_the_bottoms()
    {
        PreviewTraces traces = KiCadPreviewTraces.For(Pcb(), bottom: false, Unmoved, 1000, 500);

        Near(new PreviewTrace(new Point(0, 0), new Point(1000, 500), 5), Assert.Single(traces.ThisSide));

        Assert.Equal(4, traces.OtherSide.Count);
        Near(new PreviewTrace(new Point(100, 100), new Point(200, 100), 3), traces.OtherSide[0]);

        // No width in the data: KiCad's default track, and the layer's name in any capitals.
        Near(new PreviewTrace(new Point(300, 400), new Point(400, 400), KiCadPreviewTraces.DefaultWidthMm * 10), traces.OtherSide[1]);

        // An arc: the two straight lines through its middle.
        Near(new PreviewTrace(new Point(500, 100), new Point(550, 150), 2), traces.OtherSide[2]);
        Near(new PreviewTrace(new Point(550, 150), new Point(600, 100), 2), traces.OtherSide[3]);
    }

    // A bottom picture the other way round.
    [Fact]
    public void A_bottom_image_shows_the_bottoms_copper_and_its_other_side_is_the_tops()
    {
        PreviewTraces traces = KiCadPreviewTraces.For(Pcb(), bottom: true, Unmoved, 1000, 500);

        Assert.Equal(4, traces.ThisSide.Count);
        Near(new PreviewTrace(new Point(0, 0), new Point(1000, 500), 5), Assert.Single(traces.OtherSide));
    }

    // The alignment moves, scales and mirrors the traces as it does the copper the tab draws: across
    // the bounds 0..1, mirrored, scaled by half, moved by (20, 30) - and a width by half too.
    [Fact]
    public void The_alignment_moves_scales_and_mirrors_the_traces()
    {
        var calibration = new KiCadCalibrationValues(0.5, 0.5, 20, 30, true, false);

        PreviewTrace trace = KiCadPreviewTraces.For(Pcb(), bottom: false, calibration, 1000, 500).OtherSide[0];

        Near(new PreviewTrace(new Point(470, 80), new Point(420, 80), 1.5), trace);
    }

    // No image size, or a board with no copper: nothing to draw, never a failure.
    [Fact]
    public void Nothing_is_drawn_without_a_size_or_copper()
    {
        PreviewTraces noSize = KiCadPreviewTraces.For(Pcb(), bottom: false, Unmoved, 0, 500);
        Assert.Empty(noSize.ThisSide);
        Assert.Empty(noSize.OtherSide);
        Assert.Empty(KiCadPreviewTraces.For(new KiCadPcb(), bottom: false, Unmoved, 1000, 500).OtherSide);
    }

    // The plan "Mark all automatically" works with gives the traces - one that cannot draw, none.
    [Fact]
    public void A_plan_without_a_board_draws_nothing()
    {
        var plan = new KiCadMarkPlan("Top", null, false, Unmoved, 1000, 500, "Align KiCad on [Top] first");

        Assert.Same(PreviewTraces.None, KiCadPreviewTraces.For(plan));

        PreviewTraces traces = KiCadPreviewTraces.For(plan with { Pcb = Pcb(), Problem = null });
        Assert.Single(traces.ThisSide);
        Assert.Equal(4, traces.OtherSide.Count);
    }
}
