using Avalonia;
using Avalonia.Media;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// The KiCad overlay's blink as phases of what it already holds (code review, 2026-10-13: every
// blink tick rebuilt every blinking net's primitives). Each phase must draw exactly what the old
// primitives drew with the blink factor baked in - 1 in the shown phase, 0 in the hidden one.
// Brushes are Avalonia objects, so these run through UiTest.Run.
[Collection("HeadlessUi")]
public sealed class KiCadBlinkPhasesTests
{
    [Theory]
    [InlineData(KiCadBlinkPhase.Always, true, true)]
    [InlineData(KiCadBlinkPhase.Always, false, true)]
    [InlineData(KiCadBlinkPhase.OnPhase, true, true)]
    [InlineData(KiCadBlinkPhase.OnPhase, false, false)]
    [InlineData(KiCadBlinkPhase.OffPhase, true, false)]
    [InlineData(KiCadBlinkPhase.OffPhase, false, true)]
    public void Each_phase_draws_its_own_primitives(KiCadBlinkPhase phase, bool shown, bool drawn)
    {
        Assert.Equal(drawn, KiCadOverlayRenderControl.IsDrawnInPhase(phase, shown));
    }

    [Fact]
    public void A_blinking_nets_copper_shows_only_in_the_shown_phase_its_zone_keeps_a_faint_fill_and_pin_one_stays()
    {
        UiTest.Run(() =>
        {
            var copper = new SolidColorBrush(Colors.Orange, 0.45);
            var pinOne = new SolidColorBrush(Colors.Magenta);
            var zoneGeometry = new RectangleGeometry(new Rect(0, 0, 10, 10));

            var trace = new KiCadOverlayPrimitive
            {
                Kind = KiCadOverlayPrimitiveKind.Polyline,
                Points = [new Point(0, 0), new Point(5, 5)],
                Pen = new Pen(copper, 2),
            };

            var zone = new KiCadOverlayPrimitive
            {
                Kind = KiCadOverlayPrimitiveKind.Geometry,
                Geometry = zoneGeometry,
                Fill = new SolidColorBrush(Colors.Orange, 0.29),
                Pen = new Pen(copper, 1),
            };

            var pinOnePad = new KiCadOverlayPrimitive
            {
                Kind = KiCadOverlayPrimitiveKind.Rectangle,
                Rect = new Rect(1, 1, 2, 2),
                Fill = pinOne,
                Pen = new Pen(pinOne, 1.2),
            };

            var split = KiCadBlinkPhases.Split([trace, zone, pinOnePad], pinOne);

            Assert.Equal(4, split.Count);

            // The trace and the zone at full strength, in the shown phase.
            Assert.Equal(KiCadBlinkPhase.OnPhase, split[0].BlinkPhase);
            Assert.Same(trace.Pen, split[0].Pen);
            Assert.Equal(KiCadBlinkPhase.OnPhase, split[1].BlinkPhase);
            Assert.Same(zone.Fill, split[1].Fill);

            // In the hidden phase the zone's fill at its 0.10 floor, with no outline - what a zone
            // looked like at blink factor 0.
            Assert.Equal(KiCadBlinkPhase.OffPhase, split[2].BlinkPhase);
            Assert.Same(zoneGeometry, split[2].Geometry);
            Assert.Null(split[2].Pen);
            var faint = Assert.IsAssignableFrom<ISolidColorBrush>(split[2].Fill);
            Assert.Equal(Colors.Orange, faint.Color);
            Assert.Equal(KiCadBlinkPhases.ZoneFillFloor, faint.Opacity);

            // The pin-1 pad never blinked.
            Assert.Same(pinOnePad, split[3]);
            Assert.Equal(KiCadBlinkPhase.Always, split[3].BlinkPhase);
        });
    }

    [Fact]
    public void A_primitive_in_another_phase_keeps_everything_else()
    {
        UiTest.Run(() =>
        {
            var pen = new Pen(Brushes.Orange, 3);
            var original = new KiCadOverlayPrimitive
            {
                Kind = KiCadOverlayPrimitiveKind.Ellipse,
                Rect = new Rect(1, 2, 3, 4),
                RotationDegrees = 45,
                Pen = pen,
                Fill = Brushes.Orange,
                Start = new Point(1, 1),
                End = new Point(2, 2),
            };

            var moved = original.WithBlinkPhase(KiCadBlinkPhase.OnPhase);

            Assert.Equal(KiCadBlinkPhase.OnPhase, moved.BlinkPhase);
            Assert.Equal(original.Kind, moved.Kind);
            Assert.Equal(original.Rect, moved.Rect);
            Assert.Equal(original.RotationDegrees, moved.RotationDegrees);
            Assert.Same(original.Pen, moved.Pen);
            Assert.Same(original.Fill, moved.Fill);
            Assert.Equal(original.Start, moved.Start);
            Assert.Equal(original.End, moved.End);
        });
    }

    // Setting the phase on the control redraws what it holds - its primitives are untouched.
    [Fact]
    public void Changing_the_phase_keeps_the_primitives_the_overlay_holds()
    {
        UiTest.Run(() =>
        {
            var control = new KiCadOverlayRenderControl();
            var primitives = new List<KiCadOverlayPrimitive>
            {
                new() { Kind = KiCadOverlayPrimitiveKind.Line, BlinkPhase = KiCadBlinkPhase.OnPhase, Pen = new Pen(Brushes.Orange, 1) },
            };

            control.SetGeometry(primitives);
            control.BlinkPhaseVisible = false;
            control.BlinkPhaseVisible = true;

            Assert.Same(primitives, control.Primitives);
        });
    }
}
