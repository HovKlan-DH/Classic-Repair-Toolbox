using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// KiCadOverlayRenderControl's snapshot (2026-11-06): while the calibration box moves, the overlay is
// drawn from a picture of its transformed primitives instead of the thousands of them. A picture
// drawn when it no longer shows what is on screen would show traces that are not there, so these pin
// when it is kept and when it is let go of. Where it lands is OverlaySnapshotGeometryTests'.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class KiCadOverlaySnapshotTests
{
    private static KiCadOverlayPrimitive Trace(double y) => new()
    {
        Kind = KiCadOverlayPrimitiveKind.Polyline,
        Points = [new Point(10, y), new Point(300, y)],
        Pen = new Pen(Brushes.DeepSkyBlue, 2),
    };

    private static KiCadOverlayPrimitive Box() => new()
    {
        Kind = KiCadOverlayPrimitiveKind.Geometry,
        Geometry = new RectangleGeometry(new Rect(5, 5, 320, 200)),
        Pen = new Pen(Brushes.IndianRed, 1),
        DrawnUntransformed = true,
    };

    // A picture of a few primitives, as of a real board's thousands (PrimitivesWorthAPicture).
    private static (Window Window, KiCadOverlayRenderControl Overlay) Shown()
    {
        var overlay = new KiCadOverlayRenderControl
        {
            RenderTransform = new MatrixTransform(),
            RenderTransformOrigin = RelativePoint.TopLeft,
            PrimitivesWorthAPicture = 1,
        };

        var window = new Window { Width = 600, Height = 400, Content = overlay };
        window.Show();
        Dispatcher.UIThread.RunJobs();
        return (window, overlay);
    }

    // The calibration box is replaced on every move of a drag: the picture of the traces under it stays.
    // A trace replaced - geometry built again - is not in the picture, which is let go of.
    [Fact]
    public void The_picture_stays_while_only_the_calibration_box_is_replaced()
    {
        UiTest.Run(() =>
        {
            (Window window, KiCadOverlayRenderControl overlay) = Shown();
            KiCadOverlayPrimitive first = Trace(20);
            KiCadOverlayPrimitive second = Trace(60);

            overlay.SetGeometry([first, second, Box()]);
            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            overlay.OverlayTransform = Matrix.CreateTranslation(4, 3);
            overlay.SetGeometry([first, second, Box()]);
            Assert.True(overlay.IsDrawingFromSnapshot);

            overlay.SetGeometry([first, Trace(60), Box()]);
            Assert.False(overlay.IsDrawingFromSnapshot);

            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            overlay.ClearGeometry();
            Assert.False(overlay.IsDrawingFromSnapshot);

            window.Close();
        });
    }

    // Another zoom is another picture: the one held is not drawn - the primitives are - until the next
    // move takes a new one. Stopped, the primitives are drawn, and the picture is kept for the next move.
    [Fact]
    public void A_picture_is_drawn_only_at_the_zoom_it_was_taken_at_and_only_until_stopped()
    {
        UiTest.Run(() =>
        {
            (Window window, KiCadOverlayRenderControl overlay) = Shown();
            overlay.SetGeometry([Trace(20), Box()]);

            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            ((MatrixTransform)overlay.RenderTransform!).Matrix = new Matrix(2, 0, 0, 2, -40, -10);
            Assert.False(overlay.IsDrawingFromSnapshot);

            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            overlay.StopDrawingFromSnapshot();
            Assert.False(overlay.IsDrawingFromSnapshot);

            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            window.Close();
        });
    }

    // ###########################################################################################
    // Too few primitives for a picture to be worth it - a schematic sheet's wires: they are drawn as
    // vectors while the box moves, as sharp and as cheap (measured: 222 wires, 2.5 ms against 6 ms).
    // ###########################################################################################
    [Fact]
    public void Too_few_primitives_are_drawn_as_vectors_while_the_box_moves()
    {
        UiTest.Run(() =>
        {
            (Window window, KiCadOverlayRenderControl overlay) = Shown();
            overlay.PrimitivesWorthAPicture = 3;

            overlay.SetGeometry([Trace(20), Trace(60), Box()]);
            overlay.DrawFromSnapshot();
            Assert.False(overlay.IsDrawingFromSnapshot);

            overlay.SetGeometry([Trace(20), Trace(60), Trace(100), Box()]);
            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            window.Close();
        });
    }

    // A blinking net: the picture shows the phase it was taken in, and is not drawn in the other.
    [Fact]
    public void A_picture_of_blinking_traces_is_drawn_only_in_the_phase_it_was_taken_in()
    {
        UiTest.Run(() =>
        {
            (Window window, KiCadOverlayRenderControl overlay) = Shown();
            overlay.SetGeometry([Trace(20).WithBlinkPhase(KiCadBlinkPhase.OnPhase), Box()]);

            overlay.DrawFromSnapshot();
            Assert.True(overlay.IsDrawingFromSnapshot);

            overlay.BlinkPhaseVisible = false;
            Assert.False(overlay.IsDrawingFromSnapshot);

            overlay.BlinkPhaseVisible = true;
            Assert.True(overlay.IsDrawingFromSnapshot);

            window.Close();
        });
    }
}
