using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// PicturePreview's zoom (owner request, 2026-10-29: "maybe to zoom in/out in that small thumbnail so
// I can view relevant parts of it"). Ctrl and the wheel zoom at the pointer; the wheel alone is the
// cards' scroll and never zooms; a preview that is not zoomable never zooms. Zoomed past what its
// image can show sharp, it asks for the image sharper. Every position is PicturePreviewGeometry's,
// tested there.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class PicturePreviewTests
{
    private static (Window Window, PicturePreview Preview, List<int> Wanted) Shown(bool zoomable)
    {
        var wanted = new List<int>();
        var preview = new PicturePreview
        {
            Source = new RenderTargetBitmap(new PixelSize(80, 40), new Vector(96, 96)),
            ImageSize = new Size(4000, 2000),
            IsZoomable = zoomable,
            SharperWanted = wanted.Add,
        };

        var window = new Window { Content = preview, Width = 200, Height = 100 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, preview, wanted);
    }

    [Fact]
    public void Ctrl_and_the_wheel_zoom_and_the_wheel_alone_does_not()
    {
        UiTest.Run(() =>
        {
            (Window window, PicturePreview preview, List<int> wanted) = Shown(zoomable: true);

            window.MouseWheel(new Point(100, 50), new Vector(0, 1));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, preview.ZoomForTests);

            window.MouseWheel(new Point(100, 50), new Vector(0, 1), RawInputModifiers.Control);
            Dispatcher.UIThread.RunJobs();
            Assert.True(preview.ZoomForTests > 1);

            // Drawn wider than its 80 pixels can show sharp: the image is asked for sharper.
            Assert.NotEmpty(wanted);
            Assert.True(wanted[^1] > 80);

            preview.ShowWhole();
            Assert.Equal(1, preview.ZoomForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // A SHARP IMAGE TAKEN AWAY IS ASKED FOR AGAIN (code review, 2026-10-30): another card takes the one
    // sharp image and this preview goes back to its small one. Only zooming past the widest asked
    // before asked again, so it stayed soft at the zoom it was left at. Now its next zoom asks, a
    // smaller one too, and so does the pointer coming back - never the image going, or two zoomed
    // cards would take it from each other for ever.
    // ###########################################################################################
    [Fact]
    public void A_sharp_image_taken_away_is_asked_for_again_at_the_next_zoom_or_when_the_pointer_comes_back()
    {
        UiTest.Run(() =>
        {
            (Window window, PicturePreview preview, List<int> wanted) = Shown(zoomable: true);
            Bitmap small = preview.Source!;

            preview.ZoomAt(new Point(100, 50), 4);
            int asked = Assert.Single(wanted);
            Bitmap sharp = new RenderTargetBitmap(new PixelSize(asked, asked / 2), new Vector(96, 96));

            // The sharp image comes: nothing more is asked, zoomed a little out.
            preview.Source = sharp;
            preview.ZoomAt(new Point(100, 50), 0.9);
            Assert.Single(wanted);

            // Another card takes it - nothing asked as it goes.
            preview.Source = small;
            Assert.Single(wanted);

            // Zoomed a little further OUT, still zoomed in: asked again.
            preview.ZoomAt(new Point(100, 50), 0.9);
            Assert.Equal(2, wanted.Count);

            // Taken away once more: the pointer coming back over the preview asks - once it RESTS there
            // (code review, 2026-10-30: asked as it passed, two zoomed cards took the one sharp image
            // from each other at every pass). Passing over asks nothing.
            preview.Source = sharp;
            preview.Source = small;
            window.MouseMove(new Point(100, 50));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(2, wanted.Count);
            Assert.True(preview.IsWaitingForPointerRestForTests);

            window.MouseMove(new Point(500, 500));
            Dispatcher.UIThread.RunJobs();
            Assert.False(preview.IsWaitingForPointerRestForTests);
            Assert.Equal(2, wanted.Count);

            window.MouseMove(new Point(100, 50));
            preview.PointerRested();
            Assert.Equal(3, wanted.Count);

            window.Close();
        });
    }

    [Fact]
    public void A_preview_that_is_not_zoomable_never_zooms()
    {
        UiTest.Run(() =>
        {
            (Window window, PicturePreview preview, List<int> wanted) = Shown(zoomable: false);

            window.MouseWheel(new Point(100, 50), new Vector(0, 1), RawInputModifiers.Control);
            preview.ZoomAt(new Point(100, 50), 4);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(1, preview.ZoomForTests);
            Assert.Empty(wanted);

            window.Close();
        });
    }
}
