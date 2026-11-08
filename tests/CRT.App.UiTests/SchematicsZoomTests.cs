using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The schematic viewer's zoom limits, exercised on a real TabSchematics in a real (headless)
// window.
//
// ViewportMathTests already covers ComputeWheelZoomFactor - how far ONE wheel notch zooms.
// This covers what that number is then clamped to, which lives in TabSchematics.Viewport and
// so was previously unreachable by any test: zoom in far enough and it must stop dead at
// AppConfig.SchematicsMaxZoom; zoom out far enough and it must stop at the 1.0 baseline,
// because the image is already fitted by Stretch="Uniform" and zooming below that would show
// empty space.
//
// It also covers zoom anchoring: whatever is under the mouse pointer must still be under it
// afterwards, so zooming in on a pad puts you on that pad rather than somewhere near it. That
// is checked through TryGetSchematicsImagePixelPoint - the same mapping hover and hit testing
// use - by asking which bitmap pixel sits under the cursor before and after the zoom.
//
// ApplySchematicsZoom and TryGetSchematicsImagePixelPoint are private, so they are reached by
// reflection - the same approach ExternalTargetLauncherTests uses, and for the same reason: the
// behaviour is worth pinning and the methods are not worth widening just for a test.
// ###########################################################################################
[Collection("HeadlessUi")]
public class SchematicsZoomTests
{
    // Anchor for the zoom. Any point inside the container works; the limits do not depend on it.
    private static readonly Point ZoomCentre = new(400, 300);

    [Fact]
    public void Zooming_in_repeatedly_stops_exactly_at_the_maximum_zoom_level()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            // Far more notches than could ever be needed to reach the ceiling, so the test
            // proves it STOPS rather than merely that it climbs.
            for (int i = 0; i < 40; i++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor);

                Assert.True(
                    tab.schematicsMatrix.M11 <= AppConfig.SchematicsMaxZoom,
                    $"Zoom exceeded the maximum on notch {i + 1}: {tab.schematicsMatrix.M11}");
            }

            Assert.Equal(AppConfig.SchematicsMaxZoom, tab.schematicsMatrix.M11, precision: 6);
        });
    }

    [Fact]
    public void Zooming_out_repeatedly_stops_at_the_baseline_and_never_goes_below_it()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            // Zoom in first. Without this the test would pass even if zoom did nothing at
            // all, since the view already starts at the 1.0 baseline it is meant to land on.
            for (int i = 0; i < 5; i++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor);
            }

            Assert.True(tab.schematicsMatrix.M11 > 1.0, "Zooming in did not change the view.");

            for (int i = 0; i < 40; i++)
            {
                ApplyZoom(tab, 1.0 / AppConfig.SchematicsZoomFactor);

                Assert.True(
                    tab.schematicsMatrix.M11 >= 1.0,
                    $"Zoom fell below the baseline on notch {i + 1}: {tab.schematicsMatrix.M11}");
            }

            Assert.Equal(1.0, tab.schematicsMatrix.M11, precision: 6);
        });
    }

    [Fact]
    public void Zooming_all_the_way_in_then_all_the_way_out_returns_to_the_baseline()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            for (int i = 0; i < 40; i++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor);
            }

            Assert.Equal(AppConfig.SchematicsMaxZoom, tab.schematicsMatrix.M11, precision: 6);

            for (int i = 0; i < 40; i++)
            {
                ApplyZoom(tab, 1.0 / AppConfig.SchematicsZoomFactor);
            }

            // Below the baseline the viewer resets to Matrix.Identity outright rather than
            // clamping the scale alone, so the pan offset must be gone too - not just the zoom.
            Assert.Equal(Matrix.Identity, tab.schematicsMatrix);
        });
    }

    [Fact]
    public void A_nonsensical_zoom_factor_leaves_the_view_untouched()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            ApplyZoom(tab, AppConfig.SchematicsZoomFactor);
            Matrix afterOneNotch = tab.schematicsMatrix;

            foreach (double bad in new[] { double.NaN, double.PositiveInfinity, 0.0, -2.0 })
            {
                ApplyZoom(tab, bad);

                Assert.Equal(afterOneNotch, tab.schematicsMatrix);
            }
        });
    }

    [Fact]
    public void Zoom_does_nothing_until_a_schematic_image_is_loaded()
    {
        UiTest.Run(() =>
        {
            // No bitmap assigned: the viewer has nothing to zoom, and must not move the matrix.
            var tab = new TabSchematics();

            ApplyZoom(tab, AppConfig.SchematicsZoomFactor);

            Assert.Equal(Matrix.Identity, tab.schematicsMatrix);
        });
    }

    // ------------------------------------------------- Zoom anchoring

    // ###########################################################################################
    // The whole point of anchored zoom, checked on the axis that used to be wrong.
    //
    // TabSchematics is a split pane, so its schematics container is tall and narrow (398 x 600
    // here) while a landscape schematic fitted into it is 398 x 298.5 - it fills the width and
    // leaves an empty band below it. The clamp used to insist that band stayed below the image,
    // which forced the translation back to zero and slid the schematic out from under the
    // cursor vertically while behaving correctly horizontally.
    // ###########################################################################################
    [Theory]
    [InlineData(60, 40)]
    [InlineData(200, 200)]
    [InlineData(350, 30)]
    [InlineData(10, 290)]
    public void Zooming_in_keeps_the_same_bitmap_pixel_under_the_cursor(double x, double y)
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();
            var cursor = new Point(x, y);

            Point before = ImagePixelUnder(tab, cursor);

            for (int notch = 1; notch <= 6; notch++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor, cursor);

                Point after = ImagePixelUnder(tab, cursor);

                // Sub-pixel, not "close enough": a pad is only a handful of bitmap pixels
                // across, so drift of even one pixel per notch would compound into missing it.
                Assert.Equal(before.X, after.X, precision: 6);
                Assert.Equal(before.Y, after.Y, precision: 6);
            }
        });
    }

    [Fact]
    public void Zooming_back_out_keeps_the_same_bitmap_pixel_under_the_cursor()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            // Zoom in somewhere else first, so zooming out starts from a panned view rather
            // than from the fitted one where every anchor happens to agree.
            for (int i = 0; i < 5; i++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor, new Point(120, 240));
            }

            var cursor = new Point(300, 400);
            Point before = ImagePixelUnder(tab, cursor);

            for (int notch = 1; notch <= 3; notch++)
            {
                ApplyZoom(tab, 1.0 / AppConfig.SchematicsZoomFactor, cursor);

                Point after = ImagePixelUnder(tab, cursor);

                Assert.Equal(before.X, after.X, precision: 6);
                Assert.Equal(before.Y, after.Y, precision: 6);
            }
        });
    }

    // ###########################################################################################
    // The one place the anchor deliberately gives way, and the reason it has to.
    //
    // Zooming out shrinks the image about the cursor, so if the image is sitting with its top
    // far above the viewport, zooming out with the cursor high up drags its bottom edge upwards
    // and grows the empty band below it. Held exactly, that has no floor: a few notches of it
    // leaves a sliver of schematic at the top of an otherwise blank view.
    //
    // So the empty band is capped at the size the fitted first view already shows. Up to that
    // point zooming out holds the cursor exactly; past it the image stops rising and slides back
    // under the cursor instead - which is the same pull back towards the first view that zooming
    // all the way out ends in, just starting a little earlier.
    // ###########################################################################################
    [Fact]
    public void Zooming_out_never_shows_more_empty_space_than_the_first_view_does()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            Rect fitted = tab.GetImageContentRect();

            for (int i = 0; i < 5; i++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor, new Point(120, 240));
            }

            // High up the view, so every notch pulls the bottom edge of the image up with it.
            var cursor = new Point(300, 100);
            bool everReachedTheLimit = false;

            for (int notch = 1; notch <= 3; notch++)
            {
                ApplyZoom(tab, 1.0 / AppConfig.SchematicsZoomFactor, cursor);

                double bottomEdge =
                    tab.schematicsMatrix.M32 + (tab.schematicsMatrix.M11 * fitted.Bottom);

                Assert.True(
                    bottomEdge >= fitted.Bottom - 0.000001,
                    $"The empty band grew past the first view's on notch {notch}: {bottomEdge}.");

                everReachedTheLimit |= Math.Abs(bottomEdge - fitted.Bottom) < 0.000001;
            }

            // Guards the test itself: if the sequence above stopped short of the cap, it would
            // pass without ever exercising the rule it exists to pin down.
            Assert.True(everReachedTheLimit, "The cap was never actually reached.");
        });
    }

    // A portrait schematic letterboxes the other way round - it fills the height and leaves
    // empty space to its right - so this is the same rule on the horizontal axis.
    [Fact]
    public void Zooming_in_on_a_portrait_schematic_keeps_the_pixel_under_the_cursor()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage(new PixelSize(400, 1600));
            var cursor = new Point(100, 420);

            Point before = ImagePixelUnder(tab, cursor);

            for (int notch = 1; notch <= 6; notch++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor, cursor);

                Point after = ImagePixelUnder(tab, cursor);

                Assert.Equal(before.X, after.X, precision: 6);
                Assert.Equal(before.Y, after.Y, precision: 6);
            }
        });
    }

    [Fact]
    public void Zooming_with_the_cursor_in_the_empty_band_cannot_push_the_image_off_the_view()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            Rect fitted = tab.GetImageContentRect();

            // Below the image, in the empty band. There is nothing to hold under the cursor
            // there, so the rule that takes over is the limit: the bottom edge of the image may
            // never rise above where the fitted first view put it.
            var cursor = new Point(200, 550);

            for (int notch = 1; notch <= 8; notch++)
            {
                ApplyZoom(tab, AppConfig.SchematicsZoomFactor, cursor);

                double bottomEdge =
                    tab.schematicsMatrix.M32 + (tab.schematicsMatrix.M11 * fitted.Bottom);

                Assert.True(
                    bottomEdge >= fitted.Bottom - 0.000001,
                    $"The image was pushed up off the view on notch {notch}: {bottomEdge}.");
            }
        });
    }

    [Fact]
    public void Anchored_zooming_never_opens_a_gap_on_an_axis_the_image_fills()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateShownTabWithImage();

            Rect fitted = tab.GetImageContentRect();

            // The image fills the width, so however it is anchored there must never be blank
            // space at the left or right - that axis has no letterboxing to give away.
            foreach (Point cursor in new[] { new Point(5, 20), new Point(393, 280), new Point(200, 150) })
            {
                for (int notch = 1; notch <= 4; notch++)
                {
                    ApplyZoom(tab, AppConfig.SchematicsZoomFactor, cursor);

                    double left = tab.schematicsMatrix.M31 + (tab.schematicsMatrix.M11 * fitted.Left);
                    double right = tab.schematicsMatrix.M31 + (tab.schematicsMatrix.M11 * fitted.Right);

                    Assert.True(left <= 0.000001, $"Blank space appeared on the left: {left}.");
                    Assert.True(
                        right >= tab.SchematicsContainer.Bounds.Width - 0.000001,
                        $"Blank space appeared on the right: {right}.");
                }
            }
        });
    }

    // ###########################################################################################
    // The bitmap pixel currently under a container point, through the app's own mapping.
    // Fails the test rather than returning a flag: every point these tests use is over the
    // image, so a miss means the view moved somewhere unexpected.
    // ###########################################################################################
    private static Point ImagePixelUnder(TabSchematics tab, Point pointInContainer)
    {
        MethodInfo? method = typeof(TabSchematics).GetMethod(
            "TryGetSchematicsImagePixelPoint",
            BindingFlags.NonPublic | BindingFlags.Instance);

        Assert.NotNull(method);

        object[] arguments = { pointInContainer, new Point() };

        bool isOverTheImage = (bool)method!.Invoke(tab, arguments)!;

        Assert.True(isOverTheImage, $"{pointInContainer} is no longer over the schematic image.");

        return (Point)arguments[1];
    }

    // ###########################################################################################
    // A TabSchematics inside a shown window, laid out, with a schematic image loaded.
    //
    // The window and the layout pass matter: the clamping reads the container's bounds, so a
    // tab that was never measured would clamp against a zero-sized viewport.
    // ###########################################################################################
    private static TabSchematics CreateShownTabWithImage()
    {
        // Landscape, the common case: fitted into the split pane it fills the width and
        // letterboxes vertically.
        return CreateShownTabWithImage(new PixelSize(1200, 900));
    }

    private static TabSchematics CreateShownTabWithImage(PixelSize bitmapSize)
    {
        var tab = new TabSchematics();

        var window = new Window
        {
            Width = 800,
            Height = 600,
            Content = tab,
        };

        window.Show();
        Dispatcher.UIThread.RunJobs();

        tab.currentFullResBitmap = new WriteableBitmap(
            bitmapSize,
            new Vector(96, 96),
            PixelFormat.Bgra8888,
            AlphaFormat.Premul);

        return tab;
    }

    private static void ApplyZoom(TabSchematics tab, double zoomFactor)
    {
        ApplyZoom(tab, zoomFactor, ZoomCentre);
    }

    private static void ApplyZoom(TabSchematics tab, double zoomFactor, Point zoomCentre)
    {
        MethodInfo? method = typeof(TabSchematics).GetMethod(
            "ApplySchematicsZoom",
            BindingFlags.NonPublic | BindingFlags.Instance);

        Assert.NotNull(method);

        method!.Invoke(tab, new object[] { zoomFactor, zoomCentre });
    }
}
