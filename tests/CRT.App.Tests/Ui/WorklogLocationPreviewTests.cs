using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The full editor's "Worklog location" preview draws the entry's marked area over a scaled-down
// schematic. Its position is computed from the preview grid's OWN size, so the redraw must be
// driven by that grid - not by the window.
//
// The bug this pins: the refresh was wired to the WINDOW's SizeChanged. Dragging the editor's
// GridSplitter re-widths the preview column while the window's size never changes, so no event
// fired and the marker kept coordinates computed for the old width. It drifted away from the area
// it marks and only corrected itself when the window itself was resized.
[Collection("HeadlessUi")]
public class WorklogLocationPreviewTests
{
    // A real Bitmap, built in memory - the preview maths needs a source PixelSize and nothing
    // more, and this keeps the test off the filesystem.
    private static Bitmap CreateBitmap(int width, int height)
    {
        // WriteableBitmap needs no encoder round-trip, so nothing here is deprecated and no
        // pixels have to be produced - the preview maths reads PixelSize and nothing else.
        return new WriteableBitmap(
            new PixelSize(width, height),
            new Vector(96, 96),
            PixelFormat.Bgra8888,
            AlphaFormat.Premul);
    }

    private static WorklogEntryRecord CreateEntry() => new()
    {
        Id = 1,
        SchematicName = "Sch",
        Title = "Bad cap",
        Category = "Issue",
        State = "Open",

        // A quarter-size area sitting away from the origin, so any scaling error shows up as a
        // position error rather than being masked by a zero offset.
        AreaX = 200,
        AreaY = 100,
        AreaWidth = 100,
        AreaHeight = 50,
    };

    // Finds the marker the overlay draws, if it drew one.
    private static Rectangle? FindMarker(Window window)
    {
        var canvas = window.GetVisualDescendants()
            .OfType<Canvas>()
            .FirstOrDefault(c => c.Name == "EditorLocationPreviewOverlayCanvas");

        return canvas?.Children.OfType<Rectangle>().FirstOrDefault();
    }

    // Resizing ONLY the preview grid - exactly what the splitter does - must move the marker.
    // Before the fix the grid's own SizeChanged was not subscribed, so the marker stayed put.
    [Fact]
    public void Resizing_the_preview_area_alone_repositions_the_location_marker()
    {
        double? widthBefore = null;
        double? widthAfter = null;

        UiTest.Run(() =>
        {
            // Placement persistence off for this window only: it otherwise restores the size,
            // position and splitter ratio from the developer's REAL settings file, so every layout
            // assertion below would depend on how they last left the editor. Scoped rather than
            // assigned, so it cannot leak into other tests in the shared session.
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 900;
            window.Height = 700;

            using var bitmap = CreateBitmap(400, 200);
            window.Initialize(1, CreateEntry(), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                var grid = window.GetVisualDescendants()
                    .OfType<Grid>()
                    .First(g => g.Name == "EditorLocationPreviewGrid");

                // Pin the preview to a known width, let it lay out, and record the marker.
                grid.Width = 400;
                grid.Height = 200;
                Dispatcher.UIThread.RunJobs();
                widthBefore = FindMarker(window)?.Width;

                // Halve it - the window is never touched, which is the whole point.
                grid.Width = 200;
                Dispatcher.UIThread.RunJobs();
                widthAfter = FindMarker(window)?.Width;
            }
            finally
            {
                window.Close();
            }
        });

        Assert.NotNull(widthBefore);
        Assert.NotNull(widthAfter);

        // Halving the preview width halves the marker: it tracks the area it marks instead of
        // keeping a size computed for a width that no longer exists.
        Assert.True(
            widthAfter!.Value < widthBefore!.Value,
            $"marker did not shrink with the preview: before={widthBefore}, after={widthAfter}");
    }

    // ------------------------------------------------------------- shrink-wrapped frame

    // The preview box used to be a fixed 220px tall with a Stretch="Uniform" image centred in it,
    // so a wide schematic sat in a tall box padded top and bottom with empty Form_Bg - on a
    // typical board image, most of those 220px were blank. The border now shrink-wraps the image
    // the way the Photos thumbnails do, so the 1px frame lands on the picture rather than around
    // a mostly-empty box.
    //
    // Asserted as "the frame is no taller than the image plus its own border" rather than against
    // a pixel figure, so the test states the rule instead of a measurement that a font or padding
    // change would invalidate.
    [Fact]
    public void The_location_preview_frame_does_not_add_vertical_space_around_the_image()
    {
        double imageHeight = 0;
        double frameHeight = 0;

        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            // Deliberately WIDE and short - the aspect ratio that letterboxed worst in the fixed
            // 220px box, and therefore the one that proves the box is gone.
            using var bitmap = CreateBitmap(800, 200);
            window.Initialize(1, CreateEntry(), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                var image = window.GetVisualDescendants()
                    .OfType<Image>()
                    .First(i => i.Name == "EditorLocationPreviewImage");

                var frame = (Border)window.GetVisualDescendants()
                    .OfType<Grid>()
                    .First(g => g.Name == "EditorLocationPreviewGrid")
                    .GetVisualParent()!;

                imageHeight = image.Bounds.Height;
                frameHeight = frame.Bounds.Height;
            }
            finally
            {
                window.Close();
            }
        });

        Assert.True(imageHeight > 0, "the preview image was not laid out");

        // 1px border top and bottom is the only extra the frame may add.
        Assert.True(
            frameHeight <= imageHeight + 2.5,
            $"frame is {frameHeight}px around a {imageHeight}px image - it is padding the image out again");
    }

    // The MaxHeight still applies, so a very TALL schematic cannot push the right-hand panel into
    // a scroll of its own.
    [Fact]
    public void A_tall_schematic_is_capped_rather_than_growing_without_limit()
    {
        double imageHeight = 0;

        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap(200, 2000);
            window.Initialize(1, CreateEntry(), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                imageHeight = window.GetVisualDescendants()
                    .OfType<Image>()
                    .First(i => i.Name == "EditorLocationPreviewImage")
                    .Bounds.Height;
            }
            finally
            {
                window.Close();
            }
        });

        Assert.True(imageHeight > 0, "the preview image was not laid out");
        Assert.True(imageHeight <= 220.5, $"a tall schematic grew to {imageHeight}px, past the 220px cap");
    }
}
