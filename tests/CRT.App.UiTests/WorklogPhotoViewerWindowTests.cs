using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The photo viewer zooms (owner request, 2026-11-06: "when I click a photo in here and in the worklog, that I then can
// zoom in/out with scrollwheel") - WorklogPhotoViewerWindow's PhotoZoomView, opened from a worklog's photos, the device
// form's and a device's page. The wheel zooms at the pointer, never past whole on the way out; a double-click shows the
// photo whole again. Where each point lands is PicturePreviewGeometry's, tested on its own.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class WorklogPhotoViewerWindowTests
{
    private static void WithViewer(Action<WorklogPhotoViewerWindow, PhotoZoomView> body)
    {
        UiTest.Run(() =>
        {
            var viewer = new WorklogPhotoViewerWindow { Width = 900, Height = 700 };
            viewer.Show();

            try
            {
                PhotoZoomView view = viewer.GetControl<PhotoZoomView>("PhotoImage");
                view.Source = new WriteableBitmap(new PixelSize(400, 300), new Vector(96, 96));
                Dispatcher.UIThread.RunJobs();
                viewer.UpdateLayout();
                body(viewer, view);
            }
            finally
            {
                viewer.Close();
            }
        });
    }

    private static Point Middle(WorklogPhotoViewerWindow viewer, PhotoZoomView view) =>
        view.TranslatePoint(new Point(view.Bounds.Width / 2, view.Bounds.Height / 2), viewer)!.Value;

    [Fact]
    public void The_wheel_zooms_the_photo_in_and_out_never_past_whole()
    {
        WithViewer((viewer, view) =>
        {
            Assert.Equal(1.0, view.ZoomForTests);

            viewer.MouseWheel(Middle(viewer, view), new Vector(0, 1));
            Assert.Equal(1.25, view.ZoomForTests, 6);

            viewer.MouseWheel(Middle(viewer, view), new Vector(0, 2));
            Assert.Equal(Math.Pow(1.25, 3), view.ZoomForTests, 6);

            viewer.MouseWheel(Middle(viewer, view), new Vector(0, -10));
            Assert.Equal(1.0, view.ZoomForTests);
        });
    }

    [Fact]
    public void A_double_click_shows_the_photo_whole_again()
    {
        WithViewer((viewer, view) =>
        {
            Point middle = Middle(viewer, view);
            viewer.MouseWheel(middle, new Vector(0, 3));
            Assert.True(view.ZoomForTests > 1.0);

            viewer.MouseDown(middle, MouseButton.Left);
            viewer.MouseUp(middle, MouseButton.Left);
            viewer.MouseDown(middle, MouseButton.Left);
            viewer.MouseUp(middle, MouseButton.Left);

            Assert.Equal(1.0, view.ZoomForTests);
        });
    }

    // ###########################################################################################
    // THE ZOOMED PHOTO IS MOVED AS THE SCHEMATICS TAB MOVES ITS PICTURE (owner request, 2026-11-06: "unify how pan works,
    // and then have the same pan functionality as present in the 'Schematics' tab"; PanGesture): a drag with the right or
    // the middle button moves it, a left drag does not.
    // ###########################################################################################
    [Theory]
    [InlineData(MouseButton.Right, RawInputModifiers.RightMouseButton, true)]
    [InlineData(MouseButton.Middle, RawInputModifiers.MiddleMouseButton, true)]
    [InlineData(MouseButton.Left, RawInputModifiers.LeftMouseButton, false)]
    public void Only_a_right_or_middle_drag_moves_the_zoomed_photo(MouseButton button, RawInputModifiers held, bool moves)
    {
        WithViewer((viewer, view) =>
        {
            Point middle = Middle(viewer, view);
            viewer.MouseWheel(middle, new Vector(0, 4));

            Point centreBefore = view.CentreForTests;

            viewer.MouseDown(middle, button);
            viewer.MouseMove(middle + new Point(80, 40), held);
            viewer.MouseUp(middle + new Point(80, 40), button);

            Assert.Equal(moves, view.CentreForTests != centreBefore);
            Assert.True(view.ZoomForTests > 1.0);
        });
    }

    // ###########################################################################################
    // A drag the pointer is taken from - the window leaving the front, another control capturing it - ends as a release
    // does: the four arrows go, and the hand of a zoomed photo is back (code review, 2026-11-08: the four arrows stayed).
    // ###########################################################################################
    [Fact]
    public void A_drag_that_loses_the_pointer_gives_the_hand_cursor_back()
    {
        WithViewer((viewer, view) =>
        {
            Point middle = Middle(viewer, view);
            viewer.MouseWheel(middle, new Vector(0, 2));

            IPointer? pointer = null;
            viewer.AddHandler(InputElement.PointerPressedEvent, (_, e) => pointer = e.Pointer, Avalonia.Interactivity.RoutingStrategies.Tunnel, handledEventsToo: true);

            viewer.MouseDown(middle, MouseButton.Right);
            Assert.Equal(new Cursor(StandardCursorType.SizeAll).ToString(), view.Cursor?.ToString());

            pointer!.Capture(null);

            Assert.Equal(new Cursor(StandardCursorType.Hand).ToString(), view.Cursor?.ToString());

            // And it moves nothing any more.
            Point centre = view.CentreForTests;
            viewer.MouseMove(middle + new Point(60, 30), RawInputModifiers.RightMouseButton);
            Assert.Equal(centre, view.CentreForTests);
            viewer.MouseUp(middle + new Point(60, 30), MouseButton.Right);
        });
    }

    // ###########################################################################################
    // The photo's comment is a note with formatting (RichNote) - shown as it reads, never with its marks as typed (code
    // review, 2026-11-08: "**dead** SID" showed with its stars).
    // ###########################################################################################
    [Fact]
    public void The_comment_is_shown_as_it_reads_never_with_its_marks()
    {
        UiTest.Run(() =>
        {
            var viewer = new WorklogPhotoViewerWindow();
            viewer.Initialize("board.jpg", "**dead** SID\n", imagePath: null);

            TextBlock comment = viewer.GetControl<TextBlock>("CommentText");

            Assert.True(comment.IsVisible);
            Assert.Equal("dead SID", Handlers.Theming.BoldRuns.TextOf(comment));
            Assert.Equal(["dead"], Handlers.Theming.BoldRuns.BoldPiecesOf(comment));
        });
    }

    // A photo that could not be read is no photo to zoom: the wheel leaves the viewer as it is.
    [Fact]
    public void With_no_photo_the_wheel_does_nothing()
    {
        UiTest.Run(() =>
        {
            var viewer = new WorklogPhotoViewerWindow { Width = 900, Height = 700 };
            viewer.Initialize("gone.jpg", string.Empty, Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".jpg"));
            viewer.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                viewer.UpdateLayout();
                PhotoZoomView view = viewer.GetControl<PhotoZoomView>("PhotoImage");

                viewer.MouseWheel(Middle(viewer, view), new Vector(0, 1));

                Assert.Equal(1.0, view.ZoomForTests);
                Assert.True(viewer.GetControl<TextBlock>("PhotoUnavailableText").IsVisible);
            }
            finally
            {
                viewer.Close();
            }
        });
    }
}
