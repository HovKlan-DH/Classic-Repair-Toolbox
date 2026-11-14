using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ROOM ROUND THE "Schematics" TAB'S IMAGE (owner request, 2026-11-13: "Ideally the cropped one can be displayed nicely and
// with some padding to its borders ... and then ideally the "original" one is displayed as-is today"). Through the real
// tab: the room added only where the image's own border is short of it - and the image itself standing exactly where
// every overlay and pointer finds it, or highlights would sit beside what they mark.
//
// The headless platform keeps no pixels in a bitmap, so the image's own border is handed over here (ShowImageForTests);
// reading it off the pixels is ImageFrameGeometryTests' (the rule) and the render harness's on the owner's real images.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SchematicsImageFrameTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SchematicsImageFrameTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks"));
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static (CRT.Main Main, TabSchematics Tab) Shown(Thickness blankBorder)
    {
        var main = new CRT.Main { Width = 1280, Height = 800 };
        main.Show();
        Dispatcher.UIThread.RunJobs();

        TabSchematics tab = main.TabSchematicsControl;
        var bitmap = new WriteableBitmap(new PixelSize(1200, 900), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        tab.ShowImageForTests(bitmap, blankBorder);
        main.UpdateLayout();
        Dispatcher.UIThread.RunJobs();

        return (main, tab);
    }

    private static void Close(CRT.Main main)
    {
        main.IsApplicationShuttingDownForTests = true;
        main.Close();
    }

    // Cut to the board: the room all round, the image standing inside it - where the overlays and pointers find it.
    [Fact]
    public void An_image_cut_to_its_board_stands_inside_the_room_round_it_where_the_overlays_find_it()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = Shown(blankBorder: default);

            Thickness frame = tab.ImageFrameForTests;
            Assert.Equal(ImageFrameGeometry.Wanted, frame);

            Rect content = tab.GetImageContentRect();
            Point imageAt = tab.SchematicsImage.TranslatePoint(default, tab.SchematicsImageHost)!.Value;

            Assert.Equal(frame.Left, content.Left, precision: 3);
            Assert.Equal(frame.Top, content.Top, precision: 3);
            Assert.Equal(content.Left, imageAt.X, precision: 1);
            Assert.Equal(content.Top, imageAt.Y, precision: 1);
            Assert.Equal(content.Width, tab.SchematicsImage.Bounds.Width, precision: 1);
            Assert.Equal(content.Height, tab.SchematicsImage.Bounds.Height, precision: 1);

            // Every overlay told the same room.
            Assert.Equal(frame, tab.SchematicsHighlightsOverlay.ContentFrame);
            Assert.Equal(frame, tab.SchematicsHoverHighlightsOverlay.ContentFrame);
            Assert.Equal(frame, tab.SchematicsWorklogEntriesOverlay.ContentFrame);
            Assert.Equal(frame, tab.SchematicsWorklogEntryOverlay.ContentFrame);
            Assert.Equal(frame, tab.SchematicsLabelEditorOverlay.ContentFrame);

            Close(main);
        });
    }

    // Room of its own round the board, more than the badges need - as the owner's own images have: nothing added, the image
    // standing where it always did.
    [Fact]
    public void An_image_with_room_of_its_own_stands_where_it_always_did()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = Shown(blankBorder: new Thickness(120, 250, 120, 120));

            Assert.Equal(default, tab.ImageFrameForTests);

            Point imageAt = tab.SchematicsImage.TranslatePoint(default, tab.SchematicsImageHost)!.Value;
            Assert.Equal(new Point(0, 0), imageAt);
            Assert.Equal(new Point(0, 0), tab.GetImageContentRect().TopLeft);

            Close(main);
        });
    }

    // The panel resized: the room follows its new size, and the image with it.
    [Fact]
    public void The_room_follows_the_panels_size()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = Shown(blankBorder: new Thickness(0, 30, 0, 0));
            double before = tab.ImageFrameForTests.Top;

            main.Width = 900;
            main.Height = 500;
            main.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            // Drawn smaller, the image's own 30 rows give less, so more is added - up to the 44 wanted.
            double after = tab.ImageFrameForTests.Top;
            Assert.True(after > before, $"{after} after, {before} before");
            Assert.True(after <= ImageFrameGeometry.TopRoom);
            Assert.Equal(after, tab.SchematicsImage.TranslatePoint(default, tab.SchematicsImageHost)!.Value.Y, precision: 1);

            Close(main);
        });
    }
}
