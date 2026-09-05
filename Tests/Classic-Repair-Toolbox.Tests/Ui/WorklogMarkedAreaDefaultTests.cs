using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// Ticking "Show marked area" on a worklog that never had one.
//
// A worklog created from the oscilloscope capture flow is stored with NO area at all and parks as a
// "#N" pill in the schematic panel's corner - it was born at the bench with a probe in hand, not by
// dragging a rectangle. Ticking the box on such an entry used to leave it with a zero-sized rect,
// which draws as nothing or as a hairline and can never be grabbed and dragged into place: the
// entry looked broken with no way to fix it from the UI.
//
// These tests fail against that version. They also pin the other half of the promise - that an
// entry which ALREADY has a drawn area keeps it untouched through any tick/untick cycle, since
// silently moving a user's own marked area would be far worse than the bug being fixed.
[Collection("HeadlessUi")]
public sealed class WorklogMarkedAreaDefaultTests
{
    private const int BitmapWidth = 4000;
    private const int BitmapHeight = 3000;

    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(
            new PixelSize(BitmapWidth, BitmapHeight), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    /// <summary>An entry with no drawn area - what the oscilloscope capture flow creates.</summary>
    private static WorklogEntryRecord AreaLessEntry() => new()
    {
        Id = 7,
        SchematicName = "Sch",
        Title = "Pin 14 low",
        Category = "Issue",
        State = "Open",
        ShowMarkedArea = false,
    };

    private static void WithEditor(WorklogEntryRecord entry, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, entry, bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
            }
        });
    }

    private static CheckBox ShowMarkedAreaBox(WorklogEntryEditorWindow window) =>
        window.GetControl<CheckBox>("EditorShowMarkedAreaCheckBox");

    // The headline: ticking the box on an area-less entry must leave it with a rectangle that can
    // actually be seen and grabbed, not a zero-sized one.
    [Fact]
    public void Ticking_show_marked_area_gives_an_area_less_worklog_a_real_rectangle()
    {
        WithEditor(AreaLessEntry(), window =>
        {
            Assert.True(WorklogDefaultAreaGeometry.IsUnset(window.WorkingEntryAreaForTests));

            ShowMarkedAreaBox(window).IsChecked = true;
            Dispatcher.UIThread.RunJobs();

            var area = window.WorkingEntryAreaForTests;

            Assert.False(WorklogDefaultAreaGeometry.IsUnset(area));

            // Inside the board, so none of it is unreachable.
            Assert.True(area.X >= 0 && area.Y >= 0);
            Assert.True(area.Right <= BitmapWidth && area.Bottom <= BitmapHeight);
        });
    }

    // Bottom-right, the opposite corner from the parked pills, so a freshly placed area can never be
    // mistaken for one of them while the user is moving it into position.
    [Fact]
    public void The_new_area_appears_in_the_bottom_right_corner_away_from_the_parked_pills()
    {
        WithEditor(AreaLessEntry(), window =>
        {
            ShowMarkedAreaBox(window).IsChecked = true;
            Dispatcher.UIThread.RunJobs();

            var area = window.WorkingEntryAreaForTests;

            Assert.True(area.X > BitmapWidth / 2.0);
            Assert.True(area.Y > BitmapHeight / 2.0);
        });
    }

    // An entry the user actually drew keeps its own rectangle. Silently relocating a marked area
    // would be a worse bug than the one being fixed, so this is asserted to the exact value.
    [Fact]
    public void An_entry_that_already_has_an_area_is_never_moved()
    {
        var drawn = AreaLessEntry();
        drawn.AreaX = 120;
        drawn.AreaY = 340;
        drawn.AreaWidth = 200;
        drawn.AreaHeight = 150;
        drawn.ShowMarkedArea = true;

        WithEditor(drawn, window =>
        {
            // A full untick/retick cycle, which is the sequence that would expose any re-placement.
            ShowMarkedAreaBox(window).IsChecked = false;
            Dispatcher.UIThread.RunJobs();
            ShowMarkedAreaBox(window).IsChecked = true;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(new Rect(120, 340, 200, 150), window.WorkingEntryAreaForTests);
        });
    }

    // Unticking never invents anything: a parked entry that is toggled off stays area-less, so it
    // keeps parking rather than acquiring a rectangle nobody asked for.
    [Fact]
    public void Unticking_leaves_an_area_less_worklog_alone()
    {
        WithEditor(AreaLessEntry(), window =>
        {
            ShowMarkedAreaBox(window).IsChecked = false;
            Dispatcher.UIThread.RunJobs();

            Assert.True(WorklogDefaultAreaGeometry.IsUnset(window.WorkingEntryAreaForTests));
        });
    }

    // The seam the capture flow uses to park a brand-new worklog: it must set BOTH the checkbox and
    // the record, or the editor's own save would write back the markup default of "ticked" and the
    // entry would promise a rectangle it does not have.
    [Fact]
    public void A_new_worklog_can_be_created_parked_rather_than_marked()
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            window.InitializeForNewEntry(1, "Sch", default, null);
            window.SetShowMarkedAreaForNewEntry(false);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Assert.False(ShowMarkedAreaBox(window).IsChecked);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Why the CALLER must hand over a schematic bitmap, and the shape of the bug when it does not.
    //
    // EnsureMarkedAreaExistsWhenShown needs the board's pixel size to place a square on it, so with
    // no bitmap it can only leave the entry as it found it. That is the honest outcome here - there
    // is no board to place anything on - but it means a caller passing null silently opts the entry
    // out of the whole feature. The capture flow did exactly that, which reproduced the original
    // zero-sized-rect bug for the ONE entry kind this geometry was written for; ComponentInfoWindow
    // now resolves the tab's own full-resolution bitmap instead.
    [Fact]
    public void With_no_schematic_bitmap_ticking_the_box_cannot_invent_an_area()
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            window.Initialize(1, AreaLessEntry(), null);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                ShowMarkedAreaBox(window).IsChecked = true;
                Dispatcher.UIThread.RunJobs();

                Assert.True(WorklogDefaultAreaGeometry.IsUnset(window.WorkingEntryAreaForTests));
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The editor must reach its answer THROUGH WorklogDefaultAreaGeometry.ResolveAreaForShowing,
    // which documents itself as the single decision point for "does this entry need an area
    // inventing". It used to call IsUnset and BuildDefaultArea separately and re-derive the branch,
    // so that guarantee was false: a second caller following the comment and the editor following
    // its own copy could disagree. This asserts the editor's result IS that method's result.
    [Fact]
    public void The_editor_produces_exactly_what_the_shared_resolver_decides()
    {
        WithEditor(AreaLessEntry(), window =>
        {
            ShowMarkedAreaBox(window).IsChecked = true;
            Dispatcher.UIThread.RunJobs();

            var expected = WorklogDefaultAreaGeometry.ResolveAreaForShowing(
                default, new Size(BitmapWidth, BitmapHeight));

            Assert.Equal(expected, window.WorkingEntryAreaForTests);

            // Asserted against BuildDefaultArea as well, so this cannot pass by both sides moving
            // together: the comparison above alone stays green if the resolver is changed to return
            // something arbitrary, because the editor then returns that same arbitrary value.
            Assert.Equal(
                WorklogDefaultAreaGeometry.BuildDefaultArea(new Size(BitmapWidth, BitmapHeight)),
                window.WorkingEntryAreaForTests);
        });
    }
}
