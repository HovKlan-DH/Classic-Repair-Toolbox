using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The worklog area-marking flow on the Schematics tab: "Add worklog" starts a mode, a drag on
// the schematic defines the area, and a deliberate drag is accepted while an accidental click
// is discarded.
//
// WHY THIS FLOW SPECIFICALLY. The badge placement it feeds has been reported as a bug TWICE -
// once against the Workbooks board pane and again against the Schematics thumbnails, both times
// the same "anchored where it should be parked" defect. That makes this the flow in this tab
// most likely to break unnoticed, and it had no coverage at all.
//
// WHERE THESE STOP. CompleteDrawingWorklogEntryRectangle ends by opening the full editor through
// ShowDialog, which needs a real owner Window and cannot run headlessly. The accept/reject
// decision was therefore split into TryFinishWorklogEntryDrawing, which BOTH the shipped path
// and the test seam call - so the rule under test is the shipped rule, with only the modal
// skipped. The editor handoff itself is still verified only by running the app.
//
// Everything is in BITMAP PIXELS, the space the pointer handlers convert to before calling these
// methods - the same arrangement as LabelEditorInteractionTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public class WorklogAreaMarkingTests : IDisposable
{
    private const string Schematic = "Board top";
    private const int WorkbookId = 7;

    private readonly TempWorkspace thisWorkspace = new();

    // BeginWorklogEntryMode calls WorklogManager.PeekNextEntryId, which reads the workbook folder.
    // Pointed at a temp root so no test can touch the user's real Workbook directory.
    public WorklogAreaMarkingTests()
    {
        this.RedirectWorklogToTemp();
    }

    public void Dispose()
    {
        this.RedirectWorklogToTemp();
        this.thisWorkspace.Dispose();
    }

    private void RedirectWorklogToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
    }

    // -----------------------------------------------------------------------------------------
    // Entering and leaving the mode
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Starting_entry_mode_enters_it_for_the_given_workbook()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            Assert.True(tab.BeginWorklogEntryMode(WorkbookId));

            Assert.True(tab.IsWorklogEntryModeForTests);
            Assert.Equal(WorkbookId, tab.WorklogEntryWorkbookIdForTests);
        });
    }

    // The mode marks an area ON a schematic image. With no image loaded there is nothing to mark,
    // and the caller needs to be told so it can leave its buttons alone.
    [Fact]
    public void Entry_mode_refuses_to_start_with_no_schematic_image_loaded()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab(withBitmap: false);

            Assert.False(tab.BeginWorklogEntryMode(WorkbookId));
            Assert.False(tab.IsWorklogEntryModeForTests);
        });
    }

    // The label editor owns the same pointer handlers. Starting area-marking over it would produce
    // a mode that looks active and can never receive input - the exclusion documented on
    // BeginWorklogEntryMode, whose other half is BeginLabelEditorMode cancelling this mode.
    [Fact]
    public void Entry_mode_refuses_to_start_while_the_label_editor_is_active()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.CurrentBoardDataOverrideForTests = new BoardData();
            tab.BeginLabelEditorModeForTests();

            Assert.False(tab.BeginWorklogEntryMode(WorkbookId));
            Assert.False(tab.IsWorklogEntryModeForTests);
        });
    }

    // The other half of that exclusion: opening the label editor tears down an active marking mode
    // rather than leaving two modes fighting over the pointer.
    [Fact]
    public void Opening_the_label_editor_cancels_an_active_entry_mode()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.CurrentBoardDataOverrideForTests = new BoardData();
            tab.BeginWorklogEntryMode(WorkbookId);

            tab.BeginLabelEditorModeForTests();

            Assert.False(tab.IsWorklogEntryModeForTests);
        });
    }

    [Fact]
    public void Cancelling_entry_mode_leaves_it_and_clears_the_drawn_area()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);
            DrawArea(tab, new Point(100, 100), new Point(200, 180));

            tab.CancelWorklogEntryMode();

            Assert.False(tab.IsWorklogEntryModeForTests);
            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
            Assert.Null(tab.WorklogEntryDraftRectangleForTests);
        });
    }

    // Documented as safe to call when the mode is not running - it is reached from Escape, the
    // top bar and the editor closing, and any of those can arrive when the mode is already gone.
    [Fact]
    public void Cancelling_when_not_in_entry_mode_is_harmless()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            tab.CancelWorklogEntryMode();

            Assert.False(tab.IsWorklogEntryModeForTests);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Drawing the area
    // -----------------------------------------------------------------------------------------

    // The rubber band follows the pointer while the button is down, and is not yet the final area.
    [Fact]
    public void Dragging_shows_a_draft_rectangle_that_follows_the_pointer()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            tab.StartDrawingWorklogEntryRectangleForTests(new Point(100, 100));
            tab.UpdateDrawingWorklogEntryRectangleForTests(new Point(200, 180));

            Assert.True(tab.IsDrawingWorklogEntryRectangleForTests);
            Assert.Equal(new Rect(100, 100, 100, 80), tab.WorklogEntryDraftRectangleForTests);

            // Not committed until the button is released.
            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
        });
    }

    [Fact]
    public void Releasing_after_a_deliberate_drag_records_the_marked_area()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            bool accepted = DrawArea(tab, new Point(100, 100), new Point(200, 180));

            Assert.True(accepted);
            Assert.Equal(new Rect(100, 100, 100, 80), tab.WorklogEntryFinalRectangleForTests);
            Assert.False(tab.IsDrawingWorklogEntryRectangleForTests);
            Assert.Null(tab.WorklogEntryDraftRectangleForTests);
        });
    }

    // Dragging up-and-left is the same rectangle as dragging down-and-right. Without normalising,
    // the drag produces a negative width, which Rect represents mirrored about its origin.
    [Fact]
    public void Dragging_backwards_produces_the_same_rectangle()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            DrawArea(tab, new Point(200, 180), new Point(100, 100));

            Assert.Equal(new Rect(100, 100, 100, 80), tab.WorklogEntryFinalRectangleForTests);
        });
    }

    // A plain click is not an area. Accepting it would open the full editor on a zero-sized
    // rectangle every time the user clicked the schematic while the mode was on.
    [Fact]
    public void A_click_without_a_drag_is_discarded()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            bool accepted = DrawArea(tab, new Point(100, 100), new Point(100, 100));

            Assert.False(accepted);
            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
        });
    }

    // The 15x15 threshold (LabelEditorGeometry.IsLabelEditorRectangleTooSmall, shared with the
    // label editor). A tiny drag is a slipped click, not an intentional area.
    [Fact]
    public void A_drag_below_the_minimum_size_is_discarded()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            bool accepted = DrawArea(tab, new Point(100, 100), new Point(110, 110));

            Assert.False(accepted);
            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
        });
    }

    // A drag wide enough but only a few pixels tall is still not an area - the rule tests both
    // dimensions, not just the total. A sliver rectangle marks nothing useful on a schematic.
    [Fact]
    public void A_wide_but_flat_drag_is_discarded()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            bool accepted = DrawArea(tab, new Point(100, 100), new Point(300, 105));

            Assert.False(accepted);
            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
        });
    }

    // Just over the threshold on both axes must be ACCEPTED - the companion to the rejection
    // tests above. Without this pair, a rule that rejected everything would still pass.
    [Fact]
    public void A_drag_just_over_the_minimum_size_is_accepted()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            bool accepted = DrawArea(tab, new Point(100, 100), new Point(116, 116));

            Assert.True(accepted);
            Assert.Equal(new Rect(100, 100, 16, 16), tab.WorklogEntryFinalRectangleForTests);
        });
    }

    // A release without a preceding press is not a drag - it happens when the press landed
    // somewhere that did not start drawing.
    [Fact]
    public void Releasing_without_having_started_a_drag_does_nothing()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            Assert.False(tab.CompleteWorklogEntryDrawingWithoutEditorForTests(new Point(200, 200)));
            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
        });
    }

    // An update outside a drag must not invent a draft rectangle - the mode is live and the
    // pointer moves across the schematic constantly before the user presses anything.
    [Fact]
    public void Moving_the_pointer_without_a_drag_creates_no_draft()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);

            tab.UpdateDrawingWorklogEntryRectangleForTests(new Point(300, 300));

            Assert.Null(tab.WorklogEntryDraftRectangleForTests);
        });
    }

    // Re-entering the mode clears a previously marked area, so the second entry does not start
    // showing the first one's rectangle.
    [Fact]
    public void Restarting_entry_mode_clears_the_previous_area()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginWorklogEntryMode(WorkbookId);
            DrawArea(tab, new Point(100, 100), new Point(200, 180));

            tab.BeginWorklogEntryMode(WorkbookId + 1);

            Assert.Null(tab.WorklogEntryFinalRectangleForTests);
            Assert.Equal(WorkbookId + 1, tab.WorklogEntryWorkbookIdForTests);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Fixture
    // -----------------------------------------------------------------------------------------

    // Presses, drags and releases in one call. Returns whether the area was accepted.
    private static bool DrawArea(TabSchematics tab, Point from, Point to)
    {
        tab.StartDrawingWorklogEntryRectangleForTests(from);
        tab.UpdateDrawingWorklogEntryRectangleForTests(to);
        return tab.CompleteWorklogEntryDrawingWithoutEditorForTests(to);
    }

    // A tab with a schematic "loaded": a WriteableBitmap standing in for the decoded image, which
    // is all BeginWorklogEntryMode needs (it checks for null and reads PixelSize). WriteableBitmap
    // works headlessly - no display and no encoder involved.
    private static TabSchematics CreateTab(bool withBitmap = true)
    {
        var tab = new TabSchematics();

        if (withBitmap)
        {
            tab.currentFullResBitmap = new WriteableBitmap(
                new PixelSize(1000, 800),
                new Vector(96, 96),
                PixelFormat.Bgra8888,
                AlphaFormat.Premul);
        }

        var thumbnail = new SchematicThumbnail { Name = Schematic };
        tab.SchematicsThumbnailList.ItemsSource = new List<SchematicThumbnail> { thumbnail };
        tab.SchematicsThumbnailList.SelectedItem = thumbnail;

        return tab;
    }
}
