using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// That the Schematics tab's worklog overlay and thumbnail pills REDRAW after an entry is edited
// in the workbook already on screen.
//
// The reported bug: adding a marker, or ticking "Show marked area", did not always show up on the
// schematic image or the thumbnails. The cause was in Main.RefreshWorklogBar - the single funnel
// every worklog change passes through - which only called SetShowWorklogEntriesList inside an
// "if the shown WORKBOOK changed" branch. Editing an entry inside the workbook already displayed
// changes neither of that method's two arguments, so the overlay kept drawing the pre-edit record
// until something else happened to rebuild it. Editing from the WORKBOOKS tab was the worst case:
// its save path funnels through RefreshWorklogBar alone, so nothing redrew the schematic overlay
// at all.
//
// These tests drive TabSchematics.RefreshWorklogEntriesListForCurrentWorkbook - the method that
// branch now calls - and assert the overlay actually re-reads from disk. They fail against a
// version where that method does not exist or does not re-read.
//
// Everything is in BITMAP PIXELS, the space the overlay stores its rects in.
// ###########################################################################################
[Collection("HeadlessUi")]
public class WorklogOverlayRefreshTests : IDisposable
{
    private const string Schematic = "Board top";

    private readonly TempWorkspace thisWorkspace = new();

    public WorklogOverlayRefreshTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-detached"));
        this.thisWorkspace.Dispose();
    }

    private static TabSchematics CreateTab()
    {
        var tab = new TabSchematics
        {
            currentFullResBitmap = new WriteableBitmap(
                new PixelSize(1000, 800), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul)
        };

        var thumbnail = new SchematicThumbnail { Name = Schematic };
        tab.SchematicsThumbnailList.ItemsSource = new List<SchematicThumbnail> { thumbnail };
        tab.SchematicsThumbnailList.SelectedItem = thumbnail;

        return tab;
    }

    private static (int WorkbookId, int EntryId) CreateWorkbookWithEntry(Rect area, bool showMarkedArea)
    {
        var workbook = WorklogManager.CreateWorkbook("Commodore 64|250469", "Test", "");
        Assert.NotNull(workbook);

        var entry = WorklogManager.AddEntry(
            workbook!.Id, Schematic, area, "Bad cap", "", "Issue", "Open", Array.Empty<string>());
        Assert.NotNull(entry);

        entry!.ShowMarkedArea = showMarkedArea;
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        return (workbook.Id, entry.Id);
    }

    // The headline: an entry's area is changed on disk, and a refresh for the SAME workbook must
    // pick it up. Before the fix nothing called this on an entry edit, so the overlay kept the old
    // rectangle.
    [Fact]
    public void Editing_an_entrys_area_redraws_the_overlay_for_the_same_workbook()
    {
        UiTest.Run(() =>
        {
            var (workbookId, entryId) = CreateWorkbookWithEntry(new Rect(10, 20, 100, 80), showMarkedArea: true);

            var tab = CreateTab();
            tab.SetShowWorklogEntriesList(true, workbookId);

            var before = Assert.Single(tab.SchematicsWorklogEntriesOverlay.Entries);
            Assert.Equal(new Rect(10, 20, 100, 80), before.PixelRect);

            // The edit a user makes by dragging the marker to a new place.
            var entry = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);
            entry.AreaX = 400;
            entry.AreaY = 300;
            Assert.True(WorklogManager.UpdateEntry(workbookId, entry));

            // The workbook has NOT changed - this is the call the "same workbook" branch makes.
            tab.RefreshWorklogEntriesListForCurrentWorkbook();

            var after = Assert.Single(tab.SchematicsWorklogEntriesOverlay.Entries);
            Assert.Equal(new Rect(400, 300, 100, 80), after.PixelRect);
        });
    }

    // Ticking "Show marked area" turns a parked entry into a drawn rectangle. The overlay holds only
    // entries WITH a marked area (a parked one draws a corner pill instead), so this is observable
    // as the overlay going from empty to holding one.
    [Fact]
    public void Ticking_show_marked_area_makes_the_rectangle_appear_on_the_schematic()
    {
        UiTest.Run(() =>
        {
            var (workbookId, entryId) = CreateWorkbookWithEntry(new Rect(10, 20, 100, 80), showMarkedArea: false);

            var tab = CreateTab();
            tab.SetShowWorklogEntriesList(true, workbookId);

            Assert.Empty(tab.SchematicsWorklogEntriesOverlay.Entries);

            var entry = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);
            entry.ShowMarkedArea = true;
            Assert.True(WorklogManager.UpdateEntry(workbookId, entry));

            tab.RefreshWorklogEntriesListForCurrentWorkbook();

            Assert.Single(tab.SchematicsWorklogEntriesOverlay.Entries);
        });
    }

    // And the other way: unticking must REMOVE the rectangle, not leave it behind. A stale
    // rectangle for an entry the user just hid is the same defect in the opposite direction.
    [Fact]
    public void Unticking_show_marked_area_removes_the_rectangle_from_the_schematic()
    {
        UiTest.Run(() =>
        {
            var (workbookId, entryId) = CreateWorkbookWithEntry(new Rect(10, 20, 100, 80), showMarkedArea: true);

            var tab = CreateTab();
            tab.SetShowWorklogEntriesList(true, workbookId);

            Assert.Single(tab.SchematicsWorklogEntriesOverlay.Entries);

            var entry = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);
            entry.ShowMarkedArea = false;
            Assert.True(WorklogManager.UpdateEntry(workbookId, entry));

            tab.RefreshWorklogEntriesListForCurrentWorkbook();

            Assert.Empty(tab.SchematicsWorklogEntriesOverlay.Entries);
        });
    }

    // A newly added entry must appear without the workbook changing - the "add a worklog and it
    // shows nowhere" half of the same report.
    [Fact]
    public void An_entry_added_to_the_shown_workbook_appears_on_the_next_refresh()
    {
        UiTest.Run(() =>
        {
            var (workbookId, _) = CreateWorkbookWithEntry(new Rect(10, 20, 100, 80), showMarkedArea: true);

            var tab = CreateTab();
            tab.SetShowWorklogEntriesList(true, workbookId);

            Assert.Single(tab.SchematicsWorklogEntriesOverlay.Entries);

            Assert.NotNull(WorklogManager.AddEntry(
                workbookId, Schematic, new Rect(500, 400, 60, 60), "Second", "", "Note", "Open",
                Array.Empty<string>()));

            tab.RefreshWorklogEntriesListForCurrentWorkbook();

            Assert.Equal(2, tab.SchematicsWorklogEntriesOverlay.Entries.Count);
        });
    }

    // With "Show worklogs" switched off the refresh must stay cheap and draw nothing - it is called
    // from RefreshWorklogBar on EVERY worklog change, including for users who never turn the
    // overlay on.
    [Fact]
    public void Refreshing_with_the_overlay_switched_off_draws_nothing()
    {
        UiTest.Run(() =>
        {
            var (workbookId, _) = CreateWorkbookWithEntry(new Rect(10, 20, 100, 80), showMarkedArea: true);

            var tab = CreateTab();
            tab.SetShowWorklogEntriesList(false, workbookId);

            tab.RefreshWorklogEntriesListForCurrentWorkbook();

            Assert.Empty(tab.SchematicsWorklogEntriesOverlay.Entries);
            Assert.False(tab.SchematicsWorklogEntriesOverlay.IsVisible);
        });
    }
}
