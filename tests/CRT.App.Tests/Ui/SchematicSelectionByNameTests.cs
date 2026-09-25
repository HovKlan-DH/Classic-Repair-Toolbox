using CRT;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// TabSchematics.SelectSchematicByName - lets another tab (the Workbooks tab, specifically: see
// TabWorkbooks.PropagateSelectedSchematicToSchematicsTab) make a named schematic the one waiting
// on the Schematics tab, without the user ever clicking its thumbnail. It has to go through the
// SAME path a real click does (setting SchematicsThumbnailList.SelectedItem) rather than a
// side-channel, so the full-res image, overlays and "last schematic for this board" all follow
// exactly as they would from a click - this file only pins the selection outcome itself, since
// OnSchematicsThumbnailSelectionChanged's own image-loading behaviour is covered elsewhere
// (WorklogAreaMarkingTests's CreateTab fixture already relies on selection driving that path).
// ###########################################################################################
[Collection("HeadlessUi")]
public class SchematicSelectionByNameTests
{
    private static TabSchematics CreateTabWithThumbnails(params string[] names)
    {
        var tab = new TabSchematics();

        var thumbnails = names.Select(name => new SchematicThumbnail { Name = name }).ToList();
        foreach (var thumbnail in thumbnails)
        {
            tab.currentThumbnails.Add(thumbnail);
        }

        tab.SchematicsThumbnailList.ItemsSource = tab.currentThumbnails;
        return tab;
    }

    [Fact]
    public void Selects_the_thumbnail_matching_the_given_name()
    {
        UiTest.Run(() =>
        {
            var tab = CreateTabWithThumbnails("Sheet 1", "Sheet 2");

            tab.SelectSchematicByName("Sheet 2");

            var selected = Assert.IsType<SchematicThumbnail>(tab.SchematicsThumbnailList.SelectedItem);
            Assert.Equal("Sheet 2", selected.Name);
        });
    }

    // The workbook feature stores and compares schematic names elsewhere case-insensitively
    // (BoardDataReader's schematicsByName dictionary, RefreshBoardPreviews's own grouping) - this
    // has to match that, or a workbook entry naming "sheet 2" would silently fail to select the
    // real "Sheet 2" thumbnail.
    [Fact]
    public void Matching_is_case_insensitive()
    {
        UiTest.Run(() =>
        {
            var tab = CreateTabWithThumbnails("Sheet 1", "Sheet 2");

            tab.SelectSchematicByName("sheet 2");

            var selected = Assert.IsType<SchematicThumbnail>(tab.SchematicsThumbnailList.SelectedItem);
            Assert.Equal("Sheet 2", selected.Name);
        });
    }

    // A workbook entry can reference a schematic that was since renamed or removed from the board
    // data - dropping the request rather than throwing or clearing the current selection is the
    // same tolerance BuildSchematicPreview already applies to a missing image file.
    [Fact]
    public void An_unknown_name_leaves_the_current_selection_untouched()
    {
        UiTest.Run(() =>
        {
            var tab = CreateTabWithThumbnails("Sheet 1", "Sheet 2");
            tab.SchematicsThumbnailList.SelectedItem = tab.currentThumbnails[0];

            tab.SelectSchematicByName("Sheet that no longer exists");

            var selected = Assert.IsType<SchematicThumbnail>(tab.SchematicsThumbnailList.SelectedItem);
            Assert.Equal("Sheet 1", selected.Name);
        });
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_name_is_harmless(string? blankName)
    {
        UiTest.Run(() =>
        {
            var tab = CreateTabWithThumbnails("Sheet 1");
            tab.SchematicsThumbnailList.SelectedItem = tab.currentThumbnails[0];

            tab.SelectSchematicByName(blankName!);

            var selected = Assert.IsType<SchematicThumbnail>(tab.SchematicsThumbnailList.SelectedItem);
            Assert.Equal("Sheet 1", selected.Name);
        });
    }

    // Re-selecting the schematic that is already current must not re-trigger the full selection
    // pipeline (image reload, overlay rebuild) - SelectSchematicByName guards on this explicitly
    // rather than relying on SelectedItem's own reference equality, since a lookup returns a
    // different-but-equal-by-name match than whatever instance is already selected in principle.
    [Fact]
    public void Selecting_the_already_selected_schematic_is_a_no_op()
    {
        UiTest.Run(() =>
        {
            var tab = CreateTabWithThumbnails("Sheet 1", "Sheet 2");
            tab.SchematicsThumbnailList.SelectedItem = tab.currentThumbnails[0];

            tab.SelectSchematicByName("Sheet 1");

            Assert.Same(tab.currentThumbnails[0], tab.SchematicsThumbnailList.SelectedItem);
        });
    }

    // ###########################################################################################
    // A thumbnail DRAG-REORDER suppresses OnSchematicsThumbnailSelectionChanged for its whole
    // duration (the flag is set on drag start, cleared only on release). Assigning SelectedItem in
    // that window would move the list's highlight while the schematic image, the overlays and
    // currentFullResBitmap all stayed on the OLD schematic - the list and the picture disagreeing -
    // and the drag's own release handler reassigns SelectedItem anyway, so the change would be
    // silently discarded a moment later.
    //
    // Dropping the request is the right answer rather than deferring it: the user is reordering
    // thumbnails, not asking to view a different one, and this method's caller is a cross-tab
    // convenience rather than a command that must land.
    // ###########################################################################################
    [Fact]
    public void A_selection_during_a_thumbnail_drag_is_dropped_rather_than_desyncing_the_list()
    {
        UiTest.Run(() =>
        {
            var tab = CreateTabWithThumbnails("Sheet 1", "Sheet 2");
            tab.SchematicsThumbnailList.SelectedItem = tab.currentThumbnails[0];

            tab.SuppressThumbnailSelectionChangedForTests = true;
            tab.SelectSchematicByName("Sheet 2");

            Assert.Same(tab.currentThumbnails[0], tab.SchematicsThumbnailList.SelectedItem);

            // And it works again once the drag is over.
            tab.SuppressThumbnailSelectionChangedForTests = false;
            tab.SelectSchematicByName("Sheet 2");

            Assert.Same(tab.currentThumbnails[1], tab.SchematicsThumbnailList.SelectedItem);
        });
    }
}
