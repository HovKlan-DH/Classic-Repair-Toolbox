using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// THE TABS OF A TABLE (owner request, 2026-11-05): every sheet of a board is on one, or a maintainer could
// not see it, and a sheet in two is on one per part - a part left out would hide its rows.
// ###########################################################################################
public sealed class BoardTableTabsTests
{
    [Fact]
    public void Every_sheet_of_a_board_has_a_tab_and_no_tab_names_another()
    {
        Assert.Equal(
            BoardWorkbookSchema.AllSheets.Select(sheet => sheet.SheetName).Order(),
            BoardTableTabs.SheetsInOrder.Select(sheet => sheet.SheetName).Order());

        Assert.All(BoardTableTabs.All, tab => Assert.Contains(BoardWorkbookSchema.AllSheets, sheet => sheet.SheetName == tab.Sheet));
    }

    [Fact]
    public void A_sheet_in_two_has_a_tab_per_part_and_every_other_sheet_one_whole()
    {
        foreach (IGrouping<string, BoardTableTab> sheet in BoardTableTabs.All.GroupBy(tab => tab.Sheet))
        {
            Assert.Equal(
                BoardTableSheetParts.IsSplit(sheet.Key)
                    ? [BoardTableSheetPart.Pictures, BoardTableSheetPart.Baselines]
                    : [BoardTableSheetPart.Whole],
                sheet.Select(tab => tab.Part));
        }
    }

    // A group stands together: a table reads as its headings, each once.
    [Fact]
    public void A_groups_tabs_stand_together()
    {
        List<string> groups = BoardTableTabs.All.Select(tab => tab.Group).ToList();
        List<string> once = groups.Distinct().ToList();

        Assert.Equal(once, groups.Where((group, index) => index == 0 || groups[index - 1] != group));
    }

    // A table with no change anywhere opens on its first tab: the board's pictures, never a sheet most boards
    // have nothing on.
    [Fact]
    public void A_table_opens_on_the_schematic_images()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, new BoardData());

        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, document.Sheets[0].Name);
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, document.SheetToShow(null, BoardTableRowKinds.None).Name);
        Assert.Null(BoardTableTabs.FirstChanged(document));
    }

    // What changed is what a maintainer opening a submission wants: the first tab holding a change, by the
    // part of a sheet in two that holds it, and the two parts' counts add up to the sheet's.
    [Fact]
    public void The_first_tab_with_a_change_goes_by_the_part_that_holds_it()
    {
        var photo = new ComponentImageEntry { BoardLabel = "U1", Name = "Photo", File = "Images/photo.png" };
        var baseline = new ComponentImageEntry { BoardLabel = "U1", Pin = "2", Name = "PHI2", File = "Scope/phi2.png" };

        BoardTableDocument document = BoardTableDocument.Create(
            new BoardData { ComponentImages = [photo] },
            new BoardData { ComponentImages = [photo, baseline] });

        BoardTableTab first = Assert.IsType<BoardTableTab>(BoardTableTabs.FirstChanged(document));

        Assert.Equal(BoardWorkbookSchema.SheetComponentImages, first.Sheet);
        Assert.Equal(BoardTableSheetPart.Baselines, first.Part);

        BoardTableSheet images = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!;

        Assert.Equal(1, BoardTableSheetParts.ChangeCountOf(images, BoardTableSheetPart.Baselines));
        Assert.Equal(0, BoardTableSheetParts.ChangeCountOf(images, BoardTableSheetPart.Pictures));
        Assert.Equal(images.ChangeCount, BoardTableSheetParts.ChangeCountOf(images, BoardTableSheetPart.Whole));
    }

    // A row deleted counts where its pin put it.
    [Fact]
    public void A_deleted_baseline_counts_on_the_baselines()
    {
        var baseline = new ComponentImageEntry { BoardLabel = "U1", Pin = "2", Name = "PHI2", File = "Scope/phi2.png" };

        BoardTableDocument document = BoardTableDocument.Create(
            new BoardData { ComponentImages = [baseline] },
            new BoardData());

        BoardTableSheet images = document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!;

        Assert.Equal(1, images.ChangeCount);
        Assert.Equal(1, BoardTableSheetParts.ChangeCountOf(images, BoardTableSheetPart.Baselines));
        Assert.Equal(0, BoardTableSheetParts.ChangeCountOf(images, BoardTableSheetPart.Pictures));
    }

    // ###########################################################################################
    // A sheet by the name a person sees (code review, 2026-11-05): the list of problems over the table said "Board
    // schematics" and "Components" beside tabs reading "Schematic images" and "Basics". Its tab under its group, both
    // tabs of a sheet in two - nothing here says which part a problem's row stands in.
    // ###########################################################################################
    [Fact]
    public void A_sheet_is_named_by_its_tab_under_its_group()
    {
        Assert.Equal("Board: Schematic images", BoardTableTabs.NameOf(BoardWorkbookSchema.SheetBoardSchematics));
        Assert.Equal("KiCad: Important signals", BoardTableTabs.NameOf(BoardWorkbookSchema.SheetKiCadImportantSignals));
        Assert.Equal("Components: Basics", BoardTableTabs.NameOf(BoardWorkbookSchema.SheetComponents));
        Assert.Equal("Components: Pictures or Oscilloscope baseline", BoardTableTabs.NameOf(BoardWorkbookSchema.SheetComponentImages));

        // One no tab has keeps its own name, whatever a later sheet is called.
        Assert.Equal("No such sheet", BoardTableTabs.NameOf("No such sheet"));
    }
}
