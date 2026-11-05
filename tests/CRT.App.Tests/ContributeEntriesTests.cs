using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// THE LEFT COLUMN'S ENTRIES AND THEIR TABLES (owner question, 2026-11-03: "do we not have a direct 1:1
// mapping for everything here?"). Every sheet a board has must have an entry - a sheet added without one
// could never be seen in the Contribute tab's table, its tabs being gone.
// ###########################################################################################
public sealed class ContributeEntriesTests
{
    [Fact]
    public void Every_sheet_of_a_board_has_an_entry()
    {
        foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
        {
            Assert.NotNull(ContributeEntries.EntryOfTable(sheet.SheetName, BoardTableSheetPart.Whole));
        }
    }

    // ###########################################################################################
    // THE MAINTAINER TAB'S TABS ARE THESE ENTRIES, UNDER THE SAME NAMES (owner request, 2026-11-05): what a
    // contributor filled in is read by the words the contributor saw. A shared change: an entry or a tab
    // moved on one side only fails here.
    // ###########################################################################################
    [Fact]
    public void The_tables_tabs_are_the_entries_with_a_table_under_their_names()
    {
        static string EntryName(ContributeEntry entry) => entry switch
        {
            ContributeEntry.SchematicImages => ComponentPaneWording.SchematicImages.Replace("Board: ", string.Empty, StringComparison.Ordinal),
            ContributeEntry.ImportantSignals => ComponentPaneWording.ImportantSignals.Replace("KiCad: ", string.Empty, StringComparison.Ordinal),
            ContributeEntry.Files => "Files",
            ContributeEntry.Links => "Links",
            ContributeEntry.Credits => "Credits",
            _ => ComponentPaneWording.PageName(ContributeEntries.PageOf(entry)!.Value),
        };

        var entries = Enum.GetValues<ContributeEntry>()
            .Where(entry => ContributeEntries.TableOf(entry) is not null)
            .Select(entry => (ContributeEntries.TableOf(entry)!.Value.Sheet, ContributeEntries.TableOf(entry)!.Value.Part, Name: EntryName(entry)))
            .OrderBy(entry => entry.Sheet, StringComparer.Ordinal)
            .ThenBy(entry => entry.Part)
            .ToList();

        var tabs = BoardTableTabs.All
            .Select(tab => (tab.Sheet, tab.Part, tab.Name))
            .OrderBy(tab => tab.Sheet, StringComparer.Ordinal)
            .ThenBy(tab => tab.Part)
            .ToList();

        Assert.Equal(entries, tabs);
    }

    // A sheet in two is two entries, and each is that sheet's own part.
    [Fact]
    public void Component_images_is_two_entries()
    {
        Assert.Equal(new ContributeEntryTable(BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Pictures), ContributeEntries.TableOf(ContributeEntry.ComponentPictures));
        Assert.Equal(new ContributeEntryTable(BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Baselines), ContributeEntries.TableOf(ContributeEntry.ComponentBaselines));

        Assert.Equal(ContributeEntry.ComponentPictures, ContributeEntries.EntryOfTable(BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Pictures));
        Assert.Equal(ContributeEntry.ComponentBaselines, ContributeEntries.EntryOfTable(BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Baselines));
    }

    [Fact]
    public void Five_entries_are_no_sheet()
    {
        Assert.Equal(
            [ContributeEntry.Overview, ContributeEntry.Name, ContributeEntry.Regions, ContributeEntry.KiCadFiles, ContributeEntry.Categories],
            Enum.GetValues<ContributeEntry>().Where(entry => ContributeEntries.TableOf(entry) is null));
    }

    // Each table takes the owner back to the entry showing it: the column follows a pill or a search.
    [Fact]
    public void Each_entry_with_a_table_is_found_again_by_it()
    {
        foreach (ContributeEntry entry in Enum.GetValues<ContributeEntry>())
        {
            if (ContributeEntries.TableOf(entry) is { } table)
            {
                Assert.Equal(entry, ContributeEntries.EntryOfTable(table.Sheet, table.Part));
            }
        }
    }

    [Fact]
    public void A_page_and_a_part_are_each_an_entry_and_back()
    {
        foreach (ComponentPage page in Enum.GetValues<ComponentPage>())
        {
            Assert.Equal(page, ContributeEntries.PageOf(ContributeEntries.OfPage(page)));
        }

        foreach (PaneBoardPart part in Enum.GetValues<PaneBoardPart>())
        {
            Assert.Equal(part, ContributeEntries.PartOf(ContributeEntries.OfPart(part)));
        }

        Assert.Null(ContributeEntries.PageOf(ContributeEntry.Files));
        Assert.Null(ContributeEntries.PartOf(ContributeEntry.ComponentFiles));
    }

    private static BoardData Board() => new()
    {
        Regions = ["PAL", "NTSC"],
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", Category = "IC", Package = "DIP-40", TechnicalNameOrValue = "6510" },
            new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "PAL", TechnicalNameOrValue = "6569" },
            new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "NTSC", TechnicalNameOrValue = "6567" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "40", Name = "RES" },
        ],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/a.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/b.png" },
        ],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.com" }],
    };

    // Owner request, 2026-11-03: "Basics (show '1' when the region is not different, or '2' if it is regionalized)".
    [Fact]
    public void Basics_counts_one_or_the_two_regions_it_differs_in()
    {
        Assert.Equal(1, ContributeEntries.FactsOf(Board(), "U8", "PAL")[ComponentPage.Basics].Count);
        Assert.Equal(2, ContributeEntries.FactsOf(Board(), "U1", "PAL")[ComponentPage.Basics].Count);
    }

    [Fact]
    public void Each_page_counts_what_it_lists()
    {
        IReadOnlyDictionary<ComponentPage, ComponentPageFact> facts = ContributeEntries.FactsOf(Board(), "U8", "PAL");

        Assert.Equal(0, facts[ComponentPage.Replacements].Count);
        Assert.Equal(2, facts[ComponentPage.Pinout].Count);
        Assert.Equal(1, facts[ComponentPage.Pictures].Count);
        Assert.Equal(1, facts[ComponentPage.Baselines].Count);
        Assert.Equal(0, facts[ComponentPage.Files].Count);
        Assert.Equal(1, facts[ComponentPage.Links].Count);
        Assert.DoesNotContain(facts.Values, fact => fact.HasError);
    }

    // What Submit would refuse, on the page it is on - before the component is ever opened.
    [Fact]
    public void What_stops_Submit_is_said_of_its_page()
    {
        BoardData board = Board();
        board.Components[0] = new ComponentEntry { BoardLabel = "U8", Category = string.Empty, Package = "DIP-40" };
        board.ComponentPins[1] = new ComponentPinEntry { BoardLabel = "U8", Pin = "48", Name = "RES" };
        board.ComponentImages[1] = new ComponentImageEntry { BoardLabel = "U8", Pin = "48", Name = "VSS", File = "Scope baseline/b.png" };
        board.ComponentLinks[0] = new ComponentLinkEntry { BoardLabel = "U8", Name = "Forum", Url = "www.lemon64.com" };

        IReadOnlyDictionary<ComponentPage, ComponentPageFact> facts = ContributeEntries.FactsOf(board, "U8", "PAL");

        Assert.True(facts[ComponentPage.Basics].HasError);
        Assert.True(facts[ComponentPage.Pinout].HasError);
        Assert.True(facts[ComponentPage.Baselines].HasError);
        Assert.True(facts[ComponentPage.Links].HasError);
        Assert.False(facts[ComponentPage.Pictures].HasError);
    }

    // ###########################################################################################
    // A baseline lacking its signal name or image is an error that stops Submit, so its entry is red whatever
    // is open (code review, 2026-11-04: red only while the pane was on the component). What the pane says of
    // the same rows, ComponentPane.PagesShown.
    // ###########################################################################################
    [Theory]
    [InlineData("", "Scope baseline/b.png")]
    [InlineData("VSS", "")]
    public void A_baseline_lacking_what_it_needs_is_said_of_its_page(string name, string file)
    {
        BoardData board = Board();
        board.ComponentImages[1] = new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = name, File = file };

        Assert.True(ContributeEntries.FactsOf(board, "U8", "PAL")[ComponentPage.Baselines].HasError);
    }
}
