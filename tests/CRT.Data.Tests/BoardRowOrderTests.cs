using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardRowOrder - whether a draft holds the rows it shares with the published board in another
// order, and in which lists. A change BoardDataDiffer does not count, and a change all the same
// (owner decision, 2026-10-30: "a change of (any) order should also be submitted as a change") -
// Submit, "Review changes", retirement and the maintainer's summary ask it here.
// ###########################################################################################
public sealed class BoardRowOrderTests
{
    private static BoardData Board(params string[] pictures) => new()
    {
        Schematics = [.. pictures.Select(name => new BoardSchematicEntry { SchematicName = name, SchematicImageFile = $"{name}.png" })],
        Components = [new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "U2" }],
    };

    // ###########################################################################################
    // A component's OWN rows in another order are a change of that component (code review, 2026-11-02:
    // U8's links dragged read as unchanged in the pane). Compared among themselves, so a component only
    // moved past another is none - and capitals aside, an added or a removed row moves nothing.
    // ###########################################################################################
    [Fact]
    public void The_components_with_their_own_rows_in_another_order_are_named()
    {
        ComponentLinkEntry First(string label) => new() { BoardLabel = label, Name = "First", Url = "https://example.org/1" };
        ComponentLinkEntry Second(string label) => new() { BoardLabel = label, Name = "Second", Url = "https://example.org/2" };

        BoardData published = Board();
        published.ComponentLinks.AddRange([First("U1"), Second("U1"), First("U2"), Second("U2")]);

        BoardData draft = Board();
        draft.ComponentLinks.AddRange([First("U2"), Second("U2"), Second("u1 "), First("u1 ")]);

        Assert.Equal(["u1"], BoardRowOrder.ComponentsReordered(published, draft));

        BoardData moved = Board();
        moved.ComponentLinks.AddRange([First("U2"), Second("U2"), First("U1"), Second("U1")]);

        Assert.Empty(BoardRowOrder.ComponentsReordered(published, moved));
        Assert.Equal([BoardWorkbookSchema.SheetComponentLinks], BoardRowOrder.ReorderedSheets(published, moved));
    }

    // One component's answer is the board's answer for it - whichever list moved, in any capitals.
    [Fact]
    public void One_component_is_answered_as_the_whole_board_answers_for_it()
    {
        static ComponentLinkEntry First(string label) => new() { BoardLabel = label, Name = "One" };
        static ComponentLinkEntry Second(string label) => new() { BoardLabel = label, Name = "Two" };

        BoardData published = Board();
        published.ComponentLinks.AddRange([First("U1"), Second("U1"), First("U2"), Second("U2")]);

        BoardData draft = Board();
        draft.ComponentLinks.AddRange([First("U2"), Second("U2"), Second("u1 "), First("u1 ")]);

        Assert.True(BoardRowOrder.ComponentReordered(published, draft, "U1"));
        Assert.True(BoardRowOrder.ComponentReordered(published, draft, " u1 "));

        // Only moved past U1, its own rows in their order.
        Assert.False(BoardRowOrder.ComponentReordered(published, draft, "U2"));
        Assert.False(BoardRowOrder.ComponentReordered(published, draft, "U9"));
        Assert.False(BoardRowOrder.ComponentReordered(published, draft, "  "));
        Assert.False(BoardRowOrder.ComponentReordered(published, draft, null));

        var pins = new BoardData { ComponentPins = [new ComponentPinEntry { BoardLabel = "U3", Pin = "1" }, new ComponentPinEntry { BoardLabel = "U3", Pin = "2" }] };
        var pinsTurned = new BoardData { ComponentPins = [pins.ComponentPins[1], pins.ComponentPins[0]] };

        Assert.True(BoardRowOrder.ComponentReordered(pins, pinsTurned, "U3"));
    }

    // Every list a component has - and a baseline among the pictures is still no other order.
    [Fact]
    public void Each_of_a_components_lists_counts_for_it()
    {
        var published = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", Region = "PAL" }, new ComponentEntry { BoardLabel = "U1", Region = "NTSC" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U2", Name = "Top" }, new ComponentImageEntry { BoardLabel = "U2", Name = "Bottom" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U3", Name = "A" }, new ComponentLocalFileEntry { BoardLabel = "U3", Name = "B" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U4", Pin = "1" }, new ComponentPinEntry { BoardLabel = "U4", Pin = "2" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U5", PartNumber = "A" }, new ComponentReplacementEntry { BoardLabel = "U5", PartNumber = "B" }],
        };

        var draft = new BoardData
        {
            Components = [.. Enumerable.Reverse(published.Components)],
            ComponentImages = [.. Enumerable.Reverse(published.ComponentImages)],
            ComponentLocalFiles = [.. Enumerable.Reverse(published.ComponentLocalFiles)],
            ComponentPins = [.. Enumerable.Reverse(published.ComponentPins)],
            ComponentReplacements = [.. Enumerable.Reverse(published.ComponentReplacements)],
        };

        Assert.Equal(["U1", "U2", "U3", "U4", "U5"], BoardRowOrder.ComponentsReordered(published, draft).Order(StringComparer.Ordinal));
    }

    [Fact]
    public void Pictures_moved_are_another_order()
    {
        Assert.True(BoardRowOrder.DiffersFrom(Board("Top", "Bottom"), Board("Bottom", "Top")));
    }

    [Fact]
    public void The_same_rows_in_the_same_order_are_no_other_order()
    {
        Assert.False(BoardRowOrder.DiffersFrom(Board("Top", "Bottom"), Board("top", "Bottom")));
        Assert.Empty(BoardRowOrder.ReorderedSheets(Board("Top", "Bottom"), Board("Top", "Bottom")));
    }

    // A row added or removed is a change the differ counts - not an order of its own.
    [Fact]
    public void A_row_added_or_removed_leaves_the_others_order_alone()
    {
        Assert.False(BoardRowOrder.DiffersFrom(Board("Top", "Bottom"), Board("Top", "Side", "Bottom")));
        Assert.False(BoardRowOrder.DiffersFrom(Board("Top", "Side", "Bottom"), Board("Top", "Bottom")));
    }

    // Every sheet counts, not only the pictures - a component row dragged in the table too.
    [Fact]
    public void Components_moved_are_another_order_too()
    {
        BoardData moved = Board("Top");
        moved.Components.Reverse();

        Assert.True(BoardRowOrder.DiffersFrom(Board("Top"), moved));
    }

    // ###########################################################################################
    // The lists named, by their sheets, in the board's order of sheets - what "Review changes" and the
    // maintainer's summary say, one line each. A file moved past another category's file (2026-10-30)
    // is a list in another order.
    // ###########################################################################################
    [Fact]
    public void The_reordered_lists_are_named_by_their_sheets_in_the_boards_order()
    {
        BoardData published = new()
        {
            Schematics = Board("Top", "Bottom").Schematics,
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = "Board local files/manual.pdf" },
                new BoardLocalFileEntry { Category = "Troubleshooting", Name = "Dead test", File = "Board local files/dead.pdf" },
            ],
            KiCadImportantSignals =
            [
                new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" },
                new KiCadImportantSignalEntry { DisplayName = "Reset", KiCadNetName = "RESET" },
            ],
        };

        BoardData draft = new()
        {
            Schematics = Board("Bottom", "Top").Schematics,
            BoardLocalFiles = [published.BoardLocalFiles[1], published.BoardLocalFiles[0]],
            KiCadImportantSignals = [published.KiCadImportantSignals[1], published.KiCadImportantSignals[0]],
        };

        Assert.Equal(
            [BoardWorkbookSchema.SheetBoardSchematics, BoardWorkbookSchema.SheetBoardLocalFiles, BoardWorkbookSchema.SheetKiCadImportantSignals],
            BoardRowOrder.ReorderedSheets(published, draft));

        Assert.True(BoardRowOrder.Reordered(published.BoardLocalFiles, draft.BoardLocalFiles));
        Assert.False(BoardRowOrder.Reordered(published.BoardLinks, draft.BoardLinks));
    }

    // ###########################################################################################
    // *** THE HIGHLIGHTS' ORDER IS NO WORK. *** Nobody sees the order of the rectangles in the list - a
    // picture's marks drawn again land wherever the label editor puts them - so a draft is not
    // "changed" by it, and nothing is offered to submit for it.
    // ###########################################################################################
    [Fact]
    public void The_order_of_the_highlights_is_not_counted()
    {
        var first = new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" };
        var second = new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U2", X = "9", Y = "9", Width = "5", Height = "5" };

        BoardData published = new() { ComponentHighlights = [first, second] };
        BoardData draft = new() { ComponentHighlights = [second, first] };

        Assert.False(BoardRowOrder.DiffersFrom(published, draft));
    }

    // ###########################################################################################
    // *** THE IMAGES AS A BOARD FILE HOLDS THEM (code review, 2026-11-01). *** A file holds a board's
    // pictures first and its scope baselines after them (BoardDocumentMapping), so a baseline standing
    // among the pictures in the draft held in memory - dragged there in the table - is no order the
    // file can keep: "Review changes" said it was, the save lost it, and Submit never sent it. Moved
    // among its own kind it is an order, and the file keeps it.
    // ###########################################################################################
    [Fact]
    public void A_baseline_among_the_pictures_is_no_other_order_but_among_the_baselines_it_is()
    {
        static ComponentImageEntry Picture(string label) =>
            new() { BoardLabel = label, Name = "Photo", File = $"Component images/{label}.png" };

        static ComponentImageEntry Baseline(string label) =>
            new() { BoardLabel = label, Pin = "1", Name = "Baseline", TimeDiv = "1mS", File = $"Scope baseline/{label}.png" };

        BoardData published = new() { ComponentImages = [Picture("U7"), Picture("U8"), Baseline("U7"), Baseline("U8")] };

        BoardData amongThePictures = new() { ComponentImages = [Picture("U7"), Baseline("U7"), Picture("U8"), Baseline("U8")] };
        Assert.False(BoardRowOrder.DiffersFrom(published, amongThePictures));
        Assert.False(BoardRowOrder.DiffersFrom(amongThePictures, published));

        BoardData baselinesMoved = new() { ComponentImages = [Picture("U7"), Picture("U8"), Baseline("U8"), Baseline("U7")] };
        Assert.Equal([BoardWorkbookSchema.SheetComponentImages], BoardRowOrder.ReorderedSheets(published, baselinesMoved));

        BoardData picturesMoved = new() { ComponentImages = [Picture("U8"), Picture("U7"), Baseline("U7"), Baseline("U8")] };
        Assert.Equal([BoardWorkbookSchema.SheetComponentImages], BoardRowOrder.ReorderedSheets(published, picturesMoved));
    }

    [Theory]
    [InlineData(BoardWorkbookSchema.SheetBoardLocalFiles, "Order of the files changed")]
    [InlineData(BoardWorkbookSchema.SheetBoardLinks, "Order of the links changed")]
    [InlineData(BoardWorkbookSchema.SheetCredits, "Order of the credits changed")]
    [InlineData(BoardWorkbookSchema.SheetKiCadImportantSignals, "Order of the important KiCad signals changed")]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, "Order of the schematic images changed")]
    [InlineData("Something new", "Order of the Something new changed")]
    public void An_order_changed_is_said_in_the_lists_own_words(string sheet, string expected)
    {
        Assert.Equal(expected, BoardRowOrder.OrderChangedLine(sheet));
    }
}
