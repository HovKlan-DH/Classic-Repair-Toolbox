using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardPartEdits - what the board parts of the component pane change in a draft (CRT 4.0 step 2,
// slice 4a): the board's files, links, credits and important signals, each list replaced as a
// whole. Files, links and credits are kept grouped by category, as the pane shows them - and a list
// that differs only by its order is no change, so opening and closing a part changes nothing.
// ###########################################################################################
public sealed class BoardPartEditsTests
{
    private static BoardData Board() => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
        BoardLocalFiles =
        [
            new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" },
            new BoardLocalFileEntry { Category = "Troubleshooting", Name = "Dead test", File = "Board local files/dead.pdf" },
            new BoardLocalFileEntry { Category = "resources", Name = "Schematic notes", File = "Board local files/notes.txt" },
        ],
        BoardLinks = [new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://lemon64.com" }],
        Credits =
        [
            new CreditEntry { Category = "Schematic image", SubCategory = "PAL", NameOrHandle = "Anna", Contact = "RHD" },
            new CreditEntry { Category = "Board labelling", NameOrHandle = "Bo" },
        ],
        KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" }],
    };

    // The categories in the order they first appear, every row's place in its category kept - and
    // a category in other capitals is the same category.
    [Fact]
    public void Rows_are_grouped_by_category_in_the_order_the_categories_first_appear()
    {
        List<BoardLocalFileEntry> grouped = BoardPartEdits.Grouped(Board().BoardLocalFiles, row => row.Category);

        Assert.Equal(["Service manual", "Schematic notes", "Dead test"], grouped.Select(row => row.Name));
        Assert.Equal(["Resources", "Troubleshooting"], BoardPartEdits.CategoriesOf(Board().BoardLocalFiles, row => row.Category));
    }

    // The list as the pane shows it - grouped - is no change though the board had it in another
    // order: opening the part and leaving it makes no step.
    [Fact]
    public void A_list_that_differs_only_by_its_order_returns_the_same_board()
    {
        BoardData board = Board();
        List<BoardLocalFileEntry> shown = BoardPartEdits.Grouped(board.BoardLocalFiles, row => row.Category);

        Assert.Same(board, BoardPartEdits.WithBoardFiles(board, shown));
        Assert.Same(board, BoardPartEdits.WithBoardLinks(board, board.BoardLinks));
        Assert.Same(board, BoardPartEdits.WithCredits(board, board.Credits));
        Assert.Same(board, BoardPartEdits.WithImportantSignals(board, board.KiCadImportantSignals));
    }

    // A change writes the list grouped and trimmed, a row with nothing in it left out - and leaves
    // every other section as it was.
    [Fact]
    public void A_changed_list_is_written_grouped_and_trimmed_without_empty_rows()
    {
        BoardData board = Board();

        BoardData changed = BoardPartEdits.WithBoardFiles(board,
        [
            new BoardLocalFileEntry { Category = " Troubleshooting ", Name = " Dead test ", File = "Board local files/dead.pdf" },
            new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" },
            new BoardLocalFileEntry { Category = "  ", Name = "", File = " " },
            new BoardLocalFileEntry { Category = "Troubleshooting", Name = "Voltages", File = "Board local files/volts.png" },
        ]);

        Assert.Equal(["Dead test", "Voltages", "Service manual"], changed.BoardLocalFiles.Select(row => row.Name));
        Assert.Equal("Troubleshooting", changed.BoardLocalFiles[0].Category);
        Assert.Same(board.BoardLinks, changed.BoardLinks);
        Assert.Same(board.Credits, changed.Credits);
        Assert.Same(board.Components, changed.Components);
    }

    [Fact]
    public void Links_and_credits_are_replaced_the_same_way()
    {
        BoardData board = Board();

        BoardData links = BoardPartEdits.WithBoardLinks(board, [.. board.BoardLinks, new BoardLinkEntry { Category = "Forums", Name = "RHD", Url = "https://example.org" }]);
        Assert.Equal(["Lemon64", "RHD"], links.BoardLinks.Select(row => row.Name));

        BoardData credits = BoardPartEdits.WithCredits(board,
        [
            .. board.Credits,
            new CreditEntry { Category = "schematic image", SubCategory = "NTSC", NameOrHandle = " Cy " },
        ]);

        Assert.Equal(["Anna", "Cy", "Bo"], credits.Credits.Select(row => row.NameOrHandle));
        Assert.Same(board.BoardLocalFiles, credits.BoardLocalFiles);
    }

    // An important signal is its display name AND its net: kept in the order given, two rows may
    // share either, and a half-typed row is kept for the table's checks to point at.
    [Fact]
    public void Important_signals_keep_their_order_and_a_half_typed_row()
    {
        BoardData board = Board();

        BoardData changed = BoardPartEdits.WithImportantSignals(board,
        [
            new KiCadImportantSignalEntry { DisplayName = " Reset ", KiCadNetName = "RESET" },
            new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" },
            new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "" },
            new KiCadImportantSignalEntry { DisplayName = " ", KiCadNetName = " " },
        ]);

        Assert.Equal(["Reset|RESET", "Clock|PHI2", "Clock|"], changed.KiCadImportantSignals.Select(row => $"{row.DisplayName}|{row.KiCadNetName}"));
    }

    // Every copy of the board names every section (BoardDataCaptionTests' rule, the helper's half).
    [Fact]
    public void The_board_sections_helper_keeps_every_other_section()
    {
        BoardData board = Board();
        BoardData changed = board.WithBoardSections(credits: []);

        Assert.Empty(changed.Credits);
        Assert.Same(board.Components, changed.Components);
        Assert.Same(board.BoardLocalFiles, changed.BoardLocalFiles);
        Assert.Same(board.KiCadImportantSignals, changed.KiCadImportantSignals);
        Assert.Same(board.ComponentPins, changed.ComponentPins);
    }
}
