using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardPartEdits - what the board parts of the component pane change in a draft (CRT 4.0 step 2,
// slice 4a): the board's files, links, credits and important signals, each list replaced as a
// whole, in the order given - the board's own order, which the pane shows and a row moved changes
// (owner request, 2026-10-30: no grouping by category any more). The list as the board has it is no
// change, so opening and closing a part changes nothing.
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

    // The board's categories offered as one is typed: in the order they first appear, a category in
    // other capitals the same category.
    [Fact]
    public void The_categories_offered_are_the_boards_own_once_each()
    {
        Assert.Equal(["Resources", "Troubleshooting"], BoardPartEdits.CategoriesOf(Board().BoardLocalFiles, row => row.Category));
    }

    // The list as the board has it - which is how the pane shows it - is no change: opening the part
    // and leaving it makes no step.
    [Fact]
    public void A_list_as_the_board_has_it_returns_the_same_board()
    {
        BoardData board = Board();

        Assert.Same(board, BoardPartEdits.WithBoardFiles(board, board.BoardLocalFiles));
        Assert.Same(board, BoardPartEdits.WithBoardLinks(board, board.BoardLinks));
        Assert.Same(board, BoardPartEdits.WithCredits(board, board.Credits));
        Assert.Same(board, BoardPartEdits.WithImportantSignals(board, board.KiCadImportantSignals));
    }

    // ###########################################################################################
    // A ROW MOVES PAST ANY ROW, WHATEVER ITS CATEGORY (owner request, 2026-10-30: "please have the
    // entire line moving up/down, and not only within its own category ... I also want to be able to
    // rearrange the categories"). Only the ends stop it. The move is a change of the board: the list
    // written in its new order - and the list is NOT regrouped by category, which would undo a move
    // across categories.
    // ###########################################################################################
    [Fact]
    public void A_row_moves_past_a_row_of_another_category_and_the_move_is_a_change()
    {
        BoardData board = Board();
        List<BoardLocalFileEntry> shown = board.BoardLocalFiles;

        // Shown as the board has them: "Service manual" (Resources), "Dead test" (Troubleshooting),
        // "Schematic notes" (resources).
        Assert.False(BoardPartEdits.CanMove(shown, 0, -1));
        Assert.True(BoardPartEdits.CanMove(shown, 0, +1));
        Assert.True(BoardPartEdits.CanMove(shown, 1, -1));
        Assert.True(BoardPartEdits.CanMove(shown, 1, +1));
        Assert.True(BoardPartEdits.CanMove(shown, 2, -1));
        Assert.False(BoardPartEdits.CanMove(shown, 2, +1));
        Assert.False(BoardPartEdits.CanMove(shown, 0, +2));
        Assert.False(BoardPartEdits.CanMove(shown, 5, -1));

        // "Dead test" up past "Service manual" - into the first place, its category first.
        List<BoardLocalFileEntry> moved = BoardPartEdits.Moved(shown, 1, -1);
        Assert.Equal(["Dead test", "Service manual", "Schematic notes"], moved.Select(row => row.Name));

        BoardData changed = BoardPartEdits.WithBoardFiles(board, moved);
        Assert.NotSame(board, changed);
        Assert.Equal(["Dead test", "Service manual", "Schematic notes"], changed.BoardLocalFiles.Select(row => row.Name));

        // A move that cannot be made leaves the rows as they were.
        Assert.Equal(shown.Select(row => row.Name), BoardPartEdits.Moved(shown, 2, +1).Select(row => row.Name));
    }

    // A change writes the list in the order given and trimmed, a row with nothing in it left out -
    // and leaves every other section as it was.
    [Fact]
    public void A_changed_list_is_written_in_its_order_and_trimmed_without_empty_rows()
    {
        BoardData board = Board();

        BoardData changed = BoardPartEdits.WithBoardFiles(board,
        [
            new BoardLocalFileEntry { Category = " Troubleshooting ", Name = " Dead test ", File = "Board local files/dead.pdf" },
            new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" },
            new BoardLocalFileEntry { Category = "  ", Name = "", File = " " },
            new BoardLocalFileEntry { Category = "Troubleshooting", Name = "Voltages", File = "Board local files/volts.png" },
        ]);

        Assert.Equal(["Dead test", "Service manual", "Voltages"], changed.BoardLocalFiles.Select(row => row.Name));
        Assert.Equal("Troubleshooting", changed.BoardLocalFiles[0].Category);
        Assert.Same(board.BoardLinks, changed.BoardLinks);
        Assert.Same(board.Credits, changed.Credits);
        Assert.Same(board.Components, changed.Components);
    }

    // ###########################################################################################
    // A row with nothing in it - every value blank or spaces - is the one rule the writes leave out
    // and the pane's title does not count (code review, 2026-10-30: "Add credit" made "Credits (4)" of
    // three). One value anywhere is a row.
    // ###########################################################################################
    [Fact]
    public void A_blank_row_is_one_with_nothing_in_any_value()
    {
        Assert.True(BoardPartEdits.IsBlank(new BoardLocalFileEntry { Category = " ", Name = "", File = "\t" }));
        Assert.False(BoardPartEdits.IsBlank(new BoardLocalFileEntry { File = "Board local files/x.pdf" }));
        Assert.True(BoardPartEdits.IsBlank(new BoardLinkEntry()));
        Assert.False(BoardPartEdits.IsBlank(new BoardLinkEntry { Url = "www.lemon64.com" }));
        Assert.True(BoardPartEdits.IsBlank(new CreditEntry { Category = "  " }));
        Assert.False(BoardPartEdits.IsBlank(new CreditEntry { Contact = "RHD" }));
        Assert.True(BoardPartEdits.IsBlank(new KiCadImportantSignalEntry()));
        Assert.False(BoardPartEdits.IsBlank(new KiCadImportantSignalEntry { KiCadNetName = "PHI2" }));

        // What the writes leave out is exactly that.
        BoardData board = Board();
        BoardData credits = BoardPartEdits.WithCredits(board, [.. board.Credits, new CreditEntry { SubCategory = " " }]);
        Assert.Same(board, credits);
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

        Assert.Equal(["Anna", "Bo", "Cy"], credits.Credits.Select(row => row.NameOrHandle));
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
