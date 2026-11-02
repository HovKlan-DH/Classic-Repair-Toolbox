using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// NewBoardCredits - the rows a new board's credits start with (owner request, 2026-11-02: "Each
// individual schematic image; Board data; 'Oscilloscope baseline' for 'What' and 'PAL' or 'NTSC' for
// 'Sub-category'. Then the user can delete it if needed or he/she should input the missing contact
// details").
// ###########################################################################################
public sealed class NewBoardCreditsTests
{
    private static BoardData Board(bool regions = true, params string[] images) => new()
    {
        Regions = regions ? ["PAL", "NTSC"] : [],
        Schematics = images.Select(name => new BoardSchematicEntry { SchematicName = name }).ToList(),
    };

    private static List<string> Rows(BoardData board) =>
        board.Credits.Select(row => $"{row.Category}|{row.SubCategory}").ToList();

    [Fact]
    public void A_new_board_gets_a_row_per_image_board_data_and_a_baseline_row_per_region()
    {
        NewBoardCreditOffer offer = NewBoardCredits.Offer(Board(images: ["Top", "Bottom"]), null);

        Assert.Equal(
            ["Schematic image|Top", "Schematic image|Bottom", "Board data|", "Oscilloscope baseline|PAL", "Oscilloscope baseline|NTSC"],
            Rows(offer.Board));

        // Nobody has been named: an error of the draft's, until filled in or deleted.
        Assert.All(offer.Board.Credits, row => Assert.Equal([BoardWorkbookSchema.ColNameOrHandle], BoardRowRequired.EmptyColumns(row)));
    }

    [Fact]
    public void A_board_with_no_regions_has_one_baseline_row_with_no_sub_category()
    {
        NewBoardCreditOffer offer = NewBoardCredits.Offer(Board(regions: false), null);

        Assert.Equal(["Board data|", "Oscilloscope baseline|"], Rows(offer.Board));
    }

    // Offered once: asked again, the same board comes back, which the draft takes as no step.
    [Fact]
    public void Asked_again_with_nothing_new_it_is_the_same_board()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(images: ["Top"]), null);
        NewBoardCreditOffer again = NewBoardCredits.Offer(first.Board, first.Offered);

        Assert.Same(first.Board, again.Board);
        Assert.Equal(first.Offered, again.Offered);
    }

    // "Then the user can delete it if needed" - and it stays deleted.
    [Fact]
    public void A_row_deleted_does_not_come_back()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(images: ["Top"]), null);
        BoardData deleted = BoardPartEdits.WithCredits(first.Board, first.Board.Credits.Where(row => row.Category != "Board data").ToList());

        NewBoardCreditOffer again = NewBoardCredits.Offer(deleted, first.Offered);

        Assert.Same(deleted, again.Board);
        Assert.DoesNotContain("Board data|", Rows(again.Board));
    }

    // An image added since gets its own row, with the others of its kind.
    [Fact]
    public void An_image_added_later_gets_its_row_beside_the_others()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(images: ["Top"]), null);
        BoardData withImage = first.Board.WithSchematics([.. first.Board.Schematics, new BoardSchematicEntry { SchematicName = "Bottom" }]);

        NewBoardCreditOffer again = NewBoardCredits.Offer(withImage, first.Offered);

        Assert.Equal(
            ["Schematic image|Top", "Schematic image|Bottom", "Board data|", "Oscilloscope baseline|PAL", "Oscilloscope baseline|NTSC"],
            Rows(again.Board));
    }

    // The first image on a board that has its other rows already: the images come first.
    [Fact]
    public void The_first_image_goes_above_the_other_rows()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(regions: false), null);
        BoardData withImage = first.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Top" }]);

        Assert.Equal(
            ["Schematic image|Top", "Board data|", "Oscilloscope baseline|"],
            Rows(NewBoardCredits.Offer(withImage, first.Offered).Board));
    }

    // A row nothing was typed in follows the board: an image renamed takes its row along.
    [Fact]
    public void A_row_nothing_was_typed_in_follows_a_renamed_image()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(images: ["Top"]), null);
        BoardData renamed = first.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Upper side" }]);

        Assert.Equal(
            ["Schematic image|Upper side", "Board data|", "Oscilloscope baseline|PAL", "Oscilloscope baseline|NTSC"],
            Rows(NewBoardCredits.Offer(renamed, first.Offered).Board));
    }

    // Whether the offer took a row away - what makes its step one an Undo must put the marker back for
    // (TabDrafts.NewBoardCredits.cs). Only adding rows says no: rows undone stay gone.
    [Fact]
    public void The_offer_says_whether_a_row_followed_the_board_away()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(images: ["Top"]), null);
        Assert.False(first.RowsWent);

        BoardData renamed = first.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Upper side" }]);
        Assert.True(NewBoardCredits.Offer(renamed, first.Offered).RowsWent);
    }

    // ... but one with a name in it is the contributor's, and stays.
    [Fact]
    public void A_row_with_a_name_in_it_stays_when_its_image_goes()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(images: ["Top"]), null);
        List<CreditEntry> filled = first.Board.Credits
            .Select(row => row.Category == "Schematic image" ? new CreditEntry { Category = row.Category, SubCategory = row.SubCategory, NameOrHandle = "Anna" } : row)
            .ToList();

        BoardData imageGone = first.Board.WithBoardSections(credits: filled).WithSchematics([]);

        Assert.Contains("Schematic image|Top", Rows(NewBoardCredits.Offer(imageGone, first.Offered).Board));
    }

    // The board's regions ticked later: the row for every region gives way to one per region.
    [Fact]
    public void Regions_ticked_later_turn_the_baseline_row_into_one_per_region()
    {
        NewBoardCreditOffer first = NewBoardCredits.Offer(Board(regions: false), null);
        BoardData withRegions = first.Board.WithRegions(["PAL", "NTSC"]);

        Assert.Equal(["Board data|", "Oscilloscope baseline|PAL", "Oscilloscope baseline|NTSC"], Rows(NewBoardCredits.Offer(withRegions, first.Offered).Board));
    }

    // What "Board status" says an opening would add: what the offer adds, and nothing it does not.
    [Fact]
    public void How_many_rows_an_offer_would_add_is_counted_without_making_it()
    {
        BoardData board = Board(images: ["Top"]);

        Assert.Equal(4, NewBoardCredits.RowsToAdd(board, null));
        Assert.Empty(board.Credits);

        NewBoardCreditOffer first = NewBoardCredits.Offer(board, null);
        Assert.Equal(0, NewBoardCredits.RowsToAdd(first.Board, first.Offered));

        // Deleted, it is not offered again; an image added since is.
        BoardData deleted = BoardPartEdits.WithCredits(first.Board, first.Board.Credits.Where(row => row.Category != "Board data").ToList());
        Assert.Equal(0, NewBoardCredits.RowsToAdd(deleted, first.Offered));

        BoardData withImage = deleted.WithSchematics([.. deleted.Schematics, new BoardSchematicEntry { SchematicName = "Bottom" }]);
        Assert.Equal(1, NewBoardCredits.RowsToAdd(withImage, first.Offered));
    }

    // A credit typed in by hand counts as given - no second row, capitals aside.
    [Fact]
    public void A_row_the_credits_have_already_is_not_added_twice()
    {
        BoardData typed = Board(regions: false).WithBoardSections(credits: [new CreditEntry { Category = "board data", NameOrHandle = "Anna" }]);

        Assert.Equal(["board data|", "Oscilloscope baseline|"], Rows(NewBoardCredits.Offer(typed, null).Board));
    }
}
