using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// CreditSuggestions - the credits a draft's work deserves (CRT 4.0 step 2, 2026-10-26: the Submit
// dialog offers them). Each kind of change has its category, spelled as the board's credits already
// spell it; a change of nothing creditable offers nothing; and a credit the board already gives the
// same name is never added twice.
// ###########################################################################################
public sealed class CreditSuggestionsTests
{
    private static BoardData Published() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png", SchematicHighlightColor = "Red" }],
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" }],
        ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "PHI2", File = "Scope baseline/U8_1.png", TimeDiv = "1us" }],
        ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
        Credits = [new CreditEntry { Category = "Oscilloscope baseline", NameOrHandle = "Anna" }],
    };

    // A copy of the published board - a draft that changes nothing yet.
    private static BoardData Draft() => Published().WithBoardSections();

    [Fact]
    public void A_draft_that_changes_nothing_creditable_offers_nothing()
    {
        BoardData draft = Draft().WithBoardSections(credits: [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }]);

        Assert.Empty(CreditSuggestions.For(Published(), draft));
    }

    // A picture added is credited by its name; one only renamed or replaced is too.
    [Fact]
    public void A_picture_added_is_a_schematic_image_named_by_its_picture()
    {
        BoardData draft = Draft().WithSchematics([.. Published().Schematics, new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "bottom.png" }]);

        CreditSuggestion suggestion = Assert.Single(CreditSuggestions.For(Published(), draft));
        Assert.Equal(new CreditSuggestion("Schematic image", "Bottom"), suggestion);
    }

    // Each kind of work has its category, in a fixed order; the baselines take the board's own spelling.
    [Fact]
    public void Each_kind_of_change_has_its_category_in_the_boards_spelling()
    {
        BoardData draft = Draft()
            .WithComponentHighlights([new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "5", Y = "2", Width = "3", Height = "4" }])
            .WithComponentSections(
                componentImages: [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "PHI2", File = "Scope baseline/U8_1.png", TimeDiv = "2us" }],
                componentPins: [new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" }],
                components: [new ComponentEntry { BoardLabel = "U8", FriendlyName = "Processor", Category = "IC" }]);

        Assert.Equal(
            ["Board labelling", "Oscilloscope baseline", "Component pinouts", "Board data"],
            CreditSuggestions.For(Published(), draft).Select(suggestion => suggestion.Category));
    }

    // A picture's colours changed is the board's data - not a new schematic image.
    [Fact]
    public void A_pictures_colours_changed_are_board_data()
    {
        BoardData draft = Draft().WithSchematics([new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png", SchematicHighlightColor = "Blue" }]);

        Assert.Equal([new CreditSuggestion("Board data", "")], CreditSuggestions.For(Published(), draft));
    }

    // A baseline's note nested a level deeper is work on the baseline, as "Review changes" counts it (owner decision,
    // 2026-11-08: an indent-only change of a note "is a change like any other change") - a space at its end is not.
    [Fact]
    public void A_baselines_note_nested_deeper_is_credited_and_a_space_at_its_end_is_not()
    {
        static BoardData WithNote(BoardData board, string note) => board.WithComponentSections(
            componentImages: [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "PHI2", File = "Scope baseline/U8_1.png", TimeDiv = "1us", Note = note }]);

        BoardData published = WithNote(Published(), "- clean");

        Assert.Equal(["Oscilloscope baseline"], CreditSuggestions.For(published, WithNote(Draft(), "  - clean")).Select(suggestion => suggestion.Category));
        Assert.Empty(CreditSuggestions.For(published, WithNote(Draft(), "- clean ")));
    }

    // A note written on a board file or link is board data (code review, 2026-11-08: the keys left the note out, so a
    // draft writing only notes was offered no credit) - its line ends no change.
    [Fact]
    public void A_note_on_a_board_file_or_link_is_board_data()
    {
        BoardData published = Published().WithBoardSections(
            boardLocalFiles: [new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = "Board local files/manual.pdf", Note = "first\nsecond" }],
            boardLinks: [new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://www.lemon64.com" }]);

        BoardData fileNote = published.WithBoardSections(boardLocalFiles: [new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = "Board local files/manual.pdf", Note = "first\nsecond, page 4" }]);
        BoardData linkNote = published.WithBoardSections(boardLinks: [new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://www.lemon64.com", Note = "the repair thread" }]);
        BoardData lineEnds = published.WithBoardSections(boardLocalFiles: [new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = "Board local files/manual.pdf", Note = "first\r\nsecond" }]);

        Assert.Equal([new CreditSuggestion("Board data", "")], CreditSuggestions.For(published, fileNote));
        Assert.Equal([new CreditSuggestion("Board data", "")], CreditSuggestions.For(published, linkNote));
        Assert.Empty(CreditSuggestions.For(published, lineEnds));
    }

    // Capitals and order are no change - nothing to credit.
    [Fact]
    public void Capitals_and_order_are_no_change()
    {
        BoardData draft = Draft().WithComponentSections(components: [new ComponentEntry { BoardLabel = "u8", FriendlyName = "CPU", Category = "IC" }]);

        Assert.Empty(CreditSuggestions.For(Published(), draft));
    }

    // A new board (nothing published): every part of it is the draft's work.
    [Fact]
    public void A_new_board_is_credited_for_all_it_holds()
    {
        Assert.Equal(
            ["Schematic image", "Board labelling", "Oscilloscope baseline", "Board data"],
            CreditSuggestions.For(null, Published()).Select(suggestion => suggestion.Category));
        Assert.Equal("Top", CreditSuggestions.For(null, Published())[0].SubCategory);
    }

    // The rows to add name the person; a credit the board already gives that name is left out, and a
    // blank name adds nothing.
    [Fact]
    public void A_credit_the_board_already_gives_the_name_is_not_added_again()
    {
        BoardData draft = Draft();
        CreditSuggestion[] picked = [new("Oscilloscope baseline", ""), new("Board data", "")];

        List<CreditEntry> added = CreditSuggestions.ToAdd(draft, picked, " anna ", " RHD ");

        CreditEntry credit = Assert.Single(added);
        Assert.Equal(("Board data", "anna", "RHD"), (credit.Category, credit.NameOrHandle, credit.Contact));
        Assert.Empty(CreditSuggestions.ToAdd(draft, picked, "  ", "RHD"));
    }
}
