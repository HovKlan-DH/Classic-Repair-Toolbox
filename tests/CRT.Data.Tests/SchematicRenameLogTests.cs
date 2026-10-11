using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// SchematicRenameLog - a board's renamed schematic images, and where CRT follows a name the user's
// own data uses (owner decisions, 2026-10-10: an accepted rename "must also change for everyone
// updating to this data", kept in the board's JSON file and not in Excel - with the image file,
// "to be unambiguous" when a name is later given to another image).
public sealed class SchematicRenameLogTests
{
    private static BoardSchematicEntry Entry(string name, string file) =>
        new() { SchematicName = name, SchematicImageFile = $"Commodore/C64/250407/{file}" };

    private static BoardData Board(params BoardSchematicEntry[] schematics)
    {
        var board = new BoardData();
        board.Schematics.AddRange(schematics);
        return board;
    }

    private static SchematicImage Image(string name, string file) => new(name, $"Commodore/C64/250407/{file}");

    private static string? Where(
        string name,
        SchematicImage[] board,
        SchematicImage[]? remembered = null,
        SchematicRenameRecord[]? published = null) =>
        SchematicRenameLog.Resolver(board, remembered, published ?? [])(name);

    // ------------------------------------------------------------------ what a publish adds

    // The server's own rule for a table edit: the same image file under a new name is a rename. The
    // record names that file.
    [Fact]
    public void A_publish_that_renames_an_image_adds_it_to_the_list_with_its_file_and_date()
    {
        List<SchematicRenameRecord> log = SchematicRenameLog.ForPublish(
            [new("Old", "Older", "Commodore/C64/250407/old.png", "2026-September-1")],
            Board(Entry("Top", "top.png"), Entry("Bottom", "bottom.png")),
            Board(Entry("PCB top side", "top.png"), Entry("Bottom", "bottom.png")),
            "2026-October-10");

        Assert.Equal(
            [
                new SchematicRenameRecord("Old", "Older", "Commodore/C64/250407/old.png", "2026-September-1"),
                new SchematicRenameRecord("Top", "PCB top side", "Commodore/C64/250407/top.png", "2026-October-10"),
            ],
            log);
    }

    // A deleted image, or a new file under a new name, is no rename - nothing to follow.
    [Fact]
    public void A_deletion_or_a_new_file_under_a_new_name_is_not_a_rename()
    {
        Assert.Empty(SchematicRenameLog.ForPublish(
            [],
            Board(Entry("Top", "top.png"), Entry("Bottom", "bottom.png")),
            Board(Entry("Bottom", "bottom.png"), Entry("PCB top side", "new-top.png")),
            "2026-October-10"));
    }

    [Fact]
    public void A_publish_without_renames_keeps_the_list_as_it_was()
    {
        SchematicRenameRecord[] existing = [new("Top", "PCB top side", "Commodore/C64/250407/top.png", "2026-October-10")];

        Assert.Equal(existing, SchematicRenameLog.ForPublish(existing, Board(Entry("PCB top side", "top.png")), Board(Entry("PCB top side", "top.png")), "2026-November-1"));

        // A NEW board has nothing published to rename.
        Assert.Equal(existing, SchematicRenameLog.ForPublish(existing, null, Board(Entry("PCB top side", "top.png")), "2026-November-1"));
    }

    // ------------------------------------------------------------------ following by the REMEMBERED file

    private static readonly SchematicImage[] Before = [Image("Top", "top.png"), Image("Bottom", "bottom.png")];

    [Fact]
    public void A_remembered_name_goes_where_its_image_file_is_now()
    {
        SchematicImage[] now = [Image("PCB top side", "top.png"), Image("Bottom", "bottom.png")];

        Assert.Equal("PCB top side", Where("top", now, Before));
        Assert.Null(Where("Bottom", now, Before));
    }

    // ###########################################################################################
    // *** THE OWNER'S CASE: a name later given to ANOTHER image. *** "Top" was renamed and
    // bottom.png now carries the name "Top". What the user had on "Top" was made on top.png, so it
    // goes with top.png - not with the name.
    // ###########################################################################################
    [Fact]
    public void A_name_given_to_another_image_does_not_keep_what_was_made_on_the_first()
    {
        SchematicImage[] now = [Image("PCB top side", "top.png"), Image("Top", "bottom.png")];

        Assert.Equal("PCB top side", Where("Top", now, Before));
        Assert.Equal("Top", Where("Bottom", now, Before));
    }

    // Shown with the draft, then without it (discarded, rejected, "view boards as officially
    // published", or a rename pushed back out of BETA): the image's things come back.
    [Fact]
    public void A_rename_undone_brings_the_names_back()
    {
        SchematicImage[] renamed = [Image("PCB top side", "top.png"), Image("Bottom", "bottom.png")];

        Assert.Equal("Top", Where("PCB top side", Before, renamed));
    }

    [Fact]
    public void Two_images_swapping_names_each_keep_their_own_things()
    {
        SchematicImage[] swapped = [Image("Bottom", "top.png"), Image("Top", "bottom.png")];

        Assert.Equal("Bottom", Where("Top", swapped, Before));
        Assert.Equal("Top", Where("Bottom", swapped, Before));
    }

    // One file under two names now: which one was meant cannot be told, so nothing moves.
    [Fact]
    public void An_image_file_on_several_images_now_moves_nothing()
    {
        SchematicImage[] twice = [Image("Top A", "top.png"), Image("Top B", "top.png")];

        Assert.Null(Where("Top", twice, Before));
    }

    // The image gone from the board: the name rules decide - here the published list.
    [Fact]
    public void A_remembered_image_that_left_the_board_falls_back_to_the_published_list()
    {
        SchematicImage[] now = [Image("PCB top side", "new-top.png")];

        Assert.Null(Where("Top", now, Before));
        Assert.Equal(
            "PCB top side",
            Where("Top", now, Before, [new("Top", "PCB top side", "Commodore/C64/250407/new-top.png", "2026-October-10")]));
    }

    // ------------------------------------------------------------------ the PUBLISHED list (nothing remembered)

    [Fact]
    public void Without_a_memory_a_name_follows_the_file_its_latest_rename_names()
    {
        SchematicImage[] now = [Image("PCB top side", "top.png"), Image("Bottom", "bottom.png")];

        Assert.Equal(
            "PCB top side",
            Where("top", now, published: [new("Top", "PCB top", "Commodore/C64/250407/top.png", "2026-October-10")]));
    }

    // An older list without files (or a file no image has any more): renames followed by name.
    [Fact]
    public void Without_a_file_to_go_by_the_renames_are_followed_by_name()
    {
        SchematicImage[] now = [Image("PCB top side", "x.png")];

        Assert.Equal(
            "PCB top side",
            Where("Top", now, published: [new("Top", "PCB top", "", "2026-October-10"), new("PCB top", "PCB top side", "", "2026-November-2")]));

        // Renamed away and back again: the walk does not go round for ever.
        Assert.Null(Where("A", now, published: [new("A", "B", "", ""), new("B", "A", "", "")]));
    }

    // ###########################################################################################
    // Without a memory, a name the board HAS stays: whether an old worklog on it meant the image
    // renamed away or the one with the name now cannot be known, and nothing moves on a guess.
    // ###########################################################################################
    [Fact]
    public void Without_a_memory_a_name_the_board_has_stays()
    {
        SchematicImage[] now = [Image("PCB top side", "top.png"), Image("Top", "bottom.png")];

        Assert.Null(Where("Top", now, published: [new("Top", "PCB top side", "Commodore/C64/250407/top.png", "2026-October-10")]));
    }

    [Fact]
    public void A_blank_name_or_one_nothing_leads_from_goes_nowhere()
    {
        SchematicImage[] now = [Image("PCB top side", "top.png")];

        Assert.Null(Where("  ", now, Before));
        Assert.Null(Where("Keyboard", now, Before, [new("Top", "PCB top side", "Commodore/C64/250407/top.png", "")]));
    }

    // ------------------------------------------------------------------ comparing

    [Fact]
    public void Image_paths_compare_without_regard_to_slashes_or_capitals_and_blank_is_no_file()
    {
        Assert.True(SchematicRenameLog.SameFile("Commodore\\C64\\250407\\Top.PNG", " Commodore/C64/250407/top.png "));
        Assert.False(SchematicRenameLog.SameFile("", ""));
        Assert.False(SchematicRenameLog.SameFile("a.png", "b.png"));
    }

    [Fact]
    public void The_same_images_in_another_order_are_the_same_board_and_a_rename_is_not()
    {
        Assert.True(SchematicRenameLog.SameImages(Before, [Image("Bottom", "bottom.png"), Image("Top", "top.png")]));
        Assert.False(SchematicRenameLog.SameImages(Before, [Image("PCB top side", "top.png"), Image("Bottom", "bottom.png")]));
    }

    // ###########################################################################################
    // *** A NULL NAME OR FILE IS BLANK, NOT A CRASH (code review, 2026-10-11). *** The images CRT
    // remembers come from the user's settings file, and {"name": null} there threw in the middle of
    // the board load - every time the board was shown. The published list can carry one too.
    // ###########################################################################################
    [Fact]
    public void A_null_name_or_file_reads_as_blank()
    {
        SchematicImage nothing = new(null!, null!);

        Assert.False(SchematicRenameLog.SameImages([nothing, Image("Bottom", "bottom.png")], Before));
        Assert.True(SchematicRenameLog.SameImages([nothing], [new SchematicImage(" ", "")]));

        Assert.Equal("PCB top side", Where("Top", [Image("PCB top side", "top.png"), nothing], remembered: [nothing, Image("Top", "top.png")]));
        Assert.Null(Where("Gone", [nothing], published: [new SchematicRenameRecord("Gone", null!, null!, null!)]));
        Assert.Null(Where("Gone", [Image(null!, "top.png")], remembered: [Image("Gone", "top.png")]));
    }
}
