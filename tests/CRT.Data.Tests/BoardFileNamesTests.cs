using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The names a board gives its files in the CRT 4.0 tree: the path a row cited, without the part
// that said where it was kept - and two different files never under one name.
public sealed class BoardFileNamesTests
{
    private const string Board = "Commodore/C64/250407";
    private const string ShaA = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
    private const string ShaB = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
    private const string ShaC = "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc";

    private static BoardFileNames Names() =>
        new(Board, [Board, "Commodore/C64/250425", "Commodore/C64/250425 (short)", "Amstrad/CPC 664/MC0005A"]);

    [Theory]
    [InlineData("Commodore/C64/250407/Scope baseline/U19_1_PAL.png", "Scope baseline/U19_1_PAL.png")]
    [InlineData("Commodore/C64/250407/Board.png", "Board.png")]
    [InlineData("Commodore/Shared files/Component images/7406.png", "Component images/7406.png")]
    [InlineData("Amstrad/Shared files/Component images/Z80.png", "Component images/Z80.png")]
    [InlineData("Generic shared files/Component local files/7406.PDF", "Component local files/7406.PDF")]
    [InlineData("Commodore/C64/250425/Board layout.png", "Board layout.png")]
    [InlineData("Somewhere/else.png", "Somewhere/else.png")]
    public void A_file_is_named_without_the_folders_that_said_where_it_was_kept(string cited, string name)
    {
        Assert.Equal(name, Names().Add(cited, ShaA));
    }

    // "250425 (short)" begins with "250425" - it must be stripped as itself, not as its prefix.
    [Fact]
    public void A_board_whose_id_begins_with_another_boards_is_stripped_as_itself()
    {
        Assert.Equal("Top.png", Names().Add("Commodore/C64/250425 (short)/Top.png", ShaA));
    }

    // Matched ignoring capitals, as Windows and macOS clients find the file.
    [Fact]
    public void The_folders_are_matched_ignoring_capitals_and_the_rest_keeps_its_spelling()
    {
        Assert.Equal("Scope Baseline/U1.PNG", Names().Add("commodore/c64/250407/Scope Baseline/U1.PNG", ShaA));
        Assert.Equal("Component images/x.png", Names().Add("GENERIC SHARED FILES/Component images/x.png", ShaB));
    }

    [Theory]
    [InlineData(@"Commodore\C64\250407\Board.png")]
    [InlineData("/Commodore/C64/250407/Board.png")]
    [InlineData(" Commodore/C64/250407/./Board.png ")]
    [InlineData("Commodore/C64/250407/Pictures/../Board.png")]
    [InlineData("Commodore//C64/250407/Board.png")]
    public void Every_spelling_of_one_path_is_one_file(string cited)
    {
        BoardFileNames names = Names();

        Assert.Equal("Board.png", names.Add("Commodore/C64/250407/Board.png", ShaA));
        Assert.Equal("Board.png", names.Add(cited, ShaB));
        Assert.Single(names.NamesByPath);
    }

    // A ".." climbing out of the tree is kept, so the file is not found instead of found elsewhere.
    [Fact]
    public void A_path_climbing_out_of_the_tree_keeps_its_climb()
    {
        Assert.Equal("../outside.png", BoardFileNames.Normalise("../outside.png"));
        Assert.Equal("../outside.png", BoardFileNames.Normalise("a/../../outside.png"));
    }

    [Fact]
    public void Two_different_files_wanting_one_name_get_numbered_names_in_the_order_they_ask()
    {
        BoardFileNames names = Names();

        Assert.Equal("Component images/7406.png", names.Add("Commodore/C64/250407/Component images/7406.png", ShaA));
        Assert.Equal("Component images/7406 (2).png", names.Add("Commodore/Shared files/Component images/7406.png", ShaB));
        Assert.Equal("Component images/7406 (3).png", names.Add("Generic shared files/Component images/7406.png", ShaC));
    }

    // Names are compared ignoring capitals - "x.png" and "X.PNG" are one file on Windows.
    [Fact]
    public void Names_differing_only_in_capitals_are_one_name()
    {
        BoardFileNames names = Names();

        Assert.Equal("Pinout.png", names.Add("Commodore/C64/250407/Pinout.png", ShaA));
        Assert.Equal("PINOUT (2).PNG", names.Add("Generic shared files/PINOUT.PNG", ShaB));
    }

    // The same bytes under two paths are the same file: one name, not two.
    [Fact]
    public void Two_paths_to_the_same_bytes_share_the_name()
    {
        BoardFileNames names = Names();

        Assert.Equal("Component images/7406.png", names.Add("Commodore/Shared files/Component images/7406.png", ShaA));
        Assert.Equal("Component images/7406.png", names.Add("Generic shared files/Component images/7406.png", ShaA));
    }

    [Fact]
    public void A_name_without_an_extension_is_numbered_at_its_end()
    {
        BoardFileNames names = Names();

        names.Add("Commodore/C64/250407/README", ShaA);

        Assert.Equal("README (2)", names.Add("Generic shared files/README", ShaB));
    }

    // A rename's image the board no longer holds still needs a name - without taking one.
    [Fact]
    public void NameFor_answers_a_path_it_was_never_given_without_registering_it()
    {
        BoardFileNames names = Names();

        Assert.Equal("Old top.png", names.NameFor("Commodore/C64/250407/Old top.png"));
        Assert.Empty(names.NamesByPath);
        Assert.Equal("Old top.png", names.Add("Commodore/Shared files/Old top.png", ShaB));
        Assert.Equal(string.Empty, names.NameFor("  "));
    }

    [Fact]
    public void A_blank_path_is_refused()
    {
        Assert.Throws<ArgumentException>(() => Names().Add(" ", ShaA));
        Assert.Throws<ArgumentException>(() => Names().Assign(" ", "x.png", ShaA));
        Assert.Throws<ArgumentException>(() => Names().Assign("Commodore/C64/250407/x.png", " ", ShaA));
    }

    // A draft converted beside its published board must never give a new file a published name.
    [Fact]
    public void A_reserved_name_is_not_given_to_other_bytes()
    {
        BoardFileNames names = Names();
        names.Reserve("Board.png", ShaA);

        Assert.Equal("Board (2).png", names.Add("Commodore/C64/250407/Board.png", ShaB));
        Assert.Equal("Board.png", names.Add("Generic shared files/Board.png", ShaA));
    }

    // A file the draft replaced keeps its published name, with its new bytes.
    [Fact]
    public void An_assigned_name_is_the_path_s_from_then_on()
    {
        BoardFileNames names = Names();
        names.Reserve("Board.png", ShaA);
        names.Assign("Commodore/C64/250407/Pictures/Board.png", "Board.png", ShaB);

        Assert.Equal("Board.png", names.NameFor("Commodore/C64/250407/Pictures/Board.png"));
        Assert.Equal("Board.png", names.Add("Commodore/C64/250407/Pictures/Board.png", ShaB));
        Assert.Equal("Board (2).png", names.Add("Commodore/C64/250407/Board.png", ShaA));
    }
}
