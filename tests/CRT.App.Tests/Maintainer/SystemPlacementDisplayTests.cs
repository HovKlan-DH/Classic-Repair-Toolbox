using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// SystemPlacementDisplay - placing a new system in CRT's drop-down lists on the Systems screen
// (owner request, 2026-09-27: "The maintainer should order the new system, so it becomes visible
// in the right location for the drop-down lists. This must be done before it can be pushed to
// BETA.").
// ###########################################################################################
public sealed class SystemPlacementDisplayTests
{
    private const string C64 = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";
    private const string C128 = "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx";
    private const string C128Dcr = "Commodore/C128/250477/Data C128DCR 250477 v2.0.0.xlsx";
    private const string Spectrum = "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Issue 4B v2.0.0.xlsx";

    private const string Open128 = "Commodore/C128/310378 Open128";

    private static readonly IReadOnlyList<SystemListingRow> Listed =
    [
        new("Commodore/C64/250407", "Commodore 64", "250407 (long board)", C64),
        new("Commodore/C128/310378", "Commodore 128", "310378 (C128 & C128D)", C128),
        new("Commodore/C128/250477", "Commodore 128", "250477 (C128DCR)", C128Dcr),
        new("ZX Spectrum/Spectrum 16K-48K/Issue 4B", "ZX Spectrum 16K/48K", "Issue 4B", Spectrum),
    ];

    private static UnlistedSystemEntry Entry(SystemPlacement? placement = null, bool canPlace = true, bool inBeta = false) =>
        new(Open128, "Commodore", "C128", "310378 Open128", inBeta, canPlace, placement,
            new SystemPlacement("Commodore 128", "310378 Open128", string.Empty, C128Dcr));

    private static SystemListingAnswer Listing(params UnlistedSystemEntry[] unlisted) =>
        new(true, SystemPlacementDisplayTests.Listed, unlisted);

    // -----------------------------------------------------------------------------------
    // Where the panel starts, and what a position means
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_saved_placement_wins_over_the_servers_suggestion()
    {
        var saved = new SystemPlacement("Commodore 128", "Open128", "Notes.", C64);

        Assert.Equal(saved, SystemPlacementDisplay.Starting(SystemPlacementDisplayTests.Entry(saved)));
        Assert.Equal(C128Dcr, SystemPlacementDisplay.Starting(SystemPlacementDisplayTests.Entry()).AfterExcelDataFile);
    }

    [Theory]
    [InlineData(C64, 1)]
    [InlineData(C128Dcr, 3)]
    [InlineData(Spectrum, 4)]
    [InlineData(null, 0)]
    [InlineData("  ", 0)]
    public void The_new_row_starts_straight_after_the_row_it_was_placed_after(string? after, int expected)
    {
        Assert.Equal(expected, SystemPlacementDisplay.StartIndex(SystemPlacementDisplayTests.Listed, after, out bool isGone));
        Assert.False(isGone);
    }

    // The row is found by its SYSTEM too, and case does not matter - the server's own match
    // (MasterListing.Insert), so the screen shows the place the server would write.
    [Fact]
    public void The_row_above_is_found_by_its_system_and_regardless_of_case()
    {
        Assert.Equal(2, SystemPlacementDisplay.StartIndex(
            SystemPlacementDisplayTests.Listed, "commodore/c128/310378/Data C128 310378 v1.0.0.xlsx", out bool isGone));
        Assert.False(isGone);
    }

    // ###########################################################################################
    // *** A PLACE THAT NO LONGER EXISTS IS SAID, NOT GUESSED AT. *** The row it was placed after has
    // left the lists: the panel starts at the end and the screen says to place it again - the
    // approval would refuse it until then.
    // ###########################################################################################
    [Fact]
    public void Placed_after_a_row_that_has_left_the_lists_it_starts_at_the_end_and_says_so()
    {
        int start = SystemPlacementDisplay.StartIndex(
            SystemPlacementDisplayTests.Listed, "Commodore/VIC-20/999999/Data.xlsx", out bool isGone);

        Assert.Equal(4, start);
        Assert.True(isGone);
        Assert.Contains("Place it again", SystemPlacementDisplay.SavedState(
            SystemPlacementDisplayTests.Entry(new SystemPlacement("H", "B", string.Empty, "gone.xlsx")), afterIsGone: true));
    }

    [Theory]
    [InlineData(0, null)]
    [InlineData(1, C64)]
    [InlineData(3, C128Dcr)]
    [InlineData(4, Spectrum)]
    public void A_position_means_after_the_row_above_it_or_first(int index, string? expected)
    {
        Assert.Equal(expected, SystemPlacementDisplay.AfterAt(
            SystemPlacementDisplayTests.Listed.Select(row => row.ExcelDataFile).ToList(), index));
    }

    // Placing where the panel starts and saving without moving it saves the same place - the round
    // trip the screen performs when a maintainer accepts the suggestion.
    [Theory]
    [InlineData(null)]
    [InlineData(C64)]
    [InlineData(Spectrum)]
    public void Where_it_starts_and_what_that_position_means_agree(string? after)
    {
        int start = SystemPlacementDisplay.StartIndex(SystemPlacementDisplayTests.Listed, after, out _);

        Assert.Equal(after, SystemPlacementDisplay.AfterAt(
            SystemPlacementDisplayTests.Listed.Select(row => row.ExcelDataFile).ToList(), start));
    }

    // -----------------------------------------------------------------------------------
    // The names
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE SERVER's OWN RULE. *** CRT keys a board by "hardware name|board name", so names
    // another system is listed under would be one board twice; the words are the server's too,
    // so the screen and a refused save say the same thing.
    // ###########################################################################################
    [Fact]
    public void Names_another_system_is_listed_under_cannot_be_saved()
    {
        string? problem = SystemPlacementDisplay.NameProblem(
            SystemPlacementDisplayTests.Listed, Open128, "commodore 128", "250477 (c128dcr)", string.Empty);

        Assert.Equal(
            MasterListing.NamesTakenMessage(new MasterListingRow("Commodore 128", "250477 (C128DCR)", C128Dcr, string.Empty)),
            problem);

        Assert.Null(SystemPlacementDisplay.NameProblem(
            SystemPlacementDisplayTests.Listed, Open128, "Commodore 128", "310378 Open128", "Open-source replica."));
    }

    [Theory]
    [InlineData("", "310378 Open128")]
    [InlineData("Commodore 128", "   ")]
    public void A_blank_name_cannot_be_saved(string hardware, string board)
    {
        Assert.NotNull(SystemPlacementDisplay.NameProblem(SystemPlacementDisplayTests.Listed, Open128, hardware, board, string.Empty));
    }

    // -----------------------------------------------------------------------------------
    // Where CRT shows it
    // -----------------------------------------------------------------------------------

    private static List<(string Hardware, string Board)> WithNew(int index, string hardware, string board)
    {
        List<(string Hardware, string Board)> order = SystemPlacementDisplayTests.Listed.Select(row => (row.HardwareName, row.BoardName)).ToList();
        order.Insert(index, (hardware, board));
        return order;
    }

    [Fact]
    public void Placed_among_its_hardwares_boards_it_is_that_hardwares_next_board()
    {
        Assert.Equal(
            "Commodore 128 is hardware 2 of 3 in CRT, and 310378 Open128 its board 3 of 3.",
            SystemPlacementDisplay.WhereInCrt(SystemPlacementDisplayTests.WithNew(3, "Commodore 128", "310378 Open128"), 3));

        Assert.Equal(
            "Commodore 128 is hardware 2 of 3 in CRT, and 310378 Open128 its board 1 of 3.",
            SystemPlacementDisplayTests.WhereAt(1));
    }

    // ###########################################################################################
    // *** CRT LISTS A HARDWARE ONCE, WHERE IT FIRST APPEARS (Main.BoardSelection). *** A row dragged
    // away from its hardware's other boards still shows under that hardware - so the line says
    // what CRT will really do, which is what the maintainer is choosing.
    // ###########################################################################################
    [Fact]
    public void Placed_away_from_its_hardwares_other_boards_it_still_shows_under_that_hardware()
    {
        Assert.Equal(
            "Commodore 128 is hardware 2 of 3 in CRT, and 310378 Open128 its board 3 of 3.",
            SystemPlacementDisplayTests.WhereAt(4));
    }

    [Fact]
    public void A_new_hardware_name_is_a_new_hardware_where_it_stands()
    {
        Assert.Equal(
            "Commodore 128D is hardware 1 of 4 in CRT, and Open128 its board 1 of 1.",
            SystemPlacementDisplay.WhereInCrt(SystemPlacementDisplayTests.WithNew(0, "Commodore 128D", "Open128"), 0));

        Assert.Equal(
            "Give it a hardware name and a board name to see where CRT shows it.",
            SystemPlacementDisplay.WhereInCrt(SystemPlacementDisplayTests.WithNew(0, " ", "Open128"), 0));
    }

    private static string WhereAt(int index) =>
        SystemPlacementDisplay.WhereInCrt(SystemPlacementDisplayTests.WithNew(index, "Commodore 128", "310378 Open128"), index);

    // -----------------------------------------------------------------------------------
    // The Systems list, its badge and the review line
    // -----------------------------------------------------------------------------------

    private static SystemOverviewEntry System(string id) =>
        new(id, "M", "H", id, true, null, false, true, null, null, null, 0);

    // The ones to act on first, then CRT's own order, then everything else as the server sent it.
    [Fact]
    public void The_systems_list_puts_the_ones_waiting_for_a_place_first_then_follows_CRTs_order()
    {
        SystemListingAnswer listing = SystemPlacementDisplayTests.Listing(SystemPlacementDisplayTests.Entry());

        IReadOnlyList<SystemOverviewEntry> ordered = SystemPlacementDisplay.InListOrder(
            [
                SystemPlacementDisplayTests.System("Amiga/A500/Rev 6A"),
                SystemPlacementDisplayTests.System("ZX Spectrum/Spectrum 16K-48K/Issue 4B"),
                SystemPlacementDisplayTests.System("Commodore/C128/310378"),
                SystemPlacementDisplayTests.System(Open128),
                SystemPlacementDisplayTests.System("Commodore/C64/250407"),
                SystemPlacementDisplayTests.System("Amstrad/CPC 464/MC0001"),
            ],
            listing);

        Assert.Equal(
            [Open128, "Commodore/C64/250407", "Commodore/C128/310378", "ZX Spectrum/Spectrum 16K-48K/Issue 4B", "Amiga/A500/Rev 6A", "Amstrad/CPC 464/MC0001"],
            ordered.Select(system => system.SystemId));

        // No listing known: the server's order, untouched.
        IReadOnlyList<SystemOverviewEntry> unknown = [SystemPlacementDisplayTests.System("B/B/B"), SystemPlacementDisplayTests.System("A/A/A")];
        Assert.Same(unknown, SystemPlacementDisplay.InListOrder(unknown, null));
    }

    [Fact]
    public void A_system_waiting_for_a_place_is_marked_until_its_place_is_saved_and_no_other_is()
    {
        Assert.Equal(
            "Needs a place in the drop-down lists",
            SystemPlacementDisplay.ListMark(SystemPlacementDisplayTests.Listing(SystemPlacementDisplayTests.Entry()), Open128));

        Assert.Equal(
            "Placed - listed when published to BETA",
            SystemPlacementDisplay.ListMark(
                SystemPlacementDisplayTests.Listing(SystemPlacementDisplayTests.Entry(new SystemPlacement("H", "B", string.Empty, null))),
                Open128));

        Assert.Null(SystemPlacementDisplay.ListMark(SystemPlacementDisplayTests.Listing(SystemPlacementDisplayTests.Entry()), "Commodore/C64/250407"));
        Assert.Null(SystemPlacementDisplay.ListMark(null, Open128));
    }

    // The Systems button counts the systems THIS account can place and nobody has yet.
    [Fact]
    public void The_badge_counts_only_unplaced_systems_this_account_can_place()
    {
        SystemListingAnswer listing = SystemPlacementDisplayTests.Listing(
            SystemPlacementDisplayTests.Entry(),
            SystemPlacementDisplayTests.Entry(canPlace: false) with { SystemId = "Amstrad/CPC 6128/MC0020" },
            SystemPlacementDisplayTests.Entry(new SystemPlacement("H", "B", string.Empty, null)) with { SystemId = "Amiga/A500/Rev 6A" });

        Assert.Equal(1, SystemPlacementDisplay.NeedingPlace(listing));
        Assert.Equal(0, SystemPlacementDisplay.NeedingPlace(null));
    }

    // Said above the submission's table before the maintainer reads it - the approval would be
    // refused. Gone once a place is saved, and never for a listed system.
    [Fact]
    public void A_submission_for_an_unplaced_new_system_says_so_above_its_table()
    {
        Assert.Contains(
            "cannot be approved. Place it on the Systems screen first.",
            SystemPlacementDisplay.ReviewLine(SystemPlacementDisplayTests.Listing(SystemPlacementDisplayTests.Entry()), Open128));

        Assert.Null(SystemPlacementDisplay.ReviewLine(
            SystemPlacementDisplayTests.Listing(SystemPlacementDisplayTests.Entry(new SystemPlacement("H", "B", string.Empty, null))), Open128));
        Assert.Null(SystemPlacementDisplay.ReviewLine(SystemPlacementDisplayTests.Listing(), "Commodore/C64/250407"));
        Assert.Null(SystemPlacementDisplay.ReviewLine(null, Open128));
    }

    [Fact]
    public void A_listed_system_names_the_entry_CRT_shows_it_under()
    {
        Assert.Equal(
            "In CRT's drop-down lists as Commodore 128 / 250477 (C128DCR)",
            SystemPlacementDisplay.ListedLine(SystemPlacementDisplayTests.Listing(), "commodore/c128/250477"));

        Assert.Null(SystemPlacementDisplay.ListedLine(SystemPlacementDisplayTests.Listing(), Open128));
    }

    // What saving does depends on whether the board is already in BETA.
    [Fact]
    public void The_explanation_says_when_the_system_joins_the_lists()
    {
        Assert.Contains("saving adds it to BETA's lists at once", SystemPlacementDisplay.Explanation(SystemPlacementDisplayTests.Entry(inBeta: true)));
        Assert.Contains("cannot be published to BETA until it has a place", SystemPlacementDisplay.Explanation(SystemPlacementDisplayTests.Entry()));
    }

    // ###########################################################################################
    // *** AFTER A SAVE WITH NO ANSWER IN TWO MINUTES, THE LISTING SAYS WHETHER IT LANDED
    // (2026-09-28). *** Either the system is in the list itself (placed straight into BETA's file),
    // or it is still unlisted with exactly the place asked for - names trimmed as the server keeps
    // them. A different place, or none, has not landed.
    // ###########################################################################################
    [Fact]
    public void A_place_that_was_kept_is_found_in_the_listing_read_again()
    {
        var request = new SetPlacementRequest("Commodore/C128/310378", " Commodore 128 ", "310378", "", "Commodore/C64/C64.xlsx");
        var suggested = new SystemPlacement("x", "y", "", null);

        SystemListingAnswer Unlisted(SystemPlacement? kept) =>
            new(true, [], [new UnlistedSystemEntry("Commodore/C128/310378", "Commodore", "C128", "310378", InBeta: true, CanPlace: true, kept, suggested)]);

        Assert.True(SystemPlacementDisplay.HasLanded(request, Unlisted(new SystemPlacement("Commodore 128", "310378", "", "Commodore/C64/C64.xlsx"))));
        Assert.False(SystemPlacementDisplay.HasLanded(request, Unlisted(null)));
        Assert.False(SystemPlacementDisplay.HasLanded(request, Unlisted(new SystemPlacement("Commodore 128", "310378", "", "Commodore/VIC-20/VIC-20.xlsx"))));

        var listed = new SystemListingAnswer(true, [new SystemListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/C128.xlsx")], []);
        Assert.True(SystemPlacementDisplay.HasLanded(request, listed));
    }
}
