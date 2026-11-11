using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// WHICH REGIONS A BOARD COMES IN (owner decisions, 2026-11-01): the one rule behind the PAL/NTSC
// buttons, stored in the board file and published - and every side of it, as a change to shared data
// fails only when one side has not moved: the board file, a submission, the publish, the review, the
// draft's own bookkeeping.
// ###########################################################################################
public sealed class BoardRegionsTests
{
    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private static ComponentEntry Component(string label, string region = "") =>
        new() { BoardLabel = label, Category = "IC", Region = region };

    // What decides whether the switch is offered at all, from the rows of a board that has not said.
    [Theory]
    [InlineData("PAL", true)]
    [InlineData("NTSC", true)]
    [InlineData("pal", true)]
    [InlineData("  NTSC  ", true)]
    [InlineData("", false)]
    [InlineData("SECAM", false)]
    public void A_board_that_has_not_said_has_regions_when_a_row_says_PAL_or_NTSC(string region, bool expected)
    {
        Assert.Equal(expected, BoardRegions.Has(Board(Component("U1", region))));
    }

    [Fact]
    public void No_board_has_no_regions()
    {
        Assert.False(BoardRegions.Has(null));
        Assert.False(BoardRegions.Has(new BoardData()));
    }

    // ###########################################################################################
    // The trap the old rule held: it asked the COMPONENT rows alone, so a board whose baselines were per
    // region and whose components were not hid the buttons - and one region's baselines for good.
    // ###########################################################################################
    [Fact]
    public void A_regional_scope_baseline_pin_or_replacement_is_enough()
    {
        var withBaseline = new BoardData
        {
            Components = [Component("U1")],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", TimeDiv = "1ms", File = "a.png" }],
        };
        var withPin = new BoardData { Components = [Component("U1")], ComponentPins = [new ComponentPinEntry { BoardLabel = "U1", Region = "PAL", Pin = "1" }] };
        var withPart = new BoardData { Components = [Component("U1")], ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U1", Region = "PAL", PartNumber = "x" }] };

        Assert.True(BoardRegions.Has(withBaseline));
        Assert.True(BoardRegions.Has(withPin));
        Assert.True(BoardRegions.Has(withPart));
    }

    [Fact]
    public void A_board_that_says_so_has_them_with_no_row_saying_a_region()
    {
        BoardData board = BoardRegions.WithRegions(Board(Component("U1")), true);

        Assert.Equal(["PAL", "NTSC"], board.Regions);
        Assert.True(BoardRegions.Has(board));
    }

    [Fact]
    public void Saying_what_it_says_already_is_no_change()
    {
        BoardData with = BoardRegions.WithRegions(Board(Component("U1")), true);
        BoardData without = Board(Component("U1"));

        Assert.Same(with, BoardRegions.WithRegions(with, true));
        Assert.Same(without, BoardRegions.WithRegions(without, false));
    }

    // A row only one region shows would go on being there with no button left to reach it.
    [Fact]
    public void The_setting_cannot_be_taken_away_while_a_row_says_a_region()
    {
        BoardData board = BoardRegions.WithRegions(Board(Component("U1", "PAL")), true);

        Assert.Same(board, BoardRegions.WithRegions(board, false));
        Assert.Equal(1, BoardRegions.RegionalRows(board));
    }

    [Fact]
    public void Taken_away_with_no_row_saying_one_it_is_gone()
    {
        BoardData board = BoardRegions.WithRegions(Board(Component("U1")), true);

        BoardData without = BoardRegions.WithRegions(board, false);

        Assert.Empty(without.Regions);
        Assert.False(BoardRegions.Has(without));
    }

    // ---- every side of it ----------------------------------------------------------------------

    // A board that does not say is written byte for byte as before: nothing published needs converting.
    [Fact]
    public void The_board_file_says_regions_only_for_a_board_that_does()
    {
        BoardDocument without = BoardDocumentMapping.ToDocument(new BoardContent(Board(Component("U1")), [], []), name => name, []);
        BoardDocument with = BoardDocumentMapping.ToDocument(
            new BoardContent(BoardRegions.WithRegions(Board(Component("U1")), true), [], []), name => name, []);

        Assert.DoesNotContain("\"regions\"", DataTreeJson.Write(without));

        string written = DataTreeJson.Write(with);
        Assert.Contains("\"regions\"", written);

        BoardContent read = BoardDocumentMapping.ToContent(DataTreeJson.ReadBoard(written));
        Assert.Equal(["PAL", "NTSC"], read.Board.Regions);
        Assert.Equal(written, DataTreeJson.Write(BoardDocumentMapping.ToDocument(read, name => name, [])));
    }

    [Fact]
    public void A_submission_carries_them_to_the_publish()
    {
        SubmissionRows rows = SubmissionRowsBoard.FromBoard(BoardRegions.WithRegions(Board(Component("U1")), true));
        var published = Board(Component("U1"));

        Assert.Equal(["PAL", "NTSC"], rows.Regions);
        Assert.Equal(["PAL", "NTSC"], PublishMerge.Build(new SubmissionManifest { Rows = rows }, published).Regions);
    }

    // Null is an older client's "not sent", which keeps what is published; empty is "none".
    [Fact]
    public void A_submission_not_sending_them_keeps_what_is_published_and_an_empty_one_clears_it()
    {
        BoardData published = BoardRegions.WithRegions(Board(Component("U1")), true);

        Assert.Equal(["PAL", "NTSC"], PublishMerge.Build(new SubmissionManifest { Rows = new SubmissionRows { Regions = null } }, published).Regions);
        Assert.Empty(PublishMerge.Build(new SubmissionManifest { Rows = new SubmissionRows { Regions = [] } }, published).Regions);
    }

    [Fact]
    public void An_amendment_keeps_the_submissions_whatever_a_request_says()
    {
        var current = new SubmissionRows { Regions = ["PAL", "NTSC"] };
        var edited = new SubmissionRows { Regions = [] };

        Assert.Equal(["PAL", "NTSC"], SubmissionRowsBoard.WithTableSections(current, edited).Regions);
    }

    [Fact]
    public void The_review_and_the_boards_history_are_told()
    {
        BoardData published = Board(Component("U1"));
        BoardData submitted = BoardRegions.WithRegions(Board(Component("U1")), true);

        ReviewChangeSummary changed = ReviewSummary.Compare(published, submitted);
        Assert.True(changed.RegionsNow);
        Assert.True(changed.HasChanges);
        Assert.Contains("regions changed", changed.Describe(), StringComparison.Ordinal);

        SubmissionChanges facts = SubmissionChangeFacts.Build(changed);
        Assert.True(facts.RegionsNow);
        Assert.True(SubmissionChangeFacts.HasChanges(facts));

        Assert.Null(ReviewSummary.Compare(published, Board(Component("U1"))).RegionsNow);
    }

    // ###########################################################################################
    // A NEW BOARD SAYS IT TOO (code review, 2026-11-02): nothing to differ from, so nothing was said - and
    // a new board's "regions" reached the published file with no maintainer told, no sheet of the table
    // showing it and an amendment unable to change it. One with none says nothing, as nothing is new.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_regions_are_said_to_the_review()
    {
        ReviewChangeSummary regional = ReviewSummary.Compare(published: null, BoardRegions.WithRegions(Board(Component("U1")), true));

        Assert.True(regional.RegionsNow);
        Assert.True(SubmissionChangeFacts.Build(regional).RegionsNow);

        // By the one rule: a row saying PAL or NTSC is a board with regions too.
        Assert.True(ReviewSummary.Compare(published: null, Board(new ComponentEntry { BoardLabel = "U1", Region = "PAL" })).RegionsNow);

        Assert.Null(ReviewSummary.Compare(published: null, Board(Component("U1"))).RegionsNow);
    }

    [Fact]
    public void Review_changes_says_it()
    {
        BoardData published = Board(Component("U1"));
        BoardData draft = BoardRegions.WithRegions(Board(Component("U1")), true);

        Assert.Contains(
            DraftChangeSummary.Build(new BoardContent(published, [], []), new BoardContent(draft, [], [])).AsLines(),
            line => line.Contains(BoardRegions.ChangedLine(true), StringComparison.Ordinal));
    }

    // A draft whose only work is this has something to send, and is not retired or let go of as empty.
    [Fact]
    public void It_counts_as_a_change_of_the_draft()
    {
        Assert.True(new DraftOtherChanges(0, 0, 0, Regions: true).Any);
        Assert.Equal("regions changed", new DraftOtherChanges(0, 0, 0, Regions: true).Describe());
        Assert.Equal("1 file changed, regions changed", new DraftOtherChanges(1, 0, 0, Regions: true).Describe());
    }

    // ###########################################################################################
    // "Bring my draft up to date": the contributor's when they changed it, else the update's - the rule
    // of every other part of a draft.
    // ###########################################################################################
    [Fact]
    public void Bringing_a_draft_up_to_date_keeps_the_contributors_and_else_takes_the_updates()
    {
        BoardData plain = Board(Component("U1"));
        BoardData ticked = BoardRegions.WithRegions(Board(Component("U1")), true);

        Assert.Equal(["PAL", "NTSC"], DraftRebase.Merge(plain, ticked, plain, null, null, null).Board.Regions);
        Assert.Equal(["PAL", "NTSC"], DraftRebase.Merge(plain, plain, ticked, null, null, null).Board.Regions);
        Assert.Empty(DraftRebase.Merge(ticked, plain, ticked, null, null, null).Board.Regions);
    }

    // A file or a link of one region counts as well (owner request, 2026-11-02): "Board: Regions" cannot
    // be unticked under it, and a board that has not said has regions because of it.
    [Fact]
    public void A_file_or_a_link_saying_a_region_counts_among_the_regional_rows()
    {
        var board = new BoardData
        {
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U1", Name = "a", File = "a.pdf", Region = "PAL" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "b", Url = "https://example.org", Region = "NTSC" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "c", Name = "d", File = "d.pdf", Region = "PAL" }],
            BoardLinks = [new BoardLinkEntry { Category = "c", Name = "e", Url = "https://example.org", Region = "NTSC" }],
        };

        Assert.Equal(4, BoardRegions.RegionalRows(board));
        Assert.True(BoardRegions.Has(board));
    }

    // ---- regions of its own (2026-11-09: "do the region steps 1-9, so we are prepared for other regions") ----------------

    // ###########################################################################################
    // A BOARD MAY COME IN REGIONS OF ITS OWN, in its order - and a board in PAL and NTSC, or none, reads as it always did:
    // a row's word never makes a region of a board that has not said (a typo would be a button).
    // ###########################################################################################
    [Fact]
    public void A_board_comes_in_the_regions_it_lists_in_its_order_and_rows_say_only_PAL_and_NTSC()
    {
        Assert.Equal(["EU", "US"], BoardRegions.Of(new BoardData { Regions = [" EU ", "US", "eu", ""] }));
        Assert.Equal(["PAL", "NTSC"], BoardRegions.Of(Board(Component("U1", "NTSC"))));
        Assert.Empty(BoardRegions.Of(Board(Component("U1", "Japan"))));
        Assert.Empty(BoardRegions.Of(null));

        BoardData own = new() { Regions = ["EU", "US"], Components = [Component("U1", "us")] };
        Assert.Equal("US", BoardRegions.SpellingOf(own, " us "));
        Assert.Null(BoardRegions.SpellingOf(own, "PAL"));

        // At most six read - a seventh, refused at Submit, had no place in the buttons' grids (code review, 2026-11-09) - but
        // kept as the list is written back, moved or with one deleted (second code review, 2026-11-09).
        BoardData seven = new() { Regions = ["A", "B", "C", "D", "E", "F", "G"] };
        Assert.Equal(["A", "B", "C", "D", "E", "F"], BoardRegions.Of(seven));
        Assert.Equal(["B", "A", "C", "D", "E", "F", "G"], BoardRegions.InOrder(seven, ["B", "A", "C", "D", "E", "F"]).Regions);
        Assert.Equal(["B", "C", "D", "E", "F", "G"], BoardRegions.WithoutRegion(seven, "A").Regions);
        Assert.Equal(1, BoardRegions.RegionalRows(own));
    }

    // ###########################################################################################
    // WHAT A REGION MAY BE CALLED - the one rule the pane, the table and the server hold a name to (Security: a submission's
    // "regions" reaches the publish as sent).
    // ###########################################################################################
    [Theory]
    [InlineData("Japan", null)]
    [InlineData("EU 230V", null)]
    [InlineData("PAL-M (Brazil)", null)]
    [InlineData("", "A region needs a name.")]
    [InlineData("A name far too long to be one", "A region's name is at most 20 characters.")]
    [InlineData("EU|US", "A region's name is letters, digits, spaces and \"- + / ( ) .\" only.")]
    [InlineData("<b>EU</b>", "A region's name is letters, digits, spaces and \"- + / ( ) .\" only.")]
    [InlineData("all regions", "\"All regions\" is what a row says for every region, so it cannot be a region's name.")]
    [InlineData("pal", "The board comes in [pal] already.")]
    public void A_regions_name_is_held_to_one_rule(string name, string? problem)
    {
        Assert.Equal(problem, BoardRegions.NameProblem(name, BoardRegions.PalAndNtsc));
    }

    [Fact]
    public void A_board_comes_in_at_most_six_regions()
    {
        Assert.Equal("A board comes in at most 6 regions.", BoardRegions.NameProblem("G", ["A", "B", "C", "D", "E", "F"]));
    }

    // ###########################################################################################
    // A region is added at the end, refused as the rule says, and deleted only while no row says it and more than two are
    // left. "This board comes with multiple regions" (owner request, 2026-11-09) ticked on a board with none gives PAL and
    // NTSC; unticked, every region goes - while no row says one.
    // ###########################################################################################
    [Fact]
    public void A_region_is_added_and_deleted_and_the_tick_takes_every_region_away()
    {
        BoardData pal = BoardRegions.WithRegions(Board(Component("U1")), true);
        BoardData withJapan = BoardRegions.WithRegion(pal, " Japan ");

        Assert.Equal(["PAL", "NTSC", "Japan"], BoardRegions.Of(withJapan));
        Assert.Same(withJapan, BoardRegions.WithRegion(withJapan, "japan"));
        Assert.Same(withJapan, BoardRegions.WithRegions(withJapan, true));

        Assert.Empty(BoardRegions.Of(BoardRegions.WithRegions(withJapan, false)));

        BoardData used = withJapan.WithComponentSections(components: [Component("U1", "JAPAN")]);
        Assert.Same(used, BoardRegions.WithoutRegion(used, "Japan"));
        Assert.Same(used, BoardRegions.WithRegions(used, false));
        Assert.False(BoardRegions.CanDelete(used, "Japan"));
        Assert.True(BoardRegions.CanDelete(used, "pal"));

        Assert.Equal(["PAL", "NTSC"], BoardRegions.Of(BoardRegions.WithoutRegion(withJapan, "japan")));
        Assert.Equal(["NTSC", "Japan"], BoardRegions.Of(BoardRegions.WithoutRegion(withJapan, "PAL")));

        // Two is the fewest: neither of the last two goes, nor one the board does not have.
        Assert.False(BoardRegions.CanDelete(pal, "PAL"));
        Assert.Same(pal, BoardRegions.WithoutRegion(pal, "NTSC"));
        Assert.False(BoardRegions.CanDelete(withJapan, "EU"));
        Assert.Same(withJapan, BoardRegions.WithoutRegion(withJapan, "EU"));
    }

    // A board whose rows say PAL or NTSC, with no list of its own, has them written down as it is ticked - and keeps them.
    [Fact]
    public void The_tick_writes_down_what_the_rows_say()
    {
        BoardData worked = Board(Component("U1", "PAL"));

        BoardData ticked = BoardRegions.WithRegions(worked, true);

        Assert.Equal(["PAL", "NTSC"], ticked.Regions);
        Assert.Same(worked, BoardRegions.WithRegions(worked, false));
    }

    // ###########################################################################################
    // THE ORDER IS THE BUTTONS' ORDER (owner request, 2026-11-09: "It should be possible to reorder the list of regions
    // up/down"): taken as the board spells its regions; the same board for its own order, or for anything that is not
    // exactly its regions.
    // ###########################################################################################
    [Fact]
    public void The_regions_are_put_in_another_order_only_as_a_whole()
    {
        BoardData board = Board(Component("U1")).WithRegions(["EU", "US", "Japan"]);

        Assert.Equal(["Japan", "EU", "US"], BoardRegions.InOrder(board, ["japan", "eu", "US"]).Regions);
        Assert.True(BoardRegions.Differ(board, BoardRegions.InOrder(board, ["US", "EU", "Japan"])));

        Assert.Same(board, BoardRegions.InOrder(board, ["EU", "US", "Japan"]));
        Assert.Same(board, BoardRegions.InOrder(board, ["EU", "US"]));
        Assert.Same(board, BoardRegions.InOrder(board, ["EU", "US", "Japan", "PAL"]));
        Assert.Same(board, BoardRegions.InOrder(board, ["EU", "EU", "Japan"]));
        Assert.Same(board, BoardRegions.InOrder(board, ["EU", "US", "China"]));
    }

    // A change of the list - another name, or the same in another order - is a change of the draft, said by name.
    [Fact]
    public void A_change_of_the_list_is_a_change_said_by_its_names()
    {
        BoardData before = BoardRegions.WithRegions(Board(Component("U1")), true);
        BoardData after = BoardRegions.WithRegion(before, "Japan");

        Assert.True(BoardRegions.Differ(before, after));
        Assert.True(BoardRegions.Differ(before, before.WithRegions(["NTSC", "PAL"])));
        Assert.False(BoardRegions.Differ(before, before.WithRegions(["PAL", "NTSC"])));

        Assert.Equal("This board now comes in PAL, NTSC and Japan versions", BoardRegions.ChangedLine(BoardRegions.Of(before), BoardRegions.Of(after)));

        // The same regions in another order are said as that, never as a new set (code review, 2026-11-09).
        Assert.Equal("This board's regions are now in the order NTSC, PAL", BoardRegions.ChangedLine(["PAL", "NTSC"], ["NTSC", "PAL"]));
        Assert.Equal("This board's regions are now in the order NTSC, PAL", BoardRegions.ChangedLine(true, ["PAL", "NTSC"], ["NTSC", "PAL"]));
        Assert.Equal("This board no longer comes in EU and US versions", BoardRegions.ChangedLine(["EU", "US"], []));
        Assert.Equal("This new board comes in EU and US versions", BoardRegions.NewBoardLine(["EU", "US"]));

        // A record from before names had their own - PAL and NTSC, all there was.
        Assert.Equal("This board now comes in PAL and NTSC versions", BoardRegions.ChangedLine(true));
        Assert.Equal("This board no longer comes in PAL and NTSC versions", BoardRegions.ChangedLine(false, null, null));
        Assert.Equal("This new board comes in PAL and NTSC versions", BoardRegions.NewBoardLine(null));
    }

    // The review and the board's history carry the names, before and after - "This board now comes in ..." by name.
    [Fact]
    public void The_review_and_the_history_carry_the_names()
    {
        BoardData published = BoardRegions.WithRegions(Board(Component("U1")), true);
        BoardData submitted = BoardRegions.WithRegion(published, "Japan");

        ReviewChangeSummary changed = ReviewSummary.Compare(published, submitted);
        SubmissionChanges facts = SubmissionChangeFacts.Build(changed);

        Assert.True(changed.RegionsNow);
        Assert.Equal(["PAL", "NTSC"], changed.RegionNamesBefore);
        Assert.Equal(["PAL", "NTSC", "Japan"], changed.RegionNamesNow);
        Assert.Equal(["PAL", "NTSC", "Japan"], facts.RegionNamesNow);

        Assert.Null(ReviewSummary.Compare(published, published).RegionNamesNow);
    }

    // ###########################################################################################
    // A REGION ADDED GIVES A COMPONENT TOLD APART IN EVERY REGION A ROW IN IT (code review, 2026-11-09: it had none, so it was
    // gone from the list, the popup and the Schematics tab there). A copy of the first region's row, right after its own; its
    // pins told apart per region the first region's pins. A component of one region stays out of it, as of the others.
    // ###########################################################################################
    [Fact]
    public void A_region_added_gives_a_component_told_apart_in_every_region_a_row_in_it()
    {
        var board = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components =
            [
                new ComponentEntry { BoardLabel = "U19", Category = "IC", FriendlyName = "VIC-II", TechnicalNameOrValue = "6569", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U19", Category = "IC", FriendlyName = "VIC-II", TechnicalNameOrValue = "6567", Region = "NTSC" },
                new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
                new ComponentEntry { BoardLabel = "Y1", Category = "Misc", FriendlyName = "PAL crystal", Region = "PAL" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U19", Pin = "1", Name = "D6", Region = "PAL" },
                new ComponentPinEntry { BoardLabel = "U19", Pin = "1", Name = "D6", Region = "NTSC" },
                new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", Region = "NTSC" },
                new ComponentPinEntry { BoardLabel = "U8", Pin = "2", Name = "RDY", Region = "NTSC" },
                new ComponentPinEntry { BoardLabel = "Y1", Pin = "1", Name = "X", Region = "PAL" },
            ],
        };

        BoardData after = BoardRegions.WithRegion(board, "Japan");

        Assert.Equal(["PAL", "NTSC", "Japan"], after.Regions);
        Assert.Equal(["U19|PAL", "U19|NTSC", "U19|Japan", "U8|", "Y1|PAL"], after.Components.Select(row => $"{row.BoardLabel}|{row.Region}"));
        Assert.Equal("6569", after.Components[2].TechnicalNameOrValue);

        // U8 is in every region with its pins told apart: the first region that has any - NTSC - lends them.
        Assert.Equal(
            ["U19|1|PAL", "U19|1|NTSC", "U19|1|Japan", "U8|1|NTSC", "U8|2|NTSC", "U8|1|Japan", "U8|2|Japan", "Y1|1|PAL"],
            after.ComponentPins.Select(pin => $"{pin.BoardLabel}|{pin.Pin}|{pin.Region}"));

        // Nothing told apart per region: nothing but the region added.
        BoardData plain = BoardRegions.WithRegions(Board(Component("U1")), true);
        BoardData plainAfter = BoardRegions.WithRegion(plain, "Japan");
        Assert.Same(plain.Components, plainAfter.Components);

        // ... and the region added can be deleted again, those rows going with it (code review, 2026-11-09: they held it on
        // the board for good). The rows of the component of one region would be left in none, so they hold theirs.
        Assert.Equal(0, BoardRegions.RowsHolding(after, "Japan"));
        Assert.True(BoardRegions.CanDelete(after, "Japan"));

        BoardData back = BoardRegions.WithoutRegion(after, "Japan");
        Assert.Equal(["PAL", "NTSC"], back.Regions);
        Assert.Equal(board.Components.Select(row => $"{row.BoardLabel}|{row.Region}"), back.Components.Select(row => $"{row.BoardLabel}|{row.Region}"));
        Assert.Equal(board.ComponentPins.Select(pin => $"{pin.BoardLabel}|{pin.Pin}|{pin.Region}"), back.ComponentPins.Select(pin => $"{pin.BoardLabel}|{pin.Pin}|{pin.Region}"));

        // Y1 and its pin are PAL's alone; U8's NTSC pins go with NTSC, U8 being in every region.
        Assert.Equal(2, BoardRegions.RowsHolding(after, "PAL"));
        Assert.Equal(0, BoardRegions.RowsHolding(after, "NTSC"));

        // Every region's facts in one go, as "Board: Regions" shows them.
        // ... with how many rows go with each, said under it (code review, 2026-11-09): U19's row and pin, U8's two pins.
        // Compared on those facts: a holding carries the components holding it too (2026-11-11), a list compared apart.
        List<BoardRegions.RegionHolding> holding = BoardRegions.HoldingOf(after);
        Assert.Equal(
            [new BoardRegions.RegionHolding("PAL", 2, false, 2), new BoardRegions.RegionHolding("NTSC", 0, true, 4), new BoardRegions.RegionHolding("Japan", 0, true, 4)],
            holding.Select(region => new BoardRegions.RegionHolding(region.Region, region.Rows, region.CanDelete, region.Going)));
        // ... and the components each is applied on, those told apart in every region among them (2026-11-11).
        Assert.Equal(["U19", "Y1"], holding[0].Components);
        Assert.Equal(["U19", "U8"], holding[1].Components);
        Assert.Equal(["U19", "U8"], holding[2].Components);
        BoardData japanLink = after.WithComponentSections(componentLinks: [new ComponentLinkEntry { BoardLabel = "U19", Name = "J", Url = "https://example.org", Region = "Japan" }]);
        Assert.Equal(1, BoardRegions.RowsHolding(japanLink, "Japan"));
        Assert.False(BoardRegions.CanDelete(japanLink, "Japan"));
        Assert.Same(japanLink, BoardRegions.WithoutRegion(japanLink, "Japan"));
    }

    // ###########################################################################################
    // A ROW SAYING PAL OR NTSC KEEPS THE TICK ON A BOARD OF REGIONS OF ITS OWN TOO (code review, 2026-11-09): a board that lists
    // none is worked out from its rows, so with the list taken away that row brought PAL and NTSC back.
    // ###########################################################################################
    [Fact]
    public void A_row_saying_PAL_keeps_the_tick_on_a_board_of_regions_of_its_own()
    {
        BoardData board = Board(Component("U1", "PAL")).WithRegions(["EU", "US"]);

        Assert.Equal(1, BoardRegions.RegionalRows(board));
        Assert.Same(board, BoardRegions.WithRegions(board, false));

        BoardData clean = Board(Component("U1")).WithRegions(["EU", "US"]);
        Assert.Empty(BoardRegions.Of(BoardRegions.WithRegions(clean, false)));
    }

    // A pinout or a component told apart per region is told apart in every region the board comes in.
    [Fact]
    public void A_component_told_apart_per_region_gets_a_row_in_every_region_of_the_board()
    {
        var board = new BoardData { Regions = ["EU", "US", "Japan"], Components = [Component("U1")] };

        BoardData regional = ComponentPaneEdits.WithRegionalRows(board, "U1", "EU", true);

        Assert.Equal(["EU", "US", "Japan"], regional.Components.Select(row => row.Region));
    }

    // ###########################################################################################
    // THE CHECKS (Security: a submission's "regions" reaches the publish as sent): a name the rule refuses is an error on
    // both sides, and a file's or a link's region that is none of the board's a warning on its cell.
    // ###########################################################################################
    [Fact]
    public void The_checks_take_any_name_of_the_rule_and_refuse_the_rest()
    {
        BoardCheckRows good = BoardCheckRows.From(new SubmissionRows { Regions = ["EU", "US"] });
        BoardCheckRows bad = BoardCheckRows.From(new SubmissionRows { Regions = ["EU", "eu", "EU|US"] });

        Assert.DoesNotContain(BoardDataChecks.Check(good, files: null, BoardCheckScope.Submission), problem => problem.Code == "regions.unknown");
        Assert.Equal(2, BoardDataChecks.Check(bad, files: null, BoardCheckScope.Submission).Count(problem => problem.Code == "regions.unknown"));

        var rows = new BoardData
        {
            Regions = ["EU", "US"],
            BoardLinks =
            [
                new BoardLinkEntry { Category = "c", Name = "a", Url = "https://example.org", Region = "us" },
                new BoardLinkEntry { Category = "c", Name = "b", Url = "https://example.org", Region = "PAL" },
            ],
        };

        BoardDataProblem warned = Assert.Single(
            BoardDataChecks.Check(BoardCheckRows.From(rows), files: null, BoardCheckScope.Everything),
            problem => problem.Code == "row.region");
        Assert.Contains("Write EU or US", warned.Message, StringComparison.Ordinal);
    }
}
