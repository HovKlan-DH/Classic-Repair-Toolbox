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
}
