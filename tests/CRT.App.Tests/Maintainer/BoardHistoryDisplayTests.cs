using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// BoardHistoryDisplay - a board's History view (owner request, 2026-10-04: "it should show the
// full history of what has happened with this board, in an 'easy to overview' way... which is not
// what the current history is"). What it was not: one line per EVENT, a submission's sending and its
// decision lines apart with other things between them, and nothing about what it changed. So these
// pin the grouping (one card per submission, other events as lines, by month, newest first), the
// card's story in order, and the words for what a submission changed.
// ###########################################################################################
public sealed class BoardHistoryDisplayTests
{
    // Mid-month at noon, so a month or a day is the same in every time zone the suite runs in.
    private static readonly DateTimeOffset September = new(2026, 9, 15, 12, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset October = new(2026, 10, 15, 12, 0, 0, TimeSpan.Zero);

    private static BoardOverviewEntry Board() =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1);

    private static BoardSubmissionEntry Submission(
        long id,
        string state,
        DateTimeOffset created,
        DateTimeOffset? decided = null,
        string? comment = null,
        SubmissionChanges? changes = null,
        DateTimeOffset? discarded = null) =>
        new(id, "hest@mailscan.dk", $"Fix {id}.", state, created, decided, comment, discarded, changes);

    private static BoardDetailAnswer Detail(IReadOnlyList<BoardSubmissionEntry> submissions, IReadOnlyList<BoardHistoryEntry> history) =>
        new(BoardHistoryDisplayTests.Board(), [], [], submissions, null, history);

    // ###########################################################################################
    // *** ONE CARD PER SUBMISSION. *** Its sending and its decision are steps ON the card, not
    // lines of their own; anything else - here a publish to stable - stays a line.
    // ###########################################################################################
    [Fact]
    public void A_submissions_events_are_one_card_and_other_events_are_lines()
    {
        IReadOnlyList<HistoryMonth> months = BoardHistoryDisplay.Build(BoardHistoryDisplayTests.Detail(
            [BoardHistoryDisplayTests.Submission(9, "merged", September, September.AddDays(1))],
            [
                new BoardHistoryEntry(September.AddDays(2), BoardHistoryEvents.PublishedToProduction, "Dennis", null, "revision 2026-September-17"),
                new BoardHistoryEntry(September.AddDays(1), BoardHistoryEvents.Decided, "Anna", 9, "merged"),
                new BoardHistoryEntry(September, BoardHistoryEvents.Sent, "Bo", 9, "Fix 9."),
            ]));

        HistoryMonth month = Assert.Single(months);
        Assert.Equal("September 2026", month.Heading);

        Assert.Collection(
            month.Items,
            item => Assert.Equal("Published to the stable source", Assert.IsType<HistoryEventLine>(item).What),
            item =>
            {
                HistorySubmissionCard card = Assert.IsType<HistorySubmissionCard>(item);

                Assert.Equal("#9 - Fix 9.", card.Title);
                Assert.Equal("Published to the BETA source", card.State);
                Assert.Equal(
                    ["2026-September-15 - sent by Bo", "2026-September-16 - Published to the BETA source by Anna"],
                    card.Steps);
            });
    }

    // ###########################################################################################
    // NEWEST FIRST, BY MONTH - and a card sits at its LATEST event: sent in September, published in
    // October, it is October's.
    // ###########################################################################################
    [Fact]
    public void Months_are_newest_first_and_a_card_sits_at_its_latest_event()
    {
        IReadOnlyList<HistoryMonth> months = BoardHistoryDisplay.Build(BoardHistoryDisplayTests.Detail(
            [
                BoardHistoryDisplayTests.Submission(1, "rejected", September.AddDays(-3), September.AddDays(-2)),
                BoardHistoryDisplayTests.Submission(2, "merged", September, October),
            ],
            [new BoardHistoryEntry(September.AddDays(1), BoardHistoryEvents.MaintainerAdded, "Dennis", null, "anna@example.com")]));

        Assert.Equal(["October 2026", "September 2026"], months.Select(month => month.Heading));
        Assert.Equal(2, Assert.IsType<HistorySubmissionCard>(Assert.Single(months[0].Items)).Id);
        Assert.Equal(
            ["Maintainer added: anna@example.com", "#1 - Fix 1."],
            months[1].Items.Select(item => item is HistorySubmissionCard card ? card.Title : ((HistoryEventLine)item).What));

        Assert.Equal("15 Oct", BoardHistoryDisplay.DayOf(October));
    }

    // An event naming a submission the server did not list is never dropped - it stays a line.
    [Fact]
    public void An_event_of_a_submission_not_listed_stays_a_line()
    {
        IReadOnlyList<HistoryMonth> months = BoardHistoryDisplay.Build(BoardHistoryDisplayTests.Detail(
            [],
            [new BoardHistoryEntry(September, BoardHistoryEvents.Decided, "Anna", 4, "rejected", "Wrong board.")]));

        HistoryEventLine line = Assert.IsType<HistoryEventLine>(Assert.Single(Assert.Single(months).Items));
        Assert.Equal("#4 - Not accepted", line.What);
        Assert.Equal("Told the contributor: Wrong board.", line.Note);
    }

    // ###########################################################################################
    // The card's story: changed by a maintainer between sending and deciding, what the contributor
    // was told, and the red mark of a draft the contributor threw away.
    // ###########################################################################################
    [Fact]
    public void A_card_tells_an_amendment_what_was_told_and_a_discarded_draft()
    {
        HistorySubmissionCard card = BoardHistoryDisplay.Card(
            BoardHistoryDisplayTests.Submission(5, "rejected", September, September.AddDays(2), " Wrong revision. ", discarded: September.AddDays(3)),
            [
                new BoardHistoryEntry(September, BoardHistoryEvents.Sent, null, 5, "Fix 5."),
                new BoardHistoryEntry(September.AddDays(1), BoardHistoryEvents.Amended, "Anna", 5, "amendment 2, 1 file(s)"),
                new BoardHistoryEntry(September.AddDays(2), BoardHistoryEvents.Decided, null, 5, "rejected"),
            ]);

        Assert.Equal(
            [
                "2026-September-15 - sent by hest@mailscan.dk",
                "2026-September-16 - changed by Anna (amendment 2, 1 file(s))",
                "2026-September-17 - Not accepted"
            ],
            card.Steps);
        Assert.Equal("Told the contributor: Wrong revision.", card.Told);
        Assert.Equal(DraftDiscardWording.Mark(September.AddDays(3)), card.DraftDiscarded);
        Assert.Equal(September.AddDays(3), card.AtUtc);
        Assert.Null(card.NoSummary);
        Assert.Empty(card.Changes);
    }

    // ###########################################################################################
    // WHAT IT CHANGED: a heading per sheet, a line per kind, the counts bold ("[2]"), the names the
    // server recorded - "and N more" past them - and the files by name.
    // ###########################################################################################
    [Fact]
    public void What_a_submission_changed_reads_as_a_sheet_then_a_line_per_kind()
    {
        var changes = new SubmissionChanges(
            false,
            [
                new SectionChanges(
                    "Components", 2, 1, 1, 1,
                    ["U7", "U9"],
                    [new ChangedRowFact("U8", ["Part-number", "Description"])],
                    ["R3"],
                    [new RenamedRowFact("C10", "C51")]),
                new SectionChanges("Component images", 25, 0, 0, 0, ["U7 / 1 / Pin 1"], [], [], [])
            ],
            new FileChanges(1, 1, 0, ["Commodore/C64/250407/Images/U7.png"], ["Commodore/C64/250407/Images/U8.png"], []));

        IReadOnlyList<HistoryChangeLine> lines = BoardHistoryDisplay.ChangeLines(changes);

        Assert.Equal(
            [
                "Components",
                "[2] added: U7, U9",
                "[1] changed: U8 (Part-number, Description)",
                "[1] removed: R3",
                "[1] renamed: C10 to C51",
                "Component images",
                "[25] added: U7 / 1 / Pin 1 and 24 more",
                "Files",
                "[1] added: U7.png",
                "[1] replaced: U8.png"
            ],
            lines.Select(line => line.Text));

        Assert.Equal([true, false, false, false, false, true, false, true, false, false], lines.Select(line => line.IsHeading));

        // Only the number is bold.
        Assert.Equal(["2"], lines[1].Runs.Where(run => run.IsCount).Select(run => run.Text));
    }

    // ###########################################################################################
    // The lists a submission put in another order (2026-10-30 - order is a change): a heading "Order"
    // and a line per list, after the sheets and before the files. A submission that only moved rows
    // says that and nothing else.
    // ###########################################################################################
    [Fact]
    public void The_lists_put_in_another_order_read_under_their_own_heading()
    {
        var changes = new SubmissionChanges(
            false,
            [],
            FileChanges.None,
            [BoardWorkbookSchema.SheetBoardLocalFiles, BoardWorkbookSchema.SheetCredits]);

        IReadOnlyList<HistoryChangeLine> lines = BoardHistoryDisplay.ChangeLines(changes);

        Assert.Equal(["Order", "Order of the files changed", "Order of the credits changed"], lines.Select(line => line.Text));
        Assert.Equal([true, false, false], lines.Select(line => line.IsHeading));
    }

    // A new board is counted, not named - its first submission adds every row it has.
    [Fact]
    public void A_change_of_the_boards_regions_reads_under_its_own_heading()
    {
        var changes = new SubmissionChanges(false, [], FileChanges.None, null, RegionsNow: false);

        Assert.True(SubmissionChangeFacts.HasChanges(changes));
        Assert.Equal(["Regions", BoardRegions.ChangedLine(false)], BoardHistoryDisplay.ChangeLines(changes).Select(line => line.Text));
    }

    // ###########################################################################################
    // One rectangle moved and another resized (code review, 2026-11-02): from the review's own comparison to
    // the line, each with its own field. A record stored before - both fields on both rows - reads as stored.
    // ###########################################################################################
    [Fact]
    public void One_rectangle_moved_and_another_resized_read_each_with_its_own_field()
    {
        static ComponentHighlightEntry At(string x, string width) =>
            new() { SchematicName = "Top", BoardLabel = "U8", X = x, Y = "2", Width = width, Height = "4" };

        SubmissionChanges changes = SubmissionChangeFacts.Build(ReviewSummary.Compare(
            new BoardData { ComponentHighlights = [At("1", "3"), At("10", "3")] },
            new BoardData { ComponentHighlights = [At("5", "3"), At("10", "9")] }));

        string changed = BoardHistoryDisplay.ChangeLines(changes).Select(line => line.Text).Single(text => text.Contains("changed", StringComparison.Ordinal));

        Assert.Contains("Top / U8 (X)", changed, StringComparison.Ordinal);
        Assert.Contains("Top / U8 (Width)", changed, StringComparison.Ordinal);
    }

    [Fact]
    public void A_record_stored_before_each_rectangle_had_its_own_fields_still_reads()
    {
        SubmissionChanges stored = System.Text.Json.JsonSerializer.Deserialize<SubmissionChanges>(
            """
            {"IsNewBoard":false,"Sections":[{"Section":"Component highlights","AddedCount":0,"ChangedCount":2,"RemovedCount":0,"RenamedCount":0,
            "Added":[],"Changed":[{"Row":"Top / U8","Fields":["X","Width"]},{"Row":"Top / U8","Fields":["X","Width"]}],"Removed":[],"Renamed":[]}],
            "Files":{"AddedCount":0,"ReplacedCount":0,"RemovedCount":0,"Added":[],"Replaced":[],"Removed":[]}}
            """)!;

        Assert.Equal(
            ["Component highlights", "[2] changed: Top / U8 (X, Width), Top / U8 (X, Width)"],
            BoardHistoryDisplay.ChangeLines(stored).Select(line => line.Text));
    }

    // A new board that has regions: said as what it is - nothing "now" about a board that was not there.
    [Fact]
    public void A_new_boards_regions_read_under_their_own_heading()
    {
        var changes = new SubmissionChanges(true, [], FileChanges.None, null, RegionsNow: true);

        Assert.Equal(
            ["A new board.", "Regions", "This new board comes in PAL and NTSC versions"],
            BoardHistoryDisplay.ChangeLines(changes).Select(line => line.Text));
    }

    [Fact]
    public void A_new_board_is_counted_not_named()
    {
        var changes = new SubmissionChanges(
            true,
            [new SectionChanges("Components", 140, 0, 0, 0, ["C1", "C2"], [], [], [])],
            FileChanges.None);

        Assert.Equal(["A new board.", "Components", "[140] added"], BoardHistoryDisplay.ChangeLines(changes).Select(line => line.Text));
    }

    // ###########################################################################################
    // A card in BETA with no summary says why - published before CRT kept one - and a card never
    // published says nothing about changes at all.
    // ###########################################################################################
    [Theory]
    [InlineData("merged", true)]
    [InlineData("published", true)]
    [InlineData("pending", false)]
    [InlineData("rejected", false)]
    public void Only_a_published_card_without_a_summary_says_none_was_recorded(string state, bool says)
    {
        HistorySubmissionCard card = BoardHistoryDisplay.Card(BoardHistoryDisplayTests.Submission(3, state, September), []);

        Assert.Equal(says, card.NoSummary is not null);
        Assert.Empty(card.Changes);
    }

    [Fact]
    public void A_card_with_a_summary_shows_it_and_no_note()
    {
        var changes = new SubmissionChanges(false, [new SectionChanges("Credits", 1, 0, 0, 0, ["Code / Anna"], [], [], [])], FileChanges.None);

        HistorySubmissionCard card = BoardHistoryDisplay.Card(BoardHistoryDisplayTests.Submission(3, "merged", September, changes: changes), []);

        Assert.Null(card.NoSummary);
        Assert.Equal(["Credits", "[1] added: Code / Anna"], card.Changes.Select(line => line.Text));
    }

    [Theory]
    [InlineData(0, 0, "Nothing has happened to this board yet")]
    [InlineData(1, 0, "History - 1 submission, newest first")]
    [InlineData(3, 1, "History - 3 submissions and 1 other event, newest first")]
    [InlineData(0, 2, "History - 2 other events, newest first")]
    public void The_heading_counts_what_is_there(int submissions, int events, string heading)
    {
        Assert.Equal(heading, BoardHistoryDisplay.Heading(submissions, events));
    }

    // No user-visible text calls a contributor "their", "them" or "they" (owner request, 2026-10-01).
    [Fact]
    public void Nothing_here_calls_the_contributor_they()
    {
        HistorySubmissionCard card = BoardHistoryDisplay.Card(
            BoardHistoryDisplayTests.Submission(3, "merged", September, discarded: September),
            []);

        foreach (string text in new[] { card.NoSummary!, card.DraftDiscarded!, card.Title, card.State }.Concat(card.Steps))
            Assert.DoesNotMatch(@"\b(their|them|they)\b", text);
    }
}
