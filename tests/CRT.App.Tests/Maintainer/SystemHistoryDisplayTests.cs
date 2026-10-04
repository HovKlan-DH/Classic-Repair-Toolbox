using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// SystemHistoryDisplay - a system's History view (owner request, 2026-10-04: "it should show the
// full history of what has happened with this board, in an 'easy to overview' way... which is not
// what the current history is"). What it was not: one line per EVENT, a submission's sending and its
// decision lines apart with other things between them, and nothing about what it changed. So these
// pin the grouping (one card per submission, other events as lines, by month, newest first), the
// card's story in order, and the words for what a submission changed.
// ###########################################################################################
public sealed class SystemHistoryDisplayTests
{
    // Mid-month at noon, so a month or a day is the same in every time zone the suite runs in.
    private static readonly DateTimeOffset September = new(2026, 9, 15, 12, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset October = new(2026, 10, 15, 12, 0, 0, TimeSpan.Zero);

    private static SystemOverviewEntry System() =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1);

    private static SystemSubmissionEntry Submission(
        long id,
        string state,
        DateTimeOffset created,
        DateTimeOffset? decided = null,
        string? comment = null,
        SubmissionChanges? changes = null,
        DateTimeOffset? discarded = null) =>
        new(id, "hest@mailscan.dk", $"Fix {id}.", state, created, decided, comment, discarded, changes);

    private static SystemDetailAnswer Detail(IReadOnlyList<SystemSubmissionEntry> submissions, IReadOnlyList<SystemHistoryEntry> history) =>
        new(SystemHistoryDisplayTests.System(), [], [], submissions, null, history);

    // ###########################################################################################
    // *** ONE CARD PER SUBMISSION. *** Its sending and its decision are steps ON the card, not
    // lines of their own; anything else - here a publish to stable - stays a line.
    // ###########################################################################################
    [Fact]
    public void A_submissions_events_are_one_card_and_other_events_are_lines()
    {
        IReadOnlyList<HistoryMonth> months = SystemHistoryDisplay.Build(SystemHistoryDisplayTests.Detail(
            [SystemHistoryDisplayTests.Submission(9, "merged", September, September.AddDays(1))],
            [
                new SystemHistoryEntry(September.AddDays(2), SystemHistoryEvents.PublishedToProduction, "Dennis", null, "revision 2026-September-17"),
                new SystemHistoryEntry(September.AddDays(1), SystemHistoryEvents.Decided, "Anna", 9, "merged"),
                new SystemHistoryEntry(September, SystemHistoryEvents.Sent, "Bo", 9, "Fix 9."),
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
        IReadOnlyList<HistoryMonth> months = SystemHistoryDisplay.Build(SystemHistoryDisplayTests.Detail(
            [
                SystemHistoryDisplayTests.Submission(1, "rejected", September.AddDays(-3), September.AddDays(-2)),
                SystemHistoryDisplayTests.Submission(2, "merged", September, October),
            ],
            [new SystemHistoryEntry(September.AddDays(1), SystemHistoryEvents.MaintainerAdded, "Dennis", null, "anna@example.com")]));

        Assert.Equal(["October 2026", "September 2026"], months.Select(month => month.Heading));
        Assert.Equal(2, Assert.IsType<HistorySubmissionCard>(Assert.Single(months[0].Items)).Id);
        Assert.Equal(
            ["anna@example.com made a maintainer", "#1 - Fix 1."],
            months[1].Items.Select(item => item is HistorySubmissionCard card ? card.Title : ((HistoryEventLine)item).What));

        Assert.Equal("15 Oct", SystemHistoryDisplay.DayOf(October));
    }

    // An event naming a submission the server did not list is never dropped - it stays a line.
    [Fact]
    public void An_event_of_a_submission_not_listed_stays_a_line()
    {
        IReadOnlyList<HistoryMonth> months = SystemHistoryDisplay.Build(SystemHistoryDisplayTests.Detail(
            [],
            [new SystemHistoryEntry(September, SystemHistoryEvents.Decided, "Anna", 4, "rejected", "Wrong board.")]));

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
        HistorySubmissionCard card = SystemHistoryDisplay.Card(
            SystemHistoryDisplayTests.Submission(5, "rejected", September, September.AddDays(2), " Wrong revision. ", discarded: September.AddDays(3)),
            [
                new SystemHistoryEntry(September, SystemHistoryEvents.Sent, null, 5, "Fix 5."),
                new SystemHistoryEntry(September.AddDays(1), SystemHistoryEvents.Amended, "Anna", 5, "amendment 2, 1 file(s)"),
                new SystemHistoryEntry(September.AddDays(2), SystemHistoryEvents.Decided, null, 5, "rejected"),
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

        IReadOnlyList<HistoryChangeLine> lines = SystemHistoryDisplay.ChangeLines(changes);

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

    // A new system is counted, not named - its first submission adds every row it has.
    [Fact]
    public void A_new_system_is_counted_not_named()
    {
        var changes = new SubmissionChanges(
            true,
            [new SectionChanges("Components", 140, 0, 0, 0, ["C1", "C2"], [], [], [])],
            FileChanges.None);

        Assert.Equal(["A new system.", "Components", "[140] added"], SystemHistoryDisplay.ChangeLines(changes).Select(line => line.Text));
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
        HistorySubmissionCard card = SystemHistoryDisplay.Card(SystemHistoryDisplayTests.Submission(3, state, September), []);

        Assert.Equal(says, card.NoSummary is not null);
        Assert.Empty(card.Changes);
    }

    [Fact]
    public void A_card_with_a_summary_shows_it_and_no_note()
    {
        var changes = new SubmissionChanges(false, [new SectionChanges("Credits", 1, 0, 0, 0, ["Code / Anna"], [], [], [])], FileChanges.None);

        HistorySubmissionCard card = SystemHistoryDisplay.Card(SystemHistoryDisplayTests.Submission(3, "merged", September, changes: changes), []);

        Assert.Null(card.NoSummary);
        Assert.Equal(["Credits", "[1] added: Code / Anna"], card.Changes.Select(line => line.Text));
    }

    [Theory]
    [InlineData(0, 0, "Nothing has happened to this system yet")]
    [InlineData(1, 0, "History - 1 submission, newest first")]
    [InlineData(3, 1, "History - 3 submissions and 1 other event, newest first")]
    [InlineData(0, 2, "History - 2 other events, newest first")]
    public void The_heading_counts_what_is_there(int submissions, int events, string heading)
    {
        Assert.Equal(heading, SystemHistoryDisplay.Heading(submissions, events));
    }

    // No user-visible text calls a contributor "their", "them" or "they" (owner request, 2026-10-01).
    [Fact]
    public void Nothing_here_calls_the_contributor_they()
    {
        HistorySubmissionCard card = SystemHistoryDisplay.Card(
            SystemHistoryDisplayTests.Submission(3, "merged", September, discarded: September),
            []);

        foreach (string text in new[] { card.NoSummary!, card.DraftDiscarded!, card.Title, card.State }.Concat(card.Steps))
            Assert.DoesNotMatch(@"\b(their|them|they)\b", text);
    }
}
