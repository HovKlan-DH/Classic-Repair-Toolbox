using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// WHERE A BOARD IS, AS THREE STAGES (owner request, 2026-10-04: "where does this sit now, as I do
// not think it is in BETA nor stable?") - BoardStagesDisplay. Every stage says something: what is
// there, "Not there", or "Not known" - a blank would read as nothing to report.
// ###########################################################################################
public sealed class BoardStagesDisplayTests
{
    private static readonly DateTimeOffset Sent = new(2026, 9, 20, 10, 0, 0, TimeSpan.Zero);

    private static BoardOverviewEntry Board(
        bool inBeta = true,
        bool? inStable = true,
        bool awaiting = false,
        string? betaRevision = "2026-October-4",
        string? stableRevision = "2026-September-25",
        DateTimeOffset? published = null) =>
        new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", inBeta, inStable, awaiting, true,
            betaRevision, stableRevision, published, 0);

    private static BoardSubmissionEntry Submission(long id, string state, DateTimeOffset created, DateTimeOffset? decided = null) =>
        new(id, "c@example.com", "x", state, created, decided, null);

    // The reported board: its only submission turned down, nothing in either tree.
    [Fact]
    public void A_new_board_turned_down_says_so_and_that_neither_tree_holds_it()
    {
        IReadOnlyList<BoardStage> stages = BoardStagesDisplay.For(
            BoardStagesDisplayTests.Board(inBeta: false, inStable: false, betaRevision: null, stableRevision: null),
            [BoardStagesDisplayTests.Submission(5, "rejected", BoardStagesDisplayTests.Sent, BoardStagesDisplayTests.Sent.AddDays(1))]);

        Assert.Equal(["Submitted", "BETA", "Stable"], stages.Select(stage => stage.Label));
        Assert.Equal("#5 Not accepted", stages[0].Value);
        Assert.Equal($"decided {SubmissionReceiptPresenter.FormatDate(BoardStagesDisplayTests.Sent.AddDays(1))}", stages[0].Detail);
        Assert.Equal(BoardStageState.Reached, stages[0].State);
        Assert.Equal(("Not there", BoardStageState.NotThere), (stages[1].Value, stages[1].State));
        Assert.Equal(("Not there", BoardStageState.NotThere), (stages[2].Value, stages[2].State));
    }

    // The NEWEST submission - by when it was sent, not by its place in the list - and how many there are.
    [Fact]
    public void Submitted_is_the_newest_submission_with_when_it_was_sent_and_how_many_there_are()
    {
        BoardStage stage = BoardStagesDisplay.Submitted(
        [
            BoardStagesDisplayTests.Submission(7, "pending", BoardStagesDisplayTests.Sent.AddDays(3)),
            BoardStagesDisplayTests.Submission(4, "merged", BoardStagesDisplayTests.Sent, BoardStagesDisplayTests.Sent.AddHours(2))
        ]);

        Assert.Equal($"#7 {SubmissionReceiptPresenter.PendingWording}", stage.Value);
        Assert.Equal($"sent {SubmissionReceiptPresenter.FormatDate(BoardStagesDisplayTests.Sent.AddDays(3))} - the latest of 2", stage.Detail);
    }

    [Fact]
    public void No_submissions_is_said_and_unread_ones_are_not_guessed()
    {
        Assert.Equal(("No submissions", BoardStageState.NotThere), (BoardStagesDisplay.Submitted([]).Value, BoardStagesDisplay.Submitted([]).State));

        BoardStage unread = BoardStagesDisplay.Submitted(null);
        Assert.Equal(string.Empty, unread.Value);
        Assert.Equal(BoardStageState.NotKnown, unread.State);
    }

    // BETA ahead of stable says where it waits, by the tab's own name for that queue.
    [Fact]
    public void BETA_names_its_revision_and_whether_it_waits_to_go_to_stable()
    {
        BoardStage waiting = BoardStagesDisplay.Beta(BoardStagesDisplayTests.Board(awaiting: true));
        BoardStage level = BoardStagesDisplay.Beta(BoardStagesDisplayTests.Board(betaRevision: null));

        Assert.Equal("Revision 2026-October-4", waiting.Value);
        Assert.Equal($"ahead of stable - waiting under {MaintainerScreenWording.BetaQueueQuoted}", waiting.Detail);
        Assert.Equal("In BETA", level.Value);
        Assert.Null(level.Detail);
    }

    [Fact]
    public void Stable_names_its_revision_and_when_it_was_published_or_why_it_cannot_say()
    {
        DateTimeOffset published = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        BoardStage there = BoardStagesDisplay.Stable(BoardStagesDisplayTests.Board(published: published));
        BoardStage unknown = BoardStagesDisplay.Stable(BoardStagesDisplayTests.Board(inStable: null));

        Assert.Equal("Revision 2026-September-25", there.Value);
        Assert.Equal($"published {SubmissionReceiptPresenter.FormatDate(published)}", there.Detail);
        Assert.Equal(("Not known", BoardStageState.NotKnown), (unknown.Value, unknown.State));
        Assert.Equal("this server has no stable source to look in", unknown.Detail);
    }

    // ###########################################################################################
    // *** WHERE THE NEWEST WORK IS, IN WORDS, AND WHICH STAGE IS OUTLINED (owner request, 2026-10-09:
    // the three columns were not clear enough). *** The owner's own case first: #23 waiting for
    // review while BETA waits to go to stable - BETA is outlined, since it has to move before #23
    // can be approved, and the sentence says both.
    // ###########################################################################################
    [Fact]
    public void A_submission_waiting_behind_BETA_outlines_BETA_and_says_both()
    {
        BoardStageNow now = BoardStagesDisplay.Now(
            Board(awaiting: true),
            [Submission(23, "pending", Sent)])!;

        Assert.Equal(1, now.Stage);
        Assert.Equal(
            $"BETA is ahead of the stable source and waits under {MaintainerScreenWording.BetaQueueQuoted} to be published to stable - or pushed back. " +
            "#23 waits for review too, and can be approved once that is done.",
            now.Sentence);
    }

    [Theory]
    [InlineData("pending", true, true, false, 0, "#7 waits for review under")]
    [InlineData("changes_requested", true, true, false, 2, "BETA and the stable source hold the same - nothing is waiting. #7 is back with the contributor for changes.")]
    [InlineData("merged", true, true, true, 1, "BETA is ahead of the stable source")]
    [InlineData("published", true, true, false, 2, "BETA and the stable source hold the same - nothing is waiting.")]
    [InlineData("merged", true, false, false, 1, "In BETA only - not published to the stable source yet.")]
    [InlineData("rejected", false, false, false, 0, "Nothing of this board is published yet.")]
    public void The_stage_outlined_and_the_sentence_follow_where_the_newest_work_is(
        string latestState, bool inBeta, bool inStable, bool awaiting, int stage, string sentence)
    {
        BoardStageNow now = BoardStagesDisplay.Now(
            Board(inBeta: inBeta, inStable: inStable, awaiting: awaiting),
            [Submission(3, "published", Sent.AddDays(-9)), Submission(7, latestState, Sent)])!;

        Assert.Equal(stage, now.Stage);
        Assert.StartsWith(sentence, now.Sentence, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** EVERY SUBMISSION COUNTS, NOT ONLY THE NEWEST (code review, 2026-10-09). *** #5 waits for
    // review behind a newer #6 that was rejected: it still waits, and Submitted is outlined - it
    // was said as "nothing is waiting", with the Stable card outlined.
    // ###########################################################################################
    [Fact]
    public void A_submission_waiting_behind_a_newer_rejected_one_still_waits()
    {
        IReadOnlyList<BoardSubmissionEntry> sent =
        [
            Submission(5, "pending", Sent.AddDays(-3)),
            Submission(6, "rejected", Sent, decided: Sent.AddHours(2))
        ];

        BoardStageNow now = BoardStagesDisplay.Now(Board(), sent)!;
        Assert.Equal(0, now.Stage);
        Assert.Equal($"#5 waits for review under {MaintainerScreenWording.ContributorQueueQuoted}.", now.Sentence);

        BoardStageNow behindBeta = BoardStagesDisplay.Now(Board(awaiting: true), sent)!;
        Assert.Equal(1, behindBeta.Stage);
        Assert.EndsWith("#5 waits for review too, and can be approved once that is done.", behindBeta.Sentence, StringComparison.Ordinal);
    }

    [Fact]
    public void Several_submissions_waiting_are_named_oldest_first()
    {
        BoardStageNow now = BoardStagesDisplay.Now(
            Board(),
            [Submission(9, "returned", Sent), Submission(4, "pending", Sent.AddDays(-2)), Submission(7, "approved", Sent.AddDays(-1))])!;

        Assert.Equal($"#4, #7 and #9 wait for review under {MaintainerScreenWording.ContributorQueueQuoted}.", now.Sentence);
    }

    // Before the board's submissions are read, nothing is said - the newest work may be one of them.
    [Fact]
    public void Nothing_is_said_before_the_submissions_are_read()
    {
        Assert.Null(BoardStagesDisplay.Now(Board(awaiting: true), null));
    }

    [Fact]
    public void Each_stage_is_numbered_in_the_order_the_work_goes_through_them()
    {
        Assert.Equal(
            ["1  Submitted", "2  BETA", "3  Stable"],
            BoardStagesDisplay.For(Board(), []).Select((stage, index) => BoardStagesDisplay.NumberedLabel(index, stage.Label)));
    }
}
