using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// WHERE A SYSTEM IS, AS THREE STAGES (owner request, 2026-10-04: "where does this sit now, as I do
// not think it is in BETA nor stable?") - SystemStagesDisplay. Every stage says something: what is
// there, "Not there", or "Not known" - a blank would read as nothing to report.
// ###########################################################################################
public sealed class SystemStagesDisplayTests
{
    private static readonly DateTimeOffset Sent = new(2026, 9, 20, 10, 0, 0, TimeSpan.Zero);

    private static SystemOverviewEntry System(
        bool inBeta = true,
        bool? inStable = true,
        bool awaiting = false,
        string? betaRevision = "2026-October-4",
        string? stableRevision = "2026-September-25",
        DateTimeOffset? published = null) =>
        new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", inBeta, inStable, awaiting, true,
            betaRevision, stableRevision, published, 0);

    private static SystemSubmissionEntry Submission(long id, string state, DateTimeOffset created, DateTimeOffset? decided = null) =>
        new(id, "c@example.com", "x", state, created, decided, null);

    // The reported system: its only submission turned down, nothing in either tree.
    [Fact]
    public void A_new_system_turned_down_says_so_and_that_neither_tree_holds_it()
    {
        IReadOnlyList<SystemStage> stages = SystemStagesDisplay.For(
            SystemStagesDisplayTests.System(inBeta: false, inStable: false, betaRevision: null, stableRevision: null),
            [SystemStagesDisplayTests.Submission(5, "rejected", SystemStagesDisplayTests.Sent, SystemStagesDisplayTests.Sent.AddDays(1))]);

        Assert.Equal(["Submitted", "BETA", "Stable"], stages.Select(stage => stage.Label));
        Assert.Equal("#5 Not accepted", stages[0].Value);
        Assert.Equal($"decided {SubmissionReceiptPresenter.FormatDate(SystemStagesDisplayTests.Sent.AddDays(1))}", stages[0].Detail);
        Assert.Equal(SystemStageState.Reached, stages[0].State);
        Assert.Equal(("Not there", SystemStageState.NotThere), (stages[1].Value, stages[1].State));
        Assert.Equal(("Not there", SystemStageState.NotThere), (stages[2].Value, stages[2].State));
    }

    // The NEWEST submission - by when it was sent, not by its place in the list - and how many there are.
    [Fact]
    public void Submitted_is_the_newest_submission_with_when_it_was_sent_and_how_many_there_are()
    {
        SystemStage stage = SystemStagesDisplay.Submitted(
        [
            SystemStagesDisplayTests.Submission(7, "pending", SystemStagesDisplayTests.Sent.AddDays(3)),
            SystemStagesDisplayTests.Submission(4, "merged", SystemStagesDisplayTests.Sent, SystemStagesDisplayTests.Sent.AddHours(2))
        ]);

        Assert.Equal($"#7 {SubmissionReceiptPresenter.PendingWording}", stage.Value);
        Assert.Equal($"sent {SubmissionReceiptPresenter.FormatDate(SystemStagesDisplayTests.Sent.AddDays(3))} - the latest of 2", stage.Detail);
    }

    [Fact]
    public void No_submissions_is_said_and_unread_ones_are_not_guessed()
    {
        Assert.Equal(("No submissions", SystemStageState.NotThere), (SystemStagesDisplay.Submitted([]).Value, SystemStagesDisplay.Submitted([]).State));

        SystemStage unread = SystemStagesDisplay.Submitted(null);
        Assert.Equal(string.Empty, unread.Value);
        Assert.Equal(SystemStageState.NotKnown, unread.State);
    }

    // BETA ahead of stable says where it waits, by the tab's own name for that queue.
    [Fact]
    public void BETA_names_its_revision_and_whether_it_waits_to_go_to_stable()
    {
        SystemStage waiting = SystemStagesDisplay.Beta(SystemStagesDisplayTests.System(awaiting: true));
        SystemStage level = SystemStagesDisplay.Beta(SystemStagesDisplayTests.System(betaRevision: null));

        Assert.Equal("Revision 2026-October-4", waiting.Value);
        Assert.Equal($"ahead of stable - waiting under {MaintainerScreenWording.BetaQueueQuoted}", waiting.Detail);
        Assert.Equal("In BETA", level.Value);
        Assert.Null(level.Detail);
    }

    [Fact]
    public void Stable_names_its_revision_and_when_it_was_published_or_why_it_cannot_say()
    {
        DateTimeOffset published = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        SystemStage there = SystemStagesDisplay.Stable(SystemStagesDisplayTests.System(published: published));
        SystemStage unknown = SystemStagesDisplay.Stable(SystemStagesDisplayTests.System(inStable: null));

        Assert.Equal("Revision 2026-September-25", there.Value);
        Assert.Equal($"published {SubmissionReceiptPresenter.FormatDate(published)}", there.Detail);
        Assert.Equal(("Not known", SystemStageState.NotKnown), (unknown.Value, unknown.State));
        Assert.Equal("this server has no stable source to look in", unknown.Detail);
    }
}
