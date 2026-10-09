using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// DraftDiscardWording - "the contributor discarded their own draft" as a maintainer reads it
// (owner request, 2026-09-28). The rule the sentences keep: say what happened and what to do (ask
// the contributor, and on BETA consider pushing it back) - and never call it a withdrawal, since
// discarding a draft does not take back what was sent.
// ###########################################################################################
public sealed class DraftDiscardWordingTests
{
    private static readonly DateTimeOffset Discarded = new(2026, 9, 28, 9, 15, 0, TimeSpan.Zero);

    [Fact]
    public void The_mark_names_the_day_in_the_one_submission_date_format()
    {
        Assert.Equal("Contributor discarded the draft on 2026-September-28", DraftDiscardWording.Mark(Discarded));
    }

    [Fact]
    public void Above_the_table_it_says_to_ask_the_contributor_first_and_that_nothing_was_withdrawn()
    {
        string text = DraftDiscardWording.SubmissionWarning(Discarded, " dennis@example.com ");

        Assert.Contains("2026-September-28", text, StringComparison.Ordinal);
        Assert.Contains("check with the contributor (dennis@example.com) before approving", text, StringComparison.Ordinal);
        Assert.Contains("does not withdraw", text, StringComparison.Ordinal);
    }

    // Contributing needs no account, and a submission can carry no address at all.
    [Fact]
    public void With_no_address_it_still_reads()
    {
        string text = DraftDiscardWording.SubmissionWarning(Discarded, null);

        Assert.Contains("check with the contributor before approving", text, StringComparison.Ordinal);
        Assert.DoesNotContain("()", text, StringComparison.Ordinal);
    }

    // The owner's own advice for BETA: push it back to the queue and ask.
    [Fact]
    public void On_beta_it_names_the_contributor_and_suggests_pushing_it_back()
    {
        string text = DraftDiscardWording.BetaWarning(
            new CarriedSubmission(41, "dennis@example.com", "Change", Discarded.AddDays(-1), DraftDiscardedUtc: Discarded));

        Assert.StartsWith("dennis@example.com discarded the draft of this board on 2026-September-28", text, StringComparison.Ordinal);
        Assert.Contains("pushing it back to the queue and checking with the contributor", text, StringComparison.Ordinal);
    }

    [Fact]
    public void On_beta_with_no_address_it_says_the_contributor()
    {
        string text = DraftDiscardWording.BetaWarning(new CarriedSubmission(41, "", null, null, DraftDiscardedUtc: Discarded));

        Assert.StartsWith("The contributor discarded", text, StringComparison.Ordinal);
    }

    [Fact]
    public void The_boards_history_says_which_submission()
    {
        Assert.Equal(
            "#41 - the contributor discarded the draft",
            BoardsDisplay.HistoryWhat(new BoardHistoryEntry(Discarded, BoardHistoryEvents.DraftDiscarded, "dennis@example.com", 41, null)));
    }

    [Fact]
    public void Nothing_here_calls_it_a_withdrawal()
    {
        foreach (string text in new[]
        {
            DraftDiscardWording.Mark(Discarded),
            DraftDiscardWording.ListMark,
            DraftDiscardWording.BetaWarning(new CarriedSubmission(1, "a@b.c", null, null, Discarded)),
            DraftDiscardWording.HistoryWhat("#1")
        })
        {
            Assert.DoesNotContain("withdrew", text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("withdrawn", text, StringComparison.OrdinalIgnoreCase);
        }
    }

    // ###########################################################################################
    // *** NEVER "THEIR" OR "THEM" FOR THE CONTRIBUTOR (owner request, 2026-10-01: "never refer to
    // a contributor as "their""). *** Every sentence here is about one contributor.
    // ###########################################################################################
    [Fact]
    public void No_sentence_refers_to_the_contributor_as_their_or_them()
    {
        foreach (string text in new[]
        {
            DraftDiscardWording.Mark(Discarded),
            DraftDiscardWording.ListMark,
            DraftDiscardWording.SubmissionWarning(Discarded, "a@b.c"),
            DraftDiscardWording.SubmissionWarning(Discarded, null),
            DraftDiscardWording.BetaWarning(new CarriedSubmission(1, "a@b.c", null, null, Discarded)),
            DraftDiscardWording.BetaWarning(new CarriedSubmission(1, "", null, null, Discarded)),
            DraftDiscardWording.HistoryWhat("#1")
        })
        {
            Assert.DoesNotMatch(@"(?i)\b(their|theirs|them|they)\b", text);
        }
    }
}
