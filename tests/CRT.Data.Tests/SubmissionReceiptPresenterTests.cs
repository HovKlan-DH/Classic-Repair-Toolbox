using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The words the "my submissions" view puts on screen (NewContributeStrategy.md Phase 4, task 6).
//
// WHY THIS IS TESTED AT ALL, given it is a switch over four strings: these are the only
// explanation a contributor gets inside the application, and the server's own state values are
// database vocabulary that would actively mislead if shown raw. "pending" means QUEUED, but reads
// as "not sent yet" to somebody who has just pressed Submit. Each mapping below exists because
// the raw value says the wrong thing, so each one is worth pinning.
// ###########################################################################################
public sealed class SubmissionReceiptPresenterTests
{
    // ------------------------------------------------------------------ DescribeState

    // The four states the server actually writes (see SubmissionState on the server side), each
    // turned into something true for a reader who has never seen the schema.
    [Theory]
    [InlineData("pending", "Submitted - awaiting feedback from a maintainer")]
    [InlineData("rejected", "Not accepted")]
    [InlineData("abandoned", "Expired before it was finished")]
    [InlineData("uploading", "Never finished sending")]
    public void Each_server_state_is_described_in_the_contributors_own_terms(string state, string expected)
    {
        Assert.Equal(expected, SubmissionReceiptPresenter.DescribeState(state));
    }

    // "uploading" is the sharpest of them. To the server it means a submission that was created
    // and never finalised; to the person looking at the row, nothing is uploading any more - the
    // send failed partway. Rendering the raw word would have them waiting for a transfer that
    // stopped days ago.
    [Fact]
    public void An_unfinished_upload_is_not_described_as_still_in_progress()
    {
        string described = SubmissionReceiptPresenter.DescribeState("uploading");

        Assert.DoesNotContain("progress", described, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Never finished", described);
    }

    // Case and spacing come off a wire format and a JSON file that a user could hand-edit.
    [Theory]
    [InlineData("PENDING")]
    [InlineData("  pending  ")]
    [InlineData("Pending")]
    public void State_matching_ignores_case_and_surrounding_space(string state)
    {
        Assert.Equal("Submitted - awaiting feedback from a maintainer", SubmissionReceiptPresenter.DescribeState(state));
    }

    // ###########################################################################################
    // AN UNKNOWN STATE IS REPORTED AS UNKNOWN, never guessed at.
    //
    // Phase 5 will add states this build has never heard of. A switch defaulting to something
    // plausible - "Submitted - awaiting feedback from a maintainer", say - would tell a contributor their rejected submission
    // is still in the queue, and nothing would ever reveal the lie. Saying the server reported
    // something this version does not recognise is unhelpful but true, and it points at the real
    // fix (update the app).
    // ###########################################################################################
    [Fact]
    public void A_state_this_version_does_not_know_is_reported_rather_than_guessed()
    {
        string described = SubmissionReceiptPresenter.DescribeState("awaiting-maintainer-signoff");

        Assert.Contains("awaiting-maintainer-signoff", described);
        Assert.DoesNotContain("Submitted - awaiting feedback from a maintainer", described);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_receipt_never_checked_says_so_rather_than_inventing_a_state(string? state)
    {
        Assert.Equal("Not checked yet", SubmissionReceiptPresenter.DescribeState(state));
    }

    // ###########################################################################################
    // THE FOUR PHASE 5 REVIEW STATES - missing until 2026-09-22, and reported from live use.
    //
    // They were added to the server's vocabulary when the Maintainer tab was built and never
    // taught to DescribeState, so the very first real review round trip showed the contributor
    // "Reported as [changes_requested]" - a raw database value, underscore and all, in the one
    // place this class exists to prevent exactly that.
    //
    // "changes_requested" is the one that mattered: it is the state where somebody is being ASKED
    // TO DO SOMETHING, and it read as a fault in the application instead.
    // ###########################################################################################
    [Theory]
    [InlineData("changes_requested", "Changes requested")]
    [InlineData("approved", "Approved, waiting to be published")]
    // "merged" is the BETA data since the two-stage publish (2026-09-25); "published" is everyone's.
    [InlineData("merged", "Published to the BETA source")]
    [InlineData("published", "Published to the stable source")]
    // The server's word for a submission its contributor's newer one replaced (2026-09-26).
    [InlineData("withdrawn", "Replaced by a newer submission")]
    public void Each_REVIEW_state_is_described_in_the_contributors_own_terms(string state, string expected)
    {
        Assert.Equal(expected, SubmissionReceiptPresenter.DescribeState(state));
    }

    [Fact]
    public void No_review_state_leaks_a_RAW_DATABASE_VALUE_to_the_reader()
    {
        // The shape of the reported bug, asserted directly: no underscore, and no "Reported as"
        // fallback, for any state the server can actually write. A future state added server-side
        // without being taught here fails this immediately.
        string[] everyServerState =
        [
            "uploading", "pending", "changes_requested", "approved",
            "rejected", "withdrawn", "merged", "abandoned"
        ];

        foreach (string state in everyServerState)
        {
            string described = SubmissionReceiptPresenter.DescribeState(state);

            Assert.DoesNotContain("Reported as", described, StringComparison.Ordinal);
            Assert.DoesNotContain("_", described, StringComparison.Ordinal);
        }
    }

    // ------------------------------------------------------------------ ClassifyState

    // ###########################################################################################
    // THE COLOUR AND THE WORDS COME OFF THE SAME SWITCH (owner request, 2026-09-22).
    //
    // The submissions list paints a status-coloured edge on each card. Deciding that colour in the
    // UI would let it drift from DescribeState - a card reading "Changes requested" in green is
    // exactly the kind of thing nobody notices until a contributor acts on it.
    // ###########################################################################################
    [Theory]
    [InlineData("published")]
    [InlineData("merged")]
    [InlineData("accepted")]
    [InlineData("approved")]
    public void A_submission_that_got_IN_is_classified_as_good(string state)
    {
        Assert.Equal(SubmissionOutcomeKind.Good, SubmissionReceiptPresenter.ClassifyState(state));
    }

    [Theory]
    [InlineData("rejected")]
    [InlineData("abandoned")]
    // "uploading" belongs here rather than in Waiting: nothing is uploading any more, the send
    // failed partway, and the row is as dead as a rejection until it is sent again. Colouring it
    // as "waiting" would have the contributor waiting for a transfer that stopped days ago - the
    // same trap DescribeState's own wording exists to avoid.
    [InlineData("uploading")]
    public void A_submission_that_did_NOT_get_in_is_classified_as_bad(string state)
    {
        Assert.Equal(SubmissionOutcomeKind.Bad, SubmissionReceiptPresenter.ClassifyState(state));
    }

    // ###########################################################################################
    // *** REPLACED IS NOT REFUSED (2026-09-26). *** A submission the contributor's own newer one
    // replaced ("withdrawn" on the server) carries on in that one - painted as bad, the card read
    // as "your contribution was turned down". Neutral instead.
    // ###########################################################################################
    [Fact]
    public void A_submission_replaced_by_a_newer_one_is_not_painted_as_refused()
    {
        Assert.Equal(SubmissionOutcomeKind.Waiting, SubmissionReceiptPresenter.ClassifyState("withdrawn"));
    }

    [Fact]
    public void CHANGES_REQUESTED_is_its_own_kind_because_it_is_about_the_READER()
    {
        // The one state where the contributor has to do something. Kept distinct from Bad even
        // though both paint red, so the distinction survives for anything that needs it later -
        // a count of "things waiting on me", say.
        Assert.Equal(
            SubmissionOutcomeKind.NeedsAction,
            SubmissionReceiptPresenter.ClassifyState("changes_requested"));
    }

    [Theory]
    [InlineData("pending")]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("awaiting-maintainer-signoff")]
    public void Anything_UNDECIDED_or_unknown_is_neutral_rather_than_guessed(string? state)
    {
        // Colouring an unrecognised state as good or bad would state something this build cannot
        // know - the same reasoning DescribeState applies to its own fallback.
        Assert.Equal(SubmissionOutcomeKind.Waiting, SubmissionReceiptPresenter.ClassifyState(state));
    }

    [Fact]
    public void The_classification_agrees_with_the_WORDS_for_every_server_state()
    {
        // *** THE ANTI-DRIFT CHECK. *** Both halves read the same state string, so this pins that
        // they cannot disagree: nothing DescribeState calls "requested" may be classified Good,
        // and nothing it calls "Published" may be classified Bad.
        foreach (string state in new[] { "published", "merged", "accepted", "approved" })
        {
            Assert.DoesNotContain(
                "Not accepted", SubmissionReceiptPresenter.DescribeState(state), StringComparison.Ordinal);
        }

        Assert.Equal("Changes requested", SubmissionReceiptPresenter.DescribeState("changes_requested"));
        Assert.Equal(
            SubmissionOutcomeKind.NeedsAction,
            SubmissionReceiptPresenter.ClassifyState("changes_requested"));
    }

    // ------------------------------------------------------------------ IsStillOpen

    // What decides whether the server is asked about this row again. A decided submission can
    // never change state, so re-asking is a request per row per refresh that can only ever return
    // what is already on screen.
    [Theory]
    [InlineData("rejected")]
    [InlineData("abandoned")]
    [InlineData("accepted")]
    [InlineData("published")]
    // *** "withdrawn" WAS MISSING UNTIL 2026-09-22, and that was not cosmetic. *** Final, so
    // omitting it meant every such row was re-checked on every launch forever.
    [InlineData("withdrawn")]
    public void A_decided_submission_is_not_asked_about_again(string state)
    {
        Assert.False(SubmissionReceiptPresenter.IsStillOpen(state));
    }

    [Theory]
    [InlineData("pending")]
    [InlineData("uploading")]
    // *** DELIBERATELY STILL OPEN. *** Neither is the end of the road: a submission with changes
    // requested is re-reviewed once the contributor acts, and an approved one is still waiting to
    // be published. Treating either as decided would freeze the row and the contributor would
    // never see it move again.
    [InlineData("changes_requested")]
    [InlineData("approved")]
    // *** "merged" IS ASKED ABOUT AGAIN SINCE 2026-09-25. *** It was final until the two-stage
    // publish; now it means "in the BETA data" and moves on to "published" when the board reaches
    // production. Closing it would freeze the row at "in the BETA data" for ever.
    [InlineData("merged")]
    public void A_submission_that_can_still_move_is_asked_about(string state)
    {
        Assert.True(SubmissionReceiptPresenter.IsStillOpen(state));
    }

    // An unrecognised state is treated as still open, which is the safe direction: the cost of
    // being wrong is one extra request, whereas treating an unknown state as final would freeze a
    // row at whatever this build last understood.
    [Fact]
    public void An_unknown_state_is_treated_as_still_open()
    {
        Assert.True(SubmissionReceiptPresenter.IsStillOpen("awaiting-maintainer-signoff"));
    }

    [Fact]
    public void A_never_checked_receipt_is_still_open()
    {
        Assert.True(SubmissionReceiptPresenter.IsStillOpen(null));
        Assert.True(SubmissionReceiptPresenter.IsStillOpen(string.Empty));
    }

    // ###########################################################################################
    // *** A "MERGED" RECEIPT IS ASKED ABOUT FOR A WHILE, NOT FOR EVER (code review, 2026-09-25). ***
    // Publishing to production is off until the server is set up for it, and a board may never be
    // promoted; without a bound every merged receipt was asked about on every launch for ever.
    // ###########################################################################################
    private static readonly DateTimeOffset Launch = new(2026, 11, 1, 12, 0, 0, TimeSpan.Zero);

    // `lastChecked` defaults to the day before Launch: past the window a merged receipt is asked about
    // only every MergedLateRecheckInterval, so a receipt checked yesterday is NOT due - which is
    // what the window tests below are about.
    private static SubmissionReceipt Receipt(
        string state,
        DateTimeOffset sent,
        DateTimeOffset? decided = null,
        DateTimeOffset? lastChecked = null) => new()
    {
        SubmissionId = 1,
        LastKnownState = state,
        SentUtc = sent,
        DecidedUtc = decided,
        LastCheckedUtc = lastChecked ?? SubmissionReceiptPresenterTests.Launch.AddDays(-1)
    };

    [Fact]
    public void A_merged_receipt_is_asked_about_at_every_launch_until_the_window_after_its_decision_closes()
    {
        TimeSpan window = SubmissionReceiptPresenter.MergedRecheckWindow;
        DateTimeOffset sent = SubmissionReceiptPresenterTests.Launch - window - TimeSpan.FromDays(10);

        SubmissionReceipt recent = SubmissionReceiptPresenterTests.Receipt(
            "merged", sent, SubmissionReceiptPresenterTests.Launch - window + TimeSpan.FromDays(1));
        SubmissionReceipt old = SubmissionReceiptPresenterTests.Receipt(
            "merged", sent, SubmissionReceiptPresenterTests.Launch - window - TimeSpan.FromDays(1));

        Assert.True(SubmissionReceiptPresenter.IsStillOpen(recent, SubmissionReceiptPresenterTests.Launch));
        Assert.False(SubmissionReceiptPresenter.IsStillOpen(old, SubmissionReceiptPresenterTests.Launch));
    }

    // A receipt from before decision dates were kept falls back to when it was sent.
    [Fact]
    public void A_merged_receipt_with_no_decision_date_is_timed_from_when_it_was_sent()
    {
        TimeSpan window = SubmissionReceiptPresenter.MergedRecheckWindow;

        Assert.True(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("merged", SubmissionReceiptPresenterTests.Launch.AddDays(-2)),
            SubmissionReceiptPresenterTests.Launch));

        Assert.False(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("merged", SubmissionReceiptPresenterTests.Launch - window - TimeSpan.FromDays(1)),
            SubmissionReceiptPresenterTests.Launch));
    }

    // ###########################################################################################
    // *** PAST THE WINDOW, NOW AND THEN - NEVER NEVER (code review, 2026-09-27). *** A merged
    // submission can be taken back out of BETA long after it merged, and the old hard stop meant a
    // rollback on day 31 never reached "My submissions": the row said "Published to the BETA source"
    // for good while the server said otherwise. Now it is asked about once a
    // MergedLateRecheckInterval. Fails against the version that answered false past the window.
    // ###########################################################################################
    [Fact]
    public void A_merged_receipt_past_the_window_is_asked_about_again_once_it_has_not_been_checked_for_a_while()
    {
        DateTimeOffset longAgo = SubmissionReceiptPresenterTests.Launch.AddDays(-90);
        TimeSpan interval = SubmissionReceiptPresenter.MergedLateRecheckInterval;

        Assert.True(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("merged", longAgo, longAgo, lastChecked: SubmissionReceiptPresenterTests.Launch - interval),
            SubmissionReceiptPresenterTests.Launch));

        Assert.False(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("merged", longAgo, longAgo, lastChecked: SubmissionReceiptPresenterTests.Launch - interval + TimeSpan.FromHours(1)),
            SubmissionReceiptPresenterTests.Launch));
    }

    // A receipt never checked at all is due - there is no answer on it to keep.
    [Fact]
    public void A_merged_receipt_past_the_window_that_was_never_checked_is_asked_about()
    {
        DateTimeOffset longAgo = SubmissionReceiptPresenterTests.Launch.AddDays(-90);

        SubmissionReceipt unchecked_ = new()
        {
            SubmissionId = 1,
            LastKnownState = "merged",
            SentUtc = longAgo,
            DecidedUtc = longAgo
        };

        Assert.True(SubmissionReceiptPresenter.IsStillOpen(unchecked_, SubmissionReceiptPresenterTests.Launch));
    }

    // -----------------------------------------------------------------------------------
    // "returned" - TAKEN BACK OUT OF BETA (code review, 2026-09-27)
    //
    // The server reports it for a pending submission carrying a maintainer's reason, which only a
    // BETA rollback produces. Without these words the contributor, already told "Published to BETA
    // source", saw a bare "Waiting for review" and nothing about what had happened.
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_returned_submission_says_it_was_taken_back_out_of_BETA()
    {
        Assert.Equal(
            "Taken back out of BETA - waiting for review again",
            SubmissionReceiptPresenter.DescribeState("returned"));
    }

    // It was pulled because something needs attention, so it takes the colour that says "look".
    [Fact]
    public void A_returned_submission_is_coloured_as_needing_the_contributor()
    {
        Assert.Equal(SubmissionOutcomeKind.NeedsAction, SubmissionReceiptPresenter.ClassifyState("returned"));
    }

    // Waiting for review again, so it is still asked about - it will be decided a second time.
    [Fact]
    public void A_returned_submission_is_still_open()
    {
        Assert.True(SubmissionReceiptPresenter.IsStillOpen("returned"));
        Assert.True(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("returned", SubmissionReceiptPresenterTests.Launch.AddYears(-1)),
            SubmissionReceiptPresenterTests.Launch));
    }

    // And it is NEWS: a contributor who saw "merged" is badged when it turns into "returned".
    [Fact]
    public void A_merged_submission_turning_returned_raises_the_badge()
    {
        var receipt = new SubmissionReceipt
        {
            SubmissionId = 1,
            LastKnownState = "returned",
            AcknowledgedState = "merged",
            MaintainerComment = "The U8 pinout is wrong.",
            AcknowledgedComment = string.Empty
        };

        Assert.True(SubmissionReceiptPresenter.HasUnreadDecision(receipt));
        Assert.True(SubmissionReceiptPresenter.HasUnreadNews(receipt));
    }

    // Only "merged" is bounded: a submission still waiting for its review is asked about however
    // old it is, and a decided one never is.
    [Fact]
    public void The_window_bounds_merged_alone()
    {
        DateTimeOffset longAgo = SubmissionReceiptPresenterTests.Launch.AddYears(-1);

        Assert.True(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("pending", longAgo), SubmissionReceiptPresenterTests.Launch));
        Assert.False(SubmissionReceiptPresenter.IsStillOpen(
            SubmissionReceiptPresenterTests.Receipt("published", SubmissionReceiptPresenterTests.Launch), SubmissionReceiptPresenterTests.Launch));
    }

    // ------------------------------------------------------------------ DescribeDecided

    // ###########################################################################################
    // WHEN THE MAINTAINER REPLIED - distinct from "last checked", which is this computer's own
    // bookkeeping and says nothing about the submission (owner request, 2026-09-22).
    //
    // A comment written three weeks ago on a board the contributor has since revised means
    // something quite different from one written this morning, and the window could not tell them
    // apart: it showed only when this machine last asked.
    // ###########################################################################################
    [Fact]
    public void The_decided_line_names_the_day_the_maintainer_replied()
    {
        var decided = new DateTimeOffset(2026, 9, 22, 9, 15, 0, TimeSpan.Zero);

        string expectedDay = SubmissionReceiptPresenter.FormatDate(decided);

        Assert.Equal("Replied " + expectedDay, SubmissionReceiptPresenter.DescribeDecided(decided));
    }

    [Fact]
    public void NOTHING_DECIDED_yields_an_empty_line_rather_than_a_label_with_no_date()
    {
        // Empty so the caller can drop the line entirely. "Replied" followed by nothing would look
        // like a defect, and every submission still waiting for review is in this state.
        Assert.Equal(string.Empty, SubmissionReceiptPresenter.DescribeDecided(null));
    }

    [Fact]
    public void The_decided_line_is_LOCAL_and_carries_no_clock_time()
    {
        // Same rule as DescribeSent, and for the same reason: the two clocks belong to different
        // people in different time zones, so a time implies a precision that is not there. 23:30
        // UTC is already the next day anywhere east of it.
        var lateEvening = new DateTimeOffset(2026, 9, 21, 23, 30, 0, TimeSpan.Zero);

        string described = SubmissionReceiptPresenter.DescribeDecided(lateEvening);

        Assert.Equal(
            "Replied " + SubmissionReceiptPresenter.FormatDate(lateEvening),
            described);

        Assert.DoesNotContain(":", described, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ DescribeSent

    // ###########################################################################################
    // THE DAY, NEVER A TIME - and converted to LOCAL time before formatting.
    //
    // A stored timestamp is UTC (so a receipts file does not reorder itself across time zones),
    // but a contributor recalls "I sent that on Tuesday" in their own time. Formatting the UTC
    // value directly would show the wrong DAY for anyone far enough east or west, on exactly the
    // submissions sent late in the evening.
    // ###########################################################################################
    [Fact]
    public void The_sent_line_names_the_local_day_not_the_UTC_one()
    {
        // 23:30 UTC. Anywhere east of UTC this is already the following day locally, and that is
        // the day the person remembers.
        var lateEvening = new DateTimeOffset(2026, 9, 21, 23, 30, 0, TimeSpan.Zero);

        string described = SubmissionReceiptPresenter.DescribeSent(lateEvening);
        string expectedDay = SubmissionReceiptPresenter.FormatDate(lateEvening);

        Assert.Equal("Sent " + expectedDay, described);
    }

    [Fact]
    public void The_sent_line_carries_no_clock_time()
    {
        string described = SubmissionReceiptPresenter.DescribeSent(
            new DateTimeOffset(2026, 9, 21, 14, 45, 0, TimeSpan.Zero));

        // A time would imply a precision this does not have - the clock is the sender's own.
        Assert.DoesNotContain(":", described);
    }

    // ------------------------------------------------------------------ InDisplayOrder

    // Newest first: the thing just sent is the thing being asked about.
    [Fact]
    public void The_newest_submission_is_listed_first()
    {
        var receipts = new List<SubmissionReceipt>
        {
            new() { SubmissionId = 1, SentUtc = new DateTimeOffset(2026, 9, 1, 0, 0, 0, TimeSpan.Zero) },
            new() { SubmissionId = 2, SentUtc = new DateTimeOffset(2026, 9, 20, 0, 0, 0, TimeSpan.Zero) },
            new() { SubmissionId = 3, SentUtc = new DateTimeOffset(2026, 9, 10, 0, 0, 0, TimeSpan.Zero) },
        };

        Assert.Equal(
            new long[] { 2, 3, 1 },
            SubmissionReceiptPresenter.InDisplayOrder(receipts).Select(receipt => receipt.SubmissionId));
    }

    // Two submissions sent in the same second (a clock with coarse resolution, or two sends in
    // quick succession) still order deterministically rather than by whatever order the file
    // happened to hold, or the list would reshuffle between refreshes.
    [Fact]
    public void Receipts_sent_at_the_same_moment_still_order_deterministically()
    {
        var sameMoment = new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        var receipts = new List<SubmissionReceipt>
        {
            new() { SubmissionId = 7, SentUtc = sameMoment },
            new() { SubmissionId = 9, SentUtc = sameMoment },
            new() { SubmissionId = 8, SentUtc = sameMoment },
        };

        Assert.Equal(
            new long[] { 9, 8, 7 },
            SubmissionReceiptPresenter.InDisplayOrder(receipts).Select(receipt => receipt.SubmissionId));
    }

    [Fact]
    public void An_absent_list_orders_to_nothing_rather_than_throwing()
    {
        Assert.Empty(SubmissionReceiptPresenter.InDisplayOrder(null));
    }

    // -----------------------------------------------------------------------------------
    // UNREAD MAINTAINER FEEDBACK - what the badge on the Drafts tab counts.
    //
    // *** A COMMENT NOBODY NOTICES IS A COMMENT NOBODY READS. *** Contributing needs no account,
    // so there is no inbox and no thread: the maintainer's sentence is the entire channel back to
    // the person who did the work, and it lived behind a button nobody had a reason to press.
    // -----------------------------------------------------------------------------------

    private static SubmissionReceipt Receipt(string comment, string acknowledged) =>
        new()
        {
            SubmissionId = 7,
            MaintainerComment = comment,
            AcknowledgedComment = acknowledged
        };

    [Fact]
    public void A_comment_that_has_never_been_acknowledged_is_UNREAD()
    {
        Assert.True(SubmissionReceiptPresenter.HasUnreadComment(
            SubmissionReceiptPresenterTests.Receipt("Check the U8 highlight.", string.Empty)));
    }

    [Fact]
    public void A_comment_that_has_been_acknowledged_is_READ()
    {
        Assert.False(SubmissionReceiptPresenter.HasUnreadComment(
            SubmissionReceiptPresenterTests.Receipt("Check the U8 highlight.", "Check the U8 highlight.")));
    }

    [Fact]
    public void A_SECOND_round_of_feedback_is_unread_again()
    {
        // *** THE CASE A "SEEN" BOOLEAN GETS WRONG, and the reason this is stored as TEXT. ***
        // The contributor reads "fix U8", fixes it, resubmits, and is asked for something else.
        // With a flag, whoever writes the new comment has to remember to clear it - and the one
        // time that is forgotten, the second request arrives already marked read and is never
        // seen at all. Comparing the text makes a changed comment unread by construction.
        Assert.True(SubmissionReceiptPresenter.HasUnreadComment(
            SubmissionReceiptPresenterTests.Receipt(
                "Now the region is wrong too.", "Check the U8 highlight.")));
    }

    [Fact]
    public void NO_comment_is_not_unread()
    {
        // There is nothing to read. A badge here would send somebody to an empty box.
        Assert.False(SubmissionReceiptPresenter.HasUnreadComment(
            SubmissionReceiptPresenterTests.Receipt(string.Empty, string.Empty)));

        Assert.False(SubmissionReceiptPresenter.HasUnreadComment(
            SubmissionReceiptPresenterTests.Receipt("   ", string.Empty)));
    }

    [Fact]
    public void WHITESPACE_alone_does_not_resurface_an_acknowledged_comment()
    {
        // The same sentence with a trailing newline is the same sentence. A server or a JSON
        // round-trip that adds or drops one must not make the badge reappear - a badge that comes
        // back for no visible reason is one the contributor learns to ignore.
        Assert.False(SubmissionReceiptPresenter.HasUnreadComment(
            SubmissionReceiptPresenterTests.Receipt("Check U8.\n", "  Check U8.")));
    }

    [Fact]
    public void A_null_receipt_is_not_unread_rather_than_throwing()
    {
        // This is called while rendering rows and while counting for a badge, neither of which
        // may crash the tab.
        Assert.False(SubmissionReceiptPresenter.HasUnreadComment(null));
    }

    // ------------------------------------------------------ HasUnreadDecision (2026-09-23)

    // ###########################################################################################
    // *** THE BUG THESE EXIST FOR: AN APPROVAL WITH NO COMMENT WAS COMPLETELY SILENT. ***
    //
    // A maintainer approving good work usually types nothing - there is nothing to say - so
    // HasUnreadComment saw an empty comment and answered "not unread". The submission went to
    // "Published", the row said so, and the application told the contributor nothing at all: no
    // badge, and once the draft was gone, no Drafts tab either. The one outcome everybody is
    // actually waiting for was the one that arrived without a sound. Reported by the project owner
    // after the first real publish.
    // ###########################################################################################
    private static SubmissionReceipt Decided(string state, string acknowledgedState) =>
        new()
        {
            SubmissionId = 11,
            LastKnownState = state,
            AcknowledgedState = acknowledgedState
        };

    [Theory]
    [InlineData("published")]
    [InlineData("merged")]
    [InlineData("rejected")]
    [InlineData("changes_requested")]
    [InlineData("approved")]
    [InlineData("withdrawn")]
    [InlineData("abandoned")]
    public void A_decision_nobody_has_seen_is_UNREAD_even_with_no_comment(string state)
    {
        Assert.True(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided(state, string.Empty)));
    }

    [Fact]
    public void A_decision_already_seen_is_READ()
    {
        Assert.False(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided("published", "published")));
    }

    // ###########################################################################################
    // *** "pending" IS NOT NEWS, and this is the test that stops the badge firing on everything. ***
    // It is the state a submission is in from the moment it is sent, so treating it as a decision
    // would badge every contribution the instant it left the machine - and a badge that is always
    // lit is one nobody reads.
    // ###########################################################################################
    [Theory]
    [InlineData("pending")]
    [InlineData("uploading")]
    public void A_submission_still_waiting_is_NOT_unread(string state)
    {
        Assert.False(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided(state, string.Empty)));
    }

    [Fact]
    public void A_receipt_never_checked_against_the_server_is_NOT_unread()
    {
        // No state at all means nothing has been learned yet, which is not something to interrupt
        // somebody for.
        Assert.False(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided(string.Empty, string.Empty)));
    }

    [Fact]
    public void Reaching_PRODUCTION_after_BETA_is_unread_a_second_time()
    {
        // merged -> published is the two-stage publish's second step: the contributor's own data
        // has it now. That is news worth the badge.
        Assert.True(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided("published", "merged")));
    }

    [Fact]
    public void A_decision_that_MOVES_AGAIN_is_unread_a_second_time()
    {
        // approved -> merged is two different facts, and the second one ("it is actually live
        // now") is the one the contributor was waiting for. Storing the seen STATE rather than a
        // flag makes that work with nothing having to remember to reset anything - the same
        // reasoning the comment side uses.
        Assert.True(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided("merged", "approved")));
    }

    [Fact]
    public void An_unknown_state_is_reported_rather_than_swallowed()
    {
        // A state this build has never heard of means the server did something. Telling the
        // contributor to go and look is safer than hiding it, and matches DescribeState, which
        // reports an unknown value honestly rather than guessing.
        Assert.True(SubmissionReceiptPresenter.HasUnreadDecision(
            SubmissionReceiptPresenterTests.Decided("some_future_state", string.Empty)));
    }

    [Fact]
    public void A_null_receipt_is_not_an_unread_decision_rather_than_throwing()
    {
        Assert.False(SubmissionReceiptPresenter.HasUnreadDecision(null));
    }

    // ------------------------------------------------------------------ HasUnreadNews

    [Fact]
    public void News_is_either_a_comment_or_a_decision_or_both()
    {
        // A comment with no decision - the pre-existing case.
        Assert.True(SubmissionReceiptPresenter.HasUnreadNews(
            SubmissionReceiptPresenterTests.Receipt("Please add the region.", string.Empty)));

        // A decision with no comment - the reported case.
        Assert.True(SubmissionReceiptPresenter.HasUnreadNews(
            SubmissionReceiptPresenterTests.Decided("published", string.Empty)));

        // Both seen.
        Assert.False(SubmissionReceiptPresenter.HasUnreadNews(new SubmissionReceipt
        {
            SubmissionId = 12,
            LastKnownState = "published",
            AcknowledgedState = "published",
            MaintainerComment = "Nice work.",
            AcknowledgedComment = "Nice work."
        }));
    }

    [Fact]
    public void The_badge_counts_a_SILENTLY_PUBLISHED_submission()
    {
        // *** THE REPORTED SCREEN, as a count. *** Before the fix this returned 0 and the
        // contributor was told nothing whatsoever about work that had shipped.
        var receipts = new List<SubmissionReceipt>
        {
            SubmissionReceiptPresenterTests.Decided("published", string.Empty)
        };

        Assert.Equal(1, SubmissionReceiptPresenter.UnreadCommentCount(receipts));
    }

    [Fact]
    public void A_row_with_BOTH_a_new_comment_and_a_new_decision_counts_ONCE()
    {
        // The badge counts rows to go and look at, not items of news. Counting twice would tell
        // the contributor there are two submissions wanting attention when there is one.
        var receipts = new List<SubmissionReceipt>
        {
            new()
            {
                SubmissionId = 13,
                LastKnownState = "changes_requested",
                AcknowledgedState = string.Empty,
                MaintainerComment = "Please fix the region.",
                AcknowledgedComment = string.Empty
            }
        };

        Assert.Equal(1, SubmissionReceiptPresenter.UnreadCommentCount(receipts));
    }

    [Fact]
    public void The_badge_counts_only_the_receipts_with_unread_feedback()
    {
        var receipts = new List<SubmissionReceipt>
        {
            SubmissionReceiptPresenterTests.Receipt("Please add the region.", string.Empty),
            SubmissionReceiptPresenterTests.Receipt("Looks good.", "Looks good."),
            SubmissionReceiptPresenterTests.Receipt(string.Empty, string.Empty),
            SubmissionReceiptPresenterTests.Receipt("And the pin numbering.", "An older note.")
        };

        Assert.Equal(2, SubmissionReceiptPresenter.UnreadCommentCount(receipts));
    }

    [Fact]
    public void The_badge_counts_nothing_for_an_absent_or_empty_list()
    {
        // Zero rather than an exception: the tab asks for this on every refresh, including before
        // anything has ever been submitted.
        Assert.Equal(0, SubmissionReceiptPresenter.UnreadCommentCount(null));
        Assert.Equal(0, SubmissionReceiptPresenter.UnreadCommentCount([]));
    }

    [Fact]
    public void A_receipt_written_BEFORE_this_field_existed_reads_as_unread_once_a_comment_arrives()
    {
        // Older receipts on disk have no AcknowledgedComment, which deserialises to empty. That is
        // the correct default: the first comment to arrive for them is genuinely unseen, and
        // treating it as already-read would silently swallow the very first piece of feedback a
        // long-standing contributor ever gets.
        var receipt = new SubmissionReceipt
        {
            SubmissionId = 1,
            MaintainerComment = "Could you check the region?"
        };

        Assert.True(SubmissionReceiptPresenter.HasUnreadComment(receipt));
    }

    // ------------------------------------------------------------------ FormatDate

    // ###########################################################################################
    // THE ONE DATE SHAPE, PINNED LITERALLY (owner request, 2026-09-23).
    //
    // Every other date test in this file now asserts THROUGH FormatDate, which on its own would be
    // vacuous - the format could be changed to anything and they would all still pass together.
    // These are the tests that actually hold the shape, so exactly one place has to be right.
    // ###########################################################################################
    [Fact]
    public void A_date_reads_YEAR_then_MONTH_NAME_then_DAY()
    {
        // Local time, so the value is constructed at midday to keep it on the same calendar day in
        // every time zone this could run in - the point here is the SHAPE, and a date that slid
        // across midnight on the CI runner would fail for an unrelated reason.
        var midday = new DateTimeOffset(2026, 9, 23, 12, 0, 0, DateTimeOffset.Now.Offset);

        Assert.Equal("2026-September-23", SubmissionReceiptPresenter.FormatDate(midday));
    }

    [Fact]
    public void A_single_digit_day_carries_NO_leading_zero()
    {
        // Asked for explicitly. It is also the reason this cannot be a plain round-trip format:
        // "yyyy-MM-dd" would give 2026-09-03, and the month has to be a NAME regardless.
        var third = new DateTimeOffset(2026, 9, 3, 12, 0, 0, DateTimeOffset.Now.Offset);

        Assert.Equal("2026-September-3", SubmissionReceiptPresenter.FormatDate(third));
    }

    // ###########################################################################################
    // *** THE MONTH NAME MUST NOT FOLLOW THE MACHINE'S CULTURE. ***
    //
    // This is the bug that produced the reported "22 september 2026": on a Danish machine the old
    // CultureInfo.CurrentCulture format rendered a lower-case month, because Danish does not
    // capitalise them. The format string looked correct in the source and produced something else
    // on the project owner's own computer.
    //
    // Danish is used here deliberately - it is the locale the defect was actually seen in - and the
    // date separator differs too, so a regression to the current culture fails on more than one
    // axis.
    // ###########################################################################################
    [Fact]
    public void The_month_name_is_ENGLISH_whatever_the_machine_is_set_to()
    {
        CultureInfo original = Thread.CurrentThread.CurrentCulture;

        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("da-DK");

            var midday = new DateTimeOffset(2026, 9, 23, 12, 0, 0, DateTimeOffset.Now.Offset);

            Assert.Equal("2026-September-23", SubmissionReceiptPresenter.FormatDate(midday));
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = original;
        }
    }

    // ------------------------------------------------------------------ DescribeLastChecked

    // ###########################################################################################
    // "Last checked" is THIS COMPUTER'S bookkeeping - when CRT last asked the server, and nothing
    // about the submission itself. The project owner asked what it meant, which is reason enough to
    // pin the wording: it sits beside two dates that ARE facts about the contribution.
    // ###########################################################################################
    [Fact]
    public void The_last_checked_line_names_the_day_this_computer_last_asked()
    {
        var checkedAt = new DateTimeOffset(2026, 9, 23, 12, 0, 0, DateTimeOffset.Now.Offset);

        Assert.Equal(
            "Last checked 2026-September-23",
            SubmissionReceiptPresenter.DescribeLastChecked(checkedAt));
    }

    [Fact]
    public void A_NEVER_CHECKED_receipt_yields_an_empty_line_rather_than_a_dangling_label()
    {
        // Same rule as DescribeDecided: the caller drops the line entirely, rather than printing
        // "Last checked" followed by nothing - which reads as a defect.
        Assert.Equal(string.Empty, SubmissionReceiptPresenter.DescribeLastChecked(null));
    }

    [Fact]
    public void The_last_checked_line_carries_no_clock_time()
    {
        string described = SubmissionReceiptPresenter.DescribeLastChecked(
            new DateTimeOffset(2026, 9, 23, 14, 45, 0, TimeSpan.Zero));

        Assert.DoesNotContain(":", described, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ ClassifyState, warning

    // ###########################################################################################
    // *** "Changes requested" IS ITS OWN BUCKET, AND THAT IS WHAT THE COLOUR HANGS ON. ***
    //
    // Already covered above, but restated here as the thing the 2026-09-23 colour change depends
    // on: the card paints NeedsAction with Text_Warning_Fg and Bad with Text_Fail_Fg, so if these
    // two states ever collapsed into one bucket the two cards would go back to looking identical -
    // which is exactly what was reported.
    // ###########################################################################################
    [Fact]
    public void CHANGES_REQUESTED_and_REJECTED_are_DIFFERENT_buckets()
    {
        Assert.Equal(
            SubmissionOutcomeKind.NeedsAction,
            SubmissionReceiptPresenter.ClassifyState("changes_requested"));

        Assert.Equal(
            SubmissionOutcomeKind.Bad,
            SubmissionReceiptPresenter.ClassifyState("rejected"));

        Assert.NotEqual(
            SubmissionReceiptPresenter.ClassifyState("changes_requested"),
            SubmissionReceiptPresenter.ClassifyState("rejected"));
    }
    // A maintainer changed the submission in the Maintainer tab before deciding (2026-09-25):
    // said, so a contributor comparing what was published with what they sent knows why.
    [Fact]
    public void A_submission_a_maintainer_changed_says_so_and_one_nobody_changed_says_nothing()
    {
        Assert.StartsWith("A maintainer changed some of the details", SubmissionReceiptPresenter.DescribeAmended(true));
        Assert.Equal(string.Empty, SubmissionReceiptPresenter.DescribeAmended(false));
    }

    // ------------------------------------------------------------------ the Drafts tab's badge

    // ###########################################################################################
    // LatestForSystem - which submission a Drafts row's badge describes (owner request,
    // 2026-09-27: "When I have submitted ... I need to see that somehow"). Matched on the system id
    // a submission is sent under, and the newest one wins - a newer submission replaces an older
    // one on the server too.
    // ###########################################################################################
    private static SubmissionReceipt Receipt(long id, string systemId, string sentUtc, string state = "pending") => new()
    {
        SubmissionId = id,
        SystemId = systemId,
        SentUtc = DateTimeOffset.Parse(sentUtc, CultureInfo.InvariantCulture),
        LastKnownState = state,
    };

    [Fact]
    public void The_badge_describes_the_systems_NEWEST_submission()
    {
        SubmissionReceipt latest = SubmissionReceiptPresenter.LatestForSystem(
            [
                Receipt(3, "Commodore/C128/310378 Open128", "2026-09-20T10:00:00Z", "withdrawn"),
                Receipt(8, "Commodore/C128/310378 Open128", "2026-09-27T10:46:11Z", "merged"),
                Receipt(5, "Commodore/C128/310378 Open128", "2026-09-24T09:00:00Z", "rejected"),
            ],
            "Commodore/C128/310378 Open128")!;

        Assert.Equal(8, latest.SubmissionId);
    }

    [Fact]
    public void Another_systems_submissions_are_not_this_ones()
    {
        Assert.Null(SubmissionReceiptPresenter.LatestForSystem(
            [Receipt(1, "Commodore/C128/310378", "2026-09-27T10:00:00Z")],
            "Commodore/C128/310378 Open128"));
    }

    // The id is folder names, which compare like folder names everywhere else in the app.
    [Fact]
    public void The_system_id_matches_ignoring_case_and_outer_spaces()
    {
        Assert.NotNull(SubmissionReceiptPresenter.LatestForSystem(
            [Receipt(1, " commodore/c128/310378 OPEN128 ", "2026-09-27T10:00:00Z")],
            "Commodore/C128/310378 Open128"));
    }

    // ###########################################################################################
    // *** A NEW DRAFT DOES NOT WEAR AN OLDER SUBMISSION'S STATE (code review, 2026-09-27). *** The
    // first submission reached production and its draft was retired; the contributor starts a new
    // draft of the same board. Its row showed "Published to the stable source" beside changes never sent.
    // Only submissions sent since the draft was created describe it - and the next one it sends does.
    // ###########################################################################################
    [Fact]
    public void Submissions_sent_before_the_draft_existed_do_not_describe_it()
    {
        SubmissionReceipt[] receipts = [Receipt(8, "Commodore/C64/250407", "2026-09-20T10:00:00Z", "published")];
        DateTimeOffset draftCreated = DateTimeOffset.Parse("2026-09-25T08:00:00Z", CultureInfo.InvariantCulture);

        Assert.Null(SubmissionReceiptPresenter.LatestForSystem(receipts, "Commodore/C64/250407", draftCreated));

        SubmissionReceipt[] sentSince = [.. receipts, Receipt(9, "Commodore/C64/250407", "2026-09-26T12:00:00Z")];
        Assert.Equal(9, SubmissionReceiptPresenter.LatestForSystem(sentSince, "Commodore/C64/250407", draftCreated)!.SubmissionId);

        // A marker that does not say when it was made leaves nothing out.
        Assert.Equal(8, SubmissionReceiptPresenter.LatestForSystem(receipts, "Commodore/C64/250407", draftCreatedUtc: null)!.SubmissionId);
    }

    [Fact]
    public void A_system_never_submitted_has_no_badge()
    {
        Assert.Null(SubmissionReceiptPresenter.LatestForSystem([], "Commodore/C64/250407"));
        Assert.Null(SubmissionReceiptPresenter.LatestForSystem(null, "Commodore/C64/250407"));
        Assert.Null(SubmissionReceiptPresenter.LatestForSystem([Receipt(1, "Commodore/C64/250407", "2026-09-27T10:00:00Z")], "  "));
    }

    // Two sent in the same instant (a clock that did not move): the higher id is the later one.
    [Fact]
    public void Two_sent_at_once_are_told_apart_by_their_id()
    {
        Assert.Equal(7, SubmissionReceiptPresenter.LatestForSystem(
            [Receipt(7, "A/B/C", "2026-09-27T10:00:00Z"), Receipt(6, "A/B/C", "2026-09-27T10:00:00Z")],
            "A/B/C")!.SubmissionId);
    }

    // The tooltip carries the app's one date format - month capitalised, whatever the wording.
    [Fact]
    public void The_badges_tooltip_names_the_day_it_was_sent_in_the_one_date_format()
    {
        SubmissionReceipt receipt = Receipt(8, "A/B/C", "2026-09-27T10:46:11Z");

        string tooltip = SubmissionReceiptPresenter.DescribeLastSubmission(receipt);

        Assert.Contains(SubmissionReceiptPresenter.FormatDate(receipt.SentUtc), tooltip, StringComparison.Ordinal);
        Assert.Contains("September", tooltip, StringComparison.Ordinal);
        Assert.Contains("My submissions", tooltip, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ system names

    // ###########################################################################################
    // A real receipt carries the SYSTEM ID ("Commodore/C128/310378 Open128"), and "My submissions"
    // used to drop its last segment as a file name - reading "Commodore C128", board missing
    // (2026-09-27). A workbook path loses only its file.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore/C128/310378 Open128", "Commodore C128 310378 Open128")]
    [InlineData("Commodore/C64/250407/Data C64 250407.xlsx", "Commodore C64 250407")]
    [InlineData("  ", "(unknown system)")]
    public void A_system_is_named_in_full_whether_given_as_an_id_or_a_workbook_path(string systemId, string expected)
    {
        Assert.Equal(expected, SubmissionReceiptPresenter.DescribeSystem(systemId));
    }

    // ------------------------------------------------------------------ the switch-back-from-BETA notice

    // ###########################################################################################
    // "There should be some kind of notification for the user, when/if this gets promoted from BETA
    // to production, so the user can know he should now change the source from BETA to the normal
    // stable source" (owner, 2026-09-27). Only published submissions, only undismissed ones, and
    // only for someone actually downloading BETA.
    // ###########################################################################################
    private static SubmissionReceipt Sent(long id, string state, bool dismissed = false) => new()
    {
        SubmissionId = id,
        SystemId = "Commodore/C128/310378 Open128",
        LastKnownState = state,
        SentUtc = new DateTimeOffset(2026, 9, 27, 10, 0, 0, TimeSpan.Zero).AddMinutes(id),
        SourceNoticeDismissed = dismissed,
    };

    [Fact]
    public void A_submission_reaching_production_is_noticed_while_downloading_from_BETA()
    {
        Assert.Equal(
            [8L],
            SubmissionReceiptPresenter.NeedingSourceSwitchNotice([Sent(8, "published")], downloadingFromBeta: true)
                .Select(receipt => receipt.SubmissionId));
    }

    // In BETA only, it is not yet in the stable source - switching back would lose sight of it.
    [Theory]
    [InlineData("merged")]
    [InlineData("pending")]
    [InlineData("approved")]
    [InlineData("returned")]
    public void Nothing_short_of_production_is_noticed(string state)
    {
        Assert.Empty(SubmissionReceiptPresenter.NeedingSourceSwitchNotice([Sent(8, state)], downloadingFromBeta: true));
    }

    [Fact]
    public void Someone_on_the_normal_source_has_nothing_to_switch()
    {
        Assert.Empty(SubmissionReceiptPresenter.NeedingSourceSwitchNotice([Sent(8, "published")], downloadingFromBeta: false));
    }

    [Fact]
    public void A_dismissed_notice_is_not_shown_again()
    {
        Assert.Empty(SubmissionReceiptPresenter.NeedingSourceSwitchNotice([Sent(8, "published", dismissed: true)], downloadingFromBeta: true));
        Assert.Empty(SubmissionReceiptPresenter.NeedingSourceSwitchNotice(null, downloadingFromBeta: true));
    }

    [Fact]
    public void The_notice_names_the_whole_system_and_says_where_to_switch()
    {
        string text = SubmissionReceiptPresenter.DescribeSourceSwitchNotice([Sent(8, "published")]);

        Assert.StartsWith("Your submission for Commodore C128 310378 Open128 is now published to the stable source.", text, StringComparison.Ordinal);
        Assert.Contains("BETA source", text, StringComparison.Ordinal);
        Assert.Contains("Configuration tab", text, StringComparison.Ordinal);
    }

    // Two submissions of one system are one system; two systems are both named.
    [Fact]
    public void Several_submissions_name_each_system_once()
    {
        SubmissionReceipt other = new()
        {
            SubmissionId = 9,
            SystemId = "Commodore/C64/250407",
            LastKnownState = "published",
        };

        Assert.Contains(
            "Your submission for Commodore C128 310378 Open128 is",
            SubmissionReceiptPresenter.DescribeSourceSwitchNotice([Sent(7, "published"), Sent(8, "published")]),
            StringComparison.Ordinal);

        Assert.Contains(
            "Your submissions for Commodore C128 310378 Open128 and Commodore C64 250407 are",
            SubmissionReceiptPresenter.DescribeSourceSwitchNotice([Sent(8, "published"), other]),
            StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ Try it in BETA (2026-10-03)

    // ###########################################################################################
    // *** THE "TRY IT IN BETA" NOTICE (owner request, 2026-10-03). *** A submission accepted into the
    // BETA source, while this machine downloads from the STABLE source and the notice was not closed
    // for it - the twin of the switch-back notice above, for the step before it.
    // ###########################################################################################
    private static SubmissionReceipt InBeta(long id, string state = "merged", bool dismissed = false, string system = "Commodore/C128/310378 Open128") => new()
    {
        SubmissionId = id,
        SystemId = system,
        LastKnownState = state,
        SentUtc = new DateTimeOffset(2026, 10, 3, 10, 0, 0, TimeSpan.Zero).AddMinutes(id),
        BetaNoticeDismissed = dismissed,
    };

    [Fact]
    public void A_submission_in_BETA_is_noticed_while_downloading_from_the_stable_source()
    {
        Assert.Equal(
            [8L],
            SubmissionReceiptPresenter.NeedingBetaTryNotice([InBeta(8), InBeta(9, state: "pending")], downloadingFromBeta: false)
                .Select(receipt => receipt.SubmissionId));
    }

    // Already downloading from BETA: nothing to tick - the data comes with the next data check.
    // Closed for it: not again. Moved on (published, taken back out, rejected): gone.
    [Theory]
    [InlineData("merged", true, false)]
    [InlineData("merged", false, true)]
    [InlineData("published", false, false)]
    [InlineData("returned", false, false)]
    [InlineData("rejected", false, false)]
    public void The_BETA_notice_is_not_shown_on_BETA_once_closed_or_once_the_state_moved_on(string state, bool onBeta, bool dismissed)
    {
        Assert.Empty(SubmissionReceiptPresenter.NeedingBetaTryNotice([InBeta(8, state, dismissed)], onBeta));
    }

    [Fact]
    public void The_BETA_notice_names_the_system_and_the_check_box_as_CRT_writes_it()
    {
        Assert.Equal(
            "Your submission for Commodore C128 310378 Open128 is now in the BETA source. To try it before everyone else, " +
            "tick \"Download data from the BETA source instead of the stable source\" on the Configuration tab.",
            SubmissionReceiptPresenter.DescribeBetaTryNotice([InBeta(8)], checkDataOnLaunch: true));

        Assert.StartsWith(
            "Your submissions for Commodore C128 310378 Open128 and Commodore C64 250407 are now in the BETA source.",
            SubmissionReceiptPresenter.DescribeBetaTryNotice([InBeta(8), InBeta(9, system: "Commodore/C64/250407")], checkDataOnLaunch: true),
            StringComparison.Ordinal);
    }

    // With "Check for new or updated data at application launch" off, the BETA check box is greyed
    // out - so the notice names that one first.
    [Fact]
    public void With_data_checks_off_the_BETA_notice_says_to_tick_that_box_first()
    {
        Assert.Contains(
            "tick \"Check for new or updated data at application launch\" and then \"Download data from the BETA source instead of the stable source\"",
            SubmissionReceiptPresenter.DescribeBetaTryNotice([InBeta(8)], checkDataOnLaunch: false),
            StringComparison.Ordinal);
    }

    // The line under the submission in "My submissions", while it is in BETA and only then.
    [Fact]
    public void My_submissions_says_how_to_try_a_submission_while_it_is_in_BETA()
    {
        Assert.Equal(SubmissionReceiptPresenter.BetaTryLine, SubmissionReceiptPresenter.DescribeBetaTry(" Merged "));
        Assert.Equal(string.Empty, SubmissionReceiptPresenter.DescribeBetaTry("published"));
        Assert.Equal(string.Empty, SubmissionReceiptPresenter.DescribeBetaTry(null));
    }

    // ------------------------------------------------------------------ Already sent as it is (2026-10-03)

    // ###########################################################################################
    // *** THE SAME DRAFT CANNOT BE SENT TWICE (owner request, 2026-10-03: "It should not be
    // possible to submit the same data again"; cases agreed with the project owner). *** A receipt
    // remembers the draft's fingerprint at sending; while the draft still gives it, Submit is
    // greyed out - unless that submission never finished sending. Each test that says "not
    // blocked" also shows the blocked side, so a rule that never blocks cannot pass it.
    // ###########################################################################################
    private const string C64 = "Commodore/C64/250407";

    private static SubmissionReceipt SentAs(string fingerprint, string state, string sentUtc = "2026-10-03T10:00:00Z", long id = 1) => new()
    {
        SubmissionId = id,
        SystemId = C64,
        LastKnownState = state,
        DraftFingerprint = fingerprint,
        SentUtc = DateTimeOffset.Parse(sentUtc, CultureInfo.InvariantCulture),
    };

    // Case 1, its words.
    [Fact]
    public void The_reason_says_when_it_was_sent_and_that_something_must_change_first()
    {
        Assert.Equal(
            "You have already sent this draft as it is now, on 2026-October-3. Change something to send it again.",
            SubmissionReceiptPresenter.DescribeAlreadySent(SentAs("v1:sent", "pending")));
    }

    // Case 6.
    [Theory]
    [InlineData("pending")]
    [InlineData("approved")]
    [InlineData("merged")]
    [InlineData("returned")]
    [InlineData("changes_requested")]
    [InlineData("rejected")]
    public void A_draft_sent_as_it_is_now_cannot_be_sent_again_whatever_has_happened_to_it(string state)
    {
        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(SentAs("v1:sent", state), "v1:sent"));
    }

    // Case 7: cancelled, the connection lost, or never confirmed by the server ("" - the receipt
    // is written before the upload, and the state only after the server confirms it).
    [Theory]
    [InlineData("")]
    [InlineData("uploading")]
    [InlineData("abandoned")]
    public void A_submission_that_never_finished_sending_does_not_stop_the_same_draft_being_sent(string state)
    {
        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(SentAs("v1:sent", state), "v1:sent"));
        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(SentAs("v1:sent", "pending"), "v1:sent"));
    }

    // Case 8.
    [Fact]
    public void A_submission_sent_by_an_older_CRT_never_stops_a_draft_being_sent()
    {
        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(SentAs(string.Empty, "pending"), "v1:sent"));
        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(SentAs("v1:sent", "pending"), "v1:sent"));
    }

    // Case 9: send A, change, send B, change back to exactly A - A's fingerprint, but B is the
    // latest, and the draft differs from B.
    [Fact]
    public void Only_the_latest_submission_counts_so_a_draft_changed_back_to_an_earlier_one_can_be_sent()
    {
        SubmissionReceipt[] receipts =
        [
            SentAs("v1:A", "withdrawn", "2026-10-03T10:00:00Z", id: 1),
            SentAs("v1:B", "pending", "2026-10-03T11:00:00Z", id: 2),
        ];

        SubmissionReceipt? latest = SubmissionReceiptPresenter.LatestForSystem(receipts, C64);

        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(latest, "v1:A"));
        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(latest, "v1:B"));
    }

    // Case 10: the draft was discarded and a new one of the same board started - only what was
    // sent FROM this draft counts, as for the badge.
    [Fact]
    public void A_submission_sent_before_this_draft_was_started_does_not_stop_it_being_sent()
    {
        SubmissionReceipt[] receipts = [SentAs("v1:sent", "pending", "2026-09-01T10:00:00Z")];
        var draftCreated = DateTimeOffset.Parse("2026-10-01T10:00:00Z", CultureInfo.InvariantCulture);

        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(
            SubmissionReceiptPresenter.LatestForSystem(receipts, C64, draftCreated), "v1:sent"));

        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(
            SubmissionReceiptPresenter.LatestForSystem(receipts, C64), "v1:sent"));
    }

    // Not a case of its own: a draft whose fingerprint could not be worked out (unreadable) is
    // never blocked - an unknown answer must not stop a contributor sending.
    [Fact]
    public void A_draft_whose_fingerprint_is_unknown_is_never_blocked()
    {
        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(SentAs(string.Empty, "pending"), string.Empty));
        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(null, "v1:sent"));
        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(SentAs("v1:sent", "pending"), "v1:sent"));
    }

    // ###########################################################################################
    // *** A SUBMISSION THE SERVER NO LONGER KNOWS (2026-10-04). *** Deleted with its system, or by
    // the reset of the contribution data at go-live. The receipt keeps its last state, which used to
    // be shown for ever ("Submitted - awaiting feedback ..." about a submission nobody will see) and
    // kept the same draft from being sent again. Each test shows the known side too, so a rule that
    // ignores the mark cannot pass.
    // ###########################################################################################
    private static SubmissionReceipt Gone(string state, string fingerprint = "v1:sent") =>
        SentAs(fingerprint, state) with { NotFoundUtc = DateTimeOffset.Parse("2026-10-04T08:00:00Z", CultureInfo.InvariantCulture) };

    [Theory]
    [InlineData("pending")]
    [InlineData("merged")]
    [InlineData("changes_requested")]
    public void A_submission_the_server_no_longer_knows_says_so_in_the_neutral_colour(string state)
    {
        Assert.Equal(SubmissionReceiptPresenter.NotOnServerWording, SubmissionReceiptPresenter.DescribeReceiptState(Gone(state)));
        Assert.Equal(SubmissionOutcomeKind.Waiting, SubmissionReceiptPresenter.ClassifyReceipt(Gone(state)));

        Assert.Equal(SubmissionReceiptPresenter.DescribeState(state), SubmissionReceiptPresenter.DescribeReceiptState(SentAs("v1:sent", state)));
        Assert.Equal(SubmissionReceiptPresenter.ClassifyState(state), SubmissionReceiptPresenter.ClassifyReceipt(SentAs("v1:sent", state)));
    }

    // What was sent is gone, so the same draft is exactly what should be sent again.
    [Fact]
    public void A_submission_the_server_no_longer_knows_does_not_stop_the_same_draft_being_sent()
    {
        Assert.False(SubmissionReceiptPresenter.IsAlreadySent(Gone("pending"), "v1:sent"));
        Assert.True(SubmissionReceiptPresenter.IsAlreadySent(SentAs("v1:sent", "pending"), "v1:sent"));
    }

    // Nor is it in the BETA source any more: no "try it in BETA" line, and no notice.
    [Fact]
    public void A_submission_the_server_no_longer_knows_is_not_offered_to_try_in_BETA()
    {
        Assert.Equal(string.Empty, SubmissionReceiptPresenter.DescribeReceiptBetaTry(Gone("merged")));
        Assert.Equal(SubmissionReceiptPresenter.BetaTryLine, SubmissionReceiptPresenter.DescribeReceiptBetaTry(SentAs("v1:sent", "merged")));

        Assert.Empty(SubmissionReceiptPresenter.NeedingBetaTryNotice([Gone("merged")], downloadingFromBeta: false));
        Assert.Single(SubmissionReceiptPresenter.NeedingBetaTryNotice([SentAs("v1:sent", "merged")], downloadingFromBeta: false));
    }
}
