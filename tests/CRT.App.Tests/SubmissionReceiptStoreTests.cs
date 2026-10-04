using System;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SubmissionReceiptStore - where a contributor's submission receipts live (Phase 4, task 6).
//
// DRIVES THE REAL STATIC STATE through LoadFrom, its own test seam, exactly as UserSettingsTests
// and DataManagerTests do for theirs. NEVER call SubmissionReceiptStore.Load() from a test: it
// resolves the user's real AppData folder and would write receipts into it.
//
// The class holds static state, so this class joins its own collection to stay off the same
// thread as anything else touching it. Collections never run in parallel (xunit.runner.json), so
// that is belt and braces rather than strictly required - but the convention in CLAUDE.md is that
// any test class touching a static seam declares which one, and this is that declaration.
// ###########################################################################################
[Collection("SubmissionReceipts")]
public sealed class SubmissionReceiptStoreTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();
    private readonly string thisPath;

    public SubmissionReceiptStoreTests()
    {
        this.thisPath = Path.Combine(this.thisWorkspace.Root, "submissions.json");
        SubmissionReceiptStore.LoadFrom(this.thisPath);
    }

    public void Dispose()
    {
        // Point the store at nothing so a temp path never outlives this class.
        SubmissionReceiptStore.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static SubmissionReceipt Receipt(long id, string token = "tok", string systemId = "Commodore/C64/250407/Data.xlsx")
        => new()
        {
            SubmissionId = id,
            UploadToken = token,
            SystemId = systemId,
            Summary = "Fixed U8 pinout",
            SentUtc = new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero)
        };

    [Fact]
    public void A_recorded_receipt_survives_a_reload_from_disk()
    {
        SubmissionReceiptStore.Record(Receipt(42));

        // A fresh load from the same file is what the next application launch does.
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);

        Assert.Equal(42, stored.SubmissionId);
        Assert.Equal("tok", stored.UploadToken);
        Assert.Equal("Fixed U8 pinout", stored.Summary);
    }

    // ###########################################################################################
    // *** A RECEIPTS FILE FROM BEFORE THE RENAME KEEPS ITS MAINTAINER'S WORDS (code review,
    // 2026-09-25). *** The role was "reviewer" until then, and an earlier build wrote
    // ReviewerComment and AmendedByReviewer. A DECIDED submission is never asked about again, so a
    // comment lost on load was gone from "My submissions" for good.
    // ###########################################################################################
    [Fact]
    public void A_receipt_written_with_the_old_reviewer_names_loads_its_comment_and_amendment()
    {
        File.WriteAllText(this.thisPath, """
            [
              {
                "SubmissionId": 42,
                "UploadToken": "tok",
                "SystemId": "Commodore/C64/250407/Data.xlsx",
                "Summary": "Fixed U8 pinout",
                "LastKnownState": "rejected",
                "ReviewerComment": "Wrong pin on U8",
                "AmendedByReviewer": true
              }
            ]
            """);

        SubmissionReceiptStore.LoadFrom(this.thisPath);

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);
        Assert.Equal("Wrong pin on U8", stored.MaintainerComment);
        Assert.True(stored.AmendedByMaintainer);
    }

    // ...and the next save writes only the new names, so the old ones leave the file for good.
    [Fact]
    public void Saving_again_writes_only_the_new_maintainer_names()
    {
        File.WriteAllText(this.thisPath, """
            [ { "SubmissionId": 42, "UploadToken": "tok", "ReviewerComment": "Wrong pin on U8", "AmendedByReviewer": true } ]
            """);

        SubmissionReceiptStore.LoadFrom(this.thisPath);
        SubmissionReceiptStore.Record(Receipt(43));

        string written = File.ReadAllText(this.thisPath);
        Assert.DoesNotContain("Reviewer", written, StringComparison.Ordinal);
        Assert.Contains("\"MaintainerComment\": \"Wrong pin on U8\"", written, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // THE TOKEN IS THE WHOLE POINT OF THE FILE, so a receipt without one is refused.
    //
    // With no account, the capability token is the only thing that can prove this machine owns a
    // submission. A tokenless row could never be used to ask the server anything - it would sit in
    // the list failing silently on every refresh, looking like a server problem.
    // ###########################################################################################
    [Fact]
    public void A_receipt_with_no_token_is_refused_rather_than_stored_as_a_dead_row()
    {
        SubmissionReceiptStore.Record(Receipt(42, token: string.Empty));

        Assert.Empty(SubmissionReceiptStore.All);
    }

    // The same guard applies on the way IN from disk, not just on the way out - a hand-edited or
    // truncated file can carry a row the application never wrote.
    [Fact]
    public void A_tokenless_row_already_on_disk_is_dropped_on_load()
    {
        File.WriteAllText(this.thisPath, """
            [
              { "SubmissionId": 1, "UploadToken": "good", "SystemId": "A/B/C/D.xlsx" },
              { "SubmissionId": 2, "UploadToken": "",     "SystemId": "A/B/C/D.xlsx" }
            ]
            """);

        SubmissionReceiptStore.LoadFrom(this.thisPath);

        Assert.Equal(new long[] { 1 }, SubmissionReceiptStore.All.Select(receipt => receipt.SubmissionId));
    }

    // ###########################################################################################
    // Re-recording the same id REPLACES rather than appends. The id is the server's own, so two
    // rows carrying it could only ever disagree about the same submission.
    // ###########################################################################################
    [Fact]
    public void Recording_the_same_submission_twice_leaves_one_row()
    {
        SubmissionReceiptStore.Record(Receipt(42, token: "first"));
        SubmissionReceiptStore.Record(Receipt(42, token: "second"));

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);
        Assert.Equal("second", stored.UploadToken);
    }

    [Fact]
    public void Updating_a_state_keeps_the_token_and_everything_else_the_server_does_not_own()
    {
        SubmissionReceiptStore.Record(Receipt(42));

        var checkedAt = new DateTimeOffset(2026, 9, 22, 9, 0, 0, TimeSpan.Zero);
        SubmissionReceiptStore.UpdateState(42, "pending", "Looks good, one question about C7.", checkedAt);

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);

        Assert.Equal("pending", stored.LastKnownState);
        Assert.Equal("Looks good, one question about C7.", stored.MaintainerComment);
        Assert.Equal(checkedAt, stored.LastCheckedUtc);

        // The parts the server's reply says nothing about must survive the write-back. Losing the
        // token here would silently make the row unusable from the next refresh onwards.
        Assert.Equal("tok", stored.UploadToken);
        Assert.Equal("Fixed U8 pinout", stored.Summary);
        Assert.Equal("Commodore/C64/250407/Data.xlsx", stored.SystemId);
    }

    // A reply arriving for a receipt the user removed meanwhile must not put it back. The refresh
    // loop reads a snapshot, so this ordering is reachable in the shipped app rather than
    // theoretical.
    [Fact]
    public void A_state_update_for_an_unknown_submission_does_not_resurrect_it()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.Forget(42);

        SubmissionReceiptStore.UpdateState(42, "pending", string.Empty, DateTimeOffset.UtcNow);

        Assert.Empty(SubmissionReceiptStore.All);
    }

    [Fact]
    public void Forgetting_a_receipt_removes_it_from_disk_too()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.Record(Receipt(43));

        SubmissionReceiptStore.Forget(42);
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        Assert.Equal(new long[] { 43 }, SubmissionReceiptStore.All.Select(receipt => receipt.SubmissionId));
    }

    // ###########################################################################################
    // A BROKEN FILE COSTS THE LIST, NEVER THE APPLICATION.
    //
    // This file is read during startup. An unparseable one must degrade to an empty list rather
    // than throwing: the worst outcome of a corrupt receipts file is that a contributor cannot
    // check status in-app, and it must never be that CRT will not start or cannot submit again.
    // ###########################################################################################
    [Fact]
    public void An_unparseable_file_loads_as_empty_rather_than_throwing()
    {
        File.WriteAllText(this.thisPath, "{ this is not the json you are looking for");

        SubmissionReceiptStore.LoadFrom(this.thisPath);

        Assert.Empty(SubmissionReceiptStore.All);

        // And the store still works afterwards - a bad file must not poison the session.
        SubmissionReceiptStore.Record(Receipt(1));
        Assert.Single(SubmissionReceiptStore.All);
    }

    [Fact]
    public void A_missing_file_loads_as_empty()
    {
        SubmissionReceiptStore.LoadFrom(Path.Combine(this.thisWorkspace.Root, "never-written.json"));

        Assert.Empty(SubmissionReceiptStore.All);
    }

    // Newest first, through the store's own accessor - the view binds straight to this, so the
    // ordering has to hold here and not only in the presenter.
    [Fact]
    public void The_list_comes_back_newest_first()
    {
        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 1,
            UploadToken = "tok",
            SentUtc = new DateTimeOffset(2026, 9, 1, 0, 0, 0, TimeSpan.Zero)
        });

        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 2,
            UploadToken = "tok",
            SentUtc = new DateTimeOffset(2026, 9, 20, 0, 0, 0, TimeSpan.Zero)
        });

        Assert.Equal(
            new long[] { 2, 1 },
            SubmissionReceiptStore.All.Select(receipt => receipt.SubmissionId));
    }

    // With no path configured, the store still behaves - it simply persists nothing. This is the
    // state a test leaves behind, and a Record() that threw would turn a cleanup detail into a
    // cascade of unrelated failures.
    [Fact]
    public void With_no_path_configured_recording_persists_nothing_and_does_not_throw()
    {
        SubmissionReceiptStore.LoadFrom(string.Empty);

        SubmissionReceiptStore.Record(Receipt(42));

        Assert.Single(SubmissionReceiptStore.All);
    }

    // -----------------------------------------------------------------------------------
    // ACKNOWLEDGING MAINTAINER FEEDBACK - what clears the badge on the Drafts tab.
    // -----------------------------------------------------------------------------------

    [Fact]
    public void Acknowledging_a_comment_marks_it_read_and_SURVIVES_a_reload()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "changes_requested", "Check the U8 highlight.", DateTimeOffset.UtcNow);

        Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());

        SubmissionReceiptStore.AcknowledgeComment(42);

        Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());

        // Re-read from disk: dismissing must outlive the session, or the badge returns on the
        // next launch for feedback already dealt with.
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
    }

    [Fact]
    public void A_REFRESH_does_not_make_an_acknowledged_comment_unread_again()
    {
        // *** THE BUG THIS GUARD EXISTS FOR. *** Every refresh re-reads the SAME comment from the
        // server, so an UpdateState that reset the acknowledgement would resurrect the badge on
        // every check - and a badge that reappears for no visible reason is one the contributor
        // learns to ignore, which defeats the whole feature.
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "changes_requested", "Check the U8 highlight.", DateTimeOffset.UtcNow);
        SubmissionReceiptStore.AcknowledgeComment(42);

        SubmissionReceiptStore.UpdateState(42, "changes_requested", "Check the U8 highlight.", DateTimeOffset.UtcNow);

        Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
    }

    [Fact]
    public void A_CHANGED_comment_becomes_unread_again_without_anything_clearing_a_flag()
    {
        // A second round of review feedback. Nothing resets the acknowledgement - it is simply no
        // longer equal to the current comment, so the badge returns by construction.
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "changes_requested", "Check the U8 highlight.", DateTimeOffset.UtcNow);
        SubmissionReceiptStore.AcknowledgeComment(42);

        SubmissionReceiptStore.UpdateState(42, "changes_requested", "The region is wrong too.", DateTimeOffset.UtcNow);

        Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());
    }

    [Fact]
    public void Acknowledging_keeps_the_COMMENT_ITSELF_on_the_record()
    {
        // Marking feedback read must never hide what was said - the contributor is very likely
        // about to act on it.
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "changes_requested", "Check the U8 highlight.", DateTimeOffset.UtcNow);
        SubmissionReceiptStore.AcknowledgeComment(42);

        Assert.Equal(
            "Check the U8 highlight.",
            SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 42).MaintainerComment);
    }

    [Fact]
    public void Acknowledging_an_UNKNOWN_submission_does_nothing_rather_than_throwing()
    {
        // The receipt may have been forgotten from another surface while this window was open.
        SubmissionReceiptStore.AcknowledgeComment(999);

        Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
    }

    [Fact]
    public void Acknowledging_a_receipt_with_NO_comment_leaves_it_alone()
    {
        // Nothing to acknowledge. Writing an empty acknowledgement would be a pointless save, and
        // would also mean a later real comment compared against a value that was never shown.
        SubmissionReceiptStore.Record(Receipt(42));

        SubmissionReceiptStore.AcknowledgeComment(42);

        SubmissionReceiptStore.UpdateState(42, "changes_requested", "A first note.", DateTimeOffset.UtcNow);

        Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());
    }

    [Fact]
    public void The_DECIDED_DATE_survives_marking_the_comment_as_read()
    {
        // *** AcknowledgeComment REBUILDS THE WHOLE RECORD, so a field left out is a field ERASED.
        // *** Without carrying it, the "Replied 22 September 2026" line would vanish the instant
        // the contributor pressed "Mark as read" - losing exactly the context that tells them
        // whether the feedback predates their latest edits.
        var decided = new DateTimeOffset(2026, 9, 22, 9, 15, 0, TimeSpan.Zero);

        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(
            42, "changes_requested", "Check the U8 highlight.", DateTimeOffset.UtcNow, decided);

        SubmissionReceiptStore.AcknowledgeComment(42);

        Assert.Equal(decided, SubmissionReceiptStore.All.Single().DecidedUtc);
    }

    [Fact]
    public void A_refresh_that_reports_NO_decision_date_does_not_erase_one_already_known()
    {
        // The argument is optional and trailing, so an older caller passes nothing. "The server did
        // not tell us this time" must never wipe a date it told us last time.
        var decided = new DateTimeOffset(2026, 9, 22, 9, 15, 0, TimeSpan.Zero);

        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(
            42, "changes_requested", "Check U8.", DateTimeOffset.UtcNow, decided);

        SubmissionReceiptStore.UpdateState(
            42, "changes_requested", "Check U8.", DateTimeOffset.UtcNow);

        Assert.Equal(decided, SubmissionReceiptStore.All.Single().DecidedUtc);
    }
    // ###########################################################################################
    // "A maintainer changed it" (2026-09-25) is cached like the state: set when the server says so,
    // and never lost by a refresh that does not say, or by marking a comment read.
    // ###########################################################################################
    [Fact]
    public void The_maintainer_changed_flag_is_stored_and_kept_by_a_refresh_that_does_not_say()
    {
        SubmissionReceiptStore.Record(Receipt(42));

        SubmissionReceiptStore.UpdateState(42, "merged", "Thanks.", DateTimeOffset.UtcNow, amendedByMaintainer: true);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByMaintainer);

        SubmissionReceiptStore.UpdateState(42, "merged", "Thanks.", DateTimeOffset.UtcNow);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByMaintainer);

        SubmissionReceiptStore.AcknowledgeComment(42);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByMaintainer);

        // And it survives the file.
        SubmissionReceiptStore.LoadFrom(this.thisPath);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByMaintainer);
    }

    // ###########################################################################################
    // The "now in the online source - switch back from BETA" notice (2026-09-27). Closing it is
    // remembered per submission, on disk, and nothing that later rewrites the receipt brings it back.
    // ###########################################################################################
    [Fact]
    public void A_dismissed_source_notice_stays_dismissed_across_a_reload()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "published", string.Empty, DateTimeOffset.UtcNow);

        SubmissionReceiptStore.DismissSourceNotice([42]);
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        Assert.True(Assert.Single(SubmissionReceiptStore.All).SourceNoticeDismissed);
    }

    // Every method that rebuilds a receipt must carry the flag, or the next status check or a
    // comment being read would quietly bring the notice back.
    [Fact]
    public void A_later_state_check_or_reading_a_comment_does_not_bring_the_notice_back()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "published", "Thanks!", DateTimeOffset.UtcNow);
        SubmissionReceiptStore.DismissSourceNotice([42]);

        SubmissionReceiptStore.UpdateState(42, "published", "Thanks!", DateTimeOffset.UtcNow);
        SubmissionReceiptStore.AcknowledgeComment(42);

        Assert.True(Assert.Single(SubmissionReceiptStore.All).SourceNoticeDismissed);
    }

    // The "now in the BETA source - tick BETA to try it" notice (2026-10-03): the same, apart from
    // the switch-back notice - closing one does not close the other.
    [Fact]
    public void A_dismissed_BETA_notice_stays_dismissed_and_leaves_the_switch_back_notice_alone()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        SubmissionReceiptStore.UpdateState(42, "merged", string.Empty, DateTimeOffset.UtcNow);

        SubmissionReceiptStore.DismissBetaNotice([42]);
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        SubmissionReceipt receipt = Assert.Single(SubmissionReceiptStore.All);
        Assert.True(receipt.BetaNoticeDismissed);
        Assert.False(receipt.SourceNoticeDismissed);
    }

    [Fact]
    public void Dismissing_touches_only_the_submissions_named()
    {
        SubmissionReceiptStore.Record(Receipt(42, token: "a"));
        SubmissionReceiptStore.Record(Receipt(43, token: "b"));

        SubmissionReceiptStore.DismissSourceNotice([43]);

        Assert.False(SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 42).SourceNoticeDismissed);
        Assert.True(SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 43).SourceNoticeDismissed);
    }

    // ###########################################################################################
    // *** THE STATE THE SERVER CONFIRMED AT FINALISE (owner report, 2026-09-27). *** The receipt is
    // written before the upload, so without this it knew no state until the next launch's check,
    // and the Drafts tab's badge read "Not checked yet" straight after a successful send.
    // ###########################################################################################
    [Fact]
    public void A_finalised_submission_records_the_state_the_server_confirmed()
    {
        SubmissionReceiptStore.Record(Receipt(42));
        var checkedUtc = new DateTimeOffset(2026, 9, 27, 11, 0, 0, TimeSpan.Zero);

        SubmissionReceiptStore.RecordFinalised(
            new SubmissionResult { SubmissionId = 42, IsAccepted = true, State = "pending" }, checkedUtc);

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);
        Assert.Equal("pending", stored.LastKnownState);
        Assert.Equal(checkedUtc, stored.LastCheckedUtc);
        Assert.Equal("Submitted - awaiting feedback from a maintainer", SubmissionReceiptPresenter.DescribeState(stored.LastKnownState));
        Assert.Equal("tok", stored.UploadToken);

        // Still asked about at the next launch, and not news in the meantime.
        Assert.True(SubmissionReceiptPresenter.IsStillOpen(stored, checkedUtc));
        Assert.False(SubmissionReceiptPresenter.HasUnreadDecision(stored));
    }

    // The automatic checks refused it: recorded, and as SEEN - the send window has just shown why.
    [Fact]
    public void A_submission_refused_at_finalise_is_recorded_as_already_seen()
    {
        SubmissionReceiptStore.Record(Receipt(42));

        SubmissionReceiptStore.RecordFinalised(
            new SubmissionResult { SubmissionId = 42, IsAccepted = false, State = "rejected" }, DateTimeOffset.UtcNow);

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);
        Assert.Equal("rejected", stored.LastKnownState);
        Assert.False(SubmissionReceiptPresenter.HasUnreadDecision(stored));
        Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
    }

    // An answer that names no state (an older server) must not stamp "checked" onto an unknown one.
    [Fact]
    public void A_finalise_answer_with_no_state_changes_nothing()
    {
        SubmissionReceiptStore.Record(Receipt(42));

        SubmissionReceiptStore.RecordFinalised(
            new SubmissionResult { SubmissionId = 42, IsAccepted = true, State = "  " }, DateTimeOffset.UtcNow);
        SubmissionReceiptStore.RecordFinalised(null, DateTimeOffset.UtcNow);

        SubmissionReceipt stored = Assert.Single(SubmissionReceiptStore.All);
        Assert.Equal(string.Empty, stored.LastKnownState);
        Assert.Null(stored.LastCheckedUtc);
    }

    // ###########################################################################################
    // *** A DISCARDED DRAFT IS REMEMBERED UNTIL THE SERVER HAS BEEN TOLD (owner request,
    // 2026-09-28). *** Marked on the receipt first, so a discard made with no network is reported
    // at the next launch; marked reported once an answer finishes it - and it survives a reload.
    // ###########################################################################################
    [Fact]
    public void A_discard_waits_until_reported_and_survives_a_reload()
    {
        DateTimeOffset now = new(2026, 9, 28, 9, 0, 0, TimeSpan.Zero);

        SubmissionReceiptStore.Record(Receipt(41));
        SubmissionReceiptStore.Record(Receipt(42));

        SubmissionReceiptStore.MarkDraftDiscarded([41], now);
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        SubmissionReceipt pending = Assert.Single(SubmissionReceiptStore.PendingDraftDiscardNotices);
        Assert.Equal(41, pending.SubmissionId);
        Assert.Equal(now, pending.DraftDiscardedUtc);

        SubmissionReceiptStore.MarkDraftDiscardReported(41);
        SubmissionReceiptStore.LoadFrom(this.thisPath);

        Assert.Empty(SubmissionReceiptStore.PendingDraftDiscardNotices);
        Assert.True(SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 41).DraftDiscardReported);
        Assert.Equal(now, SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 41).DraftDiscardedUtc);
    }

    // A second discard of the same receipt (a later draft of the board) keeps the first time.
    [Fact]
    public void A_receipt_already_marked_keeps_its_first_discard_time()
    {
        DateTimeOffset first = new(2026, 9, 28, 9, 0, 0, TimeSpan.Zero);

        SubmissionReceiptStore.Record(Receipt(41));
        SubmissionReceiptStore.MarkDraftDiscarded([41], first);
        SubmissionReceiptStore.MarkDraftDiscarded([41], first.AddDays(1));

        Assert.Equal(first, Assert.Single(SubmissionReceiptStore.All).DraftDiscardedUtc);
    }

    // ###########################################################################################
    // *** EVERY FIELD SURVIVES EVERY REWRITE. *** The store rebuilds a receipt field by field (a
    // class with init-only properties has no `with`), and a field left out of one rebuild is ERASED
    // by it - the header of UpdateState says so, and it has happened before. So each rewrite runs
    // over a receipt with EVERY property set, and every property it does not mean to change must
    // come back as it was. A property added later and forgotten in one rebuild fails here by name.
    // ###########################################################################################
    [Fact]
    public void Every_receipt_field_survives_every_rewrite()
    {
        DateTimeOffset at = new(2026, 9, 20, 8, 0, 0, TimeSpan.Zero);

        SubmissionReceipt Full() => new()
        {
            SubmissionId = 7,
            UploadToken = "token",
            SystemId = "Commodore/C128/310378",
            Summary = "summary",
            SentUtc = at,
            LastKnownState = "merged",
            LastCheckedUtc = at.AddHours(1),
            MaintainerComment = "comment",
            AcknowledgedComment = "old comment",
            AcknowledgedState = "pending",
            DecidedUtc = at.AddHours(2),
            AmendedByMaintainer = true,
            SourceNoticeDismissed = false,
            BetaNoticeDismissed = false,
            DraftDiscardedUtc = at.AddHours(3),
            DraftDiscardReported = false,
            DraftFingerprint = "v1:ABC"
        };

        // Each rewrite, and the properties it is MEANT to change.
        var rewrites = new (string Name, Action Rewrite, string[] Changes)[]
        {
            ("UpdateState", () => SubmissionReceiptStore.UpdateState(7, "merged", "comment", at.AddHours(1)), ["LastCheckedUtc"]),
            ("AcknowledgeComment", () => SubmissionReceiptStore.AcknowledgeComment(7), ["AcknowledgedComment", "AcknowledgedState"]),
            ("DismissSourceNotice", () => SubmissionReceiptStore.DismissSourceNotice([7]), ["SourceNoticeDismissed"]),
            ("DismissBetaNotice", () => SubmissionReceiptStore.DismissBetaNotice([7]), ["BetaNoticeDismissed"]),
            ("MarkDraftDiscardReported", () => SubmissionReceiptStore.MarkDraftDiscardReported(7), ["DraftDiscardReported"]),
        };

        foreach ((string name, Action rewrite, string[] changes) in rewrites)
        {
            SubmissionReceiptStore.LoadFrom(this.thisPath);
            foreach (SubmissionReceipt existing in SubmissionReceiptStore.All.ToList())
                SubmissionReceiptStore.Forget(existing.SubmissionId);

            SubmissionReceipt before = Full();
            SubmissionReceiptStore.Record(before);

            rewrite();

            SubmissionReceipt after = Assert.Single(SubmissionReceiptStore.All);

            foreach (System.Reflection.PropertyInfo property in typeof(SubmissionReceipt).GetProperties())
            {
                // The legacy names are read-only shims that answer null (see the receipt's header).
                if (!property.CanWrite || property.GetGetMethod() is null || changes.Contains(property.Name))
                    continue;

                Assert.True(
                    Equals(property.GetValue(before), property.GetValue(after)),
                    $"{name} lost {property.Name}: was [{property.GetValue(before)}], now [{property.GetValue(after)}]");
            }
        }
    }

    // ###########################################################################################
    // *** SAFE FROM ANY THREAD (code review, 2026-09-29). *** The Submit dialog records and
    // finalises its receipt from the thread pool (its upload runs under Task.Run), while the launch
    // status check, the Drafts tab and the discard reporter read and write the same list on the UI
    // thread. Unguarded, two writers at once lost receipts, threw "Collection was modified" in a
    // reader, or raced on the save's temporary file. Several threads recording and updating at once
    // must lose nothing, throw nothing, and leave a file that reads back whole.
    // ###########################################################################################
    //
    // Twelve receipts per thread, not forty (2026-10-03): every Record and UpdateState saves the
    // whole file atomically (~15 ms here), so forty made this one test 7.5 s of the run. Six
    // writers still overlap on nearly every save, which is what the test needs.
    [Fact]
    public async Task Receipts_written_from_several_threads_at_once_are_all_kept()
    {
        const int threads = 6;
        const int each = 12;

        Task[] writers = Enumerable.Range(0, threads).Select(thread => Task.Run(() =>
        {
            for (int n = 0; n < each; n++)
            {
                long id = (thread * 1000) + n;

                SubmissionReceiptStore.Record(Receipt(id));
                SubmissionReceiptStore.UpdateState(id, "pending", string.Empty, DateTimeOffset.UtcNow);

                // A reader in the middle of it all, as the UI thread is.
                _ = SubmissionReceiptStore.All.Count;
                _ = SubmissionReceiptStore.UnreadCommentCount();
            }
        })).ToArray();

        await Task.WhenAll(writers);

        Assert.Equal(threads * each, SubmissionReceiptStore.All.Count);
        Assert.All(SubmissionReceiptStore.All, receipt => Assert.Equal("pending", receipt.LastKnownState));

        SubmissionReceiptStore.LoadFrom(this.thisPath);
        Assert.Equal(threads * each, SubmissionReceiptStore.All.Count);
    }
}
