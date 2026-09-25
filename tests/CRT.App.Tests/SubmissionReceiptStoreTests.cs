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
        Assert.Equal("Looks good, one question about C7.", stored.ReviewerComment);
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
    // ACKNOWLEDGING REVIEWER FEEDBACK - what clears the badge on the Drafts tab.
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
            SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 42).ReviewerComment);
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
    // "A reviewer changed it" (2026-09-25) is cached like the state: set when the server says so,
    // and never lost by a refresh that does not say, or by marking a comment read.
    // ###########################################################################################
    [Fact]
    public void The_reviewer_changed_flag_is_stored_and_kept_by_a_refresh_that_does_not_say()
    {
        SubmissionReceiptStore.Record(Receipt(42));

        SubmissionReceiptStore.UpdateState(42, "merged", "Thanks.", DateTimeOffset.UtcNow, amendedByReviewer: true);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByReviewer);

        SubmissionReceiptStore.UpdateState(42, "merged", "Thanks.", DateTimeOffset.UtcNow);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByReviewer);

        SubmissionReceiptStore.AcknowledgeComment(42);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByReviewer);

        // And it survives the file.
        SubmissionReceiptStore.LoadFrom(this.thisPath);
        Assert.True(Assert.Single(SubmissionReceiptStore.All).AmendedByReviewer);
    }
}
