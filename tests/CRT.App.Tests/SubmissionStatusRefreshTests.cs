using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SubmissionStatusRefresh - the check that now runs AT LAUNCH (owner report, 2026-09-22).
//
// *** THE BUG THIS CLASS EXISTS FOR: the check only ever ran from a button. *** Receipts were
// loaded from disk at startup, but nothing asked the server until the contributor opened "My
// submissions" and pressed Refresh. So a maintainer could request changes and the contributor would
// launch CRT to a stale cache - no badge, no comment - until they happened to press a button they
// had no reason to think was necessary. Reported after a real review round trip.
//
// Drives the REAL static receipts store through LoadFrom, its own test seam, so this class joins
// the same collection SubmissionReceiptStoreTests uses. NEVER call SubmissionReceiptStore.Load()
// from a test: it resolves the user's real AppData folder.
//
// No network anywhere (CLAUDE.md test rule 6) - the lookup is a delegate, which is exactly why it
// is a delegate.
// ###########################################################################################
[Collection("SubmissionReceipts")]
public sealed class SubmissionStatusRefreshTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private static readonly DateTimeOffset Now = new(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

    public SubmissionStatusRefreshTests()
    {
        SubmissionReceiptStore.LoadFrom(Path.Combine(this.thisWorkspace.Root, "submissions.json"));
    }

    public void Dispose()
    {
        SubmissionReceiptStore.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static void Record(long id, string state)
    {
        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = id,
            UploadToken = "token",
            BoardId = "Commodore/C64/250407/Data.xlsx",
            SentUtc = SubmissionStatusRefreshTests.Now.AddDays(-1),
            LastKnownState = state
        });
    }

    private static SubmissionStatusRefresh.StatusLookup Answering(
        string state,
        string comment,
        List<long>? asked = null)
    {
        return (id, _, _) =>
        {
            asked?.Add(id);

            return Task.FromResult<SubmissionStatus?>(new SubmissionStatus
            {
                Id = id,
                State = state,
                MaintainerComment = comment
            });
        };
    }

    [Fact]
    public async Task A_maintainers_comment_arrives_WITHOUT_anybody_pressing_refresh()
    {
        // The whole point. Launching the app is when somebody expects to be told.
        SubmissionStatusRefreshTests.Record(42, "pending");

        int changed = await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "Check the U8 highlight."),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(1, changed);

        SubmissionReceipt receipt = SubmissionReceiptStore.All.Single();

        Assert.Equal("changes_requested", receipt.LastKnownState);
        Assert.Equal("Check the U8 highlight.", receipt.MaintainerComment);

        // And the badge lights up off the back of it, which is the user-visible consequence.
        Assert.Equal(1, SubmissionReceiptStore.UnreadCommentCount());
    }

    [Fact]
    public async Task A_DECIDED_submission_is_never_asked_about_again()
    {
        // Its state cannot change, so a request could only return what is already stored. This is
        // what keeps launch cheap for a contributor with a long history.
        SubmissionStatusRefreshTests.Record(1, "published");
        SubmissionStatusRefreshTests.Record(2, "rejected");
        SubmissionStatusRefreshTests.Record(3, "pending");

        var asked = new List<long>();

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("pending", string.Empty, asked),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal([3], asked);
    }

    // ###########################################################################################
    // A submission in BETA is asked about at every launch until its board reaches production - but
    // not for ever (code review, 2026-09-25): past SubmissionReceiptPresenter.MergedRecheckWindow
    // the launch check asks only when the receipt has not been checked for
    // MergedLateRecheckInterval, so a board that is never promoted costs a request a week, not one
    // per launch - and a late BETA rollback still arrives (code review, 2026-09-27).
    // ###########################################################################################
    [Fact]
    public async Task A_submission_in_BETA_is_asked_about_at_every_launch_within_the_window_and_weekly_after()
    {
        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 7,
            UploadToken = "token",
            BoardId = "Commodore/C64/250407/Data.xlsx",
            SentUtc = SubmissionStatusRefreshTests.Now.AddDays(-3),
            DecidedUtc = SubmissionStatusRefreshTests.Now.AddDays(-2),
            LastKnownState = "merged"
        });

        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 8,
            UploadToken = "token",
            BoardId = "Commodore/C64/250407/Data.xlsx",
            SentUtc = SubmissionStatusRefreshTests.Now.AddDays(-90),
            DecidedUtc = SubmissionStatusRefreshTests.Now.AddDays(-89),
            LastKnownState = "merged",

            // Past the window and checked yesterday: not due.
            LastCheckedUtc = SubmissionStatusRefreshTests.Now.AddDays(-1)
        });

        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 9,
            UploadToken = "token",
            BoardId = "Commodore/C64/250407/Data.xlsx",
            SentUtc = SubmissionStatusRefreshTests.Now.AddDays(-90),
            DecidedUtc = SubmissionStatusRefreshTests.Now.AddDays(-89),
            LastKnownState = "merged",

            // Past the window and not checked for over a week: due.
            LastCheckedUtc = SubmissionStatusRefreshTests.Now.AddDays(-8)
        });

        var asked = new List<long>();

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("merged", string.Empty, asked),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal([7, 9], asked.Order());
    }

    // ###########################################################################################
    // *** A BETA ROLLBACK REACHES THE CONTRIBUTOR (code review, 2026-09-27). *** The server answers
    // "returned" for a merged submission a maintainer pushed back, and the receipt takes it - with
    // the maintainer's reason - so "My submissions" stops saying "Published to the BETA source" and the
    // badge tells the contributor to look.
    // ###########################################################################################
    [Fact]
    public async Task A_merged_submission_the_server_reports_returned_is_updated_and_badged()
    {
        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 11,
            UploadToken = "token",
            BoardId = "Commodore/C64/250407/Data.xlsx",
            SentUtc = SubmissionStatusRefreshTests.Now.AddDays(-5),
            DecidedUtc = SubmissionStatusRefreshTests.Now.AddDays(-4),
            LastKnownState = "merged",
            AcknowledgedState = "merged"
        });

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("returned", "The U8 pinout is wrong.", []),
            SubmissionStatusRefreshTests.Now);

        SubmissionReceipt receipt = SubmissionReceiptStore.All.Single(item => item.SubmissionId == 11);

        Assert.Equal("returned", receipt.LastKnownState);
        Assert.Equal("The U8 pinout is wrong.", receipt.MaintainerComment);
        Assert.True(SubmissionReceiptPresenter.HasUnreadNews(receipt));
    }

    [Fact]
    public async Task NOTHING_CHANGED_reports_zero_so_the_UI_is_not_rebuilt_for_nothing()
    {
        // *** WHY THE RETURN VALUE IS "CHANGED" AND NOT "CHECKED". *** UpdateState rewrites the row
        // and saves the file on every call, so counting calls would report "something happened" on
        // every single launch and make the caller redraw the tab each time.
        SubmissionStatusRefreshTests.Record(42, "pending");

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("pending", string.Empty),
            SubmissionStatusRefreshTests.Now);

        int changed = await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("pending", string.Empty),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(0, changed);
    }

    [Fact]
    public async Task A_NEW_COMMENT_on_an_unchanged_state_still_counts_as_a_change()
    {
        // A maintainer can add or reword a comment without the state moving. Comparing only the
        // state would leave that feedback sitting on disk with no badge.
        SubmissionStatusRefreshTests.Record(42, "pending");

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "First note."),
            SubmissionStatusRefreshTests.Now);

        int changed = await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "Actually, also the region."),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(1, changed);
    }

    [Fact]
    public async Task ONE_unreachable_row_does_not_abandon_the_others()
    {
        // A launch-time check runs unattended over every open submission; one failure must not
        // cost the rest. The failing row simply keeps its last known state.
        SubmissionStatusRefreshTests.Record(1, "pending");
        SubmissionStatusRefreshTests.Record(2, "pending");

        int changed = await SubmissionStatusRefresh.RefreshAsync(
            (id, _, _) => id == 1
                ? throw new HttpRequestExceptionStub()
                : Task.FromResult<SubmissionStatus?>(new SubmissionStatus
                {
                    Id = id,
                    State = "changes_requested",
                    MaintainerComment = "Have a look at U8."
                }),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(1, changed);

        Assert.Equal(
            "pending",
            SubmissionReceiptStore.All.Single(receipt => receipt.SubmissionId == 1).LastKnownState);
    }

    [Fact]
    public async Task A_NULL_answer_leaves_the_row_exactly_as_it_was()
    {
        // GetStatusAsync answers null for anything it cannot reach or is refused. That must never
        // be written over a good cached state as though it were an answer.
        SubmissionStatusRefreshTests.Record(42, "pending");

        int changed = await SubmissionStatusRefresh.RefreshAsync(
            (_, _, _) => Task.FromResult<SubmissionStatus?>(null),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(0, changed);
        Assert.Equal("pending", SubmissionReceiptStore.All.Single().LastKnownState);
    }

    // ###########################################################################################
    // *** AN UNCHANGED ANSWER WRITES NOTHING, OR NEARLY (code review, 2026-10-04). *** The minute
    // check saved the whole receipts file once per open receipt, every minute, only to move "last
    // checked" - which nothing shows. Now that date is written only once it is
    // CheckedSaveInterval stale; observed here through the file itself.
    // ###########################################################################################
    [Fact]
    public async Task An_unchanged_answer_does_not_write_the_receipts_file_again_within_the_interval()
    {
        SubmissionStatusRefreshTests.Record(42, "pending");
        await SubmissionStatusRefresh.RefreshAsync(SubmissionStatusRefreshTests.Answering("pending", string.Empty), SubmissionStatusRefreshTests.Now);

        string file = Path.Combine(this.thisWorkspace.Root, "submissions.json");
        File.Delete(file);

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("pending", string.Empty),
            SubmissionStatusRefreshTests.Now.AddMinutes(1));

        Assert.False(File.Exists(file));
        Assert.Equal(SubmissionStatusRefreshTests.Now, SubmissionReceiptStore.All.Single().LastCheckedUtc);

        // Stale enough, the date is written again - a weekly re-check of a receipt in BETA reads it.
        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("pending", string.Empty),
            SubmissionStatusRefreshTests.Now + SubmissionReceiptStore.CheckedSaveInterval);

        Assert.True(File.Exists(file));
        Assert.Equal(SubmissionStatusRefreshTests.Now + SubmissionReceiptStore.CheckedSaveInterval, SubmissionReceiptStore.All.Single().LastCheckedUtc);
    }

    // A change the row does not count as news - a decision date arriving - is still written at once.
    [Fact]
    public async Task A_decision_date_arriving_on_an_unchanged_state_is_written_at_once()
    {
        SubmissionStatusRefreshTests.Record(42, "changes_requested");
        DateTimeOffset decided = SubmissionStatusRefreshTests.Now.AddHours(-3);

        await SubmissionStatusRefresh.RefreshAsync(
            (id, _, _) => Task.FromResult<SubmissionStatus?>(new SubmissionStatus { Id = id, State = "changes_requested", DecidedUtc = decided }),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(decided, SubmissionReceiptStore.All.Single().DecidedUtc);
    }

    // ###########################################################################################
    // *** A SUBMISSION THE SERVER NO LONGER KNOWS (code review, 2026-10-04). *** One deleted with
    // its board answers 404, and was asked about every minute for ever. Now it is marked, keeps its
    // last state, and is asked about once a day - and any later answer clears the mark.
    // ###########################################################################################
    [Fact]
    public async Task A_submission_the_server_does_not_know_is_asked_about_once_a_day_and_keeps_its_state()
    {
        SubmissionStatusRefreshTests.Record(42, "pending");
        var asked = new List<long>();

        SubmissionStatusRefresh.StatusLookup notFound = (id, _, _) =>
        {
            asked.Add(id);
            throw new SubmissionNotFoundException(id);
        };

        int changed = await SubmissionStatusRefresh.RefreshAsync(notFound, SubmissionStatusRefreshTests.Now);

        // A CHANGE (code review, 2026-10-04 - it was 0): the row now reads "No longer on the
        // server", so the Drafts tab has to be drawn again.
        Assert.Equal(1, changed);
        SubmissionReceipt receipt = SubmissionReceiptStore.All.Single();
        Assert.Equal("pending", receipt.LastKnownState);
        Assert.Equal(SubmissionStatusRefreshTests.Now, receipt.NotFoundUtc);

        // A minute later, and an hour later: not asked.
        await SubmissionStatusRefresh.RefreshAsync(notFound, SubmissionStatusRefreshTests.Now.AddMinutes(1));
        await SubmissionStatusRefresh.RefreshAsync(notFound, SubmissionStatusRefreshTests.Now.AddHours(1));
        Assert.Single(asked);
        Assert.False(SubmissionStatusRefresh.ChecksPeriodically(SubmissionReceiptStore.All, SubmissionStatusRefreshTests.Now.AddHours(1), windowMinimised: false));

        // A day later: asked again - and an answer clears the mark, which is a change too.
        int knownAgain = await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("pending", string.Empty),
            SubmissionStatusRefreshTests.Now + SubmissionReceiptPresenter.NotFoundRecheckInterval);

        Assert.Equal(1, knownAgain);
        Assert.Null(SubmissionReceiptStore.All.Single().NotFoundUtc);
    }

    // ###########################################################################################
    // *** "NOT FOUND" IS NEWS ONCE (code review, 2026-10-04). *** The minute check only redraws the
    // Drafts tab when something changed, and a receipt turning "not found" did not count - so its
    // row kept saying "Submitted - awaiting feedback", with Submit greyed out, until the tab was
    // shown again. The first "not found" counts; the daily re-ask that finds the same does not.
    // ###########################################################################################
    [Fact]
    public async Task A_submission_turning_not_found_counts_as_a_change_once()
    {
        SubmissionStatusRefreshTests.Record(42, "pending");
        SubmissionStatusRefresh.StatusLookup notFound = (id, _, _) => throw new SubmissionNotFoundException(id);
        int changedSeen = 0;

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            notFound, SubmissionStatusRefreshTests.Now, onFinished: changed => changedSeen = changed);

        Assert.Equal(1, changedSeen);

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            notFound,
            SubmissionStatusRefreshTests.Now + SubmissionReceiptPresenter.NotFoundRecheckInterval,
            onFinished: changed => changedSeen = changed);

        Assert.Equal(0, changedSeen);
    }

    // ###########################################################################################
    // *** "UPDATE CRT" IS SAID, NOT SWALLOWED (code review, 2026-10-04). *** A server that no longer
    // serves this CRT answers every submission with 426. It used to be read as "unreachable": every
    // receipt froze and nothing said why. The round stops at the first such answer - the rest would
    // only hear the same - and the server's sentence is handed on to be shown; the follow-up after
    // the check still runs.
    // ###########################################################################################
    [Fact]
    public async Task An_update_CRT_answer_stops_the_check_and_hands_on_the_servers_words()
    {
        SubmissionStatusRefreshTests.Record(41, "pending");
        SubmissionStatusRefreshTests.Record(42, "pending");

        var asked = new List<long>();
        var outdated = new List<string>();
        bool finished = false;

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            (id, _, _) =>
            {
                asked.Add(id);
                throw new ClientOutdatedException("Please update CRT to version [3.2.0] or newer.");
            },
            SubmissionStatusRefreshTests.Now,
            onFinished: _ => finished = true,
            onOutdated: outdated.Add);

        Assert.Single(asked);
        Assert.Equal(["Please update CRT to version [3.2.0] or newer."], outdated);
        Assert.True(finished);
        Assert.All(SubmissionReceiptStore.All, receipt => Assert.Equal("pending", receipt.LastKnownState));
    }

    // Where the words are shown, they say what could not be done.
    [Fact]
    public void The_update_CRT_words_are_shown_as_the_reason_the_status_could_not_be_checked()
    {
        Assert.Equal(
            "The status of your submissions could not be checked: Please update CRT to version [3.2.0] or newer.",
            SubmissionStatusRefresh.DescribeOutdated(" Please update CRT to version [3.2.0] or newer. "));
    }

    // -----------------------------------------------------------------------------------
    // The launch wrapper, which must never disturb startup.
    // -----------------------------------------------------------------------------------

    [Fact]
    public async Task With_NOTHING_still_open_the_launch_check_makes_NO_REQUEST_AT_ALL()
    {
        // Somebody who has never contributed, or whose submissions are all decided, must not pay
        // for a network round trip on every launch.
        SubmissionStatusRefreshTests.Record(1, "published");

        var asked = new List<long>();

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            SubmissionStatusRefreshTests.Answering("published", string.Empty, asked),
            SubmissionStatusRefreshTests.Now);

        Assert.Empty(asked);
    }

    [Fact]
    public async Task The_launch_check_SWALLOWS_a_failure_rather_than_faulting_startup()
    {
        // *** IT IS STARTED WITHOUT BEING AWAITED. *** An exception escaping would fault a Task
        // nobody observes - reported arbitrarily late by TaskScheduler.UnobservedTaskException, or
        // never at all. A server being down must cost a stale row and nothing else.
        SubmissionStatusRefreshTests.Record(42, "pending");

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            (_, _, _) => throw new InvalidOperationException("the server is on fire"),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal("pending", SubmissionReceiptStore.All.Single().LastKnownState);
    }

    [Fact]
    public async Task The_launch_check_reports_a_change_so_the_badge_can_appear()
    {
        // The callback is what makes the Drafts tab show its badge without the user doing
        // anything - and it must fire ONLY when something actually moved.
        SubmissionStatusRefreshTests.Record(42, "pending");

        int notified = 0;

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "Please check U8."),
            SubmissionStatusRefreshTests.Now,
            onChanged: () => notified++);

        Assert.Equal(1, notified);

        // Nothing moved the second time, so nothing is redrawn.
        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "Please check U8."),
            SubmissionStatusRefreshTests.Now,
            onChanged: () => notified++);

        Assert.Equal(1, notified);
    }

    // ###########################################################################################
    // *** onFinished FIRES EVERY TIME, EVEN WITH NOTHING TO ASK (code review, 2026-09-25). ***
    //
    // Retiring a published draft was hung off onChanged, which fires only when a receipt MOVES.
    // A receipt already stored as "published" by an earlier launch is never open again, so it never
    // moves again - and a draft whose synced board arrived one launch late was never retired.
    // onFinished is the trigger for work whose precondition is a state, not a transition.
    // ("merged" was the example until 2026-09-25; since the two-stage publish it stays open until
    // the board reaches production.)
    // ###########################################################################################
    [Fact]
    public async Task The_launch_check_FINISHES_even_when_every_receipt_was_already_decided()
    {
        SubmissionStatusRefreshTests.Record(1, "published");

        var finished = new List<int>();
        int changed = 0;

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            SubmissionStatusRefreshTests.Answering("published", string.Empty),
            SubmissionStatusRefreshTests.Now,
            onChanged: () => changed++,
            onFinished: finished.Add);

        Assert.Equal([0], finished);
        Assert.Equal(0, changed);
    }

    [Fact]
    public async Task The_launch_check_FINISHES_with_the_number_that_moved()
    {
        SubmissionStatusRefreshTests.Record(42, "pending");

        var finished = new List<int>();

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            SubmissionStatusRefreshTests.Answering("merged", string.Empty),
            SubmissionStatusRefreshTests.Now,
            onFinished: finished.Add);

        Assert.Equal([1], finished);
    }

    [Fact]
    public async Task The_launch_check_FINISHES_even_when_the_server_fails()
    {
        // A server that is down must not stop local work that does not need it - retiring a draft
        // depends on the receipts already on disk and the synced board, not on this round trip.
        SubmissionStatusRefreshTests.Record(42, "pending");

        var finished = new List<int>();

        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            (_, _, _) => throw new InvalidOperationException("the server is on fire"),
            SubmissionStatusRefreshTests.Now,
            onFinished: finished.Add);

        Assert.Equal([0], finished);
    }

    [Fact]
    public async Task A_throwing_onFinished_does_not_escape_the_launch_check()
    {
        // Started without being awaited, so anything escaping would fault an unobserved Task.
        await SubmissionStatusRefresh.RefreshQuietlyAsync(
            SubmissionStatusRefreshTests.Answering("merged", string.Empty),
            SubmissionStatusRefreshTests.Now,
            onFinished: _ => throw new InvalidOperationException("retirement blew up"));
    }

    [Fact]
    public async Task A_launch_check_does_NOT_undo_a_comment_already_marked_as_read()
    {
        // The launch check writes the same comment back every time it runs. If that cleared the
        // acknowledgement, the badge would return on every single launch for feedback the
        // contributor has already dealt with.
        SubmissionStatusRefreshTests.Record(42, "pending");

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "Check U8."),
            SubmissionStatusRefreshTests.Now);

        SubmissionReceiptStore.AcknowledgeComment(42);

        await SubmissionStatusRefresh.RefreshAsync(
            SubmissionStatusRefreshTests.Answering("changes_requested", "Check U8."),
            SubmissionStatusRefreshTests.Now);

        Assert.Equal(0, SubmissionReceiptStore.UnreadCommentCount());
    }

    // A stand-in for a network failure. Named rather than using HttpRequestException directly so
    // this file needs no System.Net.Http reference for one throw.
    private sealed class HttpRequestExceptionStub : Exception
    {
    }
}
