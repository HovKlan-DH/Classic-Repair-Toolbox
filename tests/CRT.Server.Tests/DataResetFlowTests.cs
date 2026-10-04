using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Handlers.Usage;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers resetting the contribution data (owner request, 2026-10-04: "When I go-live with this, it
    // should not have old data visible ... all contributor and maintainer data should go away"),
    // Account > "Reset contribution data".
    //
    // What matters: only an administrator, only while the server's switch is on, only the counts
    // that were shown - and for every refusal, that NOTHING was deleted, no stored file touched and
    // no pending usage dropped. A reset that does go names itself in the new history.
    // ###########################################################################################
    public sealed class DataResetFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

        private static readonly DataResetCounts Before =
            new(Submissions: 14, LastSubmissionId: 52, Accounts: 6, LastAccountId: 9, Administrators: 1,
                Maintainers: 4, Invitations: 2, ProductionApprovals: 1, Systems: 7, HistoryEntries: 310,
                BoardViews: 1200, ApiUsageRows: 80);

        private static AccountRecord Account(bool administrator) =>
            new(1, "owner@example.com", "owner@example.com", "hash", "Owner",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false, DataResetFlowTests.Now, null);

        private static ReviewAccess Admin() => ReviewAccess.For(DataResetFlowTests.Account(administrator: true));

        private static ServerOptions Options(bool allow) => new() { AllowDataReset = allow };

        // What the reset touched beside the database: the stored files it asked to remove, and the
        // API usage counted in memory - one call counted before every test, so a test can see
        // whether it was dropped.
        private sealed class Calls
        {
            public Calls() => this.Usage.Count("GET", "/api/review/queue", "3.0.0", DataResetFlowTests.Now);

            public List<IReadOnlyList<long>> FilesRemovedFor { get; } = [];

            public ApiUsageCounter Usage { get; } = new();

            // Whether that one call is still counted - looked at without taking it away.
            public bool UsageKept
            {
                get
                {
                    IReadOnlyList<ApiUsageTally> counted = this.Usage.TakeAll();
                    this.Usage.PutBack(counted);

                    return counted is [var tally] && tally.Calls == 1;
                }
            }
        }

        private static Task<DataResetOutcome> ResetAsync(
            FakeDataResetStore store,
            Calls calls,
            string? fingerprint,
            ReviewAccess? access = null,
            bool allow = true,
            PublishLock? publishLock = null,
            Func<IReadOnlyList<long>, CancellationToken, Task<int>>? removeStoredFiles = null,
            bool signedOut = false,
            CancellationToken cancellationToken = default) =>
            DataResetFlow.ResetAsync(
                signedOut ? null : access ?? DataResetFlowTests.Admin(),
                fingerprint,
                DataResetFlowTests.Options(allow),
                store,
                publishLock ?? new PublishLock(),
                removeStoredFiles ?? ((ids, _) =>
                {
                    calls.FilesRemovedFor.Add(ids);
                    return Task.FromResult(23);
                }),
                calls.Usage,
                DataResetFlowTests.Now,
                NullLogger.Instance,
                cancellationToken);

        private static string Shown => DataResetRules.Fingerprint(DataResetFlowTests.Before);

        // ###########################################################################################
        // The counts are shown while the switch is off too - with the server's words for how to
        // switch it on - so the administrator can see what a reset would take before deciding.
        // ###########################################################################################
        [Fact]
        public async Task The_counts_are_shown_while_the_reset_is_switched_off_with_how_to_switch_it_on()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);

            DataResetOutcome off = await DataResetFlow.PlanAsync(DataResetFlowTests.Admin(), DataResetFlowTests.Options(allow: false), store);

            Assert.NotNull(off.Plan);
            Assert.False(off.Plan!.IsEnabled);
            Assert.Equal(DataResetRules.NotEnabledMessage, off.Plan.NotEnabledBecause);
            Assert.Contains("AllowDataReset", off.Plan.NotEnabledBecause);
            Assert.Equal((14, 6, 1, 4, 2, 7), (off.Plan.Submissions, off.Plan.Accounts, off.Plan.Administrators, off.Plan.Maintainers, off.Plan.Invitations, off.Plan.Systems));
            Assert.Equal((310, 1200, 80), (off.Plan.HistoryEntries, off.Plan.BoardViews, off.Plan.ApiUsageRows));
            Assert.Equal(DataResetFlowTests.Shown, off.Plan.Fingerprint);

            DataResetOutcome on = await DataResetFlow.PlanAsync(DataResetFlowTests.Admin(), DataResetFlowTests.Options(allow: true), store);

            Assert.True(on.Plan!.IsEnabled);
            Assert.Null(on.Plan.NotEnabledBecause);
        }

        // A maintainer who is not an administrator sees no counts and resets nothing.
        [Fact]
        public async Task Somebody_who_is_not_an_administrator_gets_neither_the_counts_nor_a_reset()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);
            var calls = new Calls();
            ReviewAccess maintainer = ReviewAccess.For(DataResetFlowTests.Account(administrator: false), ["Commodore/C64/250407"]);

            DataResetOutcome plan = await DataResetFlow.PlanAsync(maintainer, DataResetFlowTests.Options(allow: true), store);
            DataResetOutcome reset = await DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown, maintainer);
            DataResetOutcome signedOut = await DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown, signedOut: true);

            Assert.True(plan.IsForbidden);
            Assert.Null(plan.Plan);
            Assert.True(reset.IsForbidden);
            Assert.True(signedOut.IsForbidden);
            DataResetFlowTests.AssertNothingHappened(store, calls);
        }

        // ###########################################################################################
        // *** THE SWITCH. *** With AllowDataReset off, even the right fingerprint from an
        // administrator deletes nothing - the guard against a stolen session.
        // ###########################################################################################
        [Fact]
        public async Task While_the_switch_is_off_not_even_an_administrator_with_the_right_counts_deletes_anything()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);
            var calls = new Calls();

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown, allow: false);

            Assert.True(outcome.IsNotEnabled);
            Assert.Equal(DataResetRules.NotEnabledMessage, outcome.Error);
            DataResetFlowTests.AssertNothingHappened(store, calls);
        }

        [Fact]
        public async Task A_reset_of_the_counts_shown_deletes_it_all_and_starts_the_history_with_itself()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before, submissionIds: [3, 17, 52]);
            var calls = new Calls();

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown);

            Assert.Null(outcome.Error);
            Assert.Equal(1, store.Resets);

            // What went, as the answer says it - and the stored files the blob store gave up.
            DataResetAnswer answer = outcome.Answer!;
            Assert.Equal((14, 6, 4, 2, 7), (answer.SubmissionsDeleted, answer.AccountsDeleted, answer.MaintainersDeleted, answer.InvitationsDeleted, answer.SystemsDeleted));
            Assert.Equal((310, 1200, 80, 23), (answer.HistoryEntriesDeleted, answer.BoardViewsDeleted, answer.ApiUsageRowsDeleted, answer.StoredFilesRemoved));

            // The new history's first row: who, and what went.
            AuditEntry record = store.Record!;
            Assert.Equal((1L, "owner@example.com", DataResetRules.ResetAction), (record.ActorAccountId, record.ActorLabel, record.Action));
            Assert.Equal(DataResetFlowTests.Now, record.AtUtc);
            Assert.Equal(DataResetRules.Detail(DataResetFlowTests.Before), record.Detail);

            // The deleted submissions' partial uploads, and the counting not yet written, go too.
            Assert.Equal([3L, 17L, 52L], Assert.Single(calls.FilesRemovedFor));
            Assert.Empty(calls.Usage.TakeAll());
        }

        // ###########################################################################################
        // *** WHAT WAS SHOWN IS WHAT IS DELETED. *** A submission or an account arriving after the
        // counts were shown refuses the reset - the administrator looks again first.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_arriving_after_the_counts_were_shown_refuses_the_reset()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);
            var calls = new Calls();
            string shown = DataResetFlowTests.Shown;

            store.Counts = DataResetFlowTests.Before with { Submissions = 15, LastSubmissionId = 53 };

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(store, calls, shown);

            Assert.True(outcome.IsConflict);
            Assert.Equal(DataResetRules.ChangedMessage, outcome.Error);
            DataResetFlowTests.AssertNothingHappened(store, calls);

            // No fingerprint at all is no confirmation either.
            Assert.True((await DataResetFlowTests.ResetAsync(store, calls, fingerprint: null)).IsConflict);
            DataResetFlowTests.AssertNothingHappened(store, calls);
        }

        // ###########################################################################################
        // *** AND ONE ARRIVING AFTER THE EARLY CHECK (code review, 2026-10-04). *** A contributor's
        // submission commits between the flow's own count and the reset's transaction - neither
        // takes the other's lock. The counts the TRANSACTION reads are the ones held to what was
        // shown, so it is refused with nothing deleted rather than deleted unseen.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_arriving_after_the_early_check_still_refuses_the_reset()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before, submissionIds: [3, 17, 52]);

            store.ArrivesBeforeReset = () =>
            {
                store.Counts = DataResetFlowTests.Before with { Submissions = 15, LastSubmissionId = 53 };
                store.SubmissionIds = [3, 17, 52, 53];
            };

            var calls = new Calls();

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown);

            Assert.True(outcome.IsConflict);
            Assert.Equal(DataResetRules.ChangedMessage, outcome.Error);
            Assert.Equal(0, store.Resets);
            Assert.Null(store.Record);
            Assert.Empty(calls.FilesRemovedFor);

            // Nothing was reset, so the counting not yet written is kept (code review, 2026-10-04:
            // it was dropped before the transaction decided).
            Assert.True(calls.UsageKept);
        }

        // ###########################################################################################
        // A request given up once the rows are deleted still clears the deleted submissions' partial
        // uploads - nothing else can find them once their rows are gone - so the clearing is handed
        // a token that is never cancelled.
        // ###########################################################################################
        [Fact]
        public async Task A_request_cancelled_after_the_reset_still_clears_the_stored_files()
        {
            using var cancellation = new CancellationTokenSource();
            var store = new FakeDataResetStore(DataResetFlowTests.Before, submissionIds: [3]) { AfterReset = cancellation.Cancel };
            var calls = new Calls();
            bool? cancelledToken = null;

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(
                store, calls, DataResetFlowTests.Shown,
                removeStoredFiles: (ids, token) =>
                {
                    cancelledToken = token.IsCancellationRequested;
                    calls.FilesRemovedFor.Add(ids);
                    return Task.FromResult(1);
                },
                cancellationToken: cancellation.Token);

            Assert.NotNull(outcome.Answer);
            Assert.Equal([3L], Assert.Single(calls.FilesRemovedFor));
            Assert.False(cancelledToken);
        }

        // ###########################################################################################
        // The history and the board views grow by themselves while the administrator reads the
        // counts - a sign-in, a CRT reporting its views. They must not refuse the reset.
        // ###########################################################################################
        [Fact]
        public async Task History_and_board_views_arriving_meanwhile_do_not_refuse_the_reset()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);
            var calls = new Calls();
            string shown = DataResetFlowTests.Shown;

            store.Counts = DataResetFlowTests.Before with { HistoryEntries = 312, BoardViews = 1215, ApiUsageRows = 83 };

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(store, calls, shown);

            Assert.NotNull(outcome.Answer);
            Assert.Equal(1, store.Resets);
            Assert.Equal(1215, outcome.Answer!.BoardViewsDeleted);
        }

        // The rows are gone; stored files that cannot be removed now are the sweeper's later.
        [Fact]
        public async Task Stored_files_that_cannot_be_removed_now_do_not_fail_the_reset()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before, submissionIds: [3]);
            var calls = new Calls();

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(
                store, calls, DataResetFlowTests.Shown,
                removeStoredFiles: (_, _) => throw new IOException("The blob store is not mounted."));

            Assert.NotNull(outcome.Answer);
            Assert.Equal(0, outcome.Answer!.StoredFilesRemoved);
            Assert.Equal(1, store.Resets);
        }

        // A database that refuses rolls the whole transaction back: nothing deleted, and it says so.
        [Fact]
        public async Task A_database_that_refuses_the_reset_leaves_everything_and_says_nothing_was_deleted()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before) { Failure = new InvalidOperationException("Lock wait timeout.") };
            var calls = new Calls();

            DataResetOutcome outcome = await DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown);

            Assert.Null(outcome.Answer);
            Assert.Equal(DataResetRules.NotDoneMessage, outcome.Error);
            Assert.False(outcome.IsConflict || outcome.IsForbidden || outcome.IsNotEnabled);
            Assert.Empty(calls.FilesRemovedFor);
            Assert.Equal(DataResetFlowTests.Before, store.Counts);
            Assert.True(calls.UsageKept);
        }

        // ###########################################################################################
        // *** A USAGE WRITE UNDER WAY FINISHES FIRST, AND NONE STARTS UNTIL THE RESET IS DONE (code
        // review, 2026-10-04). *** A write that took its tallies just before the reset used to
        // insert them - or put them back for the next write - after it, so crt_api_calls held calls
        // from before the reset that was meant to empty it. The test holds the counter's writes the
        // way ApiUsageFlusher does while it writes.
        // ###########################################################################################
        [Fact]
        public async Task A_reset_waits_for_a_usage_write_under_way_and_none_lands_after_it()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);
            var calls = new Calls();

            IDisposable writing = await calls.Usage.HoldWritesAsync();

            Task<DataResetOutcome> reset = DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown);

            await Task.Delay(100);
            Assert.False(reset.IsCompleted);
            Assert.Equal(0, store.Resets);

            writing.Dispose();

            Assert.NotNull((await reset.WaitAsync(TimeSpan.FromSeconds(10))).Answer);
            Assert.Equal(1, store.Resets);

            // And while the transaction runs, no write can start: the reset holds the writes until
            // the counter is cleared.
            var held = new FakeDataResetStore(DataResetFlowTests.Before);
            var heldCalls = new Calls();
            Task<IDisposable>? writeDuringReset = null;
            bool startedDuringReset = true;

            held.ArrivesBeforeReset = () =>
            {
                writeDuringReset = heldCalls.Usage.HoldWritesAsync();
                startedDuringReset = writeDuringReset.IsCompleted;
            };

            Assert.NotNull((await DataResetFlowTests.ResetAsync(held, heldCalls, DataResetFlowTests.Shown)).Answer);
            Assert.False(startedDuringReset);

            // The write that waited finds nothing from before the reset to write.
            using (await writeDuringReset!.WaitAsync(TimeSpan.FromSeconds(10)))
                Assert.Empty(heldCalls.Usage.TakeAll());
        }

        // ###########################################################################################
        // *** UNDER THE PUBLISH LOCK. *** A publish, promotion or push-back half way through a
        // submission must finish before the reset deletes it - so a reset waits for the lock.
        // ###########################################################################################
        [Fact]
        public async Task A_reset_waits_for_a_publish_in_progress_to_finish()
        {
            var store = new FakeDataResetStore(DataResetFlowTests.Before);
            var calls = new Calls();
            var publishLock = new PublishLock();

            IDisposable publishing = await publishLock.EnterAsync();

            Task<DataResetOutcome> reset = DataResetFlowTests.ResetAsync(store, calls, DataResetFlowTests.Shown, publishLock: publishLock);

            await Task.Delay(100);
            Assert.False(reset.IsCompleted);
            Assert.Equal(0, store.Resets);

            publishing.Dispose();

            DataResetOutcome outcome = await reset.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.NotNull(outcome.Answer);
            Assert.Equal(1, store.Resets);
        }

        private static void AssertNothingHappened(FakeDataResetStore store, Calls calls)
        {
            Assert.Equal(0, store.Resets);
            Assert.Null(store.Record);
            Assert.Empty(calls.FilesRemovedFor);
            Assert.True(calls.UsageKept);
        }
    }
}
