using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ContributorHistory - who sent a submission and how their other submissions went, the
    // detail's `contributor` (owner request, 2026-09-26).
    // ###########################################################################################
    public sealed class ContributorHistoryTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 26, 12, 0, 0, TimeSpan.Zero);

        private static SubmissionRecord Record(long id, string? email = "dennis@example.com", long? account = null, string state = SubmissionState.Pending) =>
            new(id, "Manu1/Hardware1/Board1", account, email, "hash", "", state, null, 1, Now, null, null);

        private static AccountRecord Account(long id, string email, string name) =>
            new(id, email, email.ToLowerInvariant(), "hash", name, IsVerified: true, IsAdministrator: false, IsLocked: false, Now, null);

        // ###########################################################################################
        // The counts are the contributor's OTHER submissions, each where it ended up - and only what
        // a person decided: the automatic checks' refusal and an earlier copy the contributor's own
        // newer submission replaced are not a track record.
        // ###########################################################################################
        [Fact]
        public void The_other_submissions_are_counted_by_how_they_ended()
        {
            ContributorSubmission[] all =
            [
                new(1, SubmissionState.Merged, true),
                new(2, SubmissionState.Merged, true),
                new(3, SubmissionState.Pending, false),
                new(4, SubmissionState.Approved, true),
                new(5, SubmissionState.ChangesRequested, true),
                new(6, SubmissionState.Rejected, true),
                new(7, SubmissionState.Rejected, false),     // the automatic checks
                new(8, SubmissionState.Withdrawn, false),    // replaced by a newer one
                new(9, SubmissionState.Abandoned, false),    // never finished sending
                new(10, SubmissionState.Pending, false)      // this one
            ];

            ReviewContributorFacts facts = ContributorHistory.Facts(Record(10), all, account: null);

            Assert.Equal(new ReviewContributorFacts("dennis@example.com", null, Published: 2, Waiting: 2, ChangesRequested: 1, Rejected: 1), facts);
        }

        [Fact]
        public void A_first_contribution_has_nothing_else_to_count()
        {
            ReviewContributorFacts facts = ContributorHistory.Facts(Record(10), [new(10, SubmissionState.Pending, false)], account: null);

            Assert.Equal(new ReviewContributorFacts("dennis@example.com", null, 0, 0, 0, 0), facts);
        }

        // A signed-in contributor is named by their account - its address and display name.
        [Fact]
        public void A_signed_in_contributor_is_named_by_their_account()
        {
            ReviewContributorFacts facts = ContributorHistory.Facts(
                Record(10, email: null, account: 7), [], ContributorHistoryTests.Account(7, "anna@example.com", "Anna"));

            Assert.Equal("anna@example.com", facts.Email);
            Assert.Equal("Anna", facts.Name);
        }

        // ###########################################################################################
        // *** THE SAME CONTRIBUTOR AS EVERYWHERE ELSE. *** Through the store: the same email in any
        // case and with stray spaces is one contributor; another address, and a signed-in account's
        // submissions, are not theirs. (SubmissionReplacementRules uses the same rule.)
        // ###########################################################################################
        [Fact]
        public async Task A_contributors_submissions_are_found_by_their_email_in_any_case()
        {
            var store = new FakeSubmissionStore();

            store.Submissions[1] = Record(1, "Dennis@Example.com", state: SubmissionState.Merged);
            store.Submissions[2] = Record(2, " dennis@example.com ", state: SubmissionState.Pending);
            store.Submissions[3] = Record(3, "someone@example.com", state: SubmissionState.Merged);
            store.Submissions[4] = Record(4, null, account: 7, state: SubmissionState.Merged);
            store.Submissions[5] = Record(5, "dennis@example.com");

            ReviewContributorFacts facts = await ContributorHistory.BuildAsync(store.Submissions[5], store, new FakeAccountStore());

            // The counts; the record's list and account facts (2026-09-30) and the stable count
            // (2026-10-01) are the tests below.
            Assert.Equal(
                new ReviewContributorFacts("dennis@example.com", null, Published: 1, Waiting: 1, ChangesRequested: 0, Rejected: 0),
                facts with { SignedIn = null, AccountCreatedUtc = null, Submissions = null, PublishedToStable = null });

            Assert.Equal([2L, 1L], facts.Submissions!.Select(listed => listed.Id));
        }

        // ###########################################################################################
        // *** ONE READ OF THE SYSTEMS, HOWEVER MANY BOARDS THE CONTRIBUTOR HAS (code review,
        // 2026-10-01). *** Each distinct system used to cost its own FindSystemAsync round-trip, on
        // every submission opened in the queue. The same lesson as GET /api/review/production,
        // which asks a fixed number of queries however many systems wait.
        // ###########################################################################################
        [Fact]
        public async Task The_history_reads_the_systems_once_however_many_boards_it_lists()
        {
            var store = new FakeSubmissionStore();

            store.Submissions[1] = Record(1, state: SubmissionState.Merged) with { SystemId = "Commodore/C64/250407" };
            store.Submissions[2] = Record(2, state: SubmissionState.Merged) with { SystemId = "Commodore/C128/310378" };
            store.Submissions[3] = Record(3, state: SubmissionState.Merged) with { SystemId = "Amstrad/CPC/464" };
            store.Submissions[4] = Record(4);

            ReviewContributorFacts facts = await ContributorHistory.BuildAsync(store.Submissions[4], store, new FakeAccountStore());

            Assert.Equal(3, facts.Submissions!.Count);
            Assert.Equal(0, store.FindSystemCalls);
            Assert.Equal(1, store.ListSystemsCalls);
        }

        // A rejection counts when a maintainer made it - SetDecisionAsync records who; the automatic
        // checks' rejection at finalise is only a state.
        [Fact]
        public async Task A_maintainers_rejection_counts_and_the_automatic_checks_does_not()
        {
            var store = new FakeSubmissionStore();

            store.Submissions[1] = Record(1);
            await store.SetDecisionAsync(1, SubmissionState.Rejected, decidedByAccountId: 3, "Not this board.", Now);

            store.Submissions[2] = Record(2);
            await store.SetStateAsync(2, SubmissionState.Rejected, Now);

            store.Submissions[3] = Record(3);

            ReviewContributorFacts facts = await ContributorHistory.BuildAsync(store.Submissions[3], store, new FakeAccountStore());

            Assert.Equal(1, facts.Rejected);
        }

        // ---------------------------------------------------------------- The whole record (2026-09-30)

        // ###########################################################################################
        // *** THE LIST IS EXACTLY WHAT THE COUNTS COUNT (owner request, 2026-09-30). *** The
        // Contributor view shows the counts above the list, so a submission listed but not counted -
        // or counted but not listed - would make the two disagree on screen. Newest first; this
        // submission left out, as the counts leave it out.
        // ###########################################################################################
        [Fact]
        public void The_submissions_listed_are_exactly_the_ones_counted_newest_first()
        {
            ContributorSubmission[] all =
            [
                new(1, SubmissionState.Merged, true),
                new(3, SubmissionState.Pending, false),
                new(4, SubmissionState.Approved, true),
                new(5, SubmissionState.ChangesRequested, true),
                new(6, SubmissionState.Rejected, true),
                new(7, SubmissionState.Rejected, false),     // the automatic checks
                new(8, SubmissionState.Withdrawn, false),    // replaced by a newer one
                new(9, SubmissionState.Abandoned, false),    // never finished sending
                new(2, SubmissionState.Uploading, false),    // never finished sending
                new(10, SubmissionState.Pending, false)      // this one
            ];

            IReadOnlyList<ContributorSubmission> listed = ContributorHistory.Listed(Record(10), all);
            ReviewContributorFacts facts = ContributorHistory.Facts(Record(10), all, account: null);

            Assert.Equal([6L, 5L, 4L, 3L, 1L], listed.Select(other => other.Id));
            Assert.Equal(facts.Published + facts.Waiting + facts.ChangesRequested + facts.Rejected, listed.Count);
        }

        // A long record is cut at the limit, newest kept - the counts still say how many in all.
        [Fact]
        public void A_long_record_lists_only_the_newest()
        {
            ContributorSubmission[] all = Enumerable.Range(1, ContributorHistory.ListedSubmissions + 5)
                .Select(id => new ContributorSubmission(id, SubmissionState.Merged, true))
                .ToArray();

            IReadOnlyList<ContributorSubmission> listed = ContributorHistory.Listed(Record(999), all);

            Assert.Equal(ContributorHistory.ListedSubmissions, listed.Count);
            Assert.Equal(ContributorHistory.ListedSubmissions + 5, listed[0].Id);
        }

        // ###########################################################################################
        // Each in the word its CONTRIBUTOR is told - the Systems screen's rule, so a submission reads
        // the same on every screen: merged after its system last reached production is still
        // "merged" (in BETA), merged before it is "published", and one a BETA rollback returned is
        // "returned" rather than plain "pending".
        // ###########################################################################################
        [Fact]
        public void Each_listed_submission_is_in_the_contributors_own_words()
        {
            ContributorSubmission[] listed =
            [
                new(1, SubmissionState.Merged, true, "Commodore/C64/250407", "Old fix", Now.AddDays(-9), Now.AddDays(-8)),
                new(2, SubmissionState.Merged, true, "Commodore/C64/250407", "New fix", Now.AddDays(-2), Now.AddDays(-1)),
                new(3, SubmissionState.Pending, true, "Commodore/C128/310378", "Returned one", Now.AddDays(-5), Now.AddDays(-3), "Not right in BETA."),
                new(4, SubmissionState.Rejected, true, "Commodore/C128/310378", "Wrong board", Now.AddDays(-4), Now.AddDays(-4), "This is the 310378 board.")
            ];

            IReadOnlyList<ContributorSubmissionEntry> entries = ContributorHistory.Entries(
                listed,
                new Dictionary<string, DateTimeOffset?>
                {
                    ["Commodore/C64/250407"] = Now.AddDays(-5),
                    ["Commodore/C128/310378"] = null
                },
                new Dictionary<long, DateTimeOffset> { [3] = Now.AddDays(-3) });

            Assert.Equal(["published", "merged", "returned", "rejected"], entries.Select(entry => entry.State));
            Assert.Equal(
                new ContributorSubmissionEntry(4, "Commodore/C128/310378", "Wrong board", "rejected", Now.AddDays(-4), Now.AddDays(-4), "This is the 310378 board."),
                entries[3]);
        }

        // ###########################################################################################
        // Whether THIS submission came from an account - its address verified - or was typed in
        // without one; and since when the account exists. Through the store, as the detail builds it.
        // ###########################################################################################
        [Fact]
        public async Task The_record_says_whether_the_submission_came_from_an_account()
        {
            var store = new FakeSubmissionStore();
            var accounts = new FakeAccountStore();

            AccountRecord anna = ContributorHistoryTests.Account(7, "anna@example.com", "Anna") with { CreatedUtc = Now.AddDays(-30) };
            accounts.Accounts[anna.Id] = anna;

            store.Submissions[1] = Record(1, email: null, account: anna.Id, state: SubmissionState.Merged);
            store.Submissions[2] = Record(2, email: null, account: anna.Id);
            store.Submissions[3] = Record(3, "typed@example.com");

            ReviewContributorFacts signedIn = await ContributorHistory.BuildAsync(store.Submissions[2], store, accounts);
            ReviewContributorFacts typed = await ContributorHistory.BuildAsync(store.Submissions[3], store, accounts);

            Assert.True(signedIn.SignedIn);
            Assert.Equal(Now.AddDays(-30), signedIn.AccountCreatedUtc);
            Assert.Equal([1L], signedIn.Submissions!.Select(listed => listed.Id));

            Assert.False(typed.SignedIn);
            Assert.Null(typed.AccountCreatedUtc);
            Assert.Empty(typed.Submissions!);
        }

        // ###########################################################################################
        // *** "[1] PUBLISHED TO STABLE" (owner request, 2026-10-01). *** Published counts the BETA
        // data and the stable source together; PublishedToStable is the part whose system was
        // promoted after the decision - by the same rule that words each listed submission, so the
        // count and the list agree. This submission is left out, as every count leaves it out.
        // ###########################################################################################
        [Fact]
        public void Published_to_stable_counts_the_published_submissions_whose_system_was_promoted_since()
        {
            ContributorSubmission[] all =
            [
                new(1, SubmissionState.Merged, true, "Commodore/C64/250407", "Old fix", Now.AddDays(-9), Now.AddDays(-8)),
                new(2, SubmissionState.Merged, true, "Commodore/C64/250407", "New fix", Now.AddDays(-2), Now.AddDays(-1)),
                new(3, SubmissionState.Merged, true, "Commodore/C128/310378", "Never promoted", Now.AddDays(-4), Now.AddDays(-4)),
                new(4, SubmissionState.Rejected, true, "Commodore/C64/250407", "Turned down", Now.AddDays(-9), Now.AddDays(-9)),
                new(5, SubmissionState.Merged, true, "Commodore/C64/250407", "This one", Now.AddDays(-9), Now.AddDays(-8))
            ];

            int stable = ContributorHistory.PublishedToStable(
                Record(5),
                all,
                new Dictionary<string, DateTimeOffset?>
                {
                    ["Commodore/C64/250407"] = Now.AddDays(-5),
                    ["Commodore/C128/310378"] = null
                });

            Assert.Equal(1, stable);
        }

        // Through the store, as the detail builds it - one in stable, one in BETA only.
        [Fact]
        public async Task The_record_says_how_many_published_submissions_reached_stable()
        {
            var store = new FakeSubmissionStore();

            store.Submissions[1] = Record(1, state: SubmissionState.Merged) with { SystemId = "Commodore/C64/250407", DecidedUtc = Now.AddDays(-8) };
            store.Submissions[2] = Record(2, state: SubmissionState.Merged) with { SystemId = "Commodore/C128/310378", DecidedUtc = Now.AddDays(-1) };
            store.Submissions[3] = Record(3);

            // The C64 reached production after #1 was decided; the C128 never has.
            store.Systems["Commodore/C64/250407"] = new NewSubmission(
                "Commodore/C64/250407", "Commodore", "C64", "250407",
                null, "dennis@example.com", "192.0.2.1", "hash", "r0", "A change.", 1, [], Now, Now.AddHours(24));
            store.ProductionSystems["Commodore/C64/250407"] = new PublishedSystemRow("2026-September-21", "production-hash", Now.AddDays(-5));

            ReviewContributorFacts facts = await ContributorHistory.BuildAsync(store.Submissions[3], store, new FakeAccountStore());

            Assert.Equal(2, facts.Published);
            Assert.Equal(1, facts.PublishedToStable);
        }

        // A BETA rollback's return, as the store records it, reaches the listed word.
        [Fact]
        public async Task A_returned_submission_is_listed_as_returned()
        {
            var store = new FakeSubmissionStore();

            store.Submissions[1] = Record(1, state: SubmissionState.Pending);
            store.BetaReturns[1] = Now;
            store.Submissions[2] = Record(2);

            ReviewContributorFacts facts = await ContributorHistory.BuildAsync(store.Submissions[2], store, new FakeAccountStore());

            Assert.Equal("returned", Assert.Single(facts.Submissions!).State);
        }
    }
}
