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

            Assert.Equal(new ReviewContributorFacts("dennis@example.com", null, Published: 1, Waiting: 1, ChangesRequested: 0, Rejected: 0), facts);
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
    }
}
