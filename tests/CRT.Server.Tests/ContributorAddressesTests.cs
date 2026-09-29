using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ContributorAddresses - where a contributor is written to (code review, 2026-09-29).
    //
    // A submission sent while SIGNED IN carries the account and usually no contact address, so every
    // reader of ContactEmail alone found nobody for exactly those contributors: the production panel
    // said "(no contact address)", and the push-back and "now in production" mails skipped them.
    // ###########################################################################################
    public sealed class ContributorAddressesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 29, 12, 0, 0, TimeSpan.Zero);

        private static SubmissionRecord Submission(long id, long? accountId, string? contactEmail) =>
            new(id, "Commodore/C64/250407", accountId, contactEmail, "hash", "r0", SubmissionState.Merged,
                "A fix.", 1, ContributorAddressesTests.Now, null, ContributorAddressesTests.Now);

        private static AccountRecord Account(long id, string email) =>
            new(id, email, email, "hash", "Anna", IsVerified: true, IsAdministrator: false, IsLocked: false,
                ContributorAddressesTests.Now, null);

        [Fact]
        public void A_signed_in_contributor_is_written_to_at_their_accounts_address()
        {
            Assert.Equal(
                "anna@example.com",
                ContributorAddresses.AddressOf(
                    ContributorAddressesTests.Submission(1, 7, null),
                    ContributorAddressesTests.Account(7, "anna@example.com")));
        }

        [Fact]
        public void A_contributor_without_an_account_is_written_to_at_their_contact_address()
        {
            Assert.Equal(
                "bo@example.com",
                ContributorAddresses.AddressOf(ContributorAddressesTests.Submission(1, null, "  bo@example.com "), null));
        }

        // The account is gone (deleted since): the contact address, if any, is all there is.
        [Theory]
        [InlineData("bo@example.com", "bo@example.com")]
        [InlineData(null, "")]
        public void A_missing_account_falls_back_to_the_contact_address(string? contact, string expected)
        {
            Assert.Equal(expected, ContributorAddresses.AddressOf(ContributorAddressesTests.Submission(1, 7, contact), null));
        }

        [Fact]
        public async Task A_whole_list_is_resolved_with_one_lookup_per_account()
        {
            var accounts = new FakeAccountStore();
            accounts.Accounts[7] = ContributorAddressesTests.Account(7, "anna@example.com");

            IReadOnlyDictionary<long, string> addresses = await ContributorAddresses.ResolveAsync(
                accounts,
                [
                    ContributorAddressesTests.Submission(1, 7, null),
                    ContributorAddressesTests.Submission(2, 7, "old@example.com"),
                    ContributorAddressesTests.Submission(3, null, "bo@example.com")
                ]);

            Assert.Equal("anna@example.com", addresses[1]);
            Assert.Equal("anna@example.com", addresses[2]);
            Assert.Equal("bo@example.com", addresses[3]);
        }
    }
}
