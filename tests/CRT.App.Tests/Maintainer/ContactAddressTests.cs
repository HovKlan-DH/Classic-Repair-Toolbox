using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers ContactAddress - which email address CRT uses wherever it asks for one (owner request,
// 2026-10-01: "When I am a maintainer, and I have logged in, then I want to use that email address
// everywhere in the CRT app - e.g. for the Feedback tab or in the Draft tab").
// ###########################################################################################
public sealed class ContactAddressTests
{
    private static readonly DateTimeOffset Now = new(2026, 10, 1, 12, 0, 0, TimeSpan.Zero);

    private static ReviewSession Session(string email = "dh@example.com", DateTimeOffset? expires = null) =>
        new("token", expires ?? Now.AddDays(30), 7, email, "Dennis");

    // Signed in, the account's address wins over the typed one - and says it is the account's.
    [Fact]
    public void A_signed_in_maintainers_account_address_is_used()
    {
        ReviewSession session = ContactAddressTests.Session(" dh@example.com ");

        ContactAddress address = ContactAddress.Choose(session, "typed@example.com", Now);

        Assert.Equal("dh@example.com", address.Email);
        Assert.True(address.IsFromAccount);
        Assert.Same(session, address.Account);
    }

    // Nobody signed in: the address typed last, as it always was.
    [Fact]
    public void Without_a_sign_in_the_typed_address_is_used()
    {
        ContactAddress address = ContactAddress.Choose(null, "  typed@example.com ", Now);

        Assert.Equal("typed@example.com", address.Email);
        Assert.False(address.IsFromAccount);
        Assert.Null(address.Account);

        Assert.Equal(string.Empty, ContactAddress.Choose(null, null, Now).Email);
    }

    // ###########################################################################################
    // A session no longer usable is no sign-in: sending its token would be refused by the server
    // anyway, and the box would claim an account the submission does not go with.
    // ###########################################################################################
    [Fact]
    public void An_expired_session_counts_as_signed_out()
    {
        ContactAddress address = ContactAddress.Choose(
            ContactAddressTests.Session(expires: Now.AddSeconds(-1)), "typed@example.com", Now);

        Assert.Equal("typed@example.com", address.Email);
        Assert.False(address.IsFromAccount);
    }

    // A session without a usable address offers nothing to show - the typed one stays.
    [Fact]
    public void A_session_with_no_plausible_address_is_not_used()
    {
        ContactAddress address = ContactAddress.Choose(ContactAddressTests.Session(email: ""), "typed@example.com", Now);

        Assert.Equal("typed@example.com", address.Email);
        Assert.False(address.IsFromAccount);
    }
}
