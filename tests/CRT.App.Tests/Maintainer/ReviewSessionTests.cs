using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// ReviewSession.WithAccount (2026-10-03) - the signed-in session taking the name and address the
// server now holds, after the "Your account" window changed them or when they are read again at
// launch. The Feedback tab and the Submit dialog use them, so they must follow; the token must not
// change, or the next request would be refused.
// ###########################################################################################
public sealed class ReviewSessionTests
{
    private static readonly DateTimeOffset Expires = new(2026, 11, 2, 12, 0, 0, TimeSpan.Zero);

    private static AccountAnswer Account(long id, string email, string name) =>
        new(id, email, name, true, false, [], DateTimeOffset.UnixEpoch);

    [Fact]
    public void The_session_takes_the_new_name_and_address_and_keeps_its_token()
    {
        var session = new ReviewSession("token", ReviewSessionTests.Expires, 7, "dh@example.com", "Dennis");

        ReviewSession updated = session.WithAccount(ReviewSessionTests.Account(7, "bench@example.com", "Dennis H"));

        Assert.Equal(new ReviewSession("token", ReviewSessionTests.Expires, 7, "bench@example.com", "Dennis H"), updated);
    }

    // An answer about another account cannot be this session's - nothing changes.
    [Fact]
    public void An_answer_about_another_account_changes_nothing()
    {
        var session = new ReviewSession("token", ReviewSessionTests.Expires, 7, "dh@example.com", "Dennis");

        Assert.Same(session, session.WithAccount(ReviewSessionTests.Account(8, "anna@example.com", "Anna")));
    }

    // A session whose account id could not be read (ParseLogin's 0) takes the answer's: it came
    // back for this very token.
    [Fact]
    public void A_session_with_no_known_account_id_takes_the_answers()
    {
        var session = new ReviewSession("token", ReviewSessionTests.Expires, 0, "dh@example.com", "Dennis");

        ReviewSession updated = session.WithAccount(ReviewSessionTests.Account(7, "dh@example.com", "Dennis"));

        Assert.Equal(7, updated.AccountId);
        Assert.Equal("token", updated.BearerToken);
    }

    // ###########################################################################################
    // *** THE EXPIRY SLIDES AS THE SERVER SLIDES IT (code review, 2026-10-13). *** The server pushes
    // a session to a full lifetime from the request once less than half of it is left
    // (SessionExtensionRules); CRT kept the login's expiry, so a sign-in in daily use died 30 days
    // after it was made. AfterUseAt applies the same rule, from the moment the request was sent.
    // ###########################################################################################
    [Fact]
    public void Used_with_less_than_half_its_lifetime_left_the_expiry_moves_a_full_lifetime_from_the_request()
    {
        var signedIn = new DateTimeOffset(2026, 10, 1, 12, 0, 0, TimeSpan.Zero);
        var session = new ReviewSession("token", signedIn + SessionLifetime.Default, 7, "dh@example.com", "Dennis");

        DateTimeOffset used = signedIn.AddDays(20);
        ReviewSession slid = session.AfterUseAt(used);

        Assert.Equal(used.AddDays(30), slid.ExpiresUtc);
        Assert.Equal(session with { ExpiresUtc = used.AddDays(30) }, slid);

        // Past the login's own expiry, the slid session is still usable - which is the bug fixed.
        Assert.False(session.IsUsableAt(signedIn.AddDays(31)));
        Assert.True(slid.IsUsableAt(signedIn.AddDays(31)));
    }

    // With more than half left the server writes nothing, so neither does CRT - the same session.
    [Fact]
    public void Used_with_more_than_half_its_lifetime_left_nothing_moves()
    {
        var signedIn = new DateTimeOffset(2026, 10, 1, 12, 0, 0, TimeSpan.Zero);
        var session = new ReviewSession("token", signedIn + SessionLifetime.Default, 7, "dh@example.com", "Dennis");

        Assert.Same(session, session.AfterUseAt(signedIn.AddDays(1)));
        Assert.Same(session, session.AfterUseAt(signedIn.AddDays(15)));
    }

    // An expired session is never brought back - the server refuses to, and so does this; nor does
    // a lifetime that is not positive move anything, and the far end saturates rather than throws.
    [Fact]
    public void An_expired_session_stays_expired_and_the_far_end_does_not_throw()
    {
        var expires = new DateTimeOffset(2026, 10, 1, 12, 0, 0, TimeSpan.Zero);
        var session = new ReviewSession("token", expires, 7, "dh@example.com", "Dennis");

        Assert.Same(session, session.AfterUseAt(expires));
        Assert.Same(session, session.AfterUseAt(expires.AddDays(3)));
        Assert.Same(session, session.AfterUseAt(expires.AddDays(-1), TimeSpan.Zero));

        var nearTheEnd = new ReviewSession("token", DateTimeOffset.MaxValue.AddDays(-1), 7, "dh@example.com", "Dennis");
        Assert.Equal(DateTimeOffset.MaxValue, nearTheEnd.AfterUseAt(DateTimeOffset.MaxValue.AddDays(-2)).ExpiresUtc);
    }

    // The lifetime CRT slides by is the server's default - one constant in CRT.Data that the
    // server's ServerOptions.RefreshTokenDays takes as its default.
    [Fact]
    public void The_lifetime_is_thirty_days()
    {
        Assert.Equal(30, SessionLifetime.DefaultDays);
        Assert.Equal(TimeSpan.FromDays(30), SessionLifetime.Default);
    }
}
