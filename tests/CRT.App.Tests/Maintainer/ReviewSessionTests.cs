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
}
