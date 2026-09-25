using CRT.Maintainer.Handlers;

namespace CRT.Maintainer.Tests;

// Covers ReviewApiParser - reading the server's answers.
//
// *** THE JSON IN THIS FILE IS COPIED FROM THE SERVER'S OWN RESPONSE BUILDERS ***
// (AccountEndpoints.SessionResponse and ReviewEndpoints.ToQueueRow), down to the camelCase and
// the ISO-8601 offsets ASP.NET Core emits. Both are private, so this is a copy rather than a
// shared fixture - which means it can go stale, and that is worth stating plainly: if a field is
// renamed on the server, these tests keep passing and the app breaks at runtime.
//
// What stops that being a silent failure is ReviewApiRoutesTests' route contract plus the fact
// that a renamed field shows up as a null or a default here rather than an exception - so the
// window reports "the server sent something it did not understand" instead of crashing. Moving
// the response shapes into a shared contract type (as SubmissionContract.cs already does for the
// submission pipeline) is the real fix, and is worth doing the first time one of them changes.
//
// The parser's guiding rule: NEVER THROW. A JsonException out of a UI event handler crashes the
// app; a null lets the window say something true and actionable.
public sealed class ReviewApiParserTests
{
    // Exactly what POST /api/accounts/login answers on success.
    private const string LoginJson = """
        {"refreshToken":"abc123","expiresUtc":"2026-09-22T12:00:00+00:00",
         "account":{"id":7,"email":"dennis@example.com","displayName":"Dennis","isVerified":true}}
        """;

    // Exactly what GET /api/review/queue answers.
    private const string QueueJson = """
        {"canPublish":true,"count":1,
         "submissions":[{"id":42,"systemId":"Commodore/C64/250407","state":"pending",
                         "summary":"Corrected R12.","contactEmail":"someone@example.com",
                         "baseRevision":"r1","createdUtc":"2026-09-21T12:00:00+00:00",
                         "decidedUtc":null}]}
        """;

    // -----------------------------------------------------------------------------------
    // Login
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_login_response_yields_a_usable_session()
    {
        ReviewSession? session = ReviewApiParser.ParseLogin(ReviewApiParserTests.LoginJson);

        Assert.NotNull(session);
        Assert.Equal("abc123", session!.BearerToken);
        Assert.Equal(7, session.AccountId);
        Assert.Equal("dennis@example.com", session.Email);
        Assert.Equal("Dennis", session.DisplayName);
    }

    [Fact]
    public void The_BEARER_token_comes_from_the_refreshToken_field()
    {
        // *** THE NAMING TRAP, PINNED. *** The login endpoint calls this field `refreshToken`,
        // which reads as "exchange this for a real token". It is not: AuthenticateAsync looks up
        // that exact value as the session token. A client author who trusts the name goes hunting
        // for an access-token exchange that does not exist and gets 401s from every call.
        ReviewSession? session = ReviewApiParser.ParseLogin(ReviewApiParserTests.LoginJson);

        Assert.Equal("abc123", session!.BearerToken);
    }

    [Fact]
    public void A_login_response_with_no_token_is_not_a_session()
    {
        // Without a token there is nothing to authorise with, however complete the rest looks.
        Assert.Null(ReviewApiParser.ParseLogin("""{"expiresUtc":"2026-09-22T12:00:00+00:00","account":{"id":7}}"""));
        Assert.Null(ReviewApiParser.ParseLogin("""{"refreshToken":"","account":{"id":7}}"""));
    }

    [Fact]
    public void A_login_response_with_no_account_is_refused()
    {
        Assert.Null(ReviewApiParser.ParseLogin("""{"refreshToken":"abc123"}"""));
    }

    [Fact]
    public void An_unreadable_expiry_is_treated_as_ALREADY_EXPIRED()
    {
        // Being wrong the other way - treating it as never-expiring - would have the app keep
        // presenting a dead token and reporting 401s it could have explained.
        ReviewSession? session = ReviewApiParser.ParseLogin(
            """{"refreshToken":"abc123","account":{"id":7,"email":"a@b.c","displayName":"A"}}""");

        Assert.NotNull(session);
        Assert.False(session!.IsUsableAt(DateTimeOffset.UnixEpoch));
        Assert.False(session.IsUsableAt(DateTimeOffset.MaxValue));
    }

    // -----------------------------------------------------------------------------------
    // The queue
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_queue_response_yields_its_rows()
    {
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue(ReviewApiParserTests.QueueJson);

        Assert.NotNull(queue);
        Assert.True(queue!.CanPublish);

        ReviewQueueRow row = Assert.Single(queue.Submissions);
        Assert.Equal(42, row.Id);
        Assert.Equal("Commodore/C64/250407", row.SystemId);
        Assert.Equal("pending", row.State);
        Assert.Equal("Corrected R12.", row.Summary);
        Assert.Equal("someone@example.com", row.ContactEmail);
        Assert.Equal(new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero), row.CreatedUtc);
    }

    [Fact]
    public void An_EMPTY_queue_and_an_UNREADABLE_one_are_different_answers()
    {
        // *** THE DISTINCTION THAT MATTERS MOST HERE. *** "No submissions are waiting" and "I
        // could not ask" must never look the same to a maintainer, or a broken connection reads as
        // a clear queue and a backlog goes unnoticed for as long as nobody happens to check.
        ReviewQueueResponse? empty = ReviewApiParser.ParseQueue("""{"canPublish":false,"count":0,"submissions":[]}""");

        Assert.NotNull(empty);
        Assert.Empty(empty!.Submissions);

        Assert.Null(ReviewApiParser.ParseQueue("not json at all"));
        Assert.Null(ReviewApiParser.ParseQueue("""{"canPublish":true}"""));
    }

    [Fact]
    public void One_malformed_row_does_not_hide_the_rest_of_the_queue()
    {
        // A single bad record must not conceal every other contribution waiting behind it.
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue("""
            {"canPublish":false,"submissions":[
                {"id":1,"systemId":"A/B/C","state":"pending","summary":"First."},
                {"systemId":"no id here","state":"pending"},
                "a string where an object should be",
                {"id":3,"systemId":"D/E/F","state":"pending","summary":"Third."}]}
            """);

        Assert.NotNull(queue);
        Assert.Equal([1, 3], queue!.Submissions.Select(row => row.Id));
    }

    [Fact]
    public void A_missing_timestamp_stays_NULL_rather_than_becoming_1970()
    {
        // "Waiting since" is shown to the maintainer; a defaulted timestamp would read as a
        // submission that has been waiting fifty years.
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue("""
            {"canPublish":false,"submissions":[{"id":1,"systemId":"A/B/C","state":"pending"}]}
            """);

        Assert.Null(Assert.Single(queue!.Submissions).CreatedUtc);
    }

    [Fact]
    public void The_queue_says_whether_the_account_is_an_ADMINISTRATOR_and_defaults_to_not()
    {
        // What shows the "Maintainers" button. Defaulting the other way would offer an
        // administrator's screen to a maintainer - the server refuses it, but a button that fails
        // is its own kind of wrong. An older server that never sends the field reads as "not".
        Assert.True(ReviewApiParser.ParseQueue("""{"isAdministrator":true,"submissions":[]}""")!.IsAdministrator);
        Assert.False(ReviewApiParser.ParseQueue("""{"isAdministrator":false,"submissions":[]}""")!.IsAdministrator);
        Assert.False(ReviewApiParser.ParseQueue("""{"submissions":[]}""")!.IsAdministrator);
    }

    [Fact]
    public void A_row_says_whether_it_changes_shared_files_and_defaults_to_not()
    {
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue("""
            {"submissions":[
                {"id":1,"systemId":"A/B/C","state":"pending","touchesSharedFiles":true},
                {"id":2,"systemId":"A/B/C","state":"pending"}]}
            """);

        Assert.True(queue!.Submissions[0].TouchesSharedFiles);
        Assert.False(queue.Submissions[1].TouchesSharedFiles);
    }

    [Fact]
    public void A_missing_canPublish_defaults_to_FALSE()
    {
        // Defaulting the other way would have the app offer a publish action to somebody the
        // server never said could publish. The server refuses it anyway - authority is enforced
        // server-side - but offering an action that then fails is its own kind of wrong.
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue("""{"submissions":[]}""");

        Assert.False(queue!.CanPublish);
    }

    // -----------------------------------------------------------------------------------
    // Never throwing
    // -----------------------------------------------------------------------------------

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("<html><body>502 Bad Gateway</body></html>")]
    [InlineData("{\"unterminated\": ")]
    [InlineData("[]")]
    [InlineData("42")]
    public void Anything_that_is_not_the_expected_answer_yields_null_rather_than_throwing(string? body)
    {
        // A proxy error page, a truncated body, or a server this build was not written against.
        // All of them arrive here as "not what was expected", and a JsonException out of a UI
        // event handler crashes the app.
        Assert.Null(ReviewApiParser.ParseLogin(body));
        Assert.Null(ReviewApiParser.ParseQueue(body));
    }

    // -----------------------------------------------------------------------------------
    // Session usability
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_session_is_usable_before_it_expires()
    {
        var session = new ReviewSession("token", DateTimeOffset.UnixEpoch.AddHours(2), 1, "a@b.c", "A");

        Assert.True(session.IsUsableAt(DateTimeOffset.UnixEpoch));
    }

    [Fact]
    public void A_session_expiring_within_the_margin_is_treated_as_already_gone()
    {
        // A token that expires while a request is in flight fails with a 401 the user cannot act
        // on, and reads as the app logging them out at random. Treating an almost-expired session
        // as expired turns that into a clean, explainable sign-in.
        var session = new ReviewSession("token", DateTimeOffset.UnixEpoch.AddSeconds(30), 1, "a@b.c", "A");

        Assert.False(session.IsUsableAt(DateTimeOffset.UnixEpoch));
    }

    [Fact]
    public void A_session_with_no_token_is_never_usable()
    {
        var session = new ReviewSession("", DateTimeOffset.MaxValue, 1, "a@b.c", "A");

        Assert.False(session.IsUsableAt(DateTimeOffset.UnixEpoch));
    }

    [Fact]
    public void The_usability_check_is_TOTAL_at_both_ends_of_the_calendar()
    {
        // *** TWO BUGS LIVED HERE, BOTH ArgumentOutOfRangeException. *** Shifting either side by
        // the expiry margin overflows: `ExpiresUtc - margin` throws at MinValue (which is exactly
        // what an unreadable expiry is recorded as), and `now + margin` throws at MaxValue. This
        // method is on the path of every request and is called from UI code, where an exception
        // is a crash rather than a caught error - so it must answer for any input at all.
        //
        // Applying the margin to the DIFFERENCE is what makes that true: a TimeSpan comparison
        // cannot overflow a DateTime in either direction.
        var session = new ReviewSession("token", DateTimeOffset.MinValue, 1, "a@b.c", "A");
        var eternal = new ReviewSession("token", DateTimeOffset.MaxValue, 1, "a@b.c", "A");

        Assert.False(session.IsUsableAt(DateTimeOffset.MinValue));
        Assert.False(session.IsUsableAt(DateTimeOffset.MaxValue));
        Assert.True(eternal.IsUsableAt(DateTimeOffset.MinValue));
        Assert.False(eternal.IsUsableAt(DateTimeOffset.MaxValue));
    }
}
