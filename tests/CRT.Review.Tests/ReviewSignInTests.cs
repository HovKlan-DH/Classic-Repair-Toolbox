using CRT.Review.Handlers;

namespace CRT.Review.Tests;

// Covers the sign-in screen's contract with the server (maintainer feedback, 2026-09-22):
// no server prompt, and a working "I forgot my password".
public sealed class ReviewSignInTests
{
    [Fact]
    public void The_default_server_is_the_ONE_address_this_app_ever_talks_to()
    {
        // *** NOT A PROMPT ANY MORE. *** The sign-in screen used to ask for the server on every
        // launch - a question whose answer never changes, and which broke the moment somebody
        // typed a trailing slash or omitted the scheme.
        //
        // *** NO "/api" SUFFIX - and this test CAUGHT that being wrong. *** The first version of
        // this constant copied CRT's AppConfig.CrtServerBaseUrl, which DOES end in "/api" because
        // SubmissionClient appends only "/submissions". Every method here appends the full path,
        // so that value produced "/api/api/review/queue" - a 404 on the very first request, from
        // an address that reads correctly in the source.
        Assert.Equal("https://classic-repair-toolbox.dk", ReviewApiRoutes.DefaultBaseAddress);
        Assert.False(ReviewApiRoutes.DefaultBaseAddress.EndsWith("/api", StringComparison.Ordinal));
    }

    [Fact]
    public void The_default_server_builds_VALID_routes()
    {
        // The real check on that constant, and what makes the doubled-"/api" bug impossible to
        // reintroduce silently: these are the exact URLs the server maps.
        Assert.Equal(
            "https://classic-repair-toolbox.dk/api/review/queue",
            ReviewApiRoutes.Queue(ReviewApiRoutes.DefaultBaseAddress));

        Assert.Equal(
            "https://classic-repair-toolbox.dk/api/accounts/login",
            ReviewApiRoutes.Login(ReviewApiRoutes.DefaultBaseAddress));

        Assert.Equal(
            "https://classic-repair-toolbox.dk/api/accounts/forgot-password",
            ReviewApiRoutes.ForgotPassword(ReviewApiRoutes.DefaultBaseAddress));
    }

    [Fact]
    public void The_LOGOUT_route_matches_what_the_server_maps()
    {
        // A contract with AccountEndpoints across a process boundary, and the one that makes
        // "sign out" mean something: clearing the token locally alone would leave the session live
        // server-side for the rest of its sliding lifetime.
        Assert.Equal(
            "https://classic-repair-toolbox.dk/api/accounts/logout",
            ReviewApiRoutes.Logout(ReviewApiRoutes.DefaultBaseAddress));
    }

    [Fact]
    public void The_RESET_PASSWORD_route_matches_what_the_server_maps()
    {
        // *** THE ROUTE THE BROKEN MAIL LINK SHOULD HAVE POINTED AT ALL ALONG. *** The reset mail
        // printed ".../api/accounts/reset?token=...", which the server maps NOTHING at, so every
        // reset link 404'd from Phase 3 until it was clicked for real on 2026-09-22. The real
        // endpoint is this one, and it is a POST - completing a reset needs the new password, so
        // it can never be something a mail client follows.
        Assert.Equal(
            "https://classic-repair-toolbox.dk/api/accounts/reset-password",
            ReviewApiRoutes.ResetPassword(ReviewApiRoutes.DefaultBaseAddress));

        // The dead path, asserted as NOT what we build - a future "tidy-up" that shortens this to
        // "/reset" would recreate the original 404 exactly.
        Assert.DoesNotContain(
            "/accounts/reset?",
            ReviewApiRoutes.ResetPassword(ReviewApiRoutes.DefaultBaseAddress),
            StringComparison.Ordinal);
    }

    [Fact]
    public void A_password_rejection_reads_back_the_servers_ERRORS_ARRAY()
    {
        // *** A THIRD REFUSAL SHAPE. *** Most refusals answer {"message":...}; a rejected password
        // answers {"errors":[...]}, because AccountRules.ValidatePassword returns a list. A client
        // reading only `message` shows a generic failure in the one case where the server had
        // something genuinely useful to say.
        Assert.Equal(
            "Use at least 12 characters. Do not reuse a password from another site.",
            ReviewApiParser.ParseFirstError(
                """{"errors":["Use at least 12 characters.","Do not reuse a password from another site."]}"""));
    }

    [Fact]
    public void EVERY_password_error_is_shown_rather_than_only_the_first()
    {
        // Being told to fix one rule, fixing it, and then being told about the next is a loop
        // with no visible end - especially with a generated password manager entry.
        string? all = ReviewApiParser.ParseFirstError(
            """{"errors":["Too short.","No digit.","No symbol."]}""");

        Assert.Contains("Too short.", all!, StringComparison.Ordinal);
        Assert.Contains("No digit.", all!, StringComparison.Ordinal);
        Assert.Contains("No symbol.", all!, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("""{"errors":[]}""")]
    [InlineData("""{"errors":"not an array"}""")]
    [InlineData("""{"errors":[null,"   "]}""")]
    public void A_body_with_no_usable_ERRORS_yields_null_so_the_caller_can_fall_back(string? body)
    {
        // Null rather than an empty string, so the client substitutes its own sentence instead of
        // showing a blank line where an explanation belongs.
        Assert.Null(ReviewApiParser.ParseFirstError(body));
    }

    [Fact]
    public void The_forgot_password_route_matches_what_the_server_maps()
    {
        // A contract with AccountEndpoints.MapAccountEndpoints across a process boundary. The
        // endpoint has existed since Phase 3; only the UI for it was missing.
        Assert.Equal(
            "https://x/api/accounts/forgot-password",
            ReviewApiRoutes.ForgotPassword("https://x"));
    }

    [Theory]
    [InlineData("https://example.com")]
    [InlineData("https://example.com/")]
    [InlineData("  https://example.com//  ")]
    public void A_trailing_slash_does_not_change_the_forgot_password_route(string baseAddress)
    {
        Assert.Equal(
            "https://example.com/api/accounts/forgot-password",
            ReviewApiRoutes.ForgotPassword(baseAddress));
    }

    [Fact]
    public void A_blank_base_address_is_refused_on_the_reset_route_too()
    {
        // This request carries an email address to whatever the URL points at; a relative one
        // would send it somewhere unintended.
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.ForgotPassword(""));
    }

    [Fact]
    public void The_servers_own_MESSAGE_is_read_back_rather_than_replaced()
    {
        // The reset answer's whole content is a sentence written for the person reading it.
        // Taking the server's wording means the two cannot drift into saying different things
        // about the same event.
        Assert.Equal(
            "If that address has an account, a reset link is on its way to it.",
            ReviewApiParser.ParseMessage(
                """{"message":"If that address has an account, a reset link is on its way to it."}"""));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("""{"message":"   "}""")]
    public void A_body_with_no_usable_message_yields_NULL_so_the_caller_can_fall_back(string? body)
    {
        // Null rather than an empty string, so the client can substitute its own sentence instead
        // of showing a blank line where an answer belongs.
        Assert.Null(ReviewApiParser.ParseMessage(body));
    }
}
