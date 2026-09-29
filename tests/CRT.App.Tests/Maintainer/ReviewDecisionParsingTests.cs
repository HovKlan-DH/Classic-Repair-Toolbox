using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers the client half of task 5's decisions: the routes, and reading what the server answered.
//
// The HTTP itself is an untested I/O boundary by this project's rules, so what is covered here is
// what can actually be wrong - the URL a decision is posted to, and whether the answer is
// understood. A wrong URL on an APPROVE is the worst of these: it 404s, the maintainer tries again,
// and eventually somebody assumes the feature is broken rather than the route.
public sealed class ReviewDecisionParsingTests
{
    [Fact]
    public void Each_decision_posts_to_its_OWN_route()
    {
        // *** SEPARATE ROUTES RATHER THAN AN OUTCOME FIELD, mirroring the server. *** Approving is
        // irreversible and administrator-only; the other two are neither. A client bug that sent
        // the wrong enum value would be a publish nobody asked for, whereas a wrong URL just 404s.
        Assert.Equal(
            "https://x/api/review/submissions/42/approve",
            ReviewApiRoutes.Approve("https://x", 42));

        Assert.Equal(
            "https://x/api/review/submissions/42/reject",
            ReviewApiRoutes.Reject("https://x", 42));

        Assert.Equal(
            "https://x/api/review/submissions/42/request-changes",
            ReviewApiRoutes.RequestChanges("https://x", 42));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void A_blank_base_address_is_refused_on_a_DECISION_too(string? baseAddress)
    {
        // A relative URL would post a decision somewhere unintended. On approve that is a publish.
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.Approve(baseAddress!, 1));
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.Reject(baseAddress!, 1));
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.RequestChanges(baseAddress!, 1));
    }

    [Fact]
    public void An_APPROVAL_answer_carries_the_state_and_the_published_revision()
    {
        // The revision is what the contributor's NEXT submission diffs against, so a maintainer
        // seeing it confirmed is seeing the thing that actually matters downstream.
        ReviewDecisionResult? result = ReviewApiParser.ParseDecision(
            """{"state":"merged","revision":"2026-September-22","contentHash":"abc"}""");

        Assert.Equal("merged", result!.State);
        Assert.Equal("2026-September-22", result.Revision);
    }

    [Fact]
    public void A_REJECTION_answer_carries_no_revision_and_that_is_fine()
    {
        // Only publishing moves a revision. An empty one here is correct rather than missing.
        ReviewDecisionResult? result = ReviewApiParser.ParseDecision("""{"state":"rejected"}""");

        Assert.Equal("rejected", result!.State);
        Assert.Equal(string.Empty, result.Revision);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("""{"revision":"r1"}""")]
    public void An_answer_with_NO_STATE_is_refused_rather_than_reported_as_success(string? body)
    {
        // *** THE ONE THAT MATTERS. *** Without a state there is no way to know what happened, and
        // reporting success anyway would tell a maintainer their approval published when it may not
        // have. The window says "the server sent something this version does not understand",
        // which is true and sends them to look.
        Assert.Null(ReviewApiParser.ParseDecision(body));
    }

    [Fact]
    public void The_SERVERS_OWN_ERROR_SENTENCE_is_read_back()
    {
        // The server's wording is written for the person reading it - "write at least 10
        // characters" is actionable where a generic "the request was refused" is not.
        Assert.Equal(
            "Write at least 10 characters.",
            ReviewApiParser.ParseError("""{"error":"Write at least 10 characters."}"""));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("""{"error":"   "}""")]
    public void A_body_with_no_usable_error_yields_NULL_so_the_caller_can_fall_back(string? body)
    {
        // Null rather than an empty string, so the client can tell "the server said nothing
        // useful" from "the server said the empty string" and substitute its own message.
        Assert.Null(ReviewApiParser.ParseError(body));
    }

    [Fact]
    public void CONFLICT_and_REFUSED_are_distinct_failure_kinds()
    {
        // They send a maintainer to completely different places: a conflict means refresh and see
        // what somebody else decided, a refusal means fix what you typed and try again. Collapsing
        // them into one "it failed" would make the first look retryable.
        Assert.NotEqual(ReviewApiFailure.Conflict, ReviewApiFailure.Refused);
        Assert.NotEqual(ReviewApiFailure.Conflict, ReviewApiFailure.ServerError);
        Assert.NotEqual(ReviewApiFailure.Refused, ReviewApiFailure.ServerError);
    }
}
