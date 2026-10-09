using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ReviewApiRoutes - the half of the API client that can be wrong in a way worth testing.
//
// The socket half is an untested I/O boundary by this project's rules. What is left is URL
// building, which is exactly where an API client actually breaks: a doubled slash, a dropped
// path segment, a base address a human typed with a trailing slash. All of those produce a 404
// against a server that is working perfectly, which is the kind of fault that gets blamed on the
// server for an hour before anyone looks at the client.
public sealed class ReviewApiRoutesTests
{
    [Theory]
    [InlineData("https://example.com")]
    [InlineData("https://example.com/")]
    [InlineData("https://example.com///")]
    [InlineData("  https://example.com/  ")]
    public void A_trailing_slash_on_the_base_address_does_not_change_the_route(string baseAddress)
    {
        // Configuration a human types, so half of them will end in a slash. Uri's own combining
        // rules would quietly drop a path segment for one of the two spellings.
        Assert.Equal("https://example.com/api/review/queue", ReviewApiRoutes.Queue(baseAddress));
        Assert.Equal("https://example.com/api/health", ReviewApiRoutes.Health(baseAddress));
    }

    [Fact]
    public void A_submission_route_carries_its_id()
    {
        Assert.Equal(
            "https://example.com/api/review/submissions/42",
            ReviewApiRoutes.Submission("https://example.com", 42));
    }

    [Fact]
    public void The_routes_match_what_the_server_actually_maps()
    {
        // These strings are a contract with ReviewEndpoints.MapReviewEndpoints and
        // AccountEndpoints.MapAccountEndpoints. Written out here so that renaming a route on the
        // server without changing the client is a failing test rather than a 404 at runtime -
        // the two projects do not share a route table, and this is the cheapest thing that can
        // notice they have drifted.
        Assert.Equal("https://x/api/review/queue", ReviewApiRoutes.Queue("https://x"));
        Assert.Equal("https://x/api/review/submissions/1", ReviewApiRoutes.Submission("https://x", 1));
        Assert.Equal("https://x/api/accounts/login", ReviewApiRoutes.Login("https://x"));

        // AdminEndpoints.MapAdminEndpoints - its own group, and two POSTs with a body because a
        // board id carries slashes.
        Assert.Equal("https://x/api/admin/boards", ReviewApiRoutes.AdminBoards("https://x"));
        Assert.Equal("https://x/api/admin/accounts", ReviewApiRoutes.AdminAccounts("https://x"));
        Assert.Equal("https://x/api/admin/maintainers", ReviewApiRoutes.AdminAddMaintainer("https://x"));
        Assert.Equal("https://x/api/admin/maintainers/remove", ReviewApiRoutes.AdminRemoveMaintainer("https://x"));

        // ProductionEndpoints.MapProductionEndpoints.
        Assert.Equal("https://x/api/review/production", ReviewApiRoutes.ProductionList("https://x"));
        Assert.Equal("https://x/api/review/production/plan", ReviewApiRoutes.ProductionPlan("https://x"));
        Assert.Equal("https://x/api/review/production/publish", ReviewApiRoutes.ProductionPublish("https://x"));

        // BoardEndpoints.MapBoardEndpoints (2026-09-27) - under /api/review, not /api/admin:
        // every maintainer reads it. The detail is a POST because a board id carries slashes.
        Assert.Equal("https://x/api/review/boards", ReviewApiRoutes.Boards("https://x"));
        Assert.Equal("https://x/api/review/boards/detail", ReviewApiRoutes.BoardDetail("https://x"));

        // Inviting a new maintainer by email, withdrawing it, and accepting it (2026-09-27).
        Assert.Equal("https://x/api/admin/maintainers/invite", ReviewApiRoutes.AdminInviteMaintainer("https://x"));
        Assert.Equal("https://x/api/admin/maintainers/invitations/withdraw", ReviewApiRoutes.AdminWithdrawInvitation("https://x"));
        Assert.Equal("https://x/api/accounts/accept-invitation", ReviewApiRoutes.AcceptInvitation("https://x"));

        // The maintainer's own account (2026-10-03) - AccountEndpoints' "/me" routes, pinned on the
        // server's side by RequestBodyLimitsTests.
        Assert.Equal("https://x/api/accounts/me", ReviewApiRoutes.Me("https://x/"));
        Assert.Equal("https://x/api/accounts/me/name", ReviewApiRoutes.ChangeName("https://x"));
        Assert.Equal("https://x/api/accounts/me/email", ReviewApiRoutes.ChangeEmail("https://x"));
        Assert.Equal("https://x/api/accounts/me/email/confirm", ReviewApiRoutes.ConfirmEmailChange("https://x"));
        Assert.Equal("https://x/api/accounts/me/password", ReviewApiRoutes.ChangePassword("https://x"));

        // A new board's place in the drop-down lists (2026-09-27): GET the lists, POST a placement.
        Assert.Equal("https://x/api/review/boards/listing", ReviewApiRoutes.BoardListing("https://x"));

        // A board's Board data and Files views (2026-10-03) - the four POSTs BoardEndpoints maps,
        // pinned on the server's side by RequestBodyLimitsTests.
        Assert.Equal("https://x/api/review/boards/table", ReviewApiRoutes.BoardDataTable("https://x/"));
        Assert.Equal("https://x/api/review/boards/edit/check", ReviewApiRoutes.BoardEditCheck("https://x"));
        Assert.Equal("https://x/api/review/boards/edit", ReviewApiRoutes.BoardEdit("https://x"));
        Assert.Equal("https://x/api/review/boards/files", ReviewApiRoutes.BoardFiles("https://x"));

        // Rebuilding both checksum manifests by hand (2026-10-01) - AdminEndpoints' "/manifest/rebuild"
        // under the "/api/admin" group. Unpinned until the code review of the same day.
        Assert.Equal("https://x/api/admin/manifest/rebuild", ReviewApiRoutes.AdminRebuildManifests("https://x"));

        // Deleting a board (2026-10-03) - AdminEndpoints' "/boards/delete/plan" and "/boards/delete".
        // The server's side is pinned by RequestBodyLimitsTests, which finds both in its real route table.
        Assert.Equal("https://x/api/admin/boards/delete/plan", ReviewApiRoutes.AdminBoardDeletePlan("https://x"));
        Assert.Equal("https://x/api/admin/boards/delete", ReviewApiRoutes.AdminBoardDelete("https://x"));

        // Resetting the contribution data and the API usage (2026-10-04) - AdminEndpoints' "/reset"
        // (GET the counts, POST the reset) and "/api-usage". The server's side is pinned by
        // RequestBodyLimitsTests (the POST) and ApiUsageRulesTests' real route table.
        Assert.Equal("https://x/api/admin/reset", ReviewApiRoutes.AdminDataReset("https://x"));
        Assert.Equal("https://x/api/admin/api-usage?days=90", ReviewApiRoutes.AdminApiUsage("https://x", 90));
    }

    [Fact]
    public void A_path_in_the_base_address_is_preserved()
    {
        // The server may be hosted under a path rather than at a domain root, so a path in the base
        // address must survive the route being appended.
        Assert.Equal(
            "https://example.com/crt/api/review/queue",
            ReviewApiRoutes.Queue("https://example.com/crt"));
    }

    [Fact]
    public void An_asset_route_hangs_off_its_own_submission()
    {
        // Both asset routes are scoped to the submission rather than taking a board id, so the
        // server can derive the board from what is being reviewed instead of trusting a caller
        // to name one.
        Assert.Equal(
            "https://x/api/review/submissions/7/submitted/" + new string('a', 64),
            ReviewApiRoutes.SubmittedAsset("https://x", 7, new string('a', 64)));

        Assert.Equal(
            "https://x/api/review/submissions/7/published/Schematics/board.png",
            ReviewApiRoutes.PublishedAsset("https://x", 7, "Schematics/board.png"));
    }

    [Fact]
    public void A_published_path_escapes_each_SEGMENT_and_keeps_the_slashes()
    {
        // *** BOTH HALVES OF THIS MATTER, AND THEY PULL AGAINST EACH OTHER. *** Board data paths
        // routinely carry spaces, so an unescaped URL breaks; but escaping the whole string would
        // encode the separators too, and the server's catch-all would then receive one long
        // segment naming a file that does not exist. Neither failure is obvious from the client
        // side - both arrive as a plain 404.
        Assert.Equal(
            "https://x/api/review/submissions/1/published/Shared%20files/74LS08.png",
            ReviewApiRoutes.PublishedAsset("https://x", 1, "Shared files/74LS08.png"));
    }

    [Fact]
    public void A_HASH_character_in_a_path_cannot_truncate_the_url()
    {
        // A raw "#" turns everything after it into a fragment the server never receives, so the
        // request silently asks for a different file.
        string route = ReviewApiRoutes.PublishedAsset("https://x", 1, "Schematics/rev#2.png");

        Assert.DoesNotContain("#", route);
        Assert.EndsWith("/published/Schematics/rev%232.png", route);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void A_blank_base_address_is_refused_rather_than_producing_a_relative_url(string? baseAddress)
    {
        // A relative URL would be attempted against whatever the client's own base happened to
        // be. A request going somewhere unintended is worse than one that does not go at all -
        // and on the login route it would be a maintainer's credentials.
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.Queue(baseAddress!));
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.Submission(baseAddress!, 1));
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.Login(baseAddress!));
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.SubmittedAsset(baseAddress!, 1, "a"));
        Assert.Throws<ArgumentException>(() => ReviewApiRoutes.PublishedAsset(baseAddress!, 1, "a.png"));
    }

    // ###########################################################################################
    // A published file is fetched where CRT downloads it: the tree's public base and the path,
    // every segment escaped (board folders and "Shared files" have spaces). A base that is not an
    // absolute http(s) address gives nothing to fetch.
    // ###########################################################################################
    [Fact]
    public void A_published_file_is_addressed_as_CRT_downloads_it()
    {
        Assert.Equal(
            "https://example.org/app-data-BETA/Data/Commodore/C128/310378/Data%20C128%20310378%20v2.0.0.json",
            ReviewApiRoutes.PublishedDataFile("https://example.org/app-data-BETA/Data/", "Commodore/C128/310378/Data C128 310378 v2.0.0.json"));

        Assert.Equal(
            "https://example.org/Data/Commodore/Shared%20files/a%23b.pdf",
            ReviewApiRoutes.PublishedDataFile("https://example.org/Data", "Commodore/Shared files/a#b.pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("example.org/Data")]
    [InlineData("file:///srv/Data")]
    public void No_address_without_a_public_base(string? baseUrl)
    {
        Assert.Null(ReviewApiRoutes.PublishedDataFile(baseUrl, "Commodore/a.png"));
    }
}
