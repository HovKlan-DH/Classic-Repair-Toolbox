using System.Text.Json;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer.Tests;

// ###########################################################################################
// BOTH ENDS OF THE REVIEW API, IN ONE TEST (code review, 2026-09-25).
//
// CRT.Maintainer calls CRT.Server over HTTP, so a renamed JSON field compiles on both sides and fails
// only in a maintainer's hands. The tests on each side used to check their own half alone - the
// server's flows against fakes, this app's parser against hand-written JSON - so a field renamed
// on one side left every test green. CLAUDE.md: "A shared change needs a test that would fail if
// only one side moved."
//
// These use what each side really uses:
//   - a REQUEST is built as the maintainer app builds it (ReviewApiClient.Body) and read back with the
//     server's JSON settings into the record the server binds;
//   - an ANSWER is the CRT.Data record the server returns, serialised with the server's settings,
//     and read by the maintainer app's own parser.
// The server's settings are ReviewApiContract.ApplyWireSettings on top of ASP.NET's starting point
// (JsonSerializerDefaults.Web) - exactly what Program.cs configures.
// ###########################################################################################
public sealed class ReviewWireContractTests
{
    private static readonly JsonSerializerOptions Server = ReviewWireContractTests.ServerSettings();

    private static JsonSerializerOptions ServerSettings()
    {
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        ReviewApiContract.ApplyWireSettings(options);
        return options;
    }

    // What the server would write for this answer.
    private static string Answer(object answer) => JsonSerializer.Serialize(answer, answer.GetType(), ReviewWireContractTests.Server);

    // What the server would bind from what this app sends.
    private static T Received<T>(object request) where T : class =>
        JsonSerializer.Deserialize<T>(ReviewApiClient.Body(request), ReviewWireContractTests.Server)!;

    private static ValidationFinding Finding(ValidationSeverity severity) => new()
    {
        Severity = severity,
        Code = "image.orphan",
        Subject = "U9",
        Message = "An image references component [U9]."
    };

    // -----------------------------------------------------------------------------------
    // Requests
    // -----------------------------------------------------------------------------------

    // The example the review found: ExpectedVersion arriving as nothing turned every amendment into
    // "changed since you opened it".
    [Fact]
    public void An_amendment_arrives_with_its_version_and_its_rows()
    {
        var rows = new SubmissionRows();
        rows.Components.Add(new ComponentEntry { BoardLabel = "U8", PartNumber = "251715-01" });

        AmendRequest received = ReviewWireContractTests.Received<AmendRequest>(new AmendRequest(3, rows));

        Assert.Equal(3, received.ExpectedVersion);
        Assert.Equal("251715-01", Assert.Single(received.Rows!.Components).PartNumber);
    }

    [Fact]
    public void A_decision_arrives_with_its_comment_and_the_removals_the_maintainer_was_shown()
    {
        ReviewDecisionRequest approve = ReviewWireContractTests.Received<ReviewDecisionRequest>(
            new ReviewDecisionRequest(null, ["Commodore/C64/250407/old.png"]));

        ReviewDecisionRequest reject = ReviewWireContractTests.Received<ReviewDecisionRequest>(
            new ReviewDecisionRequest("Wrong board.", null));

        Assert.Equal(["Commodore/C64/250407/old.png"], approve.ExpectedRemovals);
        Assert.Equal("Wrong board.", reject.Comment);
        Assert.Null(reject.ExpectedRemovals);
    }

    [Fact]
    public void The_production_plan_and_publish_requests_arrive_whole()
    {
        ProductionPlanRequest plan = ReviewWireContractTests.Received<ProductionPlanRequest>(
            new ProductionPlanRequest("Commodore/C64/250407"));

        ProductionPublishRequest publish = ReviewWireContractTests.Received<ProductionPublishRequest>(
            new ProductionPublishRequest("Commodore/C64/250407", new string('a', 64), ["Commodore/Shared files/x.png"]));

        Assert.Equal("Commodore/C64/250407", plan.SystemId);
        Assert.Equal("Commodore/C64/250407", publish.SystemId);
        Assert.Equal(new string('a', 64), publish.ExpectedBetaContentHash);
        Assert.Equal(["Commodore/Shared files/x.png"], publish.ExpectedRemovals);
    }

    [Fact]
    public void A_pool_change_and_an_unused_file_removal_arrive_whole()
    {
        MaintainerChangeRequest change = ReviewWireContractTests.Received<MaintainerChangeRequest>(
            new MaintainerChangeRequest("Commodore/C64/250407", 42));

        UnusedFilesRemoveRequest remove = ReviewWireContractTests.Received<UnusedFilesRemoveRequest>(
            new UnusedFilesRemoveRequest("beta", ["Commodore/C64/250407/old.png"]));

        Assert.Equal(("Commodore/C64/250407", 42L), (change.SystemId, change.AccountId));
        Assert.Equal("beta", remove.Tree);
        Assert.Equal(["Commodore/C64/250407/old.png"], remove.Files);
    }

    // The field names on the wire are camelCase, as they were when these bodies were anonymous
    // objects - so a server not yet updated reads them the same.
    [Fact]
    public void Request_fields_keep_their_camelCase_names()
    {
        string body = ReviewApiClient.Body(new ProductionPublishRequest("S", "H", []));

        Assert.Contains("\"expectedBetaContentHash\":\"H\"", body, StringComparison.Ordinal);
        Assert.Contains("\"expectedRemovals\":[]", body, StringComparison.Ordinal);
    }

    // -----------------------------------------------------------------------------------
    // Answers
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_publish_and_a_first_approval_read_back_as_the_server_wrote_them()
    {
        ReviewDecisionResult? published = ReviewApiParser.ParseDecision(ReviewWireContractTests.Answer(
            new ReviewDecisionAnswer("merged", Revision: "2026-September-25", ContentHash: "abc", RemovedFiles: ["Commodore/C64/250407/old.png"])));

        ReviewDecisionResult? waiting = ReviewApiParser.ParseDecision(ReviewWireContractTests.Answer(
            new ReviewDecisionAnswer("approved", WaitingFor: [ApproverRole.Administrator])));

        Assert.Equal("merged", published!.State);
        Assert.Equal("2026-September-25", published.Revision);
        Assert.Equal(["Commodore/C64/250407/old.png"], published.RemovedFiles);

        Assert.Equal("approved", waiting!.State);
        Assert.Equal([ApproverRole.Administrator], waiting.WaitingFor);
    }

    [Fact]
    public void A_saved_amendment_reads_back_with_its_version_and_warnings()
    {
        ReviewAmendResult? result = ReviewApiParser.ParseAmend(ReviewWireContractTests.Answer(
            new AmendAnswer(2, [ReviewWireContractTests.Finding(ValidationSeverity.Warning)])));

        Assert.Equal(2, result!.Version);
        ReviewFindingView warning = Assert.Single(result.Warnings);
        Assert.Equal("image.orphan", warning.Code);
        Assert.False(warning.IsError);
    }

    [Fact]
    public void A_production_plan_reads_back_field_for_field()
    {
        var approval = ApprovalRules.Status([], [], ApproverRole.Maintainer);
        var file = new PromotionFile("Commodore/C64/250407/sheet.png", new string('b', 64), PromotionChange.Added, PromotionStage.Content, IsShared: false);

        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan(ReviewWireContractTests.Answer(new ProductionPlanAnswer(
            "Commodore/C64/250407",
            "2026-September-25",
            new string('c', 64),
            IsAwaitingProduction: true,
            TouchesSharedFiles: true,
            CanPublish: true,
            Refusal: null,
            Approval: approval,
            UnchangedCount: 17,
            Files: [file],
            Problems: [ReviewWireContractTests.Finding(ValidationSeverity.Error)],
            Removals: new FileRemovalPreview(["Commodore/C64/250407/old.png"], null))));

        Assert.NotNull(plan);
        Assert.Equal("Commodore/C64/250407", plan!.SystemId);
        Assert.Equal("2026-September-25", plan.BetaRevision);
        Assert.Equal(new string('c', 64), plan.BetaContentHash);
        Assert.True(plan.TouchesSharedFiles);
        Assert.True(plan.CanPublish);
        Assert.Null(plan.Refusal);
        Assert.Equal(17, plan.UnchangedCount);
        Assert.Equal(file.Path, Assert.Single(plan.Files).Path);
        Assert.True(Assert.Single(plan.Problems).IsError);
        Assert.True(plan.Approval!.CanApprove);
        Assert.Equal(["Commodore/C64/250407/old.png"], plan.Removals!.Files);
    }

    [Fact]
    public void A_production_publish_and_a_first_approval_of_one_read_back_whole()
    {
        ProductionPublishResult? published = ReviewApiParser.ParseProductionPublish(ReviewWireContractTests.Answer(
            new ProductionPublishAnswer("Commodore/C64/250407", "published", Revision: "r2", FilesCopied: 12, RemovedFiles: ["x/old.png"])));

        ProductionPublishResult? awaiting = ReviewApiParser.ParseProductionPublish(ReviewWireContractTests.Answer(
            new ProductionPublishAnswer("Commodore/C64/250407", "awaiting", WaitingFor: [ApproverRole.Maintainer])));

        Assert.Equal(("r2", 12), (published!.Revision, published.FilesCopied));
        Assert.Equal(["x/old.png"], published.RemovedFiles);
        Assert.True(awaiting!.IsAwaitingApproval);
        Assert.Equal([ApproverRole.Maintainer], awaiting.WaitingFor);
    }

    [Fact]
    public void Removing_unused_files_reads_back_what_went_and_what_stayed()
    {
        UnusedFileRemovalResult? result = ReviewApiParser.ParseUnusedFileRemoval(ReviewWireContractTests.Answer(
            new UnusedFilesRemoveAnswer("beta", ["a/old.png"], ["a/used.png"], "One file is used again.")));

        Assert.Equal("beta", result!.Tree);
        Assert.Equal(["a/old.png"], result.Removed);
        Assert.Equal(["a/used.png"], result.Kept);
        Assert.Equal("One file is used again.", result.NotDoneBecause);
    }

    // The two answers that were shared records already, read through the server's settings too -
    // including a new system's table, whose published side is null and so left out entirely.
    [Fact]
    public void The_table_and_the_unused_file_list_read_back_through_the_servers_settings()
    {
        var rows = new SubmissionRows();
        rows.Components.Add(new ComponentEntry { BoardLabel = "U8" });

        ReviewTableData? table = ReviewApiParser.ParseTable(ReviewWireContractTests.Answer(new ReviewTableData(4, null, rows)));

        UnusedFileListing? listing = ReviewApiParser.ParseUnusedFiles(ReviewWireContractTests.Answer(
            new UnusedFileListing("production", true, [], 2, 30, 900, [new UnusedFileEntry("a/old.png", 1234)])));

        Assert.Equal(4, table!.Version);
        Assert.Null(table.Published);
        Assert.Equal("U8", Assert.Single(table.Submitted.Components).BoardLabel);

        Assert.Equal("production", listing!.Tree);
        Assert.Equal(1234, Assert.Single(listing.Files).SizeBytes);
    }

    // ###########################################################################################
    // THE THREE LISTS (code review, 2026-09-25): the production list, the administrator's systems
    // and accounts. The server answered them as anonymous objects and the maintainer app read them by
    // hand-typed names, with nothing holding the two together. BetaContentHash is the one that
    // mattered: renamed on either side, every publish to production would have sent an empty hash
    // and been refused as "changed in BETA since you opened it".
    // ###########################################################################################
    [Fact]
    public void The_production_list_reads_back_field_for_field()
    {
        DateTimeOffset published = new(2026, 9, 20, 8, 30, 0, TimeSpan.Zero);
        string betaHash = new('c', 64);

        ProductionListResponse? list = ReviewApiParser.ParseProductionList(ReviewWireContractTests.Answer(
            new ProductionListAnswer(
                true,
                [new ProductionListEntry("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-25", betaHash, "2026-August-21", published)])));

        Assert.True(list!.Configured);

        ProductionSystemRow row = Assert.Single(list.Systems);
        Assert.Equal("Commodore/C64/250407", row.SystemId);
        Assert.Equal("Commodore", row.Manufacturer);
        Assert.Equal("C64", row.Hardware);
        Assert.Equal("250407", row.Board);
        Assert.Equal("2026-September-25", row.BetaRevision);
        Assert.Equal(betaHash, row.BetaContentHash);
        Assert.Equal("2026-August-21", row.ProductionRevision);
        Assert.Equal(published, row.ProductionPublishedUtc);
    }

    // "Switched off" must stay distinguishable from "nothing waiting", with nulls left out.
    [Fact]
    public void A_production_list_from_a_server_that_cannot_publish_reads_back_as_switched_off()
    {
        ProductionListResponse? list = ReviewApiParser.ParseProductionList(
            ReviewWireContractTests.Answer(new ProductionListAnswer(false, [])));

        Assert.False(list!.Configured);
        Assert.Empty(list.Systems);
    }

    [Fact]
    public void The_administrators_list_of_systems_reads_back_with_their_maintainers()
    {
        ReviewSystemsResponse? systems = ReviewApiParser.ParseSystems(ReviewWireContractTests.Answer(
            new MaintainerSystemsAnswer(
            [
                new MaintainerSystemEntry(
                    "Commodore/C64/250407", "Commodore", "C64", "250407", "2026-August-21", IsAccepting: false,
                    [new PoolMaintainerEntry(7, "Anna", "anna@example.com")])
            ])));

        ReviewSystemRow row = Assert.Single(systems!.Systems);
        Assert.Equal("Commodore/C64/250407", row.SystemId);
        Assert.Equal("Commodore", row.Manufacturer);
        Assert.Equal("C64", row.Hardware);
        Assert.Equal("250407", row.Board);
        Assert.Equal("2026-August-21", row.CurrentRevision);
        Assert.False(row.IsAccepting);
        Assert.Equal(new MaintainerRow(7, "Anna", "anna@example.com"), Assert.Single(row.Maintainers));
    }

    [Fact]
    public void The_administrators_list_of_accounts_reads_back_field_for_field()
    {
        ReviewAccountsResponse? accounts = ReviewApiParser.ParseAccounts(ReviewWireContractTests.Answer(
            new MaintainerAccountsAnswer(
            [
                new MaintainerAccountEntry(7, "anna@example.com", "Anna", IsAdministrator: true, IsVerified: true, IsLocked: false),
                new MaintainerAccountEntry(8, "bo@example.com", "Bo", IsAdministrator: false, IsVerified: false, IsLocked: true)
            ])));

        Assert.Equal(
            [
                new ReviewAccountRow(7, "anna@example.com", "Anna", IsAdministrator: true, IsVerified: true, IsLocked: false),
                new ReviewAccountRow(8, "bo@example.com", "Bo", IsAdministrator: false, IsVerified: false, IsLocked: true)
            ],
            accounts!.Accounts);
    }
}
