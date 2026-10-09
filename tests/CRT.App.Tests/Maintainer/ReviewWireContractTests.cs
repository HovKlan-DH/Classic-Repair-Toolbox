using System.Text.Json;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// BOTH ENDS OF THE REVIEW API, IN ONE TEST (code review, 2026-09-25).
//
// The Maintainer tab calls CRT.Server over HTTP, so a renamed JSON field compiles on both sides and fails
// only in a maintainer's hands. The tests on each side used to check their own half alone - the
// server's flows against fakes, the client's parser against hand-written JSON - so a field renamed
// on one side left every test green. CLAUDE.md: "A shared change needs a test that would fail if
// only one side moved."
//
// These use what each side really uses:
//   - a REQUEST is built as the Maintainer tab builds it (ReviewApiClient.Body) and read back with the
//     server's JSON settings into the record the server binds;
//   - an ANSWER is the CRT.Data record the server returns, serialised with the server's settings,
//     and read by the Maintainer tab's own parser.
// The server's settings are ReviewApiContract.ApplyWireSettings on top of ASP.NET's starting point
// (JsonSerializerDefaults.Web) - exactly what Program.cs configures.
// ###########################################################################################
public sealed class ReviewWireContractTests
{
    private static readonly JsonSerializerOptions Server = ReviewWireContractTests.ServerSettings();

    // A fixed instant, so a round-tripped timestamp is compared against a value and not against
    // "now" read twice.
    private static readonly DateTimeOffset Decided = new(2026, 9, 24, 8, 30, 0, TimeSpan.Zero);

    private static JsonSerializerOptions ServerSettings()
    {
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        ReviewApiContract.ApplyWireSettings(options);
        return options;
    }

    // What the server would write for this answer.
    private static string Answer(object answer) => JsonSerializer.Serialize(answer, answer.GetType(), ReviewWireContractTests.Server);

    // What the server would bind from what the Maintainer tab sends.
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

        Assert.Equal("Commodore/C64/250407", plan.BoardId);
        Assert.Equal("Commodore/C64/250407", publish.BoardId);
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

        Assert.Equal(("Commodore/C64/250407", 42L), (change.BoardId, change.AccountId));
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
            Removals: new FileRemovalPreview(["Commodore/C64/250407/old.png"], null),
            Carrying: [new CarriedSubmission(7, "hest@mailscan.dk", "Corrected U8.", ReviewWireContractTests.Decided)])));

        Assert.NotNull(plan);
        Assert.Equal("Commodore/C64/250407", plan!.BoardId);
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

        // ###########################################################################################
        // *** WHOSE WORK THIS CARRIES (owner request, 2026-09-27). *** The server serialises these
        // and the maintainer reads them, so a field renamed on one side alone fails HERE rather
        // than by quietly showing an empty panel - the exact failure this test file exists for.
        // ###########################################################################################
        CarriedSubmission carried = Assert.Single(plan.Carrying!);

        Assert.Equal(7, carried.Id);
        Assert.Equal("hest@mailscan.dk", carried.ContactEmail);
        Assert.Equal("Corrected U8.", carried.Comment);
        Assert.Equal(ReviewWireContractTests.Decided, carried.DecidedUtc);
    }

    // ###########################################################################################
    // *** THE FILE TREE'S HALF OF THE PLAN (owner request, 2026-09-28). *** The unchanged paths and
    // where the two trees are published. A field renamed on one side would draw a tree of only what
    // changes, or refuse to open every file - silently; it fails here instead.
    // ###########################################################################################
    [Fact]
    public void A_production_plans_unchanged_files_and_data_addresses_read_back()
    {
        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan(ReviewWireContractTests.Answer(new ProductionPlanAnswer(
            "Commodore/C64/250407",
            "2026-September-25",
            new string('c', 64),
            IsAwaitingProduction: true,
            TouchesSharedFiles: false,
            CanPublish: true,
            Refusal: null,
            Approval: null,
            UnchangedCount: 2,
            Files: [],
            Problems: [],
            Removals: null,
            UnchangedFiles: ["Commodore/C64/250407/a.png", "Commodore/Shared files/b.pdf"],
            BetaDataUrl: "https://example.org/app-data-BETA/Data",
            ProductionDataUrl: "https://example.org/app-data/Data")));

        Assert.Equal(["Commodore/C64/250407/a.png", "Commodore/Shared files/b.pdf"], plan!.UnchangedFiles);
        Assert.Equal("https://example.org/app-data-BETA/Data", plan.BetaDataUrl);
        Assert.Equal("https://example.org/app-data/Data", plan.ProductionDataUrl);
    }

    // A submission's file tree: each entry's change and source travel as NAMES, and the hash a
    // submitted file is opened by must arrive intact.
    [Fact]
    public void A_submissions_file_tree_reads_back_field_for_field()
    {
        SubmissionFilesAnswer? answer = ReviewApiParser.ParseSubmissionFiles(ReviewWireContractTests.Answer(new SubmissionFilesAnswer(
            "Commodore/C128/310378",
            [
                new BoardFileEntry("Commodore/C128/310378/a.png", BoardFileChange.Added, BoardFileSource.Submission, new string('d', 64)),
                new BoardFileEntry("Commodore/C128/310378/Data.xlsx", BoardFileChange.Changed, BoardFileSource.Beta, WrittenOnApproval: true),
                new BoardFileEntry("Commodore/C128/310378/old.png", BoardFileChange.Removed, BoardFileSource.Beta),
                new BoardFileEntry("Commodore/C128/310378/new.json", BoardFileChange.Added, BoardFileSource.NotWrittenYet, WrittenOnApproval: true)
            ],
            "https://example.org/app-data-BETA/Data")));

        Assert.NotNull(answer);
        Assert.Equal("Commodore/C128/310378", answer!.BoardId);
        Assert.Equal("https://example.org/app-data-BETA/Data", answer.BetaDataUrl);
        Assert.Equal(
            [
                new BoardFileEntry("Commodore/C128/310378/a.png", BoardFileChange.Added, BoardFileSource.Submission, new string('d', 64)),
                new BoardFileEntry("Commodore/C128/310378/Data.xlsx", BoardFileChange.Changed, BoardFileSource.Beta, WrittenOnApproval: true),
                new BoardFileEntry("Commodore/C128/310378/old.png", BoardFileChange.Removed, BoardFileSource.Beta),
                new BoardFileEntry("Commodore/C128/310378/new.json", BoardFileChange.Added, BoardFileSource.NotWrittenYet, WrittenOnApproval: true)
            ],
            answer.Files);

        // Names, not numbers, on the wire - an enum reordered on one side cannot change a meaning.
        Assert.Contains("\"notWrittenYet\"", ReviewWireContractTests.Answer(new SubmissionFilesAnswer(
            "x", [new BoardFileEntry("x/a", BoardFileChange.Added, BoardFileSource.NotWrittenYet)])), StringComparison.OrdinalIgnoreCase);
    }

    // An older server sends no `carrying` at all, which must read as "none named" rather than as a
    // parse failure that loses the whole plan.
    [Fact]
    public void A_production_plan_from_a_server_that_sends_no_carrying_still_reads()
    {
        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan(ReviewWireContractTests.Answer(new ProductionPlanAnswer(
            "Commodore/C64/250407",
            "2026-September-25",
            new string('c', 64),
            IsAwaitingProduction: true,
            TouchesSharedFiles: false,
            CanPublish: true,
            Refusal: null,
            Approval: null,
            UnchangedCount: 17,
            Files: [],
            Problems: [],
            Removals: null,
            Carrying: null)));

        Assert.NotNull(plan);
        Assert.Empty(plan!.Carrying!);
    }

    // ###########################################################################################
    // Rolling a BETA board back (2026-09-27). The server serialises these and the maintainer reads
    // them, so a renamed field fails HERE rather than by showing an empty confirmation over an
    // operation that takes a board out of BETA.
    // ###########################################################################################
    [Fact]
    public void A_beta_rollback_plan_and_its_result_read_back_field_for_field()
    {
        BetaRollbackPlanView? plan = ReviewApiParser.ParseBetaRollbackPlan(ReviewWireContractTests.Answer(
            new BetaRollbackPlanAnswer(
                "Commodore/C64/250407",
                "restore",
                Restored: ["Commodore/C64/250407/sheet.png"],
                Removed: ["Commodore/C64/250407/new.png"],
                Returning: [new CarriedSubmission(7, "hest@mailscan.dk", "Corrected U8.", ReviewWireContractTests.Decided)],
                SharedRestored: ["Commodore/Shared files/6526.png"])));

        Assert.NotNull(plan);
        Assert.Equal(BetaRollbackKind.RestoreFromProduction, plan!.Kind);
        Assert.Equal(["Commodore/C64/250407/sheet.png"], plan.Restored);
        Assert.Equal(["Commodore/C64/250407/new.png"], plan.Removed);
        Assert.Equal("hest@mailscan.dk", Assert.Single(plan.Returning).ContactEmail);

        // The shared files (code review, 2026-09-27) - a renamed field here would show a
        // confirmation that silently leaves out changes reaching every board citing them.
        Assert.Equal(["Commodore/Shared files/6526.png"], plan.SharedRestored);

        BetaRollbackResult? done = ReviewApiParser.ParseBetaRollback(ReviewWireContractTests.Answer(
            new BetaRollbackAnswer("Commodore/C64/250407", "remove", FilesRestored: 0, FilesRemoved: 4, SubmissionsReturned: 2)));

        Assert.NotNull(done);
        Assert.Equal(BetaRollbackKind.RemoveFromBeta, done!.Kind);
        Assert.Equal(4, done.FilesRemoved);
        Assert.Equal(2, done.SubmissionsReturned);
        Assert.False(done.Rejected);

        // ###########################################################################################
        // Beta > Prod's "Reject" (2026-09-28): the answer says it REJECTED. A renamed field would read
        // as false, and the Maintainer tab would tell the maintainer the server pushed back instead.
        // ###########################################################################################
        BetaRollbackResult? rejected = ReviewApiParser.ParseBetaRollback(ReviewWireContractTests.Answer(
            new BetaRollbackAnswer("Commodore/C64/250407", "restore", FilesRestored: 1, FilesRemoved: 0, SubmissionsReturned: 1, Rejected: true)));

        Assert.True(rejected!.Rejected);
    }

    // ###########################################################################################
    // *** AN UNREADABLE KIND MUST NEVER BE TAKEN AS "remove". *** That one deletes a board from
    // BETA, so the parser defaults to the restore - the safe half - rather than to whichever the
    // enum happens to declare first.
    // ###########################################################################################
    [Fact]
    public void An_unknown_rollback_kind_is_read_as_the_safe_one()
    {
        BetaRollbackPlanView? plan = ReviewApiParser.ParseBetaRollbackPlan(
            """{"boardId":"Commodore/C64/250407","kind":"something-else"}""");

        Assert.NotNull(plan);
        Assert.Equal(BetaRollbackKind.RestoreFromProduction, plan!.Kind);
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

    // ###########################################################################################
    // Rebuilding both trees' manifests (owner request, 2026-10-01). The field this would most
    // easily get wrong is `entries`: it is a NUMBER that carries -1 for a failure, so a parser
    // defaulting a missing or renamed field to 0 would report a failed rebuild as "0 files listed"
    // - a success, in the one place the administrator is looking to see whether it worked.
    //
    // A SKIPPED tree (one this server has not configured) must also survive the round trip as
    // skipped rather than as a failure; the two look the same to anyone reading only the count.
    // ###########################################################################################
    [Fact]
    public void A_manifest_rebuild_reads_back_each_trees_count_its_skip_and_its_failure()
    {
        ManifestRebuildResult? result = ReviewApiParser.ParseManifestRebuild(ReviewWireContractTests.Answer(
            new ManifestRebuildAnswer(
                "1 of 2 checksum manifests were rebuilt.",
                [
                    new ManifestRebuildEntry("beta", Skipped: false, Entries: 1180, Message: "The beta manifest was rebuilt: 1180 files listed."),
                    new ManifestRebuildEntry("production", Skipped: false, Entries: -1, Message: "The production manifest could NOT be rebuilt."),
                ])));

        Assert.Equal("1 of 2 checksum manifests were rebuilt.", result!.Headline);
        Assert.Equal(2, result.Trees.Count);

        Assert.Equal(("beta", false, 1180), (result.Trees[0].Tree, result.Trees[0].Skipped, result.Trees[0].Entries));
        Assert.Equal("The beta manifest was rebuilt: 1180 files listed.", result.Trees[0].Message);

        // -1 survives as -1, not as 0.
        Assert.Equal(-1, result.Trees[1].Entries);
        Assert.False(result.Trees[1].Skipped);

        ManifestRebuildResult? skipped = ReviewApiParser.ParseManifestRebuild(ReviewWireContractTests.Answer(
            new ManifestRebuildAnswer(
                "1 checksum manifest was rebuilt.",
                [new ManifestRebuildEntry("production", Skipped: true, Entries: 0, Message: "Not configured on this server.")])));

        Assert.True(skipped!.Trees[0].Skipped);
    }

    // ###########################################################################################
    // DELETING A BOARD (owner request, 2026-10-03). The fingerprint is the one that matters: renamed
    // on either side, every delete would send nothing back and be refused as "changed since it was
    // shown" - and the reason, which is the only thing the contributors are told.
    // ###########################################################################################
    [Fact]
    public void A_board_delete_arrives_with_its_fingerprint_and_reason()
    {
        BoardDeleteRequest received = ReviewWireContractTests.Received<BoardDeleteRequest>(
            new BoardDeleteRequest("Commodore/C64/999999", new string('f', 64), "It was a test board."));

        Assert.Equal("Commodore/C64/999999", received.BoardId);
        Assert.Equal(new string('f', 64), received.Fingerprint);
        Assert.Equal("It was a test board.", received.Reason);

        // The plan is asked for with the boards screen's own request.
        Assert.Equal(
            "Commodore/C64/999999",
            ReviewWireContractTests.Received<BoardDetailRequest>(new BoardDetailRequest("Commodore/C64/999999")).BoardId);
    }

    [Fact]
    public void A_board_delete_plan_reads_back_field_for_field()
    {
        BoardDeletePlanAnswer? plan = ReviewApiParser.ParseBoardDeletePlan(ReviewWireContractTests.Answer(
            new BoardDeletePlanAnswer(
                "Commodore/C64/999999",
                "Commodore",
                "C64",
                "999999",
                new string('f', 64),
                BetaFiles: 12,
                ProductionFiles: 11,
                ListedInBeta: true,
                ListedInProduction: false,
                HasRecord: true,
                Submissions: 5,
                Maintainers: 2,
                Invitations: 1,
                OpenSubmissions: [new BoardDeleteOpenSubmission(14, "returned", "anna@example.com", "Corrected U8.", ReviewWireContractTests.Decided)],
                BlockedBecause: "Another board uses a file.")));

        Assert.NotNull(plan);
        Assert.Equal(("Commodore/C64/999999", "Commodore", "C64", "999999"), (plan.BoardId, plan.Manufacturer, plan.Hardware, plan.Board));
        Assert.Equal(new string('f', 64), plan.Fingerprint);
        Assert.Equal((12, 11, true, false, true), (plan.BetaFiles, plan.ProductionFiles, plan.ListedInBeta, plan.ListedInProduction, plan.HasRecord));
        Assert.Equal((5, 2, 1), (plan.Submissions, plan.Maintainers, plan.Invitations));
        Assert.Equal("Another board uses a file.", plan.BlockedBecause);

        BoardDeleteOpenSubmission open = Assert.Single(plan.OpenSubmissions);
        Assert.Equal((14L, "returned", "anna@example.com", "Corrected U8."), (open.Id, open.State, open.Contributor, open.Summary));
        Assert.Equal(ReviewWireContractTests.Decided, open.CreatedUtc);

        // Nothing blocking is left out on the wire, and reads as nothing blocking.
        BoardDeletePlanAnswer? clear = ReviewApiParser.ParseBoardDeletePlan(ReviewWireContractTests.Answer(
            new BoardDeletePlanAnswer("Commodore/C64/999999", "Commodore", "C64", "999999", "f", 0, 0, false, false, true, 3, 0, 0, [])));

        Assert.Null(clear!.BlockedBecause);
        Assert.Empty(clear.OpenSubmissions);
    }

    [Fact]
    public void A_board_delete_reads_back_what_went()
    {
        BoardDeleteAnswer? answer = ReviewApiParser.ParseBoardDelete(ReviewWireContractTests.Answer(
            new BoardDeleteAnswer("Commodore/C64/999999", BetaFilesRemoved: 12, ProductionFilesRemoved: 11, SubmissionsDeleted: 5, ContributorsMailed: 2)));

        Assert.Equal("Commodore/C64/999999", answer!.BoardId);
        Assert.Equal((12, 11, 5, 2), (answer.BetaFilesRemoved, answer.ProductionFilesRemoved, answer.SubmissionsDeleted, answer.ContributorsMailed));
    }

    // The two answers that were shared records already, read through the server's settings too -
    // including a new board's table, whose published side is null and so left out entirely.
    [Fact]
    public void The_table_and_the_unused_file_list_read_back_through_the_servers_settings()
    {
        var rows = new SubmissionRows();
        rows.Components.Add(new ComponentEntry { BoardLabel = "U8" });

        ReviewTableData? table = ReviewApiParser.ParseTable(ReviewWireContractTests.Answer(new ReviewTableData(4, null, rows)));

        UnusedFileListing? listing = ReviewApiParser.ParseUnusedFiles(ReviewWireContractTests.Answer(
            new UnusedFileListing("production", true, [], 2, 30, 900, [new UnusedFileEntry("a/old.png", 1234)], "https://example.org/Data/")));

        Assert.Equal(4, table!.Version);
        Assert.Null(table.Published);
        Assert.Equal("U8", Assert.Single(table.Submitted.Components).BoardLabel);

        Assert.Equal("production", listing!.Tree);
        Assert.Equal(1234, Assert.Single(listing.Files).SizeBytes);

        // Where the tree is published, for opening a file from the list's tree (2026-10-04).
        Assert.Equal("https://example.org/Data/", listing.PublicDataUrl);
    }

    // ###########################################################################################
    // THE THREE LISTS (code review, 2026-09-25): the production list, the administrator's boards
    // and accounts. The server answered them as anonymous objects and the Maintainer tab read them by
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

        ProductionBoardRow row = Assert.Single(list.Boards);
        Assert.Equal("Commodore/C64/250407", row.BoardId);
        Assert.Equal("Commodore", row.Manufacturer);
        Assert.Equal("C64", row.Hardware);
        Assert.Equal("250407", row.Board);
        Assert.Equal("2026-September-25", row.BetaRevision);
        Assert.Equal(betaHash, row.BetaContentHash);
        Assert.Equal("2026-August-21", row.ProductionRevision);
        Assert.Equal(published, row.ProductionPublishedUtc);

        // Not said by the server - read as not said, which the BETA badge counts as yours.
        Assert.Null(row.AwaitsYou);
    }

    // ###########################################################################################
    // *** THE BETA BADGE'S FLAG (2026-09-27). *** Renamed on either side, a board this account had
    // already approved would read as waiting for it again, and the badge would count it for ever.
    // ###########################################################################################
    [Fact]
    public void Whether_a_waiting_board_waits_for_you_reads_back()
    {
        ProductionListResponse? list = ReviewApiParser.ParseProductionList(ReviewWireContractTests.Answer(
            new ProductionListAnswer(
                true,
                [
                    new ProductionListEntry("Commodore/C64/250407", "Commodore", "C64", "250407", null, new string('c', 64), null, null, AwaitsYou: false),
                    new ProductionListEntry("Commodore/C128/310378", "Commodore", "C128", "310378", null, new string('d', 64), null, null, AwaitsYou: true)
                ])));

        Assert.Equal([false, true], list!.Boards.Select(row => row.AwaitsYou));
    }

    // ###########################################################################################
    // *** WAITING FOR THE ADMINISTRATOR (2026-10-05). *** Renamed on either side, a maintainer's
    // row would fall back to "with the other approver" - the words for an approval they never gave.
    // Absent, it reads as not said.
    // ###########################################################################################
    [Fact]
    public void Whether_a_waiting_board_waits_for_the_administrator_reads_back()
    {
        ProductionListResponse? list = ReviewApiParser.ParseProductionList(ReviewWireContractTests.Answer(
            new ProductionListAnswer(
                true,
                [
                    new ProductionListEntry("Commodore/C64/250407", "Commodore", "C64", "250407", null, new string('c', 64), null, null, AwaitsYou: false, WaitsForAdministrator: true),
                    new ProductionListEntry("Commodore/C128/310378", "Commodore", "C128", "310378", null, new string('d', 64), null, null, AwaitsYou: true, WaitsForAdministrator: false),
                    new ProductionListEntry("Amstrad/CPC/464", "Amstrad", "CPC", "464", null, new string('e', 64), null, null)
                ])));

        Assert.Equal([true, false, null], list!.Boards.Select(row => row.WaitsForAdministrator));
    }

    // ###########################################################################################
    // THE "BOARDS" SCREEN (2026-09-27): the list, the detail request, and the detail - every
    // field, since each one is on screen and a renamed one would be blank there in silence.
    // ###########################################################################################
    [Fact]
    public void The_boards_list_reads_back_field_for_field()
    {
        BoardOverviewEntry sent = new(
            "Commodore/C64/250407", "Commodore", "C64", "250407",
            InBeta: true, InProduction: false, IsAwaitingProduction: true, IsAccepting: false,
            BetaRevision: "2026-September-25", ProductionRevision: "2026-May-14",
            ProductionPublishedUtc: ReviewWireContractTests.Decided, MaintainerCount: 2,
            ViewsLast30Days: 48,
            SubmissionCounts: new BoardSubmissionCounts(Total: 19, Waiting: 1, Rejected: 2, InBeta: 1, InStable: 15));

        BoardOverviewAnswer? answer = ReviewApiParser.ParseBoardOverview(
            ReviewWireContractTests.Answer(new BoardOverviewAnswer([sent])));

        Assert.Equal(sent, Assert.Single(answer!.Boards));
    }

    // How a board's submissions went (2026-10-09): absent from an older server, and then nothing is
    // said - never "0 submissions", which would be a claim about the board.
    [Fact]
    public void A_boards_submission_counts_stay_absent_from_an_older_server()
    {
        BoardOverviewEntry old = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0);

        Assert.Null(Assert.Single(ReviewApiParser.ParseBoardOverview(
            ReviewWireContractTests.Answer(new BoardOverviewAnswer([old])))!.Boards).SubmissionCounts);
    }

    // ###########################################################################################
    // Board views (2026-09-27): no count is NOT zero - absent from an older server, or when it could
    // not count, it stays null and the screen says nothing; a real 0 reads back as 0.
    // ###########################################################################################
    [Fact]
    public void A_boards_view_count_reads_back_and_its_absence_stays_absent()
    {
        BoardOverviewEntry counted = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0, ViewsLast30Days: 0);
        BoardOverviewEntry uncounted = counted with { BoardId = "Commodore/C128/310378", ViewsLast30Days = null };

        BoardOverviewAnswer? answer = ReviewApiParser.ParseBoardOverview(
            ReviewWireContractTests.Answer(new BoardOverviewAnswer([counted, uncounted])));

        Assert.Equal([(int?)0, null], answer!.Boards.Select(board => board.ViewsLast30Days));
    }

    [Fact]
    public void A_boards_view_statistics_read_back_field_for_field()
    {
        BoardOverviewEntry board = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0, 48);
        BoardViewStatistics views = new(
            12, 48, 310, 3, [new("DK", "Denmark", 120), new("DE", "Germany", 60)],
            [new(new DateOnly(2025, 10, 10), 2), new(new DateOnly(2026, 9, 26), 7)]);

        BoardDetailAnswer? read = ReviewApiParser.ParseBoardDetail(
            ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [], Views: views)));

        Assert.Equal(
            (views.Last7Days, views.Last30Days, views.Last365Days, views.FromBetaLast30Days),
            (read!.Views!.Last7Days, read.Views.Last30Days, read.Views.Last365Days, read.Views.FromBetaLast30Days));
        Assert.Equal(views.TopCountries, read.Views.TopCountries);
        Assert.Equal(views.Daily!, read.Views.Daily!);
        Assert.Equal(48, read.Board.ViewsLast30Days);

        Assert.Null(ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [])))!.Views);
    }

    // With no production tree the server says nothing about production - and that is kept, not
    // read as "not in production".
    [Fact]
    public void A_board_the_server_could_not_look_up_in_production_reads_back_as_not_said()
    {
        BoardOverviewAnswer? answer = ReviewApiParser.ParseBoardOverview(ReviewWireContractTests.Answer(
            new BoardOverviewAnswer([new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, null, false, true, null, null, null, 0)])));

        Assert.Null(Assert.Single(answer!.Boards).InProduction);
    }

    // ###########################################################################################
    // BETA's content hash (2026-10-04): what tells the Boards screen its open table and file list
    // are out of date (QueueRefreshRules.BetaBoardChanged). Arriving empty, the screen would fall
    // back on the revision date - which a second publish the same day does not move.
    // ###########################################################################################
    [Fact]
    public void A_boards_BETA_content_hash_reads_back_in_the_list_and_in_the_detail()
    {
        string hash = new('c', 64);
        BoardOverviewEntry board = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, "r2", "r1", null, 1, BetaContentHash: hash);

        BoardOverviewAnswer? list = ReviewApiParser.ParseBoardOverview(ReviewWireContractTests.Answer(new BoardOverviewAnswer([board])));
        BoardDetailAnswer? detail = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [])));

        Assert.Equal(hash, Assert.Single(list!.Boards).BetaContentHash);
        Assert.Equal(hash, detail!.Board.BetaContentHash);

        // None from the server (a board nothing has published) is none, not "".
        Assert.Null(Assert.Single(ReviewApiParser.ParseBoardOverview(ReviewWireContractTests.Answer(
            new BoardOverviewAnswer([board with { BetaContentHash = null }])))!.Boards).BetaContentHash);
    }

    [Fact]
    public void A_board_detail_request_arrives_with_its_id()
    {
        Assert.Equal(
            "Commodore/C64/250407",
            ReviewWireContractTests.Received<BoardDetailRequest>(new BoardDetailRequest("Commodore/C64/250407")).BoardId);
    }

    [Fact]
    public void A_boards_detail_reads_back_field_for_field()
    {
        BoardDetailAnswer sent = new(
            new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, "r2", "r1", ReviewWireContractTests.Decided, 1),
            [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
            [new BoardContributorEntry("hest@mailscan.dk", "Hest", Accepted: 3, Waiting: 1, ChangesRequested: 2, Rejected: 4, LastSubmittedUtc: ReviewWireContractTests.Decided)],
            [new BoardSubmissionEntry(41, "hest@mailscan.dk", "Corrected U8.", "returned", ReviewWireContractTests.Decided, ReviewWireContractTests.Decided.AddDays(1), "U7 is the wrong revision.")]);

        BoardDetailAnswer? read = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(sent));

        Assert.NotNull(read);
        Assert.Equal(sent.Board, read!.Board);
        Assert.Equal(sent.Maintainers, read.Maintainers);
        Assert.Equal(sent.Contributors, read.Contributors);
        Assert.Equal(sent.Submissions, read.Submissions);
        Assert.False(read.AddressesHidden);
    }

    // ###########################################################################################
    // NO ADDRESSES FOR A BOARD THE ACCOUNT DOES NOT MAINTAIN (owner request, 2026-10-05). The
    // server's word for it must arrive, or the tab reads a contributor without an account as one who
    // gave no address - and a maintainer's empty address must stay empty, not become "(no address)".
    // ###########################################################################################
    [Fact]
    public void A_boards_detail_sent_without_addresses_says_so_and_reads_back_without_them()
    {
        BoardDetailAnswer sent = new(
            new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, "r2", "r1", ReviewWireContractTests.Decided, 1),
            [new PoolMaintainerEntry(7, "Anna", string.Empty)],
            [new BoardContributorEntry(null, null, Accepted: 1, Waiting: 0, ChangesRequested: 0, Rejected: 0, LastSubmittedUtc: ReviewWireContractTests.Decided)],
            [new BoardSubmissionEntry(41, null, "Corrected U8.", "merged", ReviewWireContractTests.Decided, null, null)],
            AddressesHidden: true);

        BoardDetailAnswer? read = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(sent));

        Assert.True(read!.AddressesHidden);
        Assert.Equal(string.Empty, Assert.Single(read.Maintainers).Email);
        Assert.Null(Assert.Single(read.Contributors).Email);
        Assert.Null(Assert.Single(read.Submissions).ContactEmail);
    }

    // ###########################################################################################
    // WHETHER BETA'S DATA MAY BE CHANGED, IN THE DETAIL (code review, 2026-10-09) - the Boards screen
    // says why not above every view from it. Both answers arrive as sent, and an older server's detail
    // (no such fields) reads as not known - never as "may be changed".
    // ###########################################################################################
    [Fact]
    public void A_boards_detail_reads_back_whether_its_BETA_data_may_be_changed_and_why_not()
    {
        var board = new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, true, true, "r2", "r1", ReviewWireContractTests.Decided, 1);

        BoardDetailAnswer? readOnly = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(
            new BoardDetailAnswer(board, [], [], [], MayEdit: false, MayNotEditReason: "It is waiting in BETA for the stable source.")));

        Assert.Equal((false, "It is waiting in BETA for the stable source."), (readOnly!.MayEdit, readOnly.MayNotEditReason));

        BoardDetailAnswer? editable = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(
            new BoardDetailAnswer(board, [], [], [], MayEdit: true)));

        Assert.Equal((true, null), (editable!.MayEdit, editable.MayNotEditReason));

        BoardDetailAnswer? older = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [])));

        Assert.Equal((null, null), (older!.MayEdit, older.MayNotEditReason));
    }

    // ###########################################################################################
    // A BOARD'S BOARD DATA AND FILES (2026-10-03): the table the Boards screen opens on, the edit
    // sent back from it - fingerprint, description and rows - the submission it was queued as, and
    // the file listing. The fingerprint above all: arriving empty, every edit would be refused as
    // "BETA changed since you opened the table".
    // ###########################################################################################
    [Fact]
    public void A_boards_table_reads_back_with_its_fingerprint_rows_and_whether_it_may_be_changed()
    {
        var rows = new SubmissionRows { RevisionDate = "2026-August-21" };
        rows.Components.Add(new ComponentEntry { BoardLabel = "U8", PartNumber = "251715-01" });
        rows.KiCadCalibrations.Add(new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "board", OffsetX = 2 });

        BoardTableAnswer? read = ReviewApiParser.ParseBoardTable(ReviewWireContractTests.Answer(
            new BoardTableAnswer("Commodore/C64/250407", new string('f', 64), rows, MayEdit: false, "Not yours.", "https://example.org/beta")));

        Assert.NotNull(read);
        Assert.Equal("Commodore/C64/250407", read!.BoardId);
        Assert.Equal(new string('f', 64), read.Fingerprint);
        Assert.False(read.MayEdit);
        Assert.Equal("Not yours.", read.MayNotEditReason);
        Assert.Equal("https://example.org/beta", read.BetaDataUrl);
        Assert.Equal("251715-01", Assert.Single(read.Rows.Components).PartNumber);
        Assert.Equal("board", Assert.Single(read.Rows.KiCadCalibrations).CadName);
        Assert.Equal("2026-August-21", read.Rows.RevisionDate);

        Assert.True(ReviewApiParser.ParseBoardTable(ReviewWireContractTests.Answer(
            new BoardTableAnswer("Commodore/C64/250407", "f", rows, MayEdit: true)))!.MayEdit);
    }

    // The edit and its check carry one request: the fingerprint, the reason, the rows - and the list
    // of removals the check answered, which the publish must find unchanged.
    [Fact]
    public void A_boards_edit_arrives_with_its_fingerprint_reason_rows_and_the_removals_shown()
    {
        var rows = new SubmissionRows();
        rows.Components.Add(new ComponentEntry { BoardLabel = "U8", PartNumber = "251715-02" });

        BoardEditRequest received = ReviewWireContractTests.Received<BoardEditRequest>(
            new BoardEditRequest("Commodore/C64/250407", new string('f', 64), "Corrected U8.", rows, ["Commodore/C64/250407/manual.pdf"]));

        Assert.Equal("Commodore/C64/250407", received.BoardId);
        Assert.Equal(new string('f', 64), received.Fingerprint);
        Assert.Equal("Corrected U8.", received.Summary);
        Assert.Equal("251715-02", Assert.Single(received.Rows!.Components).PartNumber);
        Assert.Equal(["Commodore/C64/250407/manual.pdf"], received.ExpectedRemovals);
    }

    // ###########################################################################################
    // The check's answer: the files a publish would remove - read back as the list, an empty one as
    // empty. An answer WITHOUT the list is unreadable, never "nothing to remove": the maintainer must
    // be shown what goes, and the server refuses a publish whose list differs.
    // ###########################################################################################
    [Fact]
    public void A_boards_edit_check_reads_back_as_the_files_it_would_remove()
    {
        BoardEditCheckAnswer? read = ReviewApiParser.ParseBoardEditCheck(ReviewWireContractTests.Answer(
            new BoardEditCheckAnswer(["Commodore/C64/250407/manual.pdf"])));

        Assert.Equal(["Commodore/C64/250407/manual.pdf"], read!.Removals);
        Assert.Empty(ReviewApiParser.ParseBoardEditCheck(ReviewWireContractTests.Answer(new BoardEditCheckAnswer([])))!.Removals);
        Assert.Null(ReviewApiParser.ParseBoardEditCheck("{}"));
    }

    // ###########################################################################################
    // The publish's answer, both ways it can go: in BETA (revision, removed files) - or made into a
    // submission and not published, with the server's reason. And its warnings - an error never
    // arrives here, it refuses instead.
    // ###########################################################################################
    [Fact]
    public void A_published_boards_edit_reads_back_with_its_revision_removals_and_warnings()
    {
        BoardEditResult? read = ReviewApiParser.ParseBoardEdit(ReviewWireContractTests.Answer(
            new BoardEditAnswer(
                57,
                [ReviewWireContractTests.Finding(ValidationSeverity.Warning)],
                Published: true,
                Revision: "2026-October-03",
                RemovedFiles: ["Commodore/C64/250407/manual.pdf"])));

        Assert.NotNull(read);
        Assert.Equal(57, read!.SubmissionId);
        Assert.True(read.Published);
        Assert.Equal("2026-October-03", read.Revision);
        Assert.Equal(["Commodore/C64/250407/manual.pdf"], read.RemovedFiles);
        Assert.Null(read.NotPublishedReason);
        Assert.False(Assert.Single(read.Warnings).IsError);
        Assert.Equal("An image references component [U9].", read.Warnings[0].Message);
    }

    [Fact]
    public void A_boards_edit_made_but_not_published_reads_back_with_the_reason()
    {
        BoardEditResult? read = ReviewApiParser.ParseBoardEdit(ReviewWireContractTests.Answer(
            new BoardEditAnswer(58, [], NotPublishedReason: "BETA moved.")));

        Assert.NotNull(read);
        Assert.Equal(58, read!.SubmissionId);
        Assert.False(read.Published);
        Assert.Equal("BETA moved.", read.NotPublishedReason);
    }

    [Fact]
    public void A_boards_file_listing_reads_back_field_for_field()
    {
        BoardFilesAnswer? read = ReviewApiParser.ParseBoardFiles(ReviewWireContractTests.Answer(new BoardFilesAnswer(
            "Commodore/C64/250407",
            [
                // With its size (2026-10-04) - and the next without one, as an older server sends it.
                new BoardFileEntry("Commodore/C64/250407/manual.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta, SizeBytes: 5_242_880),
                new BoardFileEntry("Commodore/Shared files/74LS08.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta)
            ],
            "https://example.org/beta")));

        Assert.NotNull(read);
        Assert.Equal("Commodore/C64/250407", read!.BoardId);
        Assert.Equal("https://example.org/beta", read.BetaDataUrl);
        Assert.Equal(
            [
                new BoardFileEntry("Commodore/C64/250407/manual.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta, SizeBytes: 5_242_880),
                new BoardFileEntry("Commodore/Shared files/74LS08.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta)
            ],
            read.Files);
    }

    // ###########################################################################################
    // THE STABLE SOURCE'S BOARD DATA AND FILES (2026-10-04): the tree named in the request as the
    // server binds it - none (an older CRT) is BETA - and the stable source's address in both answers.
    // ###########################################################################################
    [Fact]
    public void A_board_request_names_its_tree_and_the_answers_carry_the_stable_address()
    {
        BoardDetailRequest stable = ReviewWireContractTests.Received<BoardDetailRequest>(
            new BoardDetailRequest("Commodore/C64/250407", DataTreeNames.Production));
        BoardDetailRequest older = ReviewWireContractTests.Received<BoardDetailRequest>(new BoardDetailRequest("Commodore/C64/250407"));

        Assert.True(DataTreeNames.IsProduction(stable.Tree));
        Assert.False(DataTreeNames.IsProduction(older.Tree));

        BoardTableAnswer? table = ReviewApiParser.ParseBoardTable(ReviewWireContractTests.Answer(new BoardTableAnswer(
            "Commodore/C64/250407", new string('f', 64), new SubmissionRows(), MayEdit: false, "Read-only.", BetaDataUrl: null,
            ProductionDataUrl: "https://example.org/app-data/Data/")));

        BoardFilesAnswer? files = ReviewApiParser.ParseBoardFiles(ReviewWireContractTests.Answer(new BoardFilesAnswer(
            "Commodore/C64/250407",
            [new BoardFileEntry("Commodore/C64/250407/manual.pdf", BoardFileChange.Unchanged, BoardFileSource.Production, SizeBytes: 10)],
            BetaDataUrl: null,
            ProductionDataUrl: "https://example.org/app-data/Data/")));

        Assert.Equal("https://example.org/app-data/Data/", table!.ProductionDataUrl);
        Assert.False(table.MayEdit);
        Assert.Equal("https://example.org/app-data/Data/", files!.ProductionDataUrl);
        Assert.Equal(BoardFileSource.Production, Assert.Single(files.Files).OpenFrom);
    }

    // The production plan's FileSizes (2026-10-04): path -> bytes, the keys as the paths are spelled
    // (no naming policy touches a dictionary key), read by the tab's hand-written plan parser.
    [Fact]
    public void A_production_plans_file_sizes_read_back_by_path()
    {
        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan(ReviewWireContractTests.Answer(new ProductionPlanAnswer(
            "Commodore/C128/310378",
            "2026-October-4",
            new string('c', 64),
            IsAwaitingProduction: true,
            TouchesSharedFiles: false,
            CanPublish: true,
            Refusal: null,
            Approval: null,
            UnchangedCount: 1,
            Files: [],
            Problems: [],
            Removals: null,
            FileSizes: new Dictionary<string, long>
            {
                ["Commodore/C128/310378/Scope baseline/U10_30b_NTSC.png"] = 36_147,
                ["Commodore/Shared files/Component images/6510.jpg"] = 46_182
            })));

        Assert.NotNull(plan!.FileSizes);
        Assert.Equal(36_147, plan.FileSizes!["Commodore/C128/310378/Scope baseline/U10_30b_NTSC.png"]);
        Assert.Equal(46_182, plan.FileSizes["Commodore/Shared files/Component images/6510.jpg"]);

        // An older server sends none: no sizes, not an empty claim.
        ProductionPlanView? older = ReviewApiParser.ParseProductionPlan(ReviewWireContractTests.Answer(new ProductionPlanAnswer(
            "Commodore/C128/310378", null, new string('c', 64), true, false, true, null, null, 0, [], [], null)));

        Assert.Null(older!.FileSizes);
    }

    // ###########################################################################################
    // Inviting a new maintainer by email (2026-09-27): the three requests as the server binds them,
    // the acceptance's answer, and the invitations on a board's detail.
    // ###########################################################################################
    [Fact]
    public void The_invitation_requests_arrive_as_sent()
    {
        MaintainerInviteRequest invite = ReviewWireContractTests.Received<MaintainerInviteRequest>(
            new MaintainerInviteRequest("Commodore/C64/250407", "new@example.com"));
        Assert.Equal("Commodore/C64/250407", invite.BoardId);
        Assert.Equal("new@example.com", invite.Email);

        Assert.Equal(3, ReviewWireContractTests.Received<InvitationWithdrawRequest>(new InvitationWithdrawRequest(3)).InvitationId);

        AcceptInvitationRequest accept = ReviewWireContractTests.Received<AcceptInvitationRequest>(
            new AcceptInvitationRequest("code", "Anna", "correct horse battery staple"));
        Assert.Equal(new AcceptInvitationRequest("code", "Anna", "correct horse battery staple"), accept);
    }

    // The board's history (2026-09-27), field for field - and absent from an older server stays absent.
    [Fact]
    public void A_boards_history_reads_back_field_for_field()
    {
        BoardOverviewEntry board = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0);
        BoardHistoryEntry[] history =
        [
            new(ReviewWireContractTests.Decided, BoardHistoryEvents.Decided, "Anna", 9, "rejected", "Wrong board."),
            new(ReviewWireContractTests.Decided.AddDays(-1), BoardHistoryEvents.Sent, "hest@mailscan.dk", 9, "Corrected U8."),
            new(ReviewWireContractTests.Decided.AddDays(-2), BoardHistoryEvents.MaintainerAdded, "admin@example.com", null, "anna@example.com"),
        ];

        BoardDetailAnswer? read = ReviewApiParser.ParseBoardDetail(
            ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [], null, history)));

        Assert.Equal(history, read!.History);
        Assert.Null(ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [])))!.History);
    }

    // ###########################################################################################
    // The maintainer's own account (2026-10-03): the four changes arrive as the server binds them,
    // and the account and a change's answer read back field for field - including the answer to a
    // new address, which carries NO account (it has not changed yet) and says a code was sent.
    // ###########################################################################################
    [Fact]
    public void The_account_changes_arrive_as_the_server_binds_them()
    {
        Assert.Equal(
            new ChangeNameRequest("Dennis H"),
            ReviewWireContractTests.Received<ChangeNameRequest>(new ChangeNameRequest("Dennis H")));
        Assert.Equal(
            new ChangeEmailRequest("bench@example.com"),
            ReviewWireContractTests.Received<ChangeEmailRequest>(new ChangeEmailRequest("bench@example.com")));
        Assert.Equal(
            new ConfirmEmailChangeRequest("iAXr2z0PPffPwpmzHR-bOFo5ZPCPfHEK1hZbFAKAoYQ"),
            ReviewWireContractTests.Received<ConfirmEmailChangeRequest>(new ConfirmEmailChangeRequest("iAXr2z0PPffPwpmzHR-bOFo5ZPCPfHEK1hZbFAKAoYQ")));
        Assert.Equal(
            new ChangePasswordRequest("a new one here"),
            ReviewWireContractTests.Received<ChangePasswordRequest>(new ChangePasswordRequest("a new one here")));
    }

    [Fact]
    public void The_account_reads_back_field_for_field()
    {
        var sent = new AccountAnswer(
            7, "dh@example.com", "Dennis", true, true, ["Commodore/C128/310378", "Commodore/C64/250407"],
            ReviewWireContractTests.Decided.AddYears(-1), ReviewWireContractTests.Decided);

        AccountAnswer? read = ReviewApiParser.ParseAccount(ReviewWireContractTests.Answer(sent));

        Assert.NotNull(read);
        Assert.Equal(sent.Id, read!.Id);
        Assert.Equal(sent.Email, read.Email);
        Assert.Equal(sent.DisplayName, read.DisplayName);
        Assert.Equal(sent.IsVerified, read.IsVerified);
        Assert.Equal(sent.IsAdministrator, read.IsAdministrator);
        Assert.Equal(sent.MaintainerOf, read.MaintainerOf);
        Assert.Equal(sent.CreatedUtc, read.CreatedUtc);
        Assert.Equal(sent.LastLoginUtc, read.LastLoginUtc);

        // Never signed in since it was made: no last login, and none invented.
        Assert.Null(ReviewApiParser.ParseAccount(ReviewWireContractTests.Answer(sent with { LastLoginUtc = null }))!.LastLoginUtc);
    }

    [Fact]
    public void A_changes_answer_reads_back_with_its_account_or_without_one_while_a_code_is_on_its_way()
    {
        var account = new AccountAnswer(7, "bench@example.com", "Dennis", true, false, [], ReviewWireContractTests.Decided);

        AccountChangeAnswer? changed = ReviewApiParser.ParseAccountChange(
            ReviewWireContractTests.Answer(new AccountChangeAnswer("Your email address is now bench@example.com.", account)));

        Assert.Equal("Your email address is now bench@example.com.", changed!.Message);
        Assert.Equal("bench@example.com", changed.Account!.Email);
        Assert.False(changed.CodeSent);

        AccountChangeAnswer? waiting = ReviewApiParser.ParseAccountChange(
            ReviewWireContractTests.Answer(new AccountChangeAnswer("A code is on its way.", null, CodeSent: true)));

        Assert.True(waiting!.CodeSent);
        Assert.Null(waiting.Account);
    }

    [Fact]
    public void An_accepted_invitation_reads_back_field_for_field()
    {
        AcceptInvitationAnswer sent = new("anna@example.com", ["Commodore/C64/250407", "Commodore/C128/310378"], "Your account is ready.");

        AcceptInvitationAnswer? read = ReviewApiParser.ParseAcceptInvitation(ReviewWireContractTests.Answer(sent));

        Assert.NotNull(read);
        Assert.Equal(sent.Email, read!.Email);
        Assert.Equal(sent.BoardIds, read.BoardIds);
        Assert.Equal(sent.Message, read.Message);
    }

    // For an administrator the detail carries the open invitations; for anybody else there are none
    // (null), which must read back as none rather than as an empty list claimed by the server.
    [Fact]
    public void A_boards_invitations_read_back_and_their_absence_stays_absent()
    {
        BoardOverviewEntry board = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0);
        var invitation = new MaintainerInvitationEntry(3, "new@example.com", ReviewWireContractTests.Decided, ReviewWireContractTests.Decided.AddDays(14));

        BoardDetailAnswer? forAdmin = ReviewApiParser.ParseBoardDetail(
            ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [], [invitation])));
        Assert.Equal([invitation], forAdmin!.Invitations);

        BoardDetailAnswer? forMaintainer = ReviewApiParser.ParseBoardDetail(
            ReviewWireContractTests.Answer(new BoardDetailAnswer(board, [], [], [])));
        Assert.Null(forMaintainer!.Invitations);
    }

    // ###########################################################################################
    // A new board's place in CRT's drop-down lists (2026-09-27): the lists and the boards not in
    // them, the placement a maintainer saves, and what the server says it did.
    // ###########################################################################################
    [Fact]
    public void The_drop_down_listing_reads_back_field_for_field()
    {
        BoardListingAnswer sent = new(
            HasList: true,
            [
                new BoardListingRow("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                new BoardListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
            ],
            [
                new UnlistedBoardEntry(
                    "Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128",
                    InBeta: true,
                    CanPlace: true,
                    new BoardPlacement("Commodore 128", "310378 Open128", "Open-source replica.", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),

                    // The suggestion carries the contributor's notes from "Create board" (2026-10-05)
                    // - dropped on either side, the maintainer's placement would start empty again.
                    new BoardPlacement("Commodore 128", "310378 Open128", "Sent by its contributor.", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx")),

                // Never placed, and placed FIRST by its suggestion (no row above it).
                new UnlistedBoardEntry(
                    "Amstrad/CPC 6128/MC0020", "Amstrad", "CPC 6128", "MC0020",
                    InBeta: false,
                    CanPlace: false,
                    Placement: null,
                    new BoardPlacement("CPC 6128", "MC0020", string.Empty, null)),
            ]);

        BoardListingAnswer? read = ReviewApiParser.ParseBoardListing(ReviewWireContractTests.Answer(sent));

        Assert.NotNull(read);
        Assert.True(read!.HasList);
        Assert.Equal(sent.Rows, read.Rows);
        Assert.Equal(sent.Unlisted, read.Unlisted);
    }

    [Fact]
    public void A_placement_arrives_as_the_maintainer_arranged_it()
    {
        SetPlacementRequest received = ReviewWireContractTests.Received<SetPlacementRequest>(new SetPlacementRequest(
            "Commodore/C128/310378 Open128",
            "Commodore 128",
            "310378 Open128",
            "Open-source replica.",
            "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"));

        Assert.Equal("Commodore/C128/310378 Open128", received.BoardId);
        Assert.Equal("Commodore 128", received.HardwareName);
        Assert.Equal("310378 Open128", received.BoardName);
        Assert.Equal("Open-source replica.", received.Notes);
        Assert.Equal("Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx", received.AfterExcelDataFile);

        // Placed first: no row above it, which must arrive as nothing rather than as "".
        Assert.Null(ReviewWireContractTests.Received<SetPlacementRequest>(
            new SetPlacementRequest("Amstrad/CPC 6128/MC0020", "CPC 6128", "MC0020", string.Empty, null)).AfterExcelDataFile);
    }

    // ###########################################################################################
    // Whether each source's drop-down list names a board (2026-10-04) arrives as sent - and NULL,
    // a list the server could not read, stays null rather than turning into "not listed".
    // ###########################################################################################
    [Fact]
    public void Whether_each_drop_down_list_names_a_board_arrives_including_not_known()
    {
        BoardOverviewEntry listed = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1, ListedInBeta: false, ListedInStable: true);
        BoardOverviewEntry unknown = listed with { BoardId = "Commodore/C128/310378", ListedInBeta = null, ListedInStable = null };

        BoardOverviewAnswer? read = ReviewApiParser.ParseBoardOverview(ReviewWireContractTests.Answer(new BoardOverviewAnswer([listed, unknown])));

        Assert.Equal((false, true), (read!.Boards[0].ListedInBeta, read.Boards[0].ListedInStable));
        Assert.Equal((null, null), (read.Boards[1].ListedInBeta, read.Boards[1].ListedInStable));
        Assert.Equal(listed, read.Boards[0]);
    }

    // ###########################################################################################
    // Account > "Order of boards" (2026-10-04): the order arrives as sent, every id in its place, and
    // what the server did to each list reads back - including null for a server with no stable
    // source, which must not turn into "not changed".
    // ###########################################################################################
    [Fact]
    public void The_order_of_boards_arrives_in_order_and_what_was_done_reads_back()
    {
        string[] order = ["ZX Spectrum/Spectrum 16K-48K/Issue 4B", "Commodore/C64/250407", "Commodore/C128/310378"];

        Assert.Equal(order, ReviewWireContractTests.Received<BoardOrderRequest>(new BoardOrderRequest(order)).BoardIds);

        BoardOrderAnswer both = new(true, true, null);
        BoardOrderAnswer betaOnly = new(true, null, null);
        BoardOrderAnswer failed = new(true, false, "The stable source's list could not be changed: locked.");

        Assert.Equal(both, ReviewApiParser.ParseBoardOrder(ReviewWireContractTests.Answer(both)));
        Assert.Equal(betaOnly, ReviewApiParser.ParseBoardOrder(ReviewWireContractTests.Answer(betaOnly)));
        Assert.Equal(failed, ReviewApiParser.ParseBoardOrder(ReviewWireContractTests.Answer(failed)));
    }

    // ###########################################################################################
    // What a submission changed as it went into BETA (2026-10-04) rides on its entry in the board's
    // detail - every list and field arriving as the server recorded it, and none for an entry the
    // server has no record for.
    // ###########################################################################################
    [Fact]
    public void What_a_submission_changed_arrives_with_its_entry()
    {
        var changes = new SubmissionChanges(
            false,
            [
                new SectionChanges(
                    BoardWorkbookSchema.SheetComponents, 2, 1, 1, 1,
                    ["U7", "U9"],
                    [new ChangedRowFact("U8", [BoardWorkbookSchema.ColPartNumber])],
                    ["R3"],
                    [new RenamedRowFact("C10", "C51")])
            ],
            new FileChanges(1, 0, 1, ["Commodore/C64/250407/U7.png"], [], ["Commodore/C64/250407/old.pdf"]));

        BoardOverviewEntry board = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1);

        BoardDetailAnswer? read = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(new BoardDetailAnswer(
            board,
            [],
            [],
            [
                new BoardSubmissionEntry(9, "a@example.com", "Fix.", "merged", ReviewWireContractTests.Decided, ReviewWireContractTests.Decided, null, Changes: changes),
                new BoardSubmissionEntry(8, "a@example.com", "Older.", "merged", ReviewWireContractTests.Decided, ReviewWireContractTests.Decided, null)
            ])));

        SubmissionChanges back = read!.Submissions.Single(entry => entry.Id == 9).Changes!;

        Assert.False(back.IsNewBoard);

        SectionChanges section = Assert.Single(back.Sections);
        Assert.Equal((BoardWorkbookSchema.SheetComponents, 2, 1, 1, 1), (section.Section, section.AddedCount, section.ChangedCount, section.RemovedCount, section.RenamedCount));
        Assert.Equal(["U7", "U9"], section.Added);
        Assert.Equal("U8", Assert.Single(section.Changed).Row);
        Assert.Equal([BoardWorkbookSchema.ColPartNumber], section.Changed[0].Fields);
        Assert.Equal(["R3"], section.Removed);
        Assert.Equal(new RenamedRowFact("C10", "C51"), Assert.Single(section.Renamed));
        Assert.Equal((1, 0, 1), (back.Files.AddedCount, back.Files.ReplacedCount, back.Files.RemovedCount));
        Assert.Equal(["Commodore/C64/250407/U7.png"], back.Files.Added);
        Assert.Equal(["Commodore/C64/250407/old.pdf"], back.Files.Removed);

        Assert.Null(read.Submissions.Single(entry => entry.Id == 8).Changes);
    }

    [Fact]
    public void A_saved_placement_reads_back_with_what_the_server_did()
    {
        SetPlacementAnswer sent = new(
            new BoardPlacement("Commodore 128", "310378 Open128", "Open-source replica.", null),
            ListedInBeta: true,
            "Placed, and added to BETA's drop-down lists.");

        Assert.Equal(sent, ReviewApiParser.ParseSetPlacement(ReviewWireContractTests.Answer(sent)));
    }

    // ###########################################################################################
    // The queue (2026-09-26 - an anonymous object until then), with the two badges the queue list
    // shows: whether the board is published, and whether it waits for this account.
    // ###########################################################################################
    [Fact]
    public void The_queue_reads_back_field_for_field_with_its_badges()
    {
        var created = new DateTimeOffset(2026, 9, 26, 8, 0, 0, TimeSpan.Zero);

        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue(ReviewWireContractTests.Answer(
            new ReviewQueueAnswer(
                CanPublish: true,
                IsAdministrator: true,
                Count: 2,
                Submissions:
                [
                    new ReviewQueueEntry(4, "Commodore/C64/250407", "approved", "Corrected U8.", "c@example.com", "r1", created, null, TouchesSharedFiles: true, IsNewBoard: false, AwaitsYou: true),
                    new ReviewQueueEntry(5, "Amstrad/CPC464/Z70200", "pending", null, null, "", created, null, TouchesSharedFiles: false, IsNewBoard: true, AwaitsYou: false)
                ])));

        Assert.True(queue!.CanPublish);
        Assert.True(queue.IsAdministrator);
        Assert.Equal(2, queue.Submissions.Count);

        ReviewQueueRow first = queue.Submissions[0];
        Assert.Equal(4, first.Id);
        Assert.Equal("Commodore/C64/250407", first.BoardId);
        Assert.Equal("approved", first.State);
        Assert.Equal("Corrected U8.", first.Summary);
        Assert.Equal("c@example.com", first.ContactEmail);
        Assert.Equal(created, first.CreatedUtc);
        Assert.True(first.TouchesSharedFiles);
        Assert.False(first.IsNewBoard);
        Assert.True(first.AwaitsYou);

        ReviewQueueRow second = queue.Submissions[1];
        Assert.True(second.IsNewBoard);
        Assert.False(second.AwaitsYou);
        Assert.Equal(string.Empty, second.Summary);
    }

    // ###########################################################################################
    // The contributor facts (2026-09-26): the server's record, under the detail's "contributor",
    // through the server's own settings and back out of the real parser - every field.
    // ###########################################################################################
    [Fact]
    public void A_details_contributor_reads_back_field_for_field()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(ReviewWireContractTests.Answer(new
        {
            canPublish = true,
            submission = new ReviewQueueEntry(4, "Manu1/Hardware1/Board1", "pending", "New stuff", "dh@hinet.dk", "", null, null, false),
            contributor = new ReviewContributorFacts("dh@hinet.dk", "Dennis", Published: 3, Waiting: 2, ChangesRequested: 1, Rejected: 4)
        }));

        Assert.Equal(new ReviewContributorFacts("dh@hinet.dk", "Dennis", 3, 2, 1, 4), detail!.Contributor);
    }

    // ###########################################################################################
    // *** THE CONTRIBUTOR'S WHOLE RECORD (2026-09-30), THE SAME WAY. *** The Contributor view's
    // account facts and every listed submission - its board, description, contributor-facing
    // state, dates and what they were told - out of the server's settings and back through the real
    // parser. A renamed field would leave the view silently empty; this fails instead.
    // ###########################################################################################
    [Fact]
    public void A_contributors_whole_record_reads_back_field_for_field()
    {
        var sent = new ReviewContributorFacts(
            "dh@hinet.dk",
            "Dennis",
            Published: 1,
            Waiting: 0,
            ChangesRequested: 0,
            Rejected: 1,
            SignedIn: true,
            AccountCreatedUtc: ReviewWireContractTests.Decided.AddDays(-40),
            Submissions:
            [
                new ContributorSubmissionEntry(9, "Commodore/C64/250407", "Wrong pinout", "rejected", ReviewWireContractTests.Decided.AddDays(-3), ReviewWireContractTests.Decided, "Not this board."),
                new ContributorSubmissionEntry(7, "Commodore/C128/310378", null, "published", ReviewWireContractTests.Decided.AddDays(-9), null, null)
            ],

            // "[1] published to stable" (2026-10-01, server 3.9.0).
            PublishedToStable: 1);

        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(ReviewWireContractTests.Answer(new
        {
            canPublish = true,
            submission = new ReviewQueueEntry(10, "Commodore/C64/250407", "pending", "New stuff", "dh@hinet.dk", "", null, null, false),
            contributor = sent
        }));

        ReviewContributorFacts read = detail!.Contributor!;

        Assert.Equal(sent with { Submissions = null }, read with { Submissions = null });
        Assert.Equal(sent.Submissions, read.Submissions);
    }

    // An older server sends none of it: read as "not said", which the view words as such.
    [Fact]
    public void A_contributor_from_an_older_server_reads_back_with_no_record()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(
            """
            {"canPublish":true,
             "submission":{"id":4,"boardId":"Manu1/Hardware1/Board1","state":"pending","summary":"x","contactEmail":"dh@hinet.dk","baseRevision":""},
             "contributor":{"email":"dh@hinet.dk","name":null,"published":1,"waiting":0,"changesRequested":0,"rejected":0}}
            """);

        ReviewContributorFacts read = detail!.Contributor!;

        Assert.Null(read.SignedIn);
        Assert.Null(read.AccountCreatedUtc);
        Assert.Null(read.Submissions);
        Assert.Null(read.PublishedToStable);
    }

    // A server that does not know sends no badges - read as "not said", never as "published" or
    // "not yours", so the list shows no badge rather than a wrong one.
    [Fact]
    public void A_queue_entry_without_badges_reads_back_as_not_said()
    {
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue(ReviewWireContractTests.Answer(
            new ReviewQueueAnswer(true, false, 1,
                [new ReviewQueueEntry(4, "Commodore/C64/250407", "pending", "x", "c@example.com", "r1", null, null, false)])));

        ReviewQueueRow row = Assert.Single(queue!.Submissions);
        Assert.Null(row.IsNewBoard);
        Assert.Null(row.AwaitsYou);
    }

    // "Switched off" must stay distinguishable from "nothing waiting", with nulls left out.
    [Fact]
    public void A_production_list_from_a_server_that_cannot_publish_reads_back_as_switched_off()
    {
        ProductionListResponse? list = ReviewApiParser.ParseProductionList(
            ReviewWireContractTests.Answer(new ProductionListAnswer(false, [])));

        Assert.False(list!.Configured);
        Assert.Empty(list.Boards);
    }

    [Fact]
    public void The_administrators_list_of_boards_reads_back_with_their_maintainers()
    {
        ReviewBoardsResponse? boards = ReviewApiParser.ParseBoards(ReviewWireContractTests.Answer(
            new MaintainerBoardsAnswer(
            [
                new MaintainerBoardEntry(
                    "Commodore/C64/250407", "Commodore", "C64", "250407", "2026-August-21", IsAccepting: false,
                    [new PoolMaintainerEntry(7, "Anna", "anna@example.com")])
            ])));

        ReviewBoardRow row = Assert.Single(boards!.Boards);
        Assert.Equal("Commodore/C64/250407", row.BoardId);
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

    // ###########################################################################################
    // *** "THE CONTRIBUTOR DISCARDED THEIR DRAFT" ON EVERY SCREEN IT IS SHOWN (owner request,
    // 2026-09-28). *** Four answers carry it - the queue (and the detail's `submission`), the
    // production list, the production plan's carried submissions and the Boards detail - each
    // put through the server's own JSON settings and read back by the real parser. A field renamed
    // on one side only would drop the warning in silence, which is the one thing it must not do.
    // ###########################################################################################
    [Fact]
    public void The_discarded_draft_reads_back_on_every_answer_that_carries_it()
    {
        DateTimeOffset discarded = new(2026, 9, 28, 9, 15, 0, TimeSpan.Zero);

        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue(ReviewWireContractTests.Answer(
            new ReviewQueueAnswer(true, false, 2,
            [
                new ReviewQueueEntry(4, "Commodore/C128/310378", "pending", "x", "c@example.com", "r1", null, null, false, DraftDiscardedUtc: discarded),
                new ReviewQueueEntry(5, "Commodore/C128/310378", "pending", "y", "c@example.com", "r1", null, null, false)
            ])));

        Assert.Equal([discarded, null], queue!.Submissions.Select(row => row.DraftDiscardedUtc));

        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(ReviewWireContractTests.Answer(new
        {
            canPublish = true,
            submission = new ReviewQueueEntry(4, "Commodore/C128/310378", "pending", "x", "c@example.com", "r1", null, null, false, DraftDiscardedUtc: discarded)
        }));

        Assert.Equal(discarded, detail!.Submission.DraftDiscardedUtc);

        ProductionListResponse? list = ReviewApiParser.ParseProductionList(ReviewWireContractTests.Answer(
            new ProductionListAnswer(true,
            [
                new ProductionListEntry("Commodore/C128/310378", "Commodore", "C128", "310378", null, new string('c', 64), null, null, CarriesDiscardedDraft: true),
                new ProductionListEntry("Commodore/C64/250407", "Commodore", "C64", "250407", null, new string('d', 64), null, null)
            ])));

        Assert.Equal([true, null], list!.Boards.Select(row => row.CarriesDiscardedDraft));

        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan(ReviewWireContractTests.Answer(new ProductionPlanAnswer(
            "Commodore/C128/310378",
            "2026-September-28",
            new string('c', 64),
            IsAwaitingProduction: true,
            TouchesSharedFiles: false,
            CanPublish: true,
            Refusal: null,
            Approval: null,
            UnchangedCount: 0,
            Files: [],
            Problems: [],
            Removals: null,
            Carrying: [new CarriedSubmission(4, "c@example.com", "x", discarded, DraftDiscardedUtc: discarded)])));

        Assert.Equal(discarded, Assert.Single(plan!.Carrying!).DraftDiscardedUtc);

        BoardDetailAnswer? board = ReviewApiParser.ParseBoardDetail(ReviewWireContractTests.Answer(new BoardDetailAnswer(
            new BoardOverviewEntry("Commodore/C128/310378", "Commodore", "C128", "310378", true, false, true, true, null, null, null, 0),
            [],
            [],
            [new BoardSubmissionEntry(4, "c@example.com", "x", "merged", discarded, discarded, null, DraftDiscardedUtc: discarded)])));

        Assert.Equal(discarded, Assert.Single(board!.Submissions).DraftDiscardedUtc);
    }
    // ###########################################################################################
    // *** SIGNING IN (2026-10-04). *** The session answer was an anonymous object on the server, and
    // ParseLogin read it by hand-typed names that only ReviewApiParserTests' own JSON held - so a
    // renamed field locked every maintainer out with every test green. Now it is CRT.Data's
    // SessionAnswer, through the server's settings and back out of the real parser.
    // ###########################################################################################
    [Fact]
    public void A_sign_in_answer_reads_back_as_the_session_it_describes()
    {
        DateTimeOffset expires = new(2026, 11, 3, 12, 0, 0, TimeSpan.Zero);

        ReviewSession? session = ReviewApiParser.ParseLogin(ReviewWireContractTests.Answer(
            new SessionAnswer("the-bearer-token", expires, new SessionAccountAnswer(7, "m@example.com", "Maintainer", true))));

        Assert.NotNull(session);
        Assert.Equal("the-bearer-token", session!.BearerToken);
        Assert.Equal(expires, session.ExpiresUtc);
        Assert.Equal(7, session.AccountId);
        Assert.Equal("m@example.com", session.Email);
        Assert.Equal("Maintainer", session.DisplayName);
    }

    // ###########################################################################################
    // *** ONE SUBMISSION'S DETAIL, WHOLE (2026-10-04). *** The screen a maintainer decides from was
    // an anonymous object on the server, held to the parser only field by field in the tests above
    // by anonymous objects of their own. Now it is CRT.Data's SubmissionDetailAnswer, and every
    // field the Maintainer tab reads comes back out of the real parser here.
    // ###########################################################################################
    [Fact]
    public void A_submissions_detail_reads_back_field_for_field()
    {
        var manifest = new SubmissionManifest { BoardId = "Commodore/C64/250407" };
        manifest.Files.Add(new SubmissionFile { Path = "Commodore/C64/250407/new.png", Sha256 = new string('a', 64), SizeBytes = 10 });

        var answer = new SubmissionDetailAnswer(
            CanPublish: true,
            Approval: new ApprovalStatus([ApproverRole.Maintainer], [], [ApproverRole.Maintainer], ApproverRole.Maintainer, CanApprove: true, ApprovalPublishes: true),
            Submission: new ReviewQueueEntry(4, "Commodore/C64/250407", "pending", "Corrected U8.", "c@example.com", "", ReviewWireContractTests.Decided, null, false),
            Manifest: manifest,
            Contributor: new ReviewContributorFacts("c@example.com", "C", Published: 1, Waiting: 0, ChangesRequested: 0, Rejected: 0),
            Findings: [ReviewWireContractTests.Finding(ValidationSeverity.Warning)],
            Changes: new ReviewChangeSummary(false, null, []),
            PublishedFiles: ["Commodore/C64/250407/old.png"],
            PublishedHashes: new Dictionary<string, string> { ["Commodore/C64/250407/old.png"] = new string('b', 64) },
            SchematicImages: new Dictionary<string, string> { ["Top"] = "Commodore/C64/250407/top.png" },
            SubmittedFiles: [new SubmittedFileFact("Commodore/C64/250407/new.png", new string('a', 64), 10, SubmissionFileScope.Own, true, null)],
            Removals: new FileRemovalPreview(["Commodore/C64/250407/old.png"], null),
            Amendment: new SubmissionAmendmentFact(2, "m@example.com", ReviewWireContractTests.Decided));

        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(ReviewWireContractTests.Answer(answer));

        Assert.NotNull(detail);
        Assert.True(detail!.CanPublish);
        Assert.Equal([ApproverRole.Maintainer], detail.Approval!.Required);
        Assert.Equal([ApproverRole.Maintainer], detail.Approval.WaitingFor);
        Assert.True(detail.Approval.CanApprove);
        Assert.True(detail.Approval.ApprovalPublishes);
        Assert.Equal(4, detail.Submission.Id);
        Assert.Equal("Commodore/C64/250407/new.png", Assert.Single(detail.Assets.Files).Path);
        Assert.Equal(answer.Contributor, detail.Contributor);
        Assert.Equal("image.orphan", Assert.Single(detail.Findings).Code);
        Assert.NotNull(detail.Changes);
        Assert.Equal(["Commodore/C64/250407/old.png"], detail.PublishedFiles);
        Assert.Equal(new string('b', 64), detail.PublishedHashes["Commodore/C64/250407/old.png"]);
        Assert.Equal("Commodore/C64/250407/top.png", detail.SchematicImages["Top"]);
        Assert.Equal(answer.SubmittedFiles[0], Assert.Single(detail.SubmittedFiles));
        Assert.Equal(["Commodore/C64/250407/old.png"], detail.Removals!.Files);
        Assert.Equal(new ReviewAmendmentView(2, "m@example.com", ReviewWireContractTests.Decided), detail.Amendment);
    }

    // GET /api/health (2026-10-04): the version Account > "Server version" shows.
    [Fact]
    public void The_health_answer_gives_the_server_version()
    {
        HealthStatus? health = ReviewApiParser.ParseHealth(
            ReviewWireContractTests.Answer(new HealthStatus("ok", "4.3.2", ReviewWireContractTests.Decided)));

        Assert.NotNull(health);
        Assert.Equal("ok", health!.Status);
        Assert.Equal("4.3.2", health.Version);
        Assert.Equal(ReviewWireContractTests.Decided, health.Utc);
        Assert.Null(health.ApiRevision);
    }

    // The API revision the server serves (server 4.6.0), through the server's JSON settings and the
    // tab's parser - the number the Admin line compares with this CRT's.
    [Fact]
    public void The_health_answer_gives_the_api_revision_the_server_serves()
    {
        HealthStatus? health = ReviewApiParser.ParseHealth(
            ReviewWireContractTests.Answer(new HealthStatus("ok", "4.6.0", ReviewWireContractTests.Decided, 7)));

        Assert.Equal(7, health!.ApiRevision);
    }

    // ###########################################################################################
    // RESETTING THE CONTRIBUTION DATA (owner request, 2026-10-04). The fingerprint is the field that
    // matters: renamed on either side, every reset would send nothing back and be refused as
    // "changed since the counts were shown". And IsEnabled: lost, the button would stay off for ever
    // - or, read as true from a server that says false, be offered only to be refused.
    // ###########################################################################################
    [Fact]
    public void A_reset_arrives_with_the_fingerprint_of_the_counts_shown()
    {
        DataResetRequest received = ReviewWireContractTests.Received<DataResetRequest>(new DataResetRequest(new string('a', 64)));

        Assert.Equal(new string('a', 64), received.Fingerprint);
    }

    [Fact]
    public void The_reset_counts_read_back_field_for_field()
    {
        DataResetPlanAnswer? plan = ReviewApiParser.ParseDataResetPlan(ReviewWireContractTests.Answer(
            new DataResetPlanAnswer(false, new string('f', 64), 14, 6, 1, 4, 2, 7, 310, 1200, 80, "Switched off.")));

        Assert.NotNull(plan);
        Assert.False(plan!.IsEnabled);
        Assert.Equal(new string('f', 64), plan.Fingerprint);
        Assert.Equal((14, 6, 1, 4, 2, 7), (plan.Submissions, plan.Accounts, plan.Administrators, plan.Maintainers, plan.Invitations, plan.Boards));
        Assert.Equal((310, 1200, 80), (plan.HistoryEntries, plan.BoardViews, plan.ApiUsageRows));
        Assert.Equal("Switched off.", plan.NotEnabledBecause);

        DataResetPlanAnswer? on = ReviewApiParser.ParseDataResetPlan(ReviewWireContractTests.Answer(
            new DataResetPlanAnswer(true, "f", 0, 0, 1, 0, 0, 0, 1, 0, 0)));

        Assert.True(on!.IsEnabled);
        Assert.Null(on.NotEnabledBecause);

        // Counts with no fingerprint cannot be confirmed, so they are not counts at all.
        Assert.Null(ReviewApiParser.ParseDataResetPlan(ReviewWireContractTests.Answer(
            new DataResetPlanAnswer(true, "", 1, 1, 1, 1, 1, 1, 1, 1, 1))));
    }

    [Fact]
    public void What_a_reset_deleted_reads_back_field_for_field()
    {
        DataResetAnswer? done = ReviewApiParser.ParseDataReset(ReviewWireContractTests.Answer(
            new DataResetAnswer(14, 6, 4, 2, 7, 310, 1200, 80, 23)));

        Assert.Equal(
            (14, 6, 4, 2, 7, 310, 1200, 80, 23),
            (done!.SubmissionsDeleted, done.AccountsDeleted, done.MaintainersDeleted, done.InvitationsDeleted, done.BoardsDeleted,
             done.HistoryEntriesDeleted, done.BoardViewsDeleted, done.ApiUsageRowsDeleted, done.StoredFilesRemoved));
    }

    // ###########################################################################################
    // API USAGE (owner request, 2026-10-04). NeverRetired and the versions are the fields to get
    // wrong: a route read as retirable that every CRT ever released sends, or a version list read as
    // empty, would make a route look unused when it is not - the one mistake this screen exists to
    // prevent.
    // ###########################################################################################
    [Fact]
    public void The_API_usage_reads_back_route_for_route_and_version_for_version()
    {
        ApiUsageAnswer? usage = ReviewApiParser.ParseApiUsage(ReviewWireContractTests.Answer(
            new ApiUsageAnswer(
                90,
                [
                    new ApiUsageRoute("POST", "/api/usage/check-in", "Forever", true, 900, ReviewWireContractTests.Decided,
                        [new ApiUsageVersion("2.5.0", 800, ReviewWireContractTests.Decided), new ApiUsageVersion(ApiUsageVersion.NotCrt, 100, ReviewWireContractTests.Decided)]),
                    new ApiUsageRoute("POST", "/api/review/boards/edit", "Maintainer", false, 0, null, []),
                ],
                [new ApiUsageInstallations("3.0.0", 42, 305)])));

        Assert.NotNull(usage);
        Assert.Equal(90, usage!.Days);
        Assert.Equal(2, usage.Routes.Count);

        ApiUsageRoute checkIn = usage.Routes[0];
        Assert.Equal(("POST", "/api/usage/check-in", "Forever", true, 900L), (checkIn.Method, checkIn.Route, checkIn.Area, checkIn.NeverRetired, checkIn.Calls));
        Assert.Equal(ReviewWireContractTests.Decided, checkIn.LastUtc);
        Assert.Equal(["2.5.0", ApiUsageVersion.NotCrt], checkIn.Versions.Select(version => version.Version));
        Assert.Equal(800, checkIn.Versions[0].Calls);

        ApiUsageRoute unused = usage.Routes[1];
        Assert.False(unused.NeverRetired);
        Assert.Null(unused.LastUtc);
        Assert.Empty(unused.Versions);

        Assert.Equal(new ApiUsageInstallations("3.0.0", 42, 305), Assert.Single(usage.Installations));
    }
}
