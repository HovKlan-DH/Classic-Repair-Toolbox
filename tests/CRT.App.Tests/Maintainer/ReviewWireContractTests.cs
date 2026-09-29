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
            Removals: new FileRemovalPreview(["Commodore/C64/250407/old.png"], null),
            Carrying: [new CarriedSubmission(7, "hest@mailscan.dk", "Corrected U8.", ReviewWireContractTests.Decided)])));

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
                new SystemFileEntry("Commodore/C128/310378/a.png", SystemFileChange.Added, SystemFileSource.Submission, new string('d', 64)),
                new SystemFileEntry("Commodore/C128/310378/Data.xlsx", SystemFileChange.Changed, SystemFileSource.Beta, WrittenOnApproval: true),
                new SystemFileEntry("Commodore/C128/310378/old.png", SystemFileChange.Removed, SystemFileSource.Beta),
                new SystemFileEntry("Commodore/C128/310378/new.json", SystemFileChange.Added, SystemFileSource.NotWrittenYet, WrittenOnApproval: true)
            ],
            "https://example.org/app-data-BETA/Data")));

        Assert.NotNull(answer);
        Assert.Equal("Commodore/C128/310378", answer!.SystemId);
        Assert.Equal("https://example.org/app-data-BETA/Data", answer.BetaDataUrl);
        Assert.Equal(
            [
                new SystemFileEntry("Commodore/C128/310378/a.png", SystemFileChange.Added, SystemFileSource.Submission, new string('d', 64)),
                new SystemFileEntry("Commodore/C128/310378/Data.xlsx", SystemFileChange.Changed, SystemFileSource.Beta, WrittenOnApproval: true),
                new SystemFileEntry("Commodore/C128/310378/old.png", SystemFileChange.Removed, SystemFileSource.Beta),
                new SystemFileEntry("Commodore/C128/310378/new.json", SystemFileChange.Added, SystemFileSource.NotWrittenYet, WrittenOnApproval: true)
            ],
            answer.Files);

        // Names, not numbers, on the wire - an enum reordered on one side cannot change a meaning.
        Assert.Contains("\"notWrittenYet\"", ReviewWireContractTests.Answer(new SubmissionFilesAnswer(
            "x", [new SystemFileEntry("x/a", SystemFileChange.Added, SystemFileSource.NotWrittenYet)])), StringComparison.OrdinalIgnoreCase);
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
            """{"systemId":"Commodore/C64/250407","kind":"something-else"}""");

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

        ProductionSystemRow row = Assert.Single(list.Systems);
        Assert.Equal("Commodore/C64/250407", row.SystemId);
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
    // *** THE BETA BADGE'S FLAG (2026-09-27). *** Renamed on either side, a system this account had
    // already approved would read as waiting for it again, and the badge would count it for ever.
    // ###########################################################################################
    [Fact]
    public void Whether_a_waiting_system_waits_for_you_reads_back()
    {
        ProductionListResponse? list = ReviewApiParser.ParseProductionList(ReviewWireContractTests.Answer(
            new ProductionListAnswer(
                true,
                [
                    new ProductionListEntry("Commodore/C64/250407", "Commodore", "C64", "250407", null, new string('c', 64), null, null, AwaitsYou: false),
                    new ProductionListEntry("Commodore/C128/310378", "Commodore", "C128", "310378", null, new string('d', 64), null, null, AwaitsYou: true)
                ])));

        Assert.Equal([false, true], list!.Systems.Select(row => row.AwaitsYou));
    }

    // ###########################################################################################
    // THE "SYSTEMS" SCREEN (2026-09-27): the list, the detail request, and the detail - every
    // field, since each one is on screen and a renamed one would be blank there in silence.
    // ###########################################################################################
    [Fact]
    public void The_systems_list_reads_back_field_for_field()
    {
        SystemOverviewEntry sent = new(
            "Commodore/C64/250407", "Commodore", "C64", "250407",
            InBeta: true, InProduction: false, IsAwaitingProduction: true, IsAccepting: false,
            BetaRevision: "2026-September-25", ProductionRevision: "2026-May-14",
            ProductionPublishedUtc: ReviewWireContractTests.Decided, MaintainerCount: 2,
            ViewsLast30Days: 48);

        SystemOverviewAnswer? answer = ReviewApiParser.ParseSystemOverview(
            ReviewWireContractTests.Answer(new SystemOverviewAnswer([sent])));

        Assert.Equal(sent, Assert.Single(answer!.Systems));
    }

    // ###########################################################################################
    // Board views (2026-09-27): no count is NOT zero - absent from an older server, or when it could
    // not count, it stays null and the screen says nothing; a real 0 reads back as 0.
    // ###########################################################################################
    [Fact]
    public void A_systems_view_count_reads_back_and_its_absence_stays_absent()
    {
        SystemOverviewEntry counted = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0, ViewsLast30Days: 0);
        SystemOverviewEntry uncounted = counted with { SystemId = "Commodore/C128/310378", ViewsLast30Days = null };

        SystemOverviewAnswer? answer = ReviewApiParser.ParseSystemOverview(
            ReviewWireContractTests.Answer(new SystemOverviewAnswer([counted, uncounted])));

        Assert.Equal([(int?)0, null], answer!.Systems.Select(system => system.ViewsLast30Days));
    }

    [Fact]
    public void A_systems_view_statistics_read_back_field_for_field()
    {
        SystemOverviewEntry system = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0, 48);
        BoardViewStatistics views = new(12, 48, 310, 3, [new("DK", "Denmark", 120), new("DE", "Germany", 60)]);

        SystemDetailAnswer? read = ReviewApiParser.ParseSystemDetail(
            ReviewWireContractTests.Answer(new SystemDetailAnswer(system, [], [], [], Views: views)));

        Assert.Equal(
            (views.Last7Days, views.Last30Days, views.Last365Days, views.FromBetaLast30Days),
            (read!.Views!.Last7Days, read.Views.Last30Days, read.Views.Last365Days, read.Views.FromBetaLast30Days));
        Assert.Equal(views.TopCountries, read.Views.TopCountries);
        Assert.Equal(48, read.System.ViewsLast30Days);

        Assert.Null(ReviewApiParser.ParseSystemDetail(ReviewWireContractTests.Answer(new SystemDetailAnswer(system, [], [], [])))!.Views);
    }

    // With no production tree the server says nothing about production - and that is kept, not
    // read as "not in production".
    [Fact]
    public void A_system_the_server_could_not_look_up_in_production_reads_back_as_not_said()
    {
        SystemOverviewAnswer? answer = ReviewApiParser.ParseSystemOverview(ReviewWireContractTests.Answer(
            new SystemOverviewAnswer([new SystemOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, null, false, true, null, null, null, 0)])));

        Assert.Null(Assert.Single(answer!.Systems).InProduction);
    }

    [Fact]
    public void A_system_detail_request_arrives_with_its_id()
    {
        Assert.Equal(
            "Commodore/C64/250407",
            ReviewWireContractTests.Received<SystemDetailRequest>(new SystemDetailRequest("Commodore/C64/250407")).SystemId);
    }

    [Fact]
    public void A_systems_detail_reads_back_field_for_field()
    {
        SystemDetailAnswer sent = new(
            new SystemOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, "r2", "r1", ReviewWireContractTests.Decided, 1),
            [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
            [new SystemContributorEntry("hest@mailscan.dk", "Hest", Accepted: 3, Waiting: 1, ChangesRequested: 2, Rejected: 4, LastSubmittedUtc: ReviewWireContractTests.Decided)],
            [new SystemSubmissionEntry(41, "hest@mailscan.dk", "Corrected U8.", "returned", ReviewWireContractTests.Decided, ReviewWireContractTests.Decided.AddDays(1), "U7 is the wrong revision.")]);

        SystemDetailAnswer? read = ReviewApiParser.ParseSystemDetail(ReviewWireContractTests.Answer(sent));

        Assert.NotNull(read);
        Assert.Equal(sent.System, read!.System);
        Assert.Equal(sent.Maintainers, read.Maintainers);
        Assert.Equal(sent.Contributors, read.Contributors);
        Assert.Equal(sent.Submissions, read.Submissions);
    }

    // ###########################################################################################
    // Inviting a new maintainer by email (2026-09-27): the three requests as the server binds them,
    // the acceptance's answer, and the invitations on a system's detail.
    // ###########################################################################################
    [Fact]
    public void The_invitation_requests_arrive_as_sent()
    {
        MaintainerInviteRequest invite = ReviewWireContractTests.Received<MaintainerInviteRequest>(
            new MaintainerInviteRequest("Commodore/C64/250407", "new@example.com"));
        Assert.Equal("Commodore/C64/250407", invite.SystemId);
        Assert.Equal("new@example.com", invite.Email);

        Assert.Equal(3, ReviewWireContractTests.Received<InvitationWithdrawRequest>(new InvitationWithdrawRequest(3)).InvitationId);

        AcceptInvitationRequest accept = ReviewWireContractTests.Received<AcceptInvitationRequest>(
            new AcceptInvitationRequest("code", "Anna", "correct horse battery staple"));
        Assert.Equal(new AcceptInvitationRequest("code", "Anna", "correct horse battery staple"), accept);
    }

    // The system's history (2026-09-27), field for field - and absent from an older server stays absent.
    [Fact]
    public void A_systems_history_reads_back_field_for_field()
    {
        SystemOverviewEntry system = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0);
        SystemHistoryEntry[] history =
        [
            new(ReviewWireContractTests.Decided, SystemHistoryEvents.Decided, "Anna", 9, "rejected", "Wrong board."),
            new(ReviewWireContractTests.Decided.AddDays(-1), SystemHistoryEvents.Sent, "hest@mailscan.dk", 9, "Corrected U8."),
            new(ReviewWireContractTests.Decided.AddDays(-2), SystemHistoryEvents.MaintainerAdded, "admin@example.com", null, "anna@example.com"),
        ];

        SystemDetailAnswer? read = ReviewApiParser.ParseSystemDetail(
            ReviewWireContractTests.Answer(new SystemDetailAnswer(system, [], [], [], null, history)));

        Assert.Equal(history, read!.History);
        Assert.Null(ReviewApiParser.ParseSystemDetail(ReviewWireContractTests.Answer(new SystemDetailAnswer(system, [], [], [])))!.History);
    }

    [Fact]
    public void An_accepted_invitation_reads_back_field_for_field()
    {
        AcceptInvitationAnswer sent = new("anna@example.com", ["Commodore/C64/250407", "Commodore/C128/310378"], "Your account is ready.");

        AcceptInvitationAnswer? read = ReviewApiParser.ParseAcceptInvitation(ReviewWireContractTests.Answer(sent));

        Assert.NotNull(read);
        Assert.Equal(sent.Email, read!.Email);
        Assert.Equal(sent.SystemIds, read.SystemIds);
        Assert.Equal(sent.Message, read.Message);
    }

    // For an administrator the detail carries the open invitations; for anybody else there are none
    // (null), which must read back as none rather than as an empty list claimed by the server.
    [Fact]
    public void A_systems_invitations_read_back_and_their_absence_stays_absent()
    {
        SystemOverviewEntry system = new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 0);
        var invitation = new MaintainerInvitationEntry(3, "new@example.com", ReviewWireContractTests.Decided, ReviewWireContractTests.Decided.AddDays(14));

        SystemDetailAnswer? forAdmin = ReviewApiParser.ParseSystemDetail(
            ReviewWireContractTests.Answer(new SystemDetailAnswer(system, [], [], [], [invitation])));
        Assert.Equal([invitation], forAdmin!.Invitations);

        SystemDetailAnswer? forMaintainer = ReviewApiParser.ParseSystemDetail(
            ReviewWireContractTests.Answer(new SystemDetailAnswer(system, [], [], [])));
        Assert.Null(forMaintainer!.Invitations);
    }

    // ###########################################################################################
    // A new system's place in CRT's drop-down lists (2026-09-27): the lists and the systems not in
    // them, the placement a maintainer saves, and what the server says it did.
    // ###########################################################################################
    [Fact]
    public void The_drop_down_listing_reads_back_field_for_field()
    {
        SystemListingAnswer sent = new(
            HasList: true,
            [
                new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                new SystemListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
            ],
            [
                new UnlistedSystemEntry(
                    "Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128",
                    InBeta: true,
                    CanPlace: true,
                    new SystemPlacement("Commodore 128", "310378 Open128", "Open-source replica.", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                    new SystemPlacement("Commodore 128", "310378 Open128", string.Empty, "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx")),

                // Never placed, and placed FIRST by its suggestion (no row above it).
                new UnlistedSystemEntry(
                    "Amstrad/CPC 6128/MC0020", "Amstrad", "CPC 6128", "MC0020",
                    InBeta: false,
                    CanPlace: false,
                    Placement: null,
                    new SystemPlacement("CPC 6128", "MC0020", string.Empty, null)),
            ]);

        SystemListingAnswer? read = ReviewApiParser.ParseSystemListing(ReviewWireContractTests.Answer(sent));

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

        Assert.Equal("Commodore/C128/310378 Open128", received.SystemId);
        Assert.Equal("Commodore 128", received.HardwareName);
        Assert.Equal("310378 Open128", received.BoardName);
        Assert.Equal("Open-source replica.", received.Notes);
        Assert.Equal("Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx", received.AfterExcelDataFile);

        // Placed first: no row above it, which must arrive as nothing rather than as "".
        Assert.Null(ReviewWireContractTests.Received<SetPlacementRequest>(
            new SetPlacementRequest("Amstrad/CPC 6128/MC0020", "CPC 6128", "MC0020", string.Empty, null)).AfterExcelDataFile);
    }

    [Fact]
    public void A_saved_placement_reads_back_with_what_the_server_did()
    {
        SetPlacementAnswer sent = new(
            new SystemPlacement("Commodore 128", "310378 Open128", "Open-source replica.", null),
            ListedInBeta: true,
            "Placed, and added to BETA's drop-down lists.");

        Assert.Equal(sent, ReviewApiParser.ParseSetPlacement(ReviewWireContractTests.Answer(sent)));
    }

    // ###########################################################################################
    // The queue (2026-09-26 - an anonymous object until then), with the two badges the queue list
    // shows: whether the system is published, and whether it waits for this account.
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
                    new ReviewQueueEntry(4, "Commodore/C64/250407", "approved", "Corrected U8.", "c@example.com", "r1", created, null, TouchesSharedFiles: true, IsNewSystem: false, AwaitsYou: true),
                    new ReviewQueueEntry(5, "Amstrad/CPC464/Z70200", "pending", null, null, "", created, null, TouchesSharedFiles: false, IsNewSystem: true, AwaitsYou: false)
                ])));

        Assert.True(queue!.CanPublish);
        Assert.True(queue.IsAdministrator);
        Assert.Equal(2, queue.Submissions.Count);

        ReviewQueueRow first = queue.Submissions[0];
        Assert.Equal(4, first.Id);
        Assert.Equal("Commodore/C64/250407", first.SystemId);
        Assert.Equal("approved", first.State);
        Assert.Equal("Corrected U8.", first.Summary);
        Assert.Equal("c@example.com", first.ContactEmail);
        Assert.Equal(created, first.CreatedUtc);
        Assert.True(first.TouchesSharedFiles);
        Assert.False(first.IsNewSystem);
        Assert.True(first.AwaitsYou);

        ReviewQueueRow second = queue.Submissions[1];
        Assert.True(second.IsNewSystem);
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

    // A server that does not know sends no badges - read as "not said", never as "published" or
    // "not yours", so the list shows no badge rather than a wrong one.
    [Fact]
    public void A_queue_entry_without_badges_reads_back_as_not_said()
    {
        ReviewQueueResponse? queue = ReviewApiParser.ParseQueue(ReviewWireContractTests.Answer(
            new ReviewQueueAnswer(true, false, 1,
                [new ReviewQueueEntry(4, "Commodore/C64/250407", "pending", "x", "c@example.com", "r1", null, null, false)])));

        ReviewQueueRow row = Assert.Single(queue!.Submissions);
        Assert.Null(row.IsNewSystem);
        Assert.Null(row.AwaitsYou);
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

    // ###########################################################################################
    // *** "THE CONTRIBUTOR DISCARDED THEIR DRAFT" ON EVERY SCREEN IT IS SHOWN (owner request,
    // 2026-09-28). *** Four answers carry it - the queue (and the detail's `submission`), the
    // production list, the production plan's carried submissions and the Systems detail - each
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

        Assert.Equal([true, null], list!.Systems.Select(row => row.CarriesDiscardedDraft));

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

        SystemDetailAnswer? system = ReviewApiParser.ParseSystemDetail(ReviewWireContractTests.Answer(new SystemDetailAnswer(
            new SystemOverviewEntry("Commodore/C128/310378", "Commodore", "C128", "310378", true, false, true, true, null, null, null, 0),
            [],
            [],
            [new SystemSubmissionEntry(4, "c@example.com", "x", "merged", discarded, discarded, null, DraftDiscardedUtc: discarded)])));

        Assert.Equal(discarded, Assert.Single(system!.Submissions).DraftDiscardedUtc);
    }
}
