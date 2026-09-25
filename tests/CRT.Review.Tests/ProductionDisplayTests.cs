using CRT.Review.Handlers;
using Handlers.DataHandling;

namespace CRT.Review.Tests;

// Covers ProductionDisplay and the parsers behind the "Publish to production" window
// (2026-09-25).
//
// The one rule that is a decision rather than wording is CanPress: the button follows the server's
// canPublish AND the reviewer's own "I checked it in BETA" tick - never either alone. The server
// enforces its half regardless; the tick is the human half of "only after he has checked it".
public sealed class ProductionDisplayTests
{
    private static ProductionPlanView Plan(bool canPublish = true, params PromotionFile[] files) =>
        new("Commodore/C64/250407", "2026-September-25", "hash", false, canPublish, null, 1180, files, []);

    private static PromotionFile File(string path, PromotionChange change = PromotionChange.Added, bool shared = false) =>
        new(path, new string('a', 64), change, PromotionStage.Content, shared);

    // -----------------------------------------------------------------------------------
    // The button
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_button_needs_BOTH_the_servers_yes_and_the_reviewers_tick()
    {
        Assert.True(ProductionDisplay.CanPress(ProductionDisplayTests.Plan(canPublish: true), checkedInBeta: true));

        // Not ticked: the reviewer has not said they checked it in BETA.
        Assert.False(ProductionDisplay.CanPress(ProductionDisplayTests.Plan(canPublish: true), checkedInBeta: false));

        // Ticked, but the server refuses (a shared file, a board it depends on not in production).
        Assert.False(ProductionDisplay.CanPress(ProductionDisplayTests.Plan(canPublish: false), checkedInBeta: true));

        Assert.False(ProductionDisplay.CanPress(null, checkedInBeta: true));
    }

    // -----------------------------------------------------------------------------------
    // Wording
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_summary_counts_additions_replacements_and_what_is_already_the_same()
    {
        ProductionPlanView plan = ProductionDisplayTests.Plan(
            true,
            ProductionDisplayTests.File("a.png"),
            ProductionDisplayTests.File("b.png"),
            ProductionDisplayTests.File("c.xlsx", PromotionChange.Replaced));

        Assert.Equal("2 files to add, 1 file to replace, 1,180 already the same", ProductionDisplay.PlanSummary(plan));
    }

    [Fact]
    public void A_plan_copying_nothing_says_publishing_only_records_it()
    {
        Assert.StartsWith("Production already has every file", ProductionDisplay.PlanSummary(ProductionDisplayTests.Plan()));
    }

    [Fact]
    public void A_SHARED_file_is_marked_in_its_line()
    {
        // It reaches every board that cites it - the line a reviewer most needs to notice.
        Assert.EndsWith("(shared)", ProductionDisplay.FileLine(ProductionDisplayTests.File("Commodore/Shared files/x.png", shared: true)));
        Assert.StartsWith("replaced", ProductionDisplay.FileLine(ProductionDisplayTests.File("a.png", PromotionChange.Replaced)));
    }

    [Fact]
    public void A_system_never_in_production_says_so()
    {
        var row = new ProductionSystemRow("Amstrad/CPC/464", "Amstrad", "CPC", "464", "2026-September-25", "h", null, null);

        Assert.Equal(
            "Amstrad / CPC / 464  -  BETA 2026-September-25, never published to production",
            ProductionDisplay.SystemLine(row));
    }

    [Fact]
    public void A_list_entry_puts_the_board_and_its_state_on_TWO_lines()
    {
        // On one line the list cut the state off at its width ("...never p") in the first render,
        // hiding how far production is behind - the one fact the row is there for.
        var row = new ProductionSystemRow("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-25", "h", "2026-May-14", null);

        Assert.Equal(
            "Commodore / C64 / 250407\nBETA 2026-September-25, production 2026-May-14",
            ProductionDisplay.ListEntry(row));
    }

    // -----------------------------------------------------------------------------------
    // The parsers
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_list_says_whether_the_server_can_do_it_at_all()
    {
        // "Nothing waiting" and "this server cannot publish to production" must read differently.
        ProductionListResponse? off = ReviewApiParser.ParseProductionList("""{"configured":false,"systems":[]}""");
        ProductionListResponse? on = ReviewApiParser.ParseProductionList("""
            {"configured":true,"systems":[
              {"systemId":"Commodore/C64/250407","manufacturer":"Commodore","hardware":"C64","board":"250407",
               "betaRevision":"2026-September-25","betaContentHash":"abc","productionRevision":null,"productionPublishedUtc":null}]}
            """);

        Assert.False(off!.Configured);
        Assert.True(on!.Configured);
        Assert.Equal("abc", Assert.Single(on.Systems).BetaContentHash);
    }

    [Fact]
    public void The_plan_reads_the_servers_own_PromotionFile_records_with_their_enums_by_NAME()
    {
        // The server serialises CRT.Data's PromotionFile; the enums travel as names so a member
        // added later cannot shift every value by one.
        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan("""
            {"systemId":"Commodore/C64/250407","betaRevision":"2026-September-25","betaContentHash":"abc",
             "touchesSharedFiles":true,"canPublish":false,"refusal":"Only an administrator.","unchangedCount":12,
             "files":[{"path":"Commodore/Shared files/x.png","sha256":"aa","change":"Replaced","stage":"Content","isShared":true}],
             "problems":[{"severity":1,"code":"promote.x","subject":"s","message":"Something is wrong."}]}
            """);

        Assert.NotNull(plan);
        Assert.Equal("abc", plan!.BetaContentHash);
        Assert.True(plan.TouchesSharedFiles);
        Assert.False(plan.CanPublish);
        Assert.Equal("Only an administrator.", plan.Refusal);
        Assert.Equal(12, plan.UnchangedCount);

        PromotionFile file = Assert.Single(plan.Files);
        Assert.Equal(PromotionChange.Replaced, file.Change);
        Assert.True(file.IsShared);

        Assert.Equal("Something is wrong.", Assert.Single(plan.Problems).Message);
    }

    [Fact]
    public void A_plan_with_no_canPublish_answer_is_NOT_publishable()
    {
        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan("""{"systemId":"A/B/C","files":[]}""");

        Assert.False(plan!.CanPublish);
    }

    [Fact]
    public void An_unreadable_answer_is_null()
    {
        Assert.Null(ReviewApiParser.ParseProductionList("nope"));
        Assert.Null(ReviewApiParser.ParseProductionPlan("""{"files":[]}"""));
        Assert.Null(ReviewApiParser.ParseProductionPublish("{}"));
    }
    // Removals are named only when there are any - a standing "0 to remove" would be read past on
    // the one promotion that does remove something. (2026-09-25)
    [Fact]
    public void The_summary_names_files_to_remove_only_when_there_are_any()
    {
        ProductionPlanView plan = ProductionDisplayTests.Plan(true, ProductionDisplayTests.File("Commodore/C64/250407/a.png"))
            with { Removals = new FileRemovalPreview(["Commodore/C64/250407/old.pdf", "Commodore/C64/250407/old2.pdf"], null) };

        Assert.Equal("1 file to add, 0 files to replace, 2 files to remove, 1,180 already the same", ProductionDisplay.PlanSummary(plan));
        Assert.DoesNotContain("remove", ProductionDisplay.PlanSummary(plan with { Removals = FileRemovalPreview.Nothing }), StringComparison.Ordinal);
    }
}
