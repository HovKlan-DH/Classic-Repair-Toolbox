using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ProductionDisplay and the parsers behind the BETA screen (2026-09-25; the "Publish to
// production" window until 2026-09-27).
//
// The one rule that is a decision rather than wording is CanPress: the button follows the server's
// canPublish AND the maintainer's own "I checked it in BETA" tick - never either alone. The server
// enforces its half regardless; the tick is the human half of "only after he has checked it".
public sealed class ProductionDisplayTests
{
    private static ProductionPlanView Plan(bool canPublish = true, params PromotionFile[] files) =>
        new("Commodore/C64/250407", "2026-September-25", "hash", false, canPublish, null, 1180, files, []);

    private static PromotionFile File(string path, PromotionChange change = PromotionChange.Added, bool shared = false) =>
        new(path, new string('a', 64), change, PromotionStage.Content, shared);

    // ###########################################################################################
    // THE BETA SCREEN'S LIST LINE (2026-09-27): where BETA and production stand - and, for a system
    // this account already approved, that it is with the other approver, since the row is dimmed
    // for it. Said in the queue's own words.
    // ###########################################################################################
    [Fact]
    public void A_beta_list_line_says_when_the_system_waits_for_the_other_approver()
    {
        ProductionSystemRow row = new("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-25", "hash", "2026-May-14", null);

        Assert.Equal("BETA 2026-September-25, production 2026-May-14", ProductionDisplay.ListFooter(row));
        Assert.Equal("BETA 2026-September-25, production 2026-May-14", ProductionDisplay.ListFooter(row with { AwaitsYou = true }));
        Assert.Equal(
            "BETA 2026-September-25, production 2026-May-14 - with the other approver",
            ProductionDisplay.ListFooter(row with { AwaitsYou = false }));
    }

    // -----------------------------------------------------------------------------------
    // WHOSE WORK THIS CARRIES (owner request, 2026-09-27)
    //
    // The panel showed the file copy list and nothing else, which cannot answer the question the
    // button actually asks - "whose accepted work am I about to push to every CRT user?". These
    // lines are that answer.
    // -----------------------------------------------------------------------------------

    private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

    private static CarriedSubmission Carried(
        string email = "hest@mailscan.dk",
        string? comment = "Corrected U8.",
        int daysAgo = 3) =>
        new(7, email, comment, ProductionDisplayTests.Now.AddDays(-daysAgo));

    [Fact]
    public void The_headline_counts_the_contributions_going_out()
    {
        Assert.Equal(
            "1 contribution goes out to everyone with this:",
            ProductionDisplay.CarryingHeadline([ProductionDisplayTests.Carried()]));

        Assert.Equal(
            "3 contributions go out to everyone with this:",
            ProductionDisplay.CarryingHeadline(
                [ProductionDisplayTests.Carried(), ProductionDisplayTests.Carried(), ProductionDisplayTests.Carried()]));
    }

    // ###########################################################################################
    // *** SAID EVEN WHEN IT IS NONE. *** A board brought level by hand genuinely carries no
    // contribution, and a blank space there would read as though the question had not been asked -
    // the same reasoning the removals headline already follows.
    // ###########################################################################################
    [Fact]
    public void Carrying_nothing_is_said_rather_than_left_blank()
    {
        Assert.Equal(
            "No contributions are waiting to go out with this.",
            ProductionDisplay.CarryingHeadline([]));

        Assert.Equal(
            "No contributions are waiting to go out with this.",
            ProductionDisplay.CarryingHeadline(null));
    }

    [Fact]
    public void A_carried_line_names_the_contributor_their_own_words_and_when_it_was_accepted()
    {
        Assert.Equal(
            "hest@mailscan.dk - Corrected U8. (accepted 3 days ago)",
            ProductionDisplay.CarryingLine(ProductionDisplayTests.Carried(), ProductionDisplayTests.Now));
    }

    // Neither field is guaranteed: a submission can be sent without an account and saved with no
    // description. Both say so rather than rendering an empty gap where a name should be.
    [Fact]
    public void A_carried_line_survives_a_missing_address_or_description()
    {
        Assert.StartsWith(
            "(no contact address) - (no description)",
            ProductionDisplay.CarryingLine(
                ProductionDisplayTests.Carried(email: "  ", comment: null),
                ProductionDisplayTests.Now));
    }

    // ###########################################################################################
    // *** TRUNCATED DOWNWARDS, as ReviewQueueDisplay.Waiting is. *** 29 days must not read as "a
    // month" and 23 hours must not become "a day": the coarser unit always makes a delay sound
    // shorter than it is, and this number is how long a contributor has been waiting to go live.
    // ###########################################################################################
    [Theory]
    [InlineData(0, "accepted just now")]
    [InlineData(59, "accepted 59 minutes ago")]
    [InlineData(60, "accepted 1 hour ago")]
    [InlineData(1439, "accepted 23 hours ago")]
    [InlineData(1440, "accepted 1 day ago")]
    [InlineData(41760, "accepted 29 days ago")]
    public void How_long_ago_it_was_accepted_never_rounds_up(int minutesAgo, string expected)
    {
        Assert.Equal(
            expected,
            ProductionDisplay.AcceptedAgo(ProductionDisplayTests.Now.AddMinutes(-minutesAgo), ProductionDisplayTests.Now));
    }

    // A missing timestamp says nothing rather than rendering the 1970 default as "56 years ago".
    [Fact]
    public void A_submission_with_no_decision_time_says_nothing_about_when()
    {
        Assert.Equal(string.Empty, ProductionDisplay.AcceptedAgo(null, ProductionDisplayTests.Now));

        Assert.Equal(
            "hest@mailscan.dk - Corrected U8.",
            ProductionDisplay.CarryingLine(
                ProductionDisplayTests.Carried() with { DecidedUtc = null },
                ProductionDisplayTests.Now));
    }

    // A clock difference between client and server must not read as "accepted -3 minutes ago".
    [Fact]
    public void A_future_timestamp_from_a_clock_difference_reads_as_just_now()
    {
        Assert.Equal(
            "accepted just now",
            ProductionDisplay.AcceptedAgo(ProductionDisplayTests.Now.AddMinutes(5), ProductionDisplayTests.Now));
    }

    // -----------------------------------------------------------------------------------
    // PUSHING A BOARD BACK TO THE QUEUE (owner decision, 2026-09-27)
    // -----------------------------------------------------------------------------------

    private static BetaRollbackPlanView RollBack(
        BetaRollbackKind kind = BetaRollbackKind.RestoreFromProduction,
        int restored = 2,
        int removed = 1,
        int returning = 1) =>
        new(
            "Commodore/C64/250407",
            kind,
            Enumerable.Range(0, restored).Select(index => $"board/restored{index}.png").ToList(),
            Enumerable.Range(0, removed).Select(index => $"board/removed{index}.png").ToList(),
            Enumerable.Range(0, returning)
                .Select(index => new CarriedSubmission(index + 1, $"c{index}@example.com", "A change.", ProductionDisplayTests.Now))
                .ToList());

    // The two kinds are different operations and must not read as the same one.
    [Fact]
    public void The_headline_says_which_of_the_two_operations_this_is()
    {
        Assert.Contains(
            "back to what production has",
            ProductionDisplay.RollBackHeadline(ProductionDisplayTests.RollBack()),
            StringComparison.Ordinal);

        Assert.Contains(
            "Remove this board from BETA",
            ProductionDisplay.RollBackHeadline(ProductionDisplayTests.RollBack(BetaRollbackKind.RemoveFromBeta)),
            StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** THE SENTENCE THAT STOPS SOMEONE ELSE'S WORK BEING DISCARDED UNKNOWINGLY. *** A rollback
    // is per SYSTEM, so it takes back every submission merged since the last promotion. With more
    // than one, the explanation has to SAY so - a maintainer returning one contributor's work must
    // not silently take back two others'.
    // ###########################################################################################
    [Fact]
    public void With_more_than_one_submission_the_explanation_says_all_of_them_go_back()
    {
        string many = ProductionDisplay.RollBackExplanation(ProductionDisplayTests.RollBack(returning: 3));

        Assert.Contains("ALL 3 submissions", many, StringComparison.Ordinal);
        Assert.Contains("cannot be rolled back one contribution at a time", many, StringComparison.Ordinal);

        // One reads as one, without the warning that would be noise there.
        string single = ProductionDisplay.RollBackExplanation(ProductionDisplayTests.RollBack(returning: 1));

        Assert.Contains("The submission below goes back", single, StringComparison.Ordinal);
        Assert.DoesNotContain("ALL", single, StringComparison.Ordinal);
    }

    // A system never promoted has nothing to go back to, and the explanation says that rather than
    // describing a restore that is not happening.
    [Fact]
    public void A_system_never_promoted_is_explained_as_leaving_beta()
    {
        Assert.Contains(
            "removed from BETA entirely",
            ProductionDisplay.RollBackExplanation(ProductionDisplayTests.RollBack(BetaRollbackKind.RemoveFromBeta)),
            StringComparison.Ordinal);
    }

    // The counts are real, so a maintainer sees how much moves.
    [Fact]
    public void The_explanation_counts_the_files_a_restore_would_move()
    {
        string text = ProductionDisplay.RollBackExplanation(ProductionDisplayTests.RollBack(restored: 2, removed: 1));

        Assert.Contains("2 files restored", text, StringComparison.Ordinal);
        Assert.Contains("1 file removed", text, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** SHARED FILES GET A SENTENCE OF THEIR OWN (code review, 2026-09-27). *** They reach every
    // board citing them - a different consequence from this board's own files - so they are not
    // folded into a count. Singular and plural both read as English.
    // ###########################################################################################
    [Fact]
    public void Shared_files_a_rollback_touches_are_said_separately()
    {
        BetaRollbackPlanView one = ProductionDisplayTests.RollBack() with
        {
            SharedRestored = ["Commodore/Shared files/6526.png"]
        };

        Assert.Equal(
            "Shared files, used by every board that cites them: 1 goes back to production's version.",
            ProductionDisplay.RollBackSharedFiles(one));

        BetaRollbackPlanView many = ProductionDisplayTests.RollBack() with
        {
            SharedRestored = ["a.png", "b.png"]
        };

        Assert.Equal(
            "Shared files, used by every board that cites them: 2 go back to production's version.",
            ProductionDisplay.RollBackSharedFiles(many));

        Assert.Equal(["a.png", "b.png"], ProductionDisplay.RollBackSharedPaths(many));
    }

    // No shared files, no sentence - one that says "0" would be read past on the one that matters.
    [Fact]
    public void A_rollback_touching_no_shared_file_says_nothing_about_them()
    {
        Assert.Null(ProductionDisplay.RollBackSharedFiles(ProductionDisplayTests.RollBack()));
    }

    // The finished message counts what moved and who went back - and promises nothing about shared
    // files, which a rollback only ever restores (counted in "restored").
    [Fact]
    public void A_finished_rollback_says_what_moved_and_who_went_back()
    {
        string done = ProductionDisplay.RolledBack(new BetaRollbackResult(
            "Commodore/C64/250407", BetaRollbackKind.RestoreFromProduction, 1, 0, 1));

        Assert.Equal(
            "Pushed back: 1 file restored, 0 files removed. 1 submission back in the queue, and the contributors have been told.",
            done);
    }

    // ###########################################################################################
    // "Reject" on Beta > Prod (owner request, 2026-09-28): the same rollback, said as a rejection -
    // and, taking out several contributors' work at once, saying ALL of them, as a push-back does.
    // ###########################################################################################
    [Fact]
    public void A_rejection_says_the_submissions_are_rejected_not_returned()
    {
        string one = ProductionDisplay.RejectExplanation(ProductionDisplayTests.RollBack());

        Assert.Contains("BETA goes back to the data production already has", one, StringComparison.Ordinal);
        Assert.Contains("is rejected", one, StringComparison.Ordinal);
        Assert.Contains("does not come back to the queue", one, StringComparison.Ordinal);
        Assert.Equal("Reject", ProductionDisplay.RejectConfirmButton(ProductionDisplayTests.RollBack()));

        string done = ProductionDisplay.Rejected(
            new BetaRollbackResult("Commodore/C64/250407", BetaRollbackKind.RestoreFromProduction, 1, 0, 1, Rejected: true));

        Assert.Equal("Rejected: 1 file restored, 0 files removed. 1 submission rejected, and the contributors have been told.", done);
    }

    // An older server ignores "reject" and pushes back: that is said, never a rejection claimed.
    [Fact]
    public void A_push_back_answered_to_a_rejection_says_so_and_what_to_do()
    {
        string text = ProductionDisplay.PushedBackInsteadOfRejected(
            new BetaRollbackResult("Commodore/C64/250407", BetaRollbackKind.RestoreFromProduction, 1, 0, 1));

        Assert.StartsWith("The server is older than this CRT and PUSHED Commodore/C64/250407 BACK instead of rejecting it", text, StringComparison.Ordinal);

        // There is no "CRT Maintainer" any more (2026-09-29) - the Maintainer tab is CRT.
        Assert.DoesNotContain("CRT Maintainer", text, StringComparison.Ordinal);
        Assert.Contains("1 submission back in the queue", text, StringComparison.Ordinal);
        Assert.Contains("Reject it from Contributor Submissions", text, StringComparison.Ordinal);
    }

    // The button says which operation it performs rather than a generic verb.
    [Fact]
    public void The_confirm_button_names_the_operation()
    {
        Assert.Equal("Roll back", ProductionDisplay.RollBackConfirmButton(ProductionDisplayTests.RollBack()));

        Assert.Equal(
            "Remove from BETA",
            ProductionDisplay.RollBackConfirmButton(ProductionDisplayTests.RollBack(BetaRollbackKind.RemoveFromBeta)));
    }

    // What a finished rollback says: how much moved, and that the contributors were told.
    [Fact]
    public void A_finished_rollback_reports_what_moved_and_who_was_told()
    {
        string done = ProductionDisplay.RolledBack(
            new BetaRollbackResult("Commodore/C64/250407", BetaRollbackKind.RestoreFromProduction, 2, 1, 3));

        Assert.Contains("2 files restored", done, StringComparison.Ordinal);
        Assert.Contains("1 file removed", done, StringComparison.Ordinal);
        Assert.Contains("3 submissions back in the queue", done, StringComparison.Ordinal);
        Assert.Contains("contributors have been told", done, StringComparison.Ordinal);
    }

    // A removal reports only what left - "0 files restored" would be noise.
    [Fact]
    public void A_finished_removal_reports_only_what_left_beta()
    {
        string done = ProductionDisplay.RolledBack(
            new BetaRollbackResult("Commodore/C64/250407", BetaRollbackKind.RemoveFromBeta, 0, 4, 1));

        Assert.Contains("4 files removed from BETA", done, StringComparison.Ordinal);
        Assert.DoesNotContain("restored", done, StringComparison.Ordinal);
    }

    // -----------------------------------------------------------------------------------
    // The button
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_button_needs_BOTH_the_servers_yes_and_the_maintainers_tick()
    {
        Assert.True(ProductionDisplay.CanPress(ProductionDisplayTests.Plan(canPublish: true), checkedInBeta: true));

        // Not ticked: the maintainer has not said they checked it in BETA.
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
