using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// Covers BetaRollbackPlan - rolling a BETA board back to what production holds (owner decision,
// 2026-09-27), the "push back to queue" the production window was missing.
//
// The two facts it rests on are asserted elsewhere and worth naming here: a MERGED submission
// keeps its blobs (SubmissionCollectionStates.Live contains Merged, pinned by
// SubmissionCollectionStatesTests), and production holds a COMPLETE board rather than a partial
// set (ProductionPromotionPlan). Without either, "overwrite BETA from production" would not be a
// well-defined restore.
//
// The limit that survives, and the reason the plan names its submissions: a rollback is PER
// SYSTEM. Three contributors merged since the last promotion means all three go back.
// ###########################################################################################
public sealed class BetaRollbackPlanTests
{
    private static readonly DateTimeOffset Decided = new(2026, 9, 24, 8, 30, 0, TimeSpan.Zero);

    private static CarriedSubmission Submission(long id = 7, string email = "hest@mailscan.dk") =>
        new(id, email, "Corrected U8.", BetaRollbackPlanTests.Decided);

    // Nothing matches by default, so every production file counts as needing restoring.
    private static bool Differs(string path) => false;

    // ###########################################################################################
    // *** PRODUCTION'S BYTES WIN, AND BETA-ONLY FILES GO. *** Both halves are needed. Restoring
    // alone would leave every file a rolled-back submission ADDED sitting in the tree - cited by
    // nothing, and still shipped to every BETA user.
    // ###########################################################################################
    [Fact]
    public void A_rollback_restores_productions_files_and_removes_the_ones_only_beta_has()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png", "board/u8.png", "board/new-from-the-submission.png"],
            productionOwnFiles: ["board/sheet.png", "board/u8.png"],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning: [BetaRollbackPlanTests.Submission()]);

        Assert.Equal(BetaRollbackKind.RestoreFromProduction, plan.Kind);
        Assert.Equal(["board/sheet.png", "board/u8.png"], plan.Restored);
        Assert.Equal(["board/new-from-the-submission.png"], plan.Removed);
    }

    // A file BETA already holds identically is not rewritten - the same "only what actually moved"
    // rule ProductionPromotionPlan follows in the other direction.
    [Fact]
    public void A_file_that_already_matches_production_is_not_restored_again()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png", "board/u8.png"],
            productionOwnFiles: ["board/sheet.png", "board/u8.png"],
            sameBytes: path => path == "board/sheet.png",
            returning: []);

        Assert.Equal(["board/u8.png"], plan.Restored);
        Assert.Empty(plan.Removed);
    }

    // ###########################################################################################
    // *** A SYSTEM NEVER PROMOTED IS A REMOVAL, NOT A RESTORE (owner decision, 2026-09-27). ***
    // Production has nothing to restore from, so the board leaves the BETA tree entirely. The
    // alternative - "restore nothing" - would silently leave the bad board exactly as it is while
    // reporting success.
    // ###########################################################################################
    [Fact]
    public void A_system_never_promoted_is_removed_from_beta_rather_than_restored()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png", "board/u8.png"],
            productionOwnFiles: [],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning: [BetaRollbackPlanTests.Submission()]);

        Assert.Equal(BetaRollbackKind.RemoveFromBeta, plan.Kind);
        Assert.Empty(plan.Restored);
        Assert.Equal(["board/sheet.png", "board/u8.png"], plan.Removed);
    }

    // ###########################################################################################
    // *** EVERY SUBMISSION SINCE THE LAST PROMOTION GOES BACK, and the plan says so by name. ***
    // A rollback is per SYSTEM: PublishMerge replaces rows wholesale, so nothing records whose row
    // was whose and one contributor's work cannot be picked out. Naming them is what stops the
    // feature quietly discarding two other people's accepted work.
    // ###########################################################################################
    [Fact]
    public void Every_submission_merged_since_the_last_promotion_is_named_as_returning()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png"],
            productionOwnFiles: ["board/sheet.png"],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning:
            [
                BetaRollbackPlanTests.Submission(41, "one@example.com"),
                BetaRollbackPlanTests.Submission(38, "two@example.com"),
                BetaRollbackPlanTests.Submission(35, "three@example.com"),
            ]);

        Assert.Equal([41, 38, 35], plan.Returning.Select(submission => submission.Id));
    }

    // A board already level with production changes no file. It is still a real outcome - the
    // submissions return to the queue - so the caller is told rather than being refused.
    [Fact]
    public void A_board_already_level_with_production_changes_no_file()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png"],
            productionOwnFiles: ["board/sheet.png"],
            sameBytes: _ => true,
            returning: [BetaRollbackPlanTests.Submission()]);

        Assert.True(plan.ChangesNothing);
        Assert.Single(plan.Returning);
    }

    // ###########################################################################################
    // *** TWO SPELLINGS ARE TWO FILES ON THE SERVER (code review, 2026-09-27). *** This test used
    // to assert the opposite - that "Sheet.png" and "sheet.png" were one file - which is true on a
    // Windows client and false on the AlmaLinux server whose trees a rollback writes. Paired that
    // way, BETA kept the submission's "Sheet.png" beside the restored "sheet.png" and went on
    // serving it. A rollback makes the folder IDENTICAL to production's, so the BETA-only spelling
    // leaves like any other BETA-only file.
    // ###########################################################################################
    [Fact]
    public void A_path_differing_only_by_capitalisation_is_a_different_file_and_leaves_beta()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/Sheet.png"],
            productionOwnFiles: ["board/sheet.png"],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning: []);

        Assert.Equal(["board/sheet.png"], plan.Restored);
        Assert.Equal(["board/Sheet.png"], plan.Removed);
    }

    // Two spellings in ONE tree are two files as well - the old case-insensitive Distinct dropped
    // one of them silently, so it was neither restored nor removed.
    [Fact]
    public void Two_spellings_in_one_tree_are_both_kept_in_view()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/a.png", "board/A.png"],
            productionOwnFiles: ["board/x.png"],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning: []);

        Assert.Equal(["board/A.png", "board/a.png"], plan.Removed);
    }

    // -----------------------------------------------------------------------------------
    // SHARED FILES (code review, 2026-09-27)
    //
    // Leaving a returning submission's shared change in BETA was a leak: the next promotion of ANY
    // board citing the file carried those un-reviewed bytes to production while the submission
    // that introduced them sat in the queue. Only a change the rollback can prove is its own is
    // touched, because a shared file reaches every board citing it.
    // -----------------------------------------------------------------------------------

    private static BetaRollbackPlanResult WithShared(params BetaRollbackSharedFile[] shared) =>
        BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png"],
            productionOwnFiles: ["board/sheet.png"],
            sameBytes: _ => true,
            returning: [BetaRollbackPlanTests.Submission()],
            sharedFiles: shared);

    // ###########################################################################################
    // *** AN EMPTY PRODUCTION LISTING IS NOT "NEVER PROMOTED" (code review, 2026-09-27). *** A
    // promoted system whose production folder is missing or unreadable lists nothing too - and read
    // as never promoted, its board was wiped out of BETA. Only the record decides; with it, the plan
    // does nothing at all.
    // ###########################################################################################
    [Fact]
    public void A_promoted_system_whose_production_folder_lists_nothing_is_neither_restored_nor_removed()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png", "board/Data C64 250407.xlsx"],
            productionOwnFiles: [],
            sameBytes: _ => false,
            returning: [BetaRollbackPlanTests.Submission()],
            promotedBefore: true);

        Assert.Equal(BetaRollbackKind.ProductionUnreadable, plan.Kind);
        Assert.Empty(plan.Restored);
        Assert.Empty(plan.Removed);
        Assert.Empty(plan.SharedRestored);
    }

    // The same empty listing for a system never promoted is still the removal it always was.
    [Fact]
    public void A_system_never_promoted_is_still_removed_when_production_lists_nothing()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png"],
            productionOwnFiles: [],
            sameBytes: _ => false,
            returning: [BetaRollbackPlanTests.Submission()],
            promotedBefore: false);

        Assert.Equal(BetaRollbackKind.RemoveFromBeta, plan.Kind);
        Assert.Equal(["board/sheet.png"], plan.Removed);
    }

    // BETA still holds exactly what the submission wrote, and production has an older version:
    // production's bytes go back.
    [Fact]
    public void A_shared_file_the_submission_changed_is_put_back_to_productions_bytes()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlanTests.WithShared(
            new BetaRollbackSharedFile("Commodore/Shared files/6526.png", BetaHoldsTheSubmittedBytes: true, InProduction: true, SameAsProduction: false));

        Assert.Equal(["Commodore/Shared files/6526.png"], plan.SharedRestored);
        Assert.True(plan.TouchesSharedFiles);
        Assert.False(plan.ChangesNothing);
    }

    // ###########################################################################################
    // A shared file the submission ADDED has nothing in production to go back to - and it STAYS
    // (owner decision, 2026-09-27): nothing outside the system's own folder is removed
    // automatically. Unused, it is Admin > Unused files' to clear.
    // ###########################################################################################
    [Fact]
    public void A_shared_file_the_submission_added_is_left_in_place()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlanTests.WithShared(
            new BetaRollbackSharedFile("Generic shared files/new.pdf", BetaHoldsTheSubmittedBytes: true, InProduction: false, SameAsProduction: false));

        Assert.Empty(plan.SharedRestored);
        Assert.False(plan.TouchesSharedFiles);
        Assert.DoesNotContain("Generic shared files/new.pdf", plan.Removed);
    }

    // ###########################################################################################
    // *** SOMEBODY WROTE IT AGAIN SINCE - NOT OURS TO UNDO. *** BETA no longer holds the bytes the
    // returning submission carried, so a later publish owns the current version. Reverting it would
    // silently discard that other board's accepted change.
    // ###########################################################################################
    [Fact]
    public void A_shared_file_written_again_since_is_left_alone()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlanTests.WithShared(
            new BetaRollbackSharedFile("Commodore/Shared files/6526.png", BetaHoldsTheSubmittedBytes: false, InProduction: true, SameAsProduction: false));

        Assert.False(plan.TouchesSharedFiles);
        Assert.True(plan.ChangesNothing);
    }

    // Production already has the same bytes - another board's promotion carried the change out, so
    // it is public and reviewed. Nothing to undo in BETA.
    [Fact]
    public void A_shared_file_production_already_has_is_left_alone()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlanTests.WithShared(
            new BetaRollbackSharedFile("Commodore/Shared files/6526.png", BetaHoldsTheSubmittedBytes: true, InProduction: true, SameAsProduction: true));

        Assert.False(plan.TouchesSharedFiles);
    }

    // A never-promoted system's shared changes follow the same rule: its own folder leaves BETA,
    // and a shared file it changed goes back to what production has for the OTHER boards.
    [Fact]
    public void A_never_promoted_system_still_puts_its_shared_changes_back()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png"],
            productionOwnFiles: [],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning: [BetaRollbackPlanTests.Submission()],
            sharedFiles: [new BetaRollbackSharedFile("Commodore/Shared files/6526.png", true, InProduction: true, SameAsProduction: false)]);

        Assert.Equal(BetaRollbackKind.RemoveFromBeta, plan.Kind);
        Assert.Equal(["Commodore/Shared files/6526.png"], plan.SharedRestored);
    }

    // Nothing merged since the last promotion (a promotion re-run, say) returns nobody - which is
    // correct, and must not be reported as though a contributor were affected.
    [Fact]
    public void A_rollback_that_returns_nobody_says_so()
    {
        BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
            betaOwnFiles: ["board/sheet.png"],
            productionOwnFiles: ["board/sheet.png"],
            sameBytes: BetaRollbackPlanTests.Differs,
            returning: null);

        Assert.Empty(plan.Returning);
    }
}
