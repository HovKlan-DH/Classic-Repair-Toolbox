using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers SystemDeletionWording - how Account > "Delete a system" and its confirmation read (owner
// request, 2026-10-03). The confirmation is the last thing between the administrator and a delete
// that cannot be undone, so every place the system is in must be named, and a place holding
// nothing of it must SAY so rather than be left out.
// ###########################################################################################
public sealed class SystemDeletionWordingTests
{
    private static readonly DateTimeOffset Sent = new(2026, 10, 2, 9, 0, 0, TimeSpan.Zero);

    private static SystemDeletePlanAnswer Plan(
        int betaFiles = 12,
        int productionFiles = 11,
        bool listedInBeta = true,
        bool listedInProduction = true,
        bool hasRecord = true,
        int submissions = 5,
        int maintainers = 2,
        int invitations = 0,
        params SystemDeleteOpenSubmission[] open) =>
        new("Commodore/C64/999999", "Commodore", "C64", "999999", "f",
            betaFiles, productionFiles, listedInBeta, listedInProduction, hasRecord,
            submissions, maintainers, invitations, open);

    private static SystemDeleteOpenSubmission Open(long id = 14, string state = "pending", string contributor = "anna@example.com", string? summary = "Corrected U8.") =>
        new(id, state, contributor, summary, SystemDeletionWordingTests.Sent);

    [Fact]
    public void The_headline_names_the_system_as_the_systems_screen_does()
    {
        Assert.Equal("Delete Commodore / C64 / 999999 completely?", SystemDeletionWording.Headline(SystemDeletionWordingTests.Plan()));
    }

    // Each tree, the database and the statistics, in that order - every place it is in.
    [Fact]
    public void What_goes_names_each_tree_the_database_and_the_statistics()
    {
        Assert.Equal(
            [
                "The BETA data: 12 files, and its row in the drop-down lists",
                "The stable data: 11 files, and its row in the drop-down lists",
                "The database: its record, with 5 submissions and 2 maintainers",
                "Its view statistics"
            ],
            SystemDeletionWording.WhatGoes(SystemDeletionWordingTests.Plan()));
    }

    // ###########################################################################################
    // *** THE CASE THAT PROMPTED IT: "Not published yet". *** Nothing in either tree, only a record -
    // each tree SAYS it holds nothing, so the administrator knows it was looked at.
    // ###########################################################################################
    [Fact]
    public void A_system_only_in_the_database_says_each_tree_holds_nothing_of_it()
    {
        IReadOnlyList<string> lines = SystemDeletionWording.WhatGoes(SystemDeletionWordingTests.Plan(
            betaFiles: 0, productionFiles: 0, listedInBeta: false, listedInProduction: false, submissions: 1, maintainers: 1, invitations: 1));

        Assert.Equal("The BETA data: nothing of it", lines[0]);
        Assert.Equal("The stable data: nothing of it", lines[1]);
        Assert.Equal("The database: its record, with 1 submission, 1 maintainer and 1 open invitation", lines[2]);
    }

    [Fact]
    public void A_board_with_no_record_says_so_and_a_listed_board_without_files_names_its_row()
    {
        IReadOnlyList<string> lines = SystemDeletionWording.WhatGoes(SystemDeletionWordingTests.Plan(
            betaFiles: 0, productionFiles: 1, listedInBeta: true, listedInProduction: false, hasRecord: false));

        Assert.Equal("The BETA data: its row in the drop-down lists", lines[0]);
        Assert.Equal("The stable data: 1 file", lines[1]);
        Assert.Equal("The database: nothing - it has no record there", lines[2]);
    }

    // ###########################################################################################
    // *** OPEN SUBMISSIONS: DELETED, AND EACH CONTRIBUTOR MAILED (owner decision, 2026-10-03). ***
    // The heading says both; with none open there is nothing to say and no reason is asked for.
    // ###########################################################################################
    [Fact]
    public void Open_submissions_are_announced_and_a_reason_is_asked_for_only_then()
    {
        SystemDeletePlanAnswer none = SystemDeletionWordingTests.Plan();
        SystemDeletePlanAnswer one = SystemDeletionWordingTests.Plan(open: [SystemDeletionWordingTests.Open()]);
        SystemDeletePlanAnswer two = SystemDeletionWordingTests.Plan(open: [SystemDeletionWordingTests.Open(14), SystemDeletionWordingTests.Open(15)]);

        Assert.Null(SystemDeletionWording.OpenHeading(none));
        Assert.False(SystemDeletionWording.NeedsReason(none));

        Assert.Equal(
            "1 submission to it is still open. It is deleted too, and its contributor is mailed the reason below:",
            SystemDeletionWording.OpenHeading(one));
        Assert.True(SystemDeletionWording.NeedsReason(one));

        Assert.StartsWith("2 submissions to it are still open. These are deleted too, and each contributor", SystemDeletionWording.OpenHeading(two), StringComparison.Ordinal);
    }

    // The state in CRT's own words, the same the contributor reads in "My submissions".
    [Fact]
    public void An_open_submission_reads_number_date_state_contributor_and_description()
    {
        Assert.Equal(
            $"#14 - {SubmissionReceiptPresenter.FormatDate(SystemDeletionWordingTests.Sent)} - {SubmissionReceiptPresenter.DescribeState("returned")} - anna@example.com - Corrected U8.",
            SystemDeletionWording.OpenLine(SystemDeletionWordingTests.Open(state: "returned")));

        Assert.EndsWith(
            "(no contact address) - (no description given)",
            SystemDeletionWording.OpenLine(SystemDeletionWordingTests.Open(contributor: " ", summary: null)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void What_went_is_counted_and_who_was_told_only_when_somebody_was()
    {
        Assert.Equal(
            "Commodore/C64/999999 is deleted: 12 files removed from BETA, 11 files from the stable data, 5 submissions deleted, and 2 contributors were told.",
            SystemDeletionWording.Done(new SystemDeleteAnswer("Commodore/C64/999999", 12, 11, 5, 2)));

        Assert.Equal(
            "Commodore/C64/999999 is deleted: 0 files removed from BETA, 0 files from the stable data, 1 submission deleted.",
            SystemDeletionWording.Done(new SystemDeleteAnswer("Commodore/C64/999999", 0, 0, 1, 0)));
    }

    // ###########################################################################################
    // No user-visible sentence calls a contributor "their", "them" or "they" (owner request,
    // 2026-10-01), and no button label carries "..." (owner rule, 2026-10-03).
    // ###########################################################################################
    [Fact]
    public void No_sentence_calls_a_contributor_their_and_no_button_carries_dots()
    {
        SystemDeletePlanAnswer plan = SystemDeletionWordingTests.Plan(open: [SystemDeletionWordingTests.Open(14), SystemDeletionWordingTests.Open(15)]);

        IEnumerable<string> texts =
        [
            SystemDeletionWording.Explanation,
            SystemDeletionWording.SharedFilesKept,
            SystemDeletionWording.ReasonPrompt,
            SystemDeletionWording.CannotBeUndone,
            SystemDeletionWording.OpenHeading(plan)!,
            SystemDeletionWording.OpenHeading(SystemDeletionWordingTests.Plan(open: [SystemDeletionWordingTests.Open()]))!,
            .. SystemDeletionWording.WhatGoes(plan),
        ];

        foreach (string text in texts)
        {
            string[] words = text.Split([' ', ',', '.', ':', '-'], StringSplitOptions.RemoveEmptyEntries);

            Assert.DoesNotContain(words, word => word.ToLowerInvariant() is "their" or "them" or "they");
        }

        Assert.DoesNotContain("...", SystemDeletionWording.DeleteButton, StringComparison.Ordinal);
        Assert.DoesNotContain("...", SystemDeletionWording.ConfirmButton, StringComparison.Ordinal);
    }
}
