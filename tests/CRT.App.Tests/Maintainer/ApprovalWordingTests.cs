using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ApprovalWording and the parsing behind it - the two-person approval a shared-file change
// needs (owner decision, 2026-09-25), as the Maintainer tab shows it.
//
// The rule that matters is that the Approve button says what pressing it will ACTUALLY do: publish,
// record one of two approvals, or nothing because this account's part is done. Each of those is
// read off the server's ApprovalStatus, never worked out here.
public sealed class ApprovalWordingTests
{
    private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

    private static ApprovalStatus Both(ApproverRole you, params GivenApproval[] given) =>
        ApprovalRules.Status([ApproverRole.Maintainer, ApproverRole.Administrator], given, you);

    private const long AnnaId = 7;

    private static GivenApproval Anna() => new(ApproverRole.Maintainer, "Anna (anna@example.com)", ApprovalWordingTests.Now, ApprovalWordingTests.AnnaId);

    // -----------------------------------------------------------------------------------
    // The button
    // -----------------------------------------------------------------------------------

    [Fact]
    public void An_ordinary_submission_keeps_the_one_press_publish_button()
    {
        ApprovalStatus ordinary = ApprovalRules.Status([], [], ApproverRole.Maintainer);

        Assert.Equal("Approve and publish to BETA", ApprovalWording.ApproveButton(ordinary, "BETA"));
        Assert.Null(ApprovalWording.StatusLine(ordinary));

        // An older server sends no status at all; one approval publishing is what it did.
        Assert.Equal("Approve and publish to BETA", ApprovalWording.ApproveButton(null, "BETA"));
        Assert.True(ApprovalWording.CanApprove(null));
    }

    [Fact]
    public void The_FIRST_of_two_approvals_says_the_other_must_approve_too()
    {
        ApprovalStatus first = ApprovalWordingTests.Both(ApproverRole.Maintainer);

        Assert.Equal("Approve - the administrator must approve too", ApprovalWording.ApproveButton(first, "BETA"));
        Assert.True(ApprovalWording.CanApprove(first));
    }

    [Fact]
    public void The_SECOND_approval_says_it_publishes()
    {
        ApprovalStatus second = ApprovalWordingTests.Both(ApproverRole.Administrator, ApprovalWordingTests.Anna());

        Assert.Equal("Approve and publish to production", ApprovalWording.ApproveButton(second, "production"));
    }

    [Fact]
    public void An_account_whose_part_is_done_sees_who_it_waits_for_and_cannot_press()
    {
        // Anna approved; this is Anna's screen.
        ApprovalStatus done = ApprovalRules.Status(
            [ApproverRole.Maintainer, ApproverRole.Administrator], [ApprovalWordingTests.Anna()], ApproverRole.Maintainer, ApprovalWordingTests.AnnaId);

        Assert.Equal("You approved - waiting for the administrator", ApprovalWording.ApproveButton(done, "BETA"));
        Assert.False(ApprovalWording.CanApprove(done));
    }

    // ###########################################################################################
    // *** A BOARD'S SECOND MAINTAINER IS NOT TOLD THEY APPROVED (code review, 2026-09-25). *** Their
    // part is done too - a second maintainer approval counts for nothing - but by a colleague, and
    // the button says so rather than reading as their own approval.
    // ###########################################################################################
    [Fact]
    public void A_colleagues_approval_in_the_same_role_is_not_shown_as_yours()
    {
        ApprovalStatus colleague = ApprovalRules.Status(
            [ApproverRole.Maintainer, ApproverRole.Administrator], [ApprovalWordingTests.Anna()], ApproverRole.Maintainer, yourAccountId: 8);

        Assert.Equal("Another maintainer approved - waiting for the administrator", ApprovalWording.ApproveButton(colleague, "BETA"));
        Assert.False(ApprovalWording.CanApprove(colleague));

        GivenApproval admin = new(ApproverRole.Administrator, "Admin (admin@example.com)", ApprovalWordingTests.Now, 1);
        ApprovalStatus secondAdmin = ApprovalRules.Status(
            [ApproverRole.Maintainer, ApproverRole.Administrator], [admin], ApproverRole.Administrator, yourAccountId: 2);

        Assert.Equal("Another administrator approved - waiting for a maintainer of this board", ApprovalWording.ApproveButton(secondAdmin, "production"));
    }

    // -----------------------------------------------------------------------------------
    // The line above the buttons
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_status_line_names_both_roles_and_who_has_approved()
    {
        string? line = ApprovalWording.StatusLine(ApprovalWordingTests.Both(ApproverRole.Administrator, ApprovalWordingTests.Anna()));

        Assert.NotNull(line);
        Assert.Contains("a maintainer of this board AND the administrator", line, StringComparison.Ordinal);
        Assert.Contains("Anna (anna@example.com) (maintainer)", line, StringComparison.Ordinal);
    }

    [Fact]
    public void A_board_with_no_maintainers_says_the_administrator_alone_decides()
    {
        string? line = ApprovalWording.StatusLine(
            ApprovalRules.Status([ApproverRole.Administrator], [], ApproverRole.Administrator));

        Assert.Contains("needs the administrator", line, StringComparison.Ordinal);
    }

    [Fact]
    public void A_recorded_approval_says_it_was_NOT_published_yet()
    {
        Assert.Equal(
            "Your approval is recorded. It is published to BETA once the administrator has approved too.",
            ApprovalWording.Recorded([ApproverRole.Administrator], "BETA"));
    }

    // -----------------------------------------------------------------------------------
    // Reading it off the server's answers
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_detail_carries_the_servers_approval_status_as_the_same_record()
    {
        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission("""
            {"canPublish":true,
             "submission":{"id":42,"systemId":"Commodore/C64/250407","state":"approved","summary":"x","touchesSharedFiles":true},
             "findings":[],
             "approval":{"required":["Maintainer","Administrator"],
                         "given":[{"role":"Maintainer","by":"Anna (anna@example.com)","atUtc":"2026-09-25T12:00:00+00:00"}],
                         "waitingFor":["Administrator"],"yourRole":"Administrator","canApprove":true,"approvalPublishes":true}}
            """);

        Assert.NotNull(detail);
        ApprovalStatus approval = detail!.Approval!;
        Assert.Equal([ApproverRole.Administrator], approval.WaitingFor);
        Assert.True(approval.ApprovalPublishes);
        Assert.Equal("Anna (anna@example.com)", Assert.Single(approval.Given).By);
    }

    [Fact]
    public void A_recorded_approval_reads_as_recorded_rather_than_published()
    {
        ReviewDecisionResult? result = ReviewApiParser.ParseDecision("""{"state":"approved","waitingFor":["Administrator"]}""");

        Assert.Equal([ApproverRole.Administrator], result!.WaitingFor);
        Assert.Equal(
            "Your approval is recorded. It is published to BETA once the administrator has approved too.",
            ReviewDecisionWording.Describe(ReviewDecisionKind.Approve, result));
    }

    [Fact]
    public void A_production_approval_that_published_nothing_says_so()
    {
        ProductionPublishResult? result = ReviewApiParser.ParseProductionPublish(
            """{"systemId":"Commodore/C64/250407","state":"awaiting","waitingFor":["Maintainer"]}""");

        Assert.True(result!.IsAwaitingApproval);
        Assert.Equal([ApproverRole.Maintainer], result.WaitingFor);
    }

    [Fact]
    public void The_queue_row_says_when_one_of_two_approvals_is_given()
    {
        var row = new ReviewQueueRow(42, "Commodore/C64/250407", "approved", "Corrected R12.", "c@example.com", null, TouchesSharedFiles: true);

        // No wait is known here, so the footer begins with the shared-file note - its own line in
        // the queue row since 2026-09-26, and capitalised as one.
        Assert.Equal(
            "Replaces a shared file - one of two approvals given",
            ReviewQueueDisplay.Footer(row, ApprovalWordingTests.Now));
    }
}
