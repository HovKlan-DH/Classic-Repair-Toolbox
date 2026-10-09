using System;
using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers ApprovalRules - who must approve before a publish (owner decision, 2026-09-25):
    // one maintainer normally; a maintainer AND the administrator when a shared file changes.
    // ###########################################################################################
    public sealed class ApprovalRulesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private static GivenApproval By(ApproverRole role) => new(role, role.ToString(), ApprovalRulesTests.Now);

        // -----------------------------------------------------------------------------------
        // What is required
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Nothing_shared_needs_ANY_one_approval()
        {
            Assert.Empty(ApprovalRules.Required(touchesSharedFiles: false, boardHasMaintainers: true));
            Assert.Empty(ApprovalRules.Required(touchesSharedFiles: false, boardHasMaintainers: false));
        }

        [Fact]
        public void A_shared_change_needs_a_maintainer_AND_the_administrator()
        {
            Assert.Equal(
                [ApproverRole.Maintainer, ApproverRole.Administrator],
                ApprovalRules.Required(touchesSharedFiles: true, boardHasMaintainers: true));
        }

        [Fact]
        public void A_shared_change_on_a_board_with_NO_maintainers_needs_the_administrator_alone()
        {
            // There is nobody to ask for the other half.
            Assert.Equal([ApproverRole.Administrator], ApprovalRules.Required(touchesSharedFiles: true, boardHasMaintainers: false));
        }

        // -----------------------------------------------------------------------------------
        // The ordinary case
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(ApproverRole.Maintainer)]
        [InlineData(ApproverRole.Administrator)]
        public void Without_shared_files_either_role_publishes_alone(ApproverRole role)
        {
            ApprovalStatus status = ApprovalRules.Status([], [], role);

            Assert.True(status.CanApprove);
            Assert.True(status.ApprovalPublishes);
            Assert.Empty(status.WaitingFor);
        }

        [Fact]
        public void An_account_with_no_role_can_approve_nothing()
        {
            ApprovalStatus status = ApprovalRules.Status([], [], yourRole: null);

            Assert.False(status.CanApprove);
            Assert.False(status.ApprovalPublishes);
        }

        // -----------------------------------------------------------------------------------
        // Both must approve
        // -----------------------------------------------------------------------------------

        private static readonly IReadOnlyList<ApproverRole> Both = [ApproverRole.Maintainer, ApproverRole.Administrator];

        [Theory]
        [InlineData(ApproverRole.Maintainer)]
        [InlineData(ApproverRole.Administrator)]
        public void The_FIRST_approval_of_a_shared_change_does_not_publish(ApproverRole first)
        {
            ApprovalStatus status = ApprovalRules.Status(ApprovalRulesTests.Both, [], first);

            Assert.True(status.CanApprove);
            Assert.False(status.ApprovalPublishes);
        }

        [Fact]
        public void The_SECOND_role_publishes_in_either_order()
        {
            ApprovalStatus adminAfterMaintainer = ApprovalRules.Status(
                ApprovalRulesTests.Both, [ApprovalRulesTests.By(ApproverRole.Maintainer)], ApproverRole.Administrator);

            ApprovalStatus maintainerAfterAdmin = ApprovalRules.Status(
                ApprovalRulesTests.Both, [ApprovalRulesTests.By(ApproverRole.Administrator)], ApproverRole.Maintainer);

            Assert.True(adminAfterMaintainer.ApprovalPublishes);
            Assert.True(maintainerAfterAdmin.ApprovalPublishes);
        }

        [Fact]
        public void The_same_role_cannot_approve_twice_and_the_status_says_who_is_awaited()
        {
            // A second maintainer of the same board does not count as the administrator.
            ApprovalStatus status = ApprovalRules.Status(
                ApprovalRulesTests.Both, [ApprovalRulesTests.By(ApproverRole.Maintainer)], ApproverRole.Maintainer);

            Assert.False(status.CanApprove);
            Assert.False(status.ApprovalPublishes);
            Assert.Equal([ApproverRole.Administrator], status.WaitingFor);
        }

        // ###########################################################################################
        // WHO gave an approval, not only in which role (code review, 2026-09-25): a board's second
        // maintainer cannot add a second maintainer approval, but did not give the first one either.
        // ###########################################################################################
        [Fact]
        public void YouApproved_is_true_only_for_the_account_that_gave_the_approval()
        {
            GivenApproval anna = new(ApproverRole.Maintainer, "Anna", DateTimeOffset.UnixEpoch, AccountId: 7);

            ApprovalStatus annas = ApprovalRules.Status(ApprovalRulesTests.Both, [anna], ApproverRole.Maintainer, yourAccountId: 7);
            ApprovalStatus bos = ApprovalRules.Status(ApprovalRulesTests.Both, [anna], ApproverRole.Maintainer, yourAccountId: 8);
            ApprovalStatus unknown = ApprovalRules.Status(ApprovalRulesTests.Both, [anna], ApproverRole.Maintainer);

            Assert.True(annas.YouApproved);
            Assert.False(bos.YouApproved);
            Assert.False(unknown.YouApproved);

            // Neither may add a second maintainer approval.
            Assert.False(annas.CanApprove);
            Assert.False(bos.CanApprove);
        }

        // Every maintainer is sent the list of approvals; who gave each is the label, not an id.
        [Fact]
        public void The_approvers_account_id_stays_off_the_wire()
        {
            GivenApproval anna = new(ApproverRole.Maintainer, "Anna", DateTimeOffset.UnixEpoch, AccountId: 7);

            string json = JsonSerializer.Serialize(anna, new JsonSerializerOptions(JsonSerializerDefaults.Web));

            Assert.DoesNotContain("accountId", json, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void The_administrator_alone_publishes_when_the_board_has_no_maintainers()
        {
            ApprovalStatus status = ApprovalRules.Status([ApproverRole.Administrator], [], ApproverRole.Administrator);

            Assert.True(status.ApprovalPublishes);
        }

        // ###########################################################################################
        // *** WHAT IS REQUIRED CAN SHRINK AFTER AN APPROVAL, AND THE ITEM MUST NOT STICK (code review,
        // 2026-09-25). *** The administrator approves a shared change first; the board's only
        // maintainer then leaves its pool, so the item needs the administrator alone - who has already
        // approved. Nothing is waited for, yet the administrator was refused as "already approved"
        // and no other role existed to press the button: the item sat in 'approved' for ever.
        // ###########################################################################################
        [Fact]
        public void An_item_whose_every_required_approval_is_given_publishes_on_the_next_approval()
        {
            GivenApproval admin = new(ApproverRole.Administrator, "Admin", ApprovalRulesTests.Now, AccountId: 1);

            ApprovalStatus status = ApprovalRules.Status([ApproverRole.Administrator], [admin], ApproverRole.Administrator, yourAccountId: 1);

            Assert.Empty(status.WaitingFor);
            Assert.True(status.CanApprove);
            Assert.True(status.ApprovalPublishes);
        }

        // The same shrink seen by an account whose role is no longer asked for changes nothing: only
        // an approver the item still needs completes it.
        [Fact]
        public void A_role_the_item_no_longer_needs_still_cannot_approve_it()
        {
            ApprovalStatus status = ApprovalRules.Status(
                [ApproverRole.Administrator], [ApprovalRulesTests.By(ApproverRole.Administrator)], ApproverRole.Maintainer);

            Assert.False(status.CanApprove);
            Assert.False(status.ApprovalPublishes);
        }

        // ###########################################################################################
        // *** THE REQUIREMENT SHRANK TO "ANY ONE APPROVAL" (2026-09-27). *** A maintainer gave the
        // first of two approvals for ADDING a shared file, which now needs only one. That approval is
        // enough: the maintainer who gave it can press Approve again and it publishes, rather than
        // being told "already approved" of an item only the administrator could then finish.
        // ###########################################################################################
        [Fact]
        public void An_item_already_approved_once_when_one_approval_became_enough_is_finished_by_its_approver()
        {
            GivenApproval maintainer = new(ApproverRole.Maintainer, "Anna", ApprovalRulesTests.Now, AccountId: 7);

            ApprovalStatus status = ApprovalRules.Status([], [maintainer], ApproverRole.Maintainer, yourAccountId: 7);

            Assert.True(status.CanApprove);
            Assert.True(status.ApprovalPublishes);
        }

        // -----------------------------------------------------------------------------------
        // On the wire
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_status_round_trips_with_its_roles_as_NAMES()
        {
            // The server serialises this record and the Maintainer tab reads it back as the
            // same record; roles travel as names so a member added later cannot shift the others.
            ApprovalStatus status = ApprovalRules.Status(
                ApprovalRulesTests.Both, [ApprovalRulesTests.By(ApproverRole.Maintainer)], ApproverRole.Administrator);

            var web = new JsonSerializerOptions(JsonSerializerDefaults.Web);
            string json = JsonSerializer.Serialize(status, web);

            Assert.Contains("\"Administrator\"", json, StringComparison.Ordinal);

            ApprovalStatus? back = JsonSerializer.Deserialize<ApprovalStatus>(json, web);

            Assert.NotNull(back);
            Assert.Equal(status.WaitingFor, back!.WaitingFor);
            Assert.Equal(ApproverRole.Maintainer, Assert.Single(back.Given).Role);
            Assert.True(back.ApprovalPublishes);
        }
    }
}
