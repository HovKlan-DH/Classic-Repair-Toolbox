using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the two-person approval reads in the Maintainer tab (owner decision,
    // 2026-09-25): a change to a shared file needs the board's maintainer AND the administrator,
    // for the BETA publish and for the production one.
    //
    // The server sends CRT.Data's ApprovalStatus - who must approve, who has, and what THIS
    // account's approval would do - and these sentences are written from it alone, so the screen
    // cannot promise a publish the server would not perform. Pure, so the words are tested.
    // ###########################################################################################
    public static class ApprovalWording
    {
        // ###########################################################################################
        // The line above the buttons, or null for an ordinary item that one approval publishes -
        // nothing to explain there.
        // ###########################################################################################
        public static string? StatusLine(ApprovalStatus? status) => ApprovalWording.StatusLine(status, canPublish: true);

        // ###########################################################################################
        // *** AN ACCOUNT THAT MAY NOT PUBLISH IS TOLD SO IN THIS LINE (2026-09-30). *** It was the
        // Approve button's TOOLTIP, and a tooltip at the bottom of the window is placed over the
        // pointer and takes the click meant for the button (see TabMaintainer's ShowDecisionPanel) -
        // so the button carries none, and the reason is said where it is always visible.
        // ###########################################################################################
        public static string? StatusLine(ApprovalStatus? status, bool canPublish)
        {
            if (!canPublish)
                return ApprovalWording.NotAMaintainer;

            if (status is null || status.Required.Count == 0)
                return null;

            // The administrator ALONE has three causes the status cannot tell apart - nobody
            // maintains the board, only the administrator is in its pool, or only administrators
            // publish to stable - so no reason is given (code review, 2026-10-05: "nobody reviews
            // this board yet" sat beside a Maintainer view naming the administrator).
            string needs = status.Required.Count > 1
                ? "Replaces a shared file other boards may use, so it needs a maintainer of this board AND the administrator."
                : "Replaces a shared file other boards may use, so it needs the administrator's approval.";

            string given = status.Given.Count == 0
                ? "Nobody has approved it yet."
                : "Approved so far by " + string.Join(" and ", status.Given.Select(approval => $"{approval.By} ({ApprovalWording.RoleName(approval.Role)})")) + ".";

            return $"{needs} {given}";
        }

        // ###########################################################################################
        // The Approve button's text. `target` names where the publish goes ("BETA", "production").
        // It says what pressing it will actually do: publish, record one of two approvals, or
        // nothing because this account's part is done.
        // ###########################################################################################
        public static string ApproveButton(ApprovalStatus? status, string target)
        {
            if (status is null || status.ApprovalPublishes)
                return $"Approve and publish to {target}";

            if (status.CanApprove)
                return $"Approve - {ApprovalWording.Others(status)} must approve too";

            string waiting = $"waiting for {ApprovalWording.Names(status.WaitingFor)}";

            // This account's part is done - by this account, or by somebody else in the same role
            // (a board's second maintainer). Said apart, so nobody reads a colleague's approval as
            // their own (code review, 2026-09-25). Who it was is on the status line.
            if (status.YouApproved)
                return $"You approved - {waiting}";

            if (status.YourRole is not null && status.Given.Any(approval => approval.Role == status.YourRole))
            {
                return status.YourRole == ApproverRole.Administrator
                    ? $"Another administrator approved - {waiting}"
                    : $"Another maintainer approved - {waiting}";
            }

            return $"Waiting for {ApprovalWording.Names(status.WaitingFor)}";
        }

        public static bool CanApprove(ApprovalStatus? status) => status is null || status.CanApprove;

        public const string NotAMaintainer = "This account is not a maintainer of this board, so it cannot approve this. Ask the administrator.";

        // ###########################################################################################
        // What the maintainer is told after an approval that did NOT publish.
        // ###########################################################################################
        public static string Recorded(IReadOnlyList<ApproverRole>? waitingFor, string target)
        {
            string who = waitingFor is { Count: > 0 } ? ApprovalWording.Names(waitingFor) : "the other approver";

            return $"Your approval is recorded. It is published to {target} once {who} has approved too.";
        }

        public static string RoleName(ApproverRole role) =>
            role == ApproverRole.Administrator ? "administrator" : "maintainer";

        // "the administrator", "a maintainer of this board", or both.
        public static string Names(IEnumerable<ApproverRole> roles) =>
            string.Join(" and ", roles.Select(role => role == ApproverRole.Administrator ? "the administrator" : "a maintainer of this board"));

        // The roles still needed besides this account's own.
        private static string Others(ApprovalStatus status) =>
            ApprovalWording.Names(status.WaitingFor.Where(role => role != status.YourRole));
    }
}
