using System;
using System.Collections.Generic;
using System.Globalization;
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

        // ###########################################################################################
        // *** WHICH OTHER BOARDS SEE A REPLACED SHARED FILE (code review, 2026-10-13). *** The two
        // approvers of a shared-file replacement decided without being told which boards cite the
        // file - "other boards may use" and nothing more. The server names them on each replaced
        // shared file of the submission's file tree (BoardFileEntry.AlsoUsedBy); this is the one line
        // for the approval status: "Also used by 3 other boards: Commodore / C64 / 250407, ...". Null
        // when no file names any.
        // ###########################################################################################
        public static string? AlsoUsedByLine(IEnumerable<BoardFileEntry>? files)
        {
            List<string> boards = (files ?? [])
                .Where(file => file is not null)
                .SelectMany(file => file.AlsoUsedBy ?? [])
                .Where(board => !string.IsNullOrWhiteSpace(board))
                .Select(board => board.Trim())
                .Distinct(StringComparer.Ordinal)
                .Order(StringComparer.Ordinal)
                .ToList();

            if (boards.Count == 0)
                return null;

            string count = boards.Count == 1 ? "1 other board" : $"{boards.Count} other boards";

            return $"Also used by {count}: {string.Join(", ", boards)}.";
        }

        // The status line with the boards line after it - either alone when the other is missing.
        public static string? WithAlsoUsedBy(string? statusLine, string? alsoUsedByLine) =>
            statusLine is null ? alsoUsedByLine
            : alsoUsedByLine is null ? statusLine
            : $"{statusLine} {alsoUsedByLine}";

        public static bool CanApprove(ApprovalStatus? status) => status is null || status.CanApprove;

        // Also the reason the submission's table is read-only for such an account (code review,
        // 2026-10-13) - the server takes no change from it either - so it names both.
        public const string NotAMaintainer = "This account is not a maintainer of this board, so it cannot approve or change this submission. Ask the administrator.";

        // ###########################################################################################
        // Above a submission made from an older revision of the board than BETA holds now (code review,
        // 2026-10-13; ReviewQueueEntry.BasedOnOlderRevision) - in the amber panel of what to settle
        // before approving. Rows published since the contributor's copy show as removed, and an
        // approval would remove them.
        // ###########################################################################################
        public const string MadeFromOlderData =
            "This was made from older data than BETA has now - rows published since may show as removed. Compare carefully.";

        // ###########################################################################################
        // Above a submission the same contributor has sent a NEWER one beside (owner decision,
        // 2026-10-13; ReviewQueueEntry.NewerSubmissionId) - in the same amber panel. A submission a
        // maintainer has amended, or approved once of two, is not replaced by the newer one, so both
        // wait, and approving this one publishes the older work. Names the contributor's role, never
        // "their" (owner rule).
        // ###########################################################################################
        public static string NewerFromSameContributor(long? newerId, DateTimeOffset? sentUtc)
        {
            string which = newerId is long id ? string.Create(CultureInfo.InvariantCulture, $" (#{id})") : string.Empty;
            string when = sentUtc is DateTimeOffset sent ? $" on {SubmissionReceiptPresenter.FormatDate(sent)}" : string.Empty;

            return $"The same contributor sent a newer submission for this board{when}{which}, which is waiting too. " +
                "Approving this one publishes the older work - look at the newer one first.";
        }

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
