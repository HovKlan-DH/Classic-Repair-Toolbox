using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHO MUST APPROVE before something is published (owner decision, 2026-09-25): "in case of
    // changes to any shared file, then both the maintainer and the admin should approve before
    // publishing to BETA or production. If there is no shared files changed, then normal maintainer
    // is sufficient."
    //
    //   - Nothing shared changes     -> ONE approval, from a maintainer of the board or an
    //                                   administrator. The publish happens on that approval.
    //   - A shared file changes      -> a maintainer of the board AND an administrator, in either
    //                                   order. The first approval is recorded and the item waits;
    //                                   the second one publishes.
    //   - ...and the board has no
    //     maintainers yet              -> the administrator alone: there is nobody to ask for the
    //                                   other half, and the administrator is in every pool anyway.
    //
    // The same rule for BOTH publishes: a submission to BETA, and a board's promotion to
    // production. A "no" never needs two people - either of them may reject or return a
    // submission on their own.
    //
    // Pure, and in CRT.Data with its status record so the server computes it and the review
    // application reads the SAME record off the wire - the SubmittedFileFact rule.
    // ###########################################################################################
    public static class ApprovalRules
    {
        // An empty list means "any one approval": the ordinary case.
        public static IReadOnlyList<ApproverRole> Required(bool touchesSharedFiles, bool boardHasMaintainers)
        {
            if (!touchesSharedFiles)
                return [];

            return boardHasMaintainers
                ? [ApproverRole.Maintainer, ApproverRole.Administrator]
                : [ApproverRole.Administrator];
        }

        // ###########################################################################################
        // Where an item stands for ONE account: what is needed, what has been given, and whether
        // this account's approval would be the one that publishes.
        //
        // `yourAccountId` says whether THIS account gave an approval already, as opposed to someone
        // else in the same role (code review, 2026-09-25). A board's second maintainer cannot add a
        // second maintainer approval either way, but was told "You have already approved this" for
        // an approval a colleague gave.
        // ###########################################################################################
        public static ApprovalStatus Status(
            IReadOnlyList<ApproverRole> required,
            IReadOnlyList<GivenApproval> given,
            ApproverRole? yourRole,
            long? yourAccountId = null)
        {
            ArgumentNullException.ThrowIfNull(required);
            ArgumentNullException.ThrowIfNull(given);

            bool alreadyGiven = yourRole is not null && given.Any(approval => approval.Role == yourRole);

            // Needed: any role will do for an ordinary item; otherwise the role must be one asked for.
            bool needed = yourRole is not null && (required.Count == 0 || required.Contains(yourRole.Value));

            IReadOnlyList<ApproverRole> waitingFor = required
                .Where(role => given.All(approval => approval.Role != role))
                .ToList();

            // *** EVERY REQUIRED APPROVAL ALREADY GIVEN, yet nothing published (code review,
            // 2026-09-25). *** Only possible when the requirement SHRANK after an approval - the
            // board's last maintainer left its pool, so a shared change needs the administrator
            // alone, and the administrator approved first. Nothing is waited for, so the next
            // approval by a role the item needs completes it; refusing it as "already approved"
            // left the item waiting for nobody, for ever.
            //
            // The same when it shrank to "any one approval" (2026-09-27): a submission given the
            // first of two approvals for ADDING a shared file - which now needs one - already has
            // its approval. Whoever approves next in a role it needs, its first approver included,
            // publishes it; otherwise that maintainer was told "already approved" of an item only
            // the administrator could then finish. The approvals table ignores a repeat in a role.
            bool satisfied = given.Count > 0 && waitingFor.Count == 0;

            bool canApprove = needed && (!alreadyGiven || satisfied);

            bool completes = canApprove &&
                (required.Count == 0 || waitingFor.All(role => role == yourRole));

            bool youApproved = yourAccountId is not null && given.Any(approval => approval.AccountId == yourAccountId);

            return new ApprovalStatus(required, given, waitingFor, yourRole, canApprove, completes, youApproved);
        }
    }

    [JsonConverter(typeof(JsonStringEnumConverter<ApproverRole>))]
    public enum ApproverRole
    {
        Maintainer,
        Administrator
    }

    // One approval already recorded. By is a readable label (name and address), stored with the
    // approval so it survives the account being renamed or removed - the audit table's rule.
    // AccountId is who gave it, for "did YOU approve this"; null once the account is deleted. It
    // stays on the server - every maintainer is sent the list, and the label says enough.
    public sealed record GivenApproval(
        ApproverRole Role,
        string By,
        DateTimeOffset AtUtc,
        [property: JsonIgnore] long? AccountId = null);

    // ###########################################################################################
    // An item's approval position, as one account sees it. On the wire as this very record.
    //
    //   Required        - empty for "any one approval"; otherwise the roles that must all approve.
    //   Given           - what has been recorded so far.
    //   WaitingFor      - the required roles not yet given.
    //   YourRole        - the role this account approves as here, or null when it may not.
    //   CanApprove      - this account may add its approval now (not already given in its role).
    //   ApprovalPublishes - this account's approval would complete it, and so publish.
    //   YouApproved     - THIS account gave one of the approvals in Given. False when somebody
    //                     else gave the approval in its role, which is why it cannot approve.
    // ###########################################################################################
    public sealed record ApprovalStatus(
        IReadOnlyList<ApproverRole> Required,
        IReadOnlyList<GivenApproval> Given,
        IReadOnlyList<ApproverRole> WaitingFor,
        ApproverRole? YourRole,
        bool CanApprove,
        bool ApprovalPublishes,
        bool YouApproved = false);
}
