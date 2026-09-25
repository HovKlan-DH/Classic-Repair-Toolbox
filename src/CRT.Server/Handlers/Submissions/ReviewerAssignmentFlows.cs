using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The administrator's side of Phase 6 roles (2026-09-25): who reviews which system.
    //
    // Every rule is here, pure, taking its stores as arguments - the same shape SubmissionFlows
    // and AccountFlows have, for the same reason: AdminEndpoints is a rim, and a rule that lives in
    // a handler body is a rule no test can reach.
    //
    // *** ADMINISTRATOR ONLY, AND THERE IS NO WAY TO BECOME ONE HERE. *** Granting administrator
    // stays a by-hand SQL step (DEPLOYMENT.md), deliberately: an endpoint that grants it is an
    // endpoint that can be abused to grant it. These flows hand out REVIEWER rights, which are
    // bounded to a system.
    //
    // *** REMOVAL TAKES EFFECT ON THE NEXT REQUEST. *** Nothing here has to do anything for that:
    // ReviewEndpoints.AuthoriseAsync reads the pool on every call. A test proves it by removing a
    // reviewer and asking again with the same token.
    // ###########################################################################################
    public static class ReviewerAssignmentFlows
    {
        public const int AccountListLimit = 500;

        // ###########################################################################################
        // Every system the administrator can assign reviewers to: the `systems` rows, unioned with
        // the boards in the data tree. The tree is what makes a shipped board that nobody has ever
        // submitted to assignable at all.
        // ###########################################################################################
        public static async Task<IReadOnlyList<SystemWithReviewers>> ListSystemsAsync(
            IReadOnlyList<PublishedSystemLister.KnownSystem> inTree,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inTree);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            IReadOnlyList<SystemRecord> rows = await submissions.ListSystemsAsync(cancellationToken);
            IReadOnlyList<ReviewerRecord> reviewers = await accounts.ListReviewersAsync(cancellationToken);

            ILookup<string, ReviewerRecord> bySystem =
                reviewers.ToLookup(reviewer => reviewer.SystemId, StringComparer.Ordinal);

            var merged = new Dictionary<string, SystemWithReviewers>(StringComparer.Ordinal);

            // The database row is the truth where it exists: it carries the revision and whether
            // the board is accepting.
            foreach (SystemRecord row in rows)
            {
                merged[row.SystemId] = new SystemWithReviewers(
                    row.SystemId, row.Manufacturer, row.Hardware, row.Board,
                    row.CurrentRevision, row.IsAccepting, bySystem[row.SystemId].ToList());
            }

            foreach (PublishedSystemLister.KnownSystem system in inTree)
            {
                if (merged.ContainsKey(system.SystemId))
                    continue;

                merged[system.SystemId] = new SystemWithReviewers(
                    system.SystemId, system.Manufacturer, system.Hardware, system.Board,
                    CurrentRevision: null, IsAccepting: true, bySystem[system.SystemId].ToList());
            }

            return merged.Values.OrderBy(system => system.SystemId, StringComparer.Ordinal).ToList();
        }

        // ###########################################################################################
        // Puts an account into a system's pool.
        //
        // The account must be VERIFIED and not LOCKED - ReviewAuthority refuses either whatever the
        // pool says, so granting to one would produce a reviewer who cannot review and a puzzled
        // administrator. An ADMINISTRATOR is refused too: they are in every pool already, and a
        // row for them would be a second source of the same truth.
        //
        // The system must be one that EXISTS - a `systems` row or a board in the tree. A pool row
        // for a system that is neither would grant authority over something that could only come
        // into being through a submission nobody has reviewed.
        // ###########################################################################################
        public static async Task<ReviewerAssignmentOutcome> AddAsync(
            ReviewAccess? actor,
            string? systemId,
            long accountId,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inTree,
            IAccountStore accounts,
            ISubmissionStore submissions,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inTree);
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(submissions);

            if (!ReviewAuthority.CanAdminister(actor))
                return ReviewerAssignmentOutcome.Forbidden();

            (string Manufacturer, string Hardware, string Board)? identity =
                await ReviewerAssignmentFlows.FindSystemAsync(systemId, inTree, submissions, cancellationToken);

            if (identity is null)
                return ReviewerAssignmentOutcome.NotFound("No such system.");

            AccountRecord? target = await accounts.FindByIdAsync(accountId, cancellationToken);

            string? refusal = ReviewerAssignmentRules.WhyNotGrantable(target);

            if (refusal is not null)
                return target is null ? ReviewerAssignmentOutcome.NotFound(refusal) : ReviewerAssignmentOutcome.Refused(refusal);

            // The pool table's foreign key needs the row; a shipped board nobody has submitted to
            // has none yet. 'shipped' because it was in the tree before any submission.
            await submissions.EnsureSystemAsync(
                systemId!, identity.Value.Manufacturer, identity.Value.Hardware, identity.Value.Board,
                SystemDescriptorRules.SystemOrigin.Shipped, now, cancellationToken);

            await accounts.AddReviewerAsync(systemId!, accountId, actor!.Account.Id, now, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    actor.Account.Id,
                    actor.Account.Email,
                    ReviewerAssignmentFlows.GrantedAction,
                    systemId,
                    $"account {accountId} ({target!.Email})",
                    now),
                cancellationToken);

            return ReviewerAssignmentOutcome.Done();
        }

        // ###########################################################################################
        // Takes an account out of a system's pool. Removing somebody who is not in it is not an
        // error - the state asked for is the state that exists.
        // ###########################################################################################
        public static async Task<ReviewerAssignmentOutcome> RemoveAsync(
            ReviewAccess? actor,
            string? systemId,
            long accountId,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanAdminister(actor))
                return ReviewerAssignmentOutcome.Forbidden();

            if (string.IsNullOrWhiteSpace(systemId))
                return ReviewerAssignmentOutcome.NotFound("No such system.");

            await accounts.RemoveReviewerAsync(systemId, accountId, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    actor!.Account.Id,
                    actor.Account.Email,
                    ReviewerAssignmentFlows.RevokedAction,
                    systemId,
                    $"account {accountId}",
                    now),
                cancellationToken);

            return ReviewerAssignmentOutcome.Done();
        }

        public const string GrantedAction = "reviewer.granted";
        public const string RevokedAction = "reviewer.revoked";

        // A system's three name parts, from its row or from the tree; null when it is neither.
        private static async Task<(string Manufacturer, string Hardware, string Board)?> FindSystemAsync(
            string? systemId,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inTree,
            ISubmissionStore submissions,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(systemId))
                return null;

            SystemRecord? row = (await submissions.ListSystemsAsync(cancellationToken))
                .FirstOrDefault(system => string.Equals(system.SystemId, systemId, StringComparison.Ordinal));

            if (row is not null)
                return (row.Manufacturer, row.Hardware, row.Board);

            PublishedSystemLister.KnownSystem? known = inTree
                .FirstOrDefault(system => string.Equals(system.SystemId, systemId, StringComparison.Ordinal));

            return known is null ? null : (known.Manufacturer, known.Hardware, known.Board);
        }
    }

    // ###########################################################################################
    // Why an account may not be made a reviewer, or null when it may. Pure, so the sentences the
    // administrator reads are tested.
    // ###########################################################################################
    public static class ReviewerAssignmentRules
    {
        public static string? WhyNotGrantable(AccountRecord? account)
        {
            if (account is null)
                return "No such account.";

            if (!account.IsVerified)
                return $"{account.Email} has not verified their address yet, so they cannot review anything.";

            if (account.IsLocked)
                return $"{account.Email} is locked. Unlock the account before making them a reviewer.";

            if (account.IsAdministrator)
                return $"{account.Email} is an administrator and already reviews every system.";

            return null;
        }
    }

    public sealed record SystemWithReviewers(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? CurrentRevision,
        bool IsAccepting,
        IReadOnlyList<ReviewerRecord> Reviewers);

    public sealed record ReviewerAssignmentOutcome(
        bool IsDone,
        string Error,
        bool IsForbidden = false,
        bool IsNotFound = false)
    {
        public static ReviewerAssignmentOutcome Done() => new(true, string.Empty);

        public static ReviewerAssignmentOutcome Forbidden() =>
            new(false, "Only an administrator can change who reviews a system.", IsForbidden: true);

        public static ReviewerAssignmentOutcome NotFound(string error) => new(false, error, IsNotFound: true);

        public static ReviewerAssignmentOutcome Refused(string error) => new(false, error);
    }
}
