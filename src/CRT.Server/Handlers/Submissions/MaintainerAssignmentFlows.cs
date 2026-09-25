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
    // endpoint that can be abused to grant it. These flows hand out MAINTAINER rights, which are
    // bounded to a system.
    //
    // *** REMOVAL TAKES EFFECT ON THE NEXT REQUEST. *** Nothing here has to do anything for that:
    // ReviewEndpoints.AuthoriseAsync reads the pool on every call. A test proves it by removing a
    // maintainer and asking again with the same token.
    // ###########################################################################################
    public static class MaintainerAssignmentFlows
    {
        public const int AccountListLimit = 500;

        // ###########################################################################################
        // Every system the administrator can assign maintainers to: the `systems` rows, unioned with
        // the boards in the data tree. The tree is what makes a shipped board that nobody has ever
        // submitted to assignable at all.
        // ###########################################################################################
        public static async Task<IReadOnlyList<SystemWithMaintainers>> ListSystemsAsync(
            IReadOnlyList<PublishedSystemLister.KnownSystem> inTree,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inTree);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            IReadOnlyList<SystemRecord> rows = await submissions.ListSystemsAsync(cancellationToken);
            IReadOnlyList<MaintainerRecord> maintainers = await accounts.ListMaintainersAsync(cancellationToken);

            ILookup<string, MaintainerRecord> bySystem =
                maintainers.ToLookup(maintainer => maintainer.SystemId, StringComparer.Ordinal);

            var merged = new Dictionary<string, SystemWithMaintainers>(StringComparer.Ordinal);

            // The database row is the truth where it exists: it carries the revision and whether
            // the board is accepting.
            foreach (SystemRecord row in rows)
            {
                merged[row.SystemId] = new SystemWithMaintainers(
                    row.SystemId, row.Manufacturer, row.Hardware, row.Board,
                    row.CurrentRevision, row.IsAccepting, bySystem[row.SystemId].ToList());
            }

            foreach (PublishedSystemLister.KnownSystem system in inTree)
            {
                if (merged.ContainsKey(system.SystemId))
                    continue;

                merged[system.SystemId] = new SystemWithMaintainers(
                    system.SystemId, system.Manufacturer, system.Hardware, system.Board,
                    CurrentRevision: null, IsAccepting: true, bySystem[system.SystemId].ToList());
            }

            return merged.Values.OrderBy(system => system.SystemId, StringComparer.Ordinal).ToList();
        }

        // ###########################################################################################
        // Puts an account into a system's pool.
        //
        // The account must be VERIFIED and not LOCKED - ReviewAuthority refuses either whatever the
        // pool says, so granting to one would produce a maintainer who cannot review and a puzzled
        // administrator. An ADMINISTRATOR is refused too: they are in every pool already, and a
        // row for them would be a second source of the same truth.
        //
        // The system must be one that EXISTS - a `systems` row or a board in the tree. A pool row
        // for a system that is neither would grant authority over something that could only come
        // into being through a submission nobody has reviewed.
        // ###########################################################################################
        public static async Task<MaintainerAssignmentOutcome> AddAsync(
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
                return MaintainerAssignmentOutcome.Forbidden();

            (string Manufacturer, string Hardware, string Board)? identity =
                await MaintainerAssignmentFlows.FindSystemAsync(systemId, inTree, submissions, cancellationToken);

            if (identity is null)
                return MaintainerAssignmentOutcome.NotFound("No such system.");

            AccountRecord? target = await accounts.FindByIdAsync(accountId, cancellationToken);

            string? refusal = MaintainerAssignmentRules.WhyNotGrantable(target);

            if (refusal is not null)
                return target is null ? MaintainerAssignmentOutcome.NotFound(refusal) : MaintainerAssignmentOutcome.Refused(refusal);

            // The pool table's foreign key needs the row; a shipped board nobody has submitted to
            // has none yet. 'shipped' because it was in the tree before any submission.
            await submissions.EnsureSystemAsync(
                systemId!, identity.Value.Manufacturer, identity.Value.Hardware, identity.Value.Board,
                SystemDescriptorRules.SystemOrigin.Shipped, now, cancellationToken);

            await accounts.AddMaintainerAsync(systemId!, accountId, actor!.Account.Id, now, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    actor.Account.Id,
                    actor.Account.Email,
                    MaintainerAssignmentFlows.GrantedAction,
                    systemId,
                    $"account {accountId} ({target!.Email})",
                    now),
                cancellationToken);

            return MaintainerAssignmentOutcome.Done();
        }

        // ###########################################################################################
        // Takes an account out of a system's pool. Removing somebody who is not in it is not an
        // error - the state asked for is the state that exists.
        // ###########################################################################################
        public static async Task<MaintainerAssignmentOutcome> RemoveAsync(
            ReviewAccess? actor,
            string? systemId,
            long accountId,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanAdminister(actor))
                return MaintainerAssignmentOutcome.Forbidden();

            if (string.IsNullOrWhiteSpace(systemId))
                return MaintainerAssignmentOutcome.NotFound("No such system.");

            await accounts.RemoveMaintainerAsync(systemId, accountId, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    actor!.Account.Id,
                    actor.Account.Email,
                    MaintainerAssignmentFlows.RevokedAction,
                    systemId,
                    $"account {accountId}",
                    now),
                cancellationToken);

            return MaintainerAssignmentOutcome.Done();
        }

        public const string GrantedAction = "maintainer.granted";
        public const string RevokedAction = "maintainer.revoked";

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
    // Why an account may not be made a maintainer, or null when it may. Pure, so the sentences the
    // administrator reads are tested.
    // ###########################################################################################
    public static class MaintainerAssignmentRules
    {
        public static string? WhyNotGrantable(AccountRecord? account)
        {
            if (account is null)
                return "No such account.";

            if (!account.IsVerified)
                return $"{account.Email} has not verified their address yet, so they cannot review anything.";

            if (account.IsLocked)
                return $"{account.Email} is locked. Unlock the account before making them a maintainer.";

            if (account.IsAdministrator)
                return $"{account.Email} is an administrator and already reviews every system.";

            return null;
        }
    }

    public sealed record SystemWithMaintainers(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? CurrentRevision,
        bool IsAccepting,
        IReadOnlyList<MaintainerRecord> Maintainers);

    public sealed record MaintainerAssignmentOutcome(
        bool IsDone,
        string Error,
        bool IsForbidden = false,
        bool IsNotFound = false)
    {
        public static MaintainerAssignmentOutcome Done() => new(true, string.Empty);

        public static MaintainerAssignmentOutcome Forbidden() =>
            new(false, "Only an administrator can change who reviews a system.", IsForbidden: true);

        public static MaintainerAssignmentOutcome NotFound(string error) => new(false, error, IsNotFound: true);

        public static MaintainerAssignmentOutcome Refused(string error) => new(false, error);
    }
}
