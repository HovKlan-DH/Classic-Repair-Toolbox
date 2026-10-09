using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The administrator's side of Phase 6 roles (2026-09-25): who reviews which board.
    //
    // Every rule is here, pure, taking its stores as arguments - the same shape SubmissionFlows
    // and AccountFlows have, for the same reason: AdminEndpoints is a rim, and a rule that lives in
    // a handler body is a rule no test can reach.
    //
    // *** ADMINISTRATOR ONLY, AND THERE IS NO WAY TO BECOME ONE HERE. *** Granting administrator
    // stays a by-hand SQL step (INSTALLING.md), deliberately: an endpoint that grants it is an
    // endpoint that can be abused to grant it. These flows hand out MAINTAINER rights, which are
    // bounded to a board.
    //
    // *** REMOVAL TAKES EFFECT ON THE NEXT REQUEST. *** Nothing here has to do anything for that:
    // ReviewEndpoints.AuthoriseAsync reads the pool on every call. A test proves it by removing a
    // maintainer and asking again with the same token.
    // ###########################################################################################
    public static class MaintainerAssignmentFlows
    {
        public const int AccountListLimit = 500;

        // ###########################################################################################
        // Every board the administrator can assign maintainers to: the `boards` rows, unioned with
        // the boards in the data tree. The tree is what makes a shipped board that nobody has ever
        // submitted to assignable at all.
        // ###########################################################################################
        public static async Task<IReadOnlyList<BoardWithMaintainers>> ListBoardsAsync(
            IReadOnlyList<PublishedBoardLister.KnownBoard> inTree,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inTree);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            IReadOnlyList<BoardRecord> rows = await submissions.ListBoardsAsync(cancellationToken);
            IReadOnlyList<MaintainerRecord> maintainers = await accounts.ListMaintainersAsync(cancellationToken);

            ILookup<string, MaintainerRecord> byBoard =
                maintainers.ToLookup(maintainer => maintainer.BoardId, StringComparer.Ordinal);

            var merged = new Dictionary<string, BoardWithMaintainers>(StringComparer.Ordinal);

            // The database row is the truth where it exists: it carries the revision and whether
            // the board is accepting.
            foreach (BoardRecord row in rows)
            {
                merged[row.BoardId] = new BoardWithMaintainers(
                    row.BoardId, row.Manufacturer, row.Hardware, row.Board,
                    row.CurrentRevision, row.IsAccepting, byBoard[row.BoardId].ToList());
            }

            foreach (PublishedBoardLister.KnownBoard board in inTree)
            {
                if (merged.ContainsKey(board.BoardId))
                    continue;

                merged[board.BoardId] = new BoardWithMaintainers(
                    board.BoardId, board.Manufacturer, board.Hardware, board.Board,
                    CurrentRevision: null, IsAccepting: true, byBoard[board.BoardId].ToList());
            }

            return merged.Values.OrderBy(board => board.BoardId, StringComparer.Ordinal).ToList();
        }

        // ###########################################################################################
        // Puts an account into a board's pool.
        //
        // The account must be VERIFIED and not LOCKED - ReviewAuthority refuses either whatever the
        // pool says, so granting to one would produce a maintainer who cannot review and a puzzled
        // administrator.
        //
        // *** AN ADMINISTRATOR MAY BE PUT IN A POOL (owner request, 2026-10-05: "I would like to be
        // able to set myself (admin) as a maintainer, so others can see that this is me maintaining
        // these systems"). *** It was refused until then as "a second source of the same truth". The
        // row grants nothing more - an administrator reviews and publishes every board anyway, and
        // approves as the administrator (ReviewAuthority.RoleIn) - so what it adds is being NAMED as
        // the board's maintainer, and that board's "submission waiting" mail (SubmissionRouting,
        // one mail however many roles). It never counts as the maintainer half of a two-person
        // approval (ReviewAuthority.CanGiveMaintainerApproval), or a shared-file change on a board
        // they alone maintain would wait for a second approval nobody can give.
        //
        // The board must be one that EXISTS - a `boards` row or a board in the tree. A pool row
        // for a board that is neither would grant authority over something that could only come
        // into being through a submission nobody has reviewed.
        // ###########################################################################################
        public static async Task<MaintainerAssignmentOutcome> AddAsync(
            ReviewAccess? actor,
            string? boardId,
            long accountId,
            IReadOnlyList<PublishedBoardLister.KnownBoard> inTree,
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
                await MaintainerAssignmentFlows.FindBoardAsync(boardId, inTree, submissions, cancellationToken);

            if (identity is null)
                return MaintainerAssignmentOutcome.NotFound("No such board.");

            AccountRecord? target = await accounts.FindByIdAsync(accountId, cancellationToken);

            string? refusal = MaintainerAssignmentRules.WhyNotGrantable(target);

            if (refusal is not null)
                return target is null ? MaintainerAssignmentOutcome.NotFound(refusal) : MaintainerAssignmentOutcome.Refused(refusal);

            // The pool table's foreign key needs the row; a shipped board nobody has submitted to
            // has none yet. 'shipped' because it was in the tree before any submission.
            await submissions.EnsureBoardAsync(
                boardId!, identity.Value.Manufacturer, identity.Value.Hardware, identity.Value.Board,
                BoardDescriptorRules.BoardOrigin.Shipped, now, cancellationToken);

            await accounts.AddMaintainerAsync(boardId!, accountId, actor!.Account.Id, now, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    actor.Account.Id,
                    actor.Account.Email,
                    MaintainerAssignmentFlows.GrantedAction,
                    boardId,
                    $"account {accountId} ({target!.Email})",
                    now),
                cancellationToken);

            return MaintainerAssignmentOutcome.Done();
        }

        // ###########################################################################################
        // Takes an account out of a board's pool. Removing somebody who is not in it is not an
        // error - the state asked for is the state that exists.
        // ###########################################################################################
        public static async Task<MaintainerAssignmentOutcome> RemoveAsync(
            ReviewAccess? actor,
            string? boardId,
            long accountId,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanAdminister(actor))
                return MaintainerAssignmentOutcome.Forbidden();

            if (string.IsNullOrWhiteSpace(boardId))
                return MaintainerAssignmentOutcome.NotFound("No such board.");

            // The address, like a grant's, so the board's history can say who was removed.
            AccountRecord? removed = await accounts.FindByIdAsync(accountId, cancellationToken);

            await accounts.RemoveMaintainerAsync(boardId, accountId, cancellationToken);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    actor!.Account.Id,
                    actor.Account.Email,
                    MaintainerAssignmentFlows.RevokedAction,
                    boardId,
                    removed is null ? $"account {accountId}" : $"account {accountId} ({removed.Email})",
                    now),
                cancellationToken);

            return MaintainerAssignmentOutcome.Done();
        }

        public const string GrantedAction = "maintainer.granted";
        public const string RevokedAction = "maintainer.revoked";

        // A board's three name parts, from its row or from the tree; null when it is neither.
        internal static async Task<(string Manufacturer, string Hardware, string Board)?> FindBoardAsync(
            string? boardId,
            IReadOnlyList<PublishedBoardLister.KnownBoard> inTree,
            ISubmissionStore submissions,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(boardId))
                return null;

            BoardRecord? row = (await submissions.ListBoardsAsync(cancellationToken))
                .FirstOrDefault(board => string.Equals(board.BoardId, boardId, StringComparison.Ordinal));

            if (row is not null)
                return (row.Manufacturer, row.Hardware, row.Board);

            PublishedBoardLister.KnownBoard? known = inTree
                .FirstOrDefault(board => string.Equals(board.BoardId, boardId, StringComparison.Ordinal));

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

            // An administrator may be granted (2026-10-05) - see AddAsync.
            return null;
        }
    }

    public sealed record BoardWithMaintainers(
        string BoardId,
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
            new(false, "Only an administrator can change who reviews a board.", IsForbidden: true);

        public static MaintainerAssignmentOutcome NotFound(string error) => new(false, error, IsNotFound: true);

        public static MaintainerAssignmentOutcome Refused(string error) => new(false, error);
    }
}
