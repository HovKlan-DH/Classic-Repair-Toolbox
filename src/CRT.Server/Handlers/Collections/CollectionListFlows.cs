using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // THE WEBSITE'S "YOUR COLLECTIONS" (step 4, 2026-11-26; owner decisions, 2026-11-25: "A signed-in user's web list: their
    // own collections and those they commented on"; nobody but the administrator sees a list of all collections). Pure, as
    // CollectionFlows is: stores handed in.
    //
    //   - Shared by you: the reader's collections not deleted, newest first.
    //   - You commented on: other people's with a comment of the reader's on them, the last commented first - only while
    //     shown, so a locked owner's, and one its owner stopped sharing, are left out (CollectionFlows.FindShownAsync's rule).
    //   - The administrator alone, besides: every collection not deleted (at most AllShownAtMost), each with its owner's name
    //     and address and whether the owner is locked - or has no account at all - the newest comments on them
    //     (NewestCommentsShownAtMost), hidden ones and a locked account's too - where moderating starts - and the accounts
    //     locked, to unlock (AccountLockFlows).
    // ###########################################################################################
    public static class CollectionListFlows
    {
        public const int AllShownAtMost = 500;

        public const int NewestCommentsShownAtMost = 50;

        public static async Task<CollectionListView> ListForAsync(
            AccountRecord reader,
            ICollectionStore store,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(reader);

            // A collection whose time is up is gone from that moment, as its page says (CollectionFlows.IsGone; code review,
            // 2026-11-27: listed until the hourly sweep, linking to a page saying it was deleted).
            List<CollectionRecord> Shown(IReadOnlyList<CollectionRecord> collections) =>
                collections.Where(collection => !CollectionFlows.IsGone(collection, now)).ToList();

            IReadOnlyList<CollectionRecord> own = Shown(await store.ListLiveForAccountAsync(reader.Id, cancellationToken));
            IReadOnlyList<CollectionRecord> commented = Shown(await store.ListLiveCommentedOnAsync(reader.Id, cancellationToken));
            IReadOnlyList<CollectionRecord> allLive = reader.IsAdministrator
                ? await store.ListAllLiveAsync(CollectionListFlows.AllShownAtMost, cancellationToken)
                : [];
            IReadOnlyList<CollectionRecord> all = Shown(allLive);

            // The newest comments leave out those on a collection whose time is up - a link to a page saying it was deleted,
            // where nothing can be moderated (code review, 2026-11-27).
            HashSet<string> gone = allLive.Where(collection => CollectionFlows.IsGone(collection, now)).Select(collection => collection.Id).ToHashSet(StringComparer.Ordinal);

            IReadOnlyDictionary<long, AccountRecord> owners = await accounts.FindByIdsAsync(
                commented.Concat(all).Select(collection => collection.AccountId).Distinct().ToList(),
                cancellationToken);

            // Nor one its owner stopped sharing - offline to everybody but its owner and the administrator (migration 0028).
            List<CollectionRecord> commentedShown = commented
                .Where(collection => !collection.IsSharingStopped &&
                                     owners.TryGetValue(collection.AccountId, out AccountRecord? owner) && !owner.IsLocked)
                .ToList();

            if (!reader.IsAdministrator)
                return new CollectionListView(own, commentedShown, null, null);

            List<CollectionListed> listed = all
                .Select(collection => owners.TryGetValue(collection.AccountId, out AccountRecord? owner)
                    ? new CollectionListed(collection, owner.DisplayName, owner.Email, owner.IsLocked, owner.IsAdministrator)
                    : new CollectionListed(collection, CollectionListFlows.NoOwner, null, OwnerLocked: false, OwnerGone: true))
                .ToList();

            IReadOnlyList<AccountRecord> locked = await AccountLockFlows.ListLockedAsync(reader, accounts, cancellationToken);

            return new CollectionListView(
                own,
                commentedShown,
                listed,
                (await store.ListNewestCommentsAsync(CollectionListFlows.NewestCommentsShownAtMost, cancellationToken))
                    .Where(comment => !gone.Contains(comment.CollectionId))
                    .ToList(),
                locked.Select(account => new LockedAccount(account.Id, account.DisplayName, account.Email)).ToList());
        }

        // An owner whose account is gone - its collections are shown to nobody but the administrator.
        public const string NoOwner = "(no account)";
    }

    // ###########################################################################################
    // What "Your collections" shows a reader: All, NewestComments and Locked only to the administrator (null for anybody else).
    // ###########################################################################################
    public sealed record CollectionListView(
        IReadOnlyList<CollectionRecord> Own,
        IReadOnlyList<CollectionRecord> CommentedOn,
        IReadOnlyList<CollectionListed>? All,
        IReadOnlyList<CollectionCommentRecord>? NewestComments,
        IReadOnlyList<LockedAccount>? Locked = null);

    // A collection in the administrator's list, with whose it is - a locked owner's, and one whose owner has no account
    // (`OwnerGone`), is offline for everybody else; an administrator's account is never offered to be locked (AccountLockFlows).
    public sealed record CollectionListed(
        CollectionRecord Collection,
        string OwnerName,
        string? OwnerEmail,
        bool OwnerLocked,
        bool OwnerIsAdministrator = false,
        bool OwnerGone = false)
    {
        public bool IsOffline => this.OwnerLocked || this.OwnerGone;
    }

    // An account locked, as the administrator's list names it - to unlock, or to ask about before locking.
    public sealed record LockedAccount(long Id, string Name, string Email);
}
