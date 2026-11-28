using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // THE WEBSITE'S LIST OF COLLECTIONS - THE ADMINISTRATOR'S ALONE (owner decisions, 2026-11-25: nobody but the administrator
    // sees a list of all collections; 2026-11-28, as the accounts for everybody went: "I (as web admin) still want to be able to
    // see all collections, so I can control it looks correct, but do remove comment possibilities etc."). Every collection not
    // deleted, newest first (at most AllShownAtMost), one whose time is up left out - it is gone from that moment, as its page
    // says. Nobody else signs in on the website, so nobody else gets a list. Pure, as CollectionFlows is: the store handed in.
    // ###########################################################################################
    public static class CollectionListFlows
    {
        public const int AllShownAtMost = 500;

        // Null for anybody but the administrator.
        public static async Task<IReadOnlyList<CollectionRecord>?> ListForAsync(
            AccountRecord reader,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(reader);

            if (!reader.IsAdministrator)
                return null;

            return (await store.ListAllLiveAsync(CollectionListFlows.AllShownAtMost, cancellationToken))
                .Where(collection => !CollectionFlows.IsGone(collection, now))
                .ToList();
        }
    }
}
