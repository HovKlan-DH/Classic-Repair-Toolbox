using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // Where shared oscilloscope collections are kept (migration 0025, CollectionFlows). MySqlCollectionStore is the database;
    // the tests use a fake. The pictures' bytes are CollectionFiles', on disk.
    // ###########################################################################################
    public interface ICollectionStore
    {
        // A collection with its captures - deleted or not - or null when there is none of that id.
        Task<CollectionRecord?> FindAsync(string id, CancellationToken cancellationToken = default);

        // The account's collections not deleted, newest first - without their captures (ImageCount says how many), with how
        // many comments their pages show and the newest one's id (CommentCount, LatestCommentId).
        Task<IReadOnlyList<CollectionRecord>> ListLiveForAccountAsync(long accountId, CancellationToken cancellationToken = default);

        // How many collections the account made since `sinceUtc`, deleted ones too - the day's limit.
        Task<int> CountCreatedSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);

        // Stores a new collection and its captures. False when the id is taken (CollectionFlows draws another).
        Task<bool> AddAsync(CollectionRecord collection, CancellationToken cancellationToken = default);

        // Replaces a collection's words, date and captures - only while it is not deleted; false when it is (deleted
        // meanwhile). A date moved later than it was asks for the reminder again (reminded_utc cleared).
        Task<bool> ReplaceContentAsync(CollectionRecord collection, CancellationToken cancellationToken = default);

        // Deletes a collection: its captures' rows go, its words are emptied, and the row stays so the page can say when it
        // went. False when it was deleted already.
        Task<bool> MarkDeletedAsync(string id, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // The ids of the collections whose time is up at `nowUtc` and that are not deleted yet.
        Task<IReadOnlyList<string>> ListExpiredAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default);

        // The collections not deleted that go by `byUtc` and whose owner has not been reminded yet - without their captures.
        Task<IReadOnlyList<CollectionRecord>> ListToRemindAsync(DateTimeOffset byUtc, CancellationToken cancellationToken = default);

        Task MarkRemindedAsync(string id, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // ---------------------------------------------------------------------------------------
        // Uploads: which account sent which picture.
        // ---------------------------------------------------------------------------------------

        // The account's uploads among `sha256s`, by checksum.
        Task<IReadOnlyDictionary<string, CollectionUploadRecord>> FindUploadsAsync(long accountId, IReadOnlyCollection<string> sha256s, CancellationToken cancellationToken = default);

        // Records an upload - none when the account sent the same picture before (its first time is kept). True when recorded.
        Task<bool> AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default);

        // How many pictures the account uploaded since `sinceUtc` - the day's limit.
        Task<int> CountUploadsSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);

        // The bytes the pictures take all together, each picture once (CollectionsMaxStoredBytes).
        Task<long> StoredBytesAsync(CancellationToken cancellationToken = default);

        // Deletes every upload made before `uploadedBeforeUtc` that no collection of its account shows any more (deleted ones
        // show none), and answers the checksums no upload is left of - the pictures whose bytes can go.
        Task<IReadOnlyList<string>> DeleteUnusedUploadsAsync(DateTimeOffset uploadedBeforeUtc, CancellationToken cancellationToken = default);

        // ---------------------------------------------------------------------------------------
        // Comments (step 3, migration 0026). A collection deleted takes its comments (MarkDeletedAsync).
        // ---------------------------------------------------------------------------------------

        // The comments not deleted on a collection, oldest first, each with its writer's name - hidden ones too.
        Task<IReadOnlyList<CollectionCommentRecord>> ListCommentsAsync(string collectionId, CancellationToken cancellationToken = default);

        Task<CollectionCommentRecord?> FindCommentAsync(long commentId, CancellationToken cancellationToken = default);

        // Its id.
        Task<long> AddCommentAsync(string collectionId, string imageSha256, long accountId, string body, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        Task<bool> EditCommentAsync(long commentId, string body, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        Task<bool> DeleteCommentAsync(long commentId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        Task<bool> SetCommentHiddenAsync(long commentId, bool hidden, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // How many comments the account wrote since `sinceUtc` - the day's limit.
        Task<int> CountCommentsSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // A comment as kept, with its writer's name and whether the writer is an administrator (joined from accounts).
    // ###########################################################################################
    public sealed record CollectionCommentRecord(
        long Id,
        string CollectionId,
        string ImageSha256,
        long AccountId,
        string AuthorName,
        bool AuthorIsAdministrator,
        string Body,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? EditedUtc,
        DateTimeOffset? HiddenUtc)
    {
        public bool IsHidden => this.HiddenUtc is not null;
    }

    public static class CollectionDeletedReasons
    {
        public const string Expired = "expired";

        public const string Owner = "owner";

        public const string Administrator = "administrator";
    }

    // ###########################################################################################
    // A collection as kept. The images are the contract's own entries with the picture's size (from its upload).
    // ###########################################################################################
    public sealed record CollectionRecord(
        string Id,
        long AccountId,
        string HardwareName,
        string BoardName,
        string? Region,
        string? Title,
        string? Description,
        IReadOnlyList<string> ForumLinks,
        DateTimeOffset CreatedUtc,
        DateTimeOffset UpdatedUtc,
        DateTimeOffset DeleteAtUtc,
        DateTimeOffset? RemindedUtc,
        DateTimeOffset? DeletedUtc,
        string? DeletedReason,
        IReadOnlyList<CollectionImageRecord> Images,
        int ImageCount,
        int CommentCount = 0,
        long LatestCommentId = 0)
    {
        public bool IsDeleted => this.DeletedUtc is not null;
    }

    public sealed record CollectionImageRecord(CollectionImageEntry Entry, int Width, int Height);

    public sealed record CollectionUploadRecord(
        long AccountId,
        string Sha256,
        int ByteCount,
        int Width,
        int Height,
        DateTimeOffset UploadedUtc);
}
