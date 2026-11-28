using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // Where shared oscilloscope collections are kept (migrations 0025-0031, CollectionFlows). MySqlCollectionStore is the
    // database; the tests use a fake. The pictures' bytes are CollectionFiles', on disk.
    //
    // Since server 17.0.0 a collection is owned by the HASH OF ITS KEY (owner decision, 2026-11-28: no accounts for everybody;
    // migration 0031) and counted by the internet connection it came from - see CollectionRules.IsOwnerKey.
    // ###########################################################################################
    public interface ICollectionStore
    {
        // A collection with its captures - deleted or not - or null when there is none of that id.
        Task<CollectionRecord?> FindAsync(string id, CancellationToken cancellationToken = default);

        // The same, its Images only the rows of the capture `sha256` names (none when it shows no such capture) - what a
        // picture's every request asks (code review, 2026-11-27: each read every capture of the collection, 300 rows a picture
        // on a page of 300).
        Task<CollectionRecord?> FindWithImageAsync(string id, string sha256, CancellationToken cancellationToken = default);

        // How many collections were first shared from `connection` since `sinceUtc`, deleted ones too - the day's limit.
        Task<int> CountCreatedSinceAsync(string connection, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);

        // Stores a new collection and its captures. False when the id is taken (CollectionFlows draws another).
        Task<bool> AddAsync(CollectionRecord collection, CancellationToken cancellationToken = default);

        // The collection not deleted that the key `ownerKeyHash` owns - a first share sent again, its answer lost the first
        // time. Null when there is none.
        Task<CollectionRecord?> FindLiveByOwnerAsync(string ownerKeyHash, CancellationToken cancellationToken = default);

        // Replaces a collection's words, date and captures - only while it is not deleted; false when it is (deleted
        // meanwhile). A collection whose sharing was stopped is shared again by it (unshared_utc cleared) - "Share" brings the
        // page back under its link.
        Task<bool> ReplaceContentAsync(CollectionRecord collection, CancellationToken cancellationToken = default);

        // Stops sharing a collection (`unsharedUtc` set - its page offline) or shares it again (null), its captures kept either
        // way (migration 0028). False when it is deleted.
        Task<bool> SetUnsharedAsync(string id, DateTimeOffset? unsharedUtc, CancellationToken cancellationToken = default);

        // Takes one capture off a collection (server 16.1.0) - its deletion date, its sharing, its words and its other captures
        // left as they are; updated_utc set to `whenUtc`. True also when the capture was not on it; false when it is deleted.
        Task<bool> RemoveImageAsync(string id, string sha256, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // Deletes a collection: its captures' rows go, its words are emptied, and the row stays so the page can say when it
        // went. False when it was deleted already.
        Task<bool> MarkDeletedAsync(string id, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // The ids of the collections whose time is up at `nowUtc` and that are not deleted yet.
        Task<IReadOnlyList<string>> ListExpiredAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default);

        // Every collection not deleted, newest first, at most `count` - without their captures (ImageCount says how many) - the
        // administrator's list on the website.
        Task<IReadOnlyList<CollectionRecord>> ListAllLiveAsync(int count, CancellationToken cancellationToken = default);

        // ---------------------------------------------------------------------------------------
        // Uploads: which key sent which picture, from which connection.
        // ---------------------------------------------------------------------------------------

        // The key's uploads among `sha256s`, by checksum.
        Task<IReadOnlyDictionary<string, CollectionUploadRecord>> FindUploadsAsync(string ownerKeyHash, IReadOnlyCollection<string> sha256s, CancellationToken cancellationToken = default);

        // Records an upload - or, for a picture the key sent before, moves that upload's time to this one's (never back), so the
        // sweep's day counts from the last send.
        Task AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default);

        // How many pictures were uploaded from `connection` since `sinceUtc` - the day's limit.
        Task<int> CountUploadsSinceAsync(string connection, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);

        // The bytes the pictures take all together, each picture once (CollectionsMaxStoredBytes).
        Task<long> StoredBytesAsync(CancellationToken cancellationToken = default);

        // Deletes every upload made before `uploadedBeforeUtc` that no collection of its owner shows any more (deleted ones show
        // none) - its key's, or for an upload from before server 17.0.0 its account's - and answers the checksums no upload is
        // left of: the pictures whose bytes can go.
        Task<IReadOnlyList<string>> DeleteUnusedUploadsAsync(DateTimeOffset uploadedBeforeUtc, CancellationToken cancellationToken = default);

        // Whether any upload of the picture is kept now - asked again right before its bytes are deleted, as another key may
        // have sent the same picture since DeleteUnusedUploadsAsync answered.
        Task<bool> IsUploadedAsync(string sha256, CancellationToken cancellationToken = default);
    }

    public static class CollectionDeletedReasons
    {
        public const string Expired = "expired";

        public const string Owner = "owner";

        public const string Administrator = "administrator";
    }

    // ###########################################################################################
    // A collection as kept. The images are the contract's own entries with the picture's size (from its upload).
    // `OwnerKeyHash`: the hash of the key that owns it (null for one shared with an account before server 17.0.0, which nobody
    // owns any more). `Connection`: the internet connection it was first shared from (the day's limit).
    // `UnsharedUtc`: when its owner stopped sharing it - its page offline until "Share" brings it back (migration 0028).
    // `Components`: what CRT's component info window said of each component captured (migration 0029).
    // `Oscilloscopes`: the oscilloscopes the captures were taken with, each once (migration 0030).
    // ###########################################################################################
    public sealed record CollectionRecord(
        string Id,
        string? OwnerKeyHash,
        string HardwareName,
        string BoardName,
        string? Region,
        string? Title,
        string? Description,
        IReadOnlyList<string> ForumLinks,
        DateTimeOffset CreatedUtc,
        DateTimeOffset UpdatedUtc,
        DateTimeOffset DeleteAtUtc,
        DateTimeOffset? DeletedUtc,
        string? DeletedReason,
        IReadOnlyList<CollectionImageRecord> Images,
        int ImageCount,
        string? Connection = null,
        DateTimeOffset? UnsharedUtc = null,
        IReadOnlyList<CollectionComponent>? Components = null,
        IReadOnlyList<string>? Oscilloscopes = null)
    {
        public bool IsDeleted => this.DeletedUtc is not null;

        public bool IsSharingStopped => this.UnsharedUtc is not null;

        // The facts of the component a capture is of, or null.
        public CollectionComponent? ComponentOf(CollectionImageEntry image) =>
            this.Components?.FirstOrDefault(component => string.Equals(component.Label, image.ComponentLabel, StringComparison.OrdinalIgnoreCase));
    }

    public sealed record CollectionImageRecord(CollectionImageEntry Entry, int Width, int Height);

    public sealed record CollectionUploadRecord(
        string OwnerKeyHash,
        string Sha256,
        int ByteCount,
        int Width,
        int Height,
        DateTimeOffset UploadedUtc,
        string? Connection = null);
}
