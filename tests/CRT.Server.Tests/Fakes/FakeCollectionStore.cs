using CRT.Server.Handlers.Collections;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // ICollectionStore in memory - the rules MySqlCollectionStore keeps in SQL, kept here in plain code: an id taken is no
    // add, a deleted collection takes no change, and an upload no collection of its key shows is deleted a day after it came.
    // ###########################################################################################
    public sealed class FakeCollectionStore : ICollectionStore
    {
        public Dictionary<string, CollectionRecord> Collections { get; } = new(StringComparer.Ordinal);

        public List<CollectionUploadRecord> Uploads { get; } = [];

        public Task<CollectionRecord?> FindAsync(string id, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.TryGetValue(id, out CollectionRecord? collection) ? collection : null);

        public Task<CollectionRecord?> FindWithImageAsync(string id, string sha256, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(id, out CollectionRecord? collection))
                return Task.FromResult<CollectionRecord?>(null);

            // As the database's: a capture, or a capture's pinout drawing (migration 0029).
            List<CollectionImageRecord> images = collection.Images.Where(image => image.Entry.Sha256 == sha256 || image.Entry.PinoutSha256 == sha256).ToList();

            return Task.FromResult<CollectionRecord?>(collection with { Images = images, ImageCount = images.Count });
        }

        public Task<IReadOnlyList<CollectionRecord>> ListAllLiveAsync(int count, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionRecord>>(this.Collections.Values
                .Where(collection => !collection.IsDeleted)
                .OrderByDescending(collection => collection.CreatedUtc)
                .Take(count)
                .Select(collection => collection with { Images = [], ImageCount = collection.Images.Count })
                .ToList());

        // Run as a share counts the day's collections - another request coming at that very moment.
        public Action? WhileCountingCreated { get; set; }

        public Task<int> CountCreatedSinceAsync(string connection, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            Action? meanwhile = this.WhileCountingCreated;

            this.WhileCountingCreated = null;
            meanwhile?.Invoke();

            return Task.FromResult(this.Collections.Values.Count(collection => collection.Connection == connection && collection.CreatedUtc > sinceUtc));
        }

        public Task<bool> AddAsync(CollectionRecord collection, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.TryAdd(collection.Id, collection));

        public Task<CollectionRecord?> FindLiveByOwnerAsync(string ownerKeyHash, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.Values
                .Where(collection => collection.OwnerKeyHash == ownerKeyHash && !collection.IsDeleted)
                .OrderByDescending(collection => collection.CreatedUtc)
                .FirstOrDefault());

        public Task<bool> ReplaceContentAsync(CollectionRecord collection, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(collection.Id, out CollectionRecord? existing) || existing.IsDeleted)
                return Task.FromResult(false);

            // As the database's UPDATE: its owner, where and when it was first shared stay, and a share brings one whose sharing
            // was stopped back (migration 0028).
            this.Collections[collection.Id] = collection with
            {
                OwnerKeyHash = existing.OwnerKeyHash,
                Connection = existing.Connection,
                CreatedUtc = existing.CreatedUtc,
                UnsharedUtc = null
            };

            return Task.FromResult(true);
        }

        // As the database's: stopped again keeps the first stop's time.
        public Task<bool> SetUnsharedAsync(string id, DateTimeOffset? unsharedUtc, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(id, out CollectionRecord? existing) || existing.IsDeleted)
                return Task.FromResult(false);

            this.Collections[id] = existing with { UnsharedUtc = unsharedUtc is null ? null : existing.UnsharedUtc ?? unsharedUtc };

            return Task.FromResult(true);
        }

        public Task<bool> RemoveImageAsync(string id, string sha256, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(id, out CollectionRecord? existing) || existing.IsDeleted)
                return Task.FromResult(false);

            this.Collections[id] = existing with
            {
                UpdatedUtc = whenUtc,
                Images = existing.Images.Where(image => !string.Equals(image.Entry.Sha256, sha256, StringComparison.Ordinal)).ToList()
            };

            return Task.FromResult(true);
        }

        public Task<bool> MarkDeletedAsync(string id, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(id, out CollectionRecord? existing) || existing.IsDeleted)
                return Task.FromResult(false);

            this.Collections[id] = existing with
            {
                DeletedUtc = whenUtc,
                DeletedReason = reason,
                Title = null,
                Description = null,
                ForumLinks = [],
                Components = null,
                Oscilloscopes = null,
                Images = [],
                ImageCount = 0
            };

            return Task.FromResult(true);
        }

        public Task<IReadOnlyList<string>> ListExpiredAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<string>>(this.Collections.Values
                .Where(collection => !collection.IsDeleted && collection.DeleteAtUtc <= nowUtc)
                .Select(collection => collection.Id)
                .ToList());

        public Task<IReadOnlyDictionary<string, CollectionUploadRecord>> FindUploadsAsync(
            string ownerKeyHash,
            IReadOnlyCollection<string> sha256s,
            CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyDictionary<string, CollectionUploadRecord>>(this.Uploads
                .Where(upload => upload.OwnerKeyHash == ownerKeyHash && sha256s.Contains(upload.Sha256))
                .ToDictionary(upload => upload.Sha256, StringComparer.Ordinal));

        public Task AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default)
        {
            int earlier = this.Uploads.FindIndex(row => row.OwnerKeyHash == upload.OwnerKeyHash && row.Sha256 == upload.Sha256);

            if (earlier < 0)
                this.Uploads.Add(upload);
            else if (upload.UploadedUtc > this.Uploads[earlier].UploadedUtc)
                this.Uploads[earlier] = this.Uploads[earlier] with { UploadedUtc = upload.UploadedUtc };

            return Task.CompletedTask;
        }

        public Task<int> CountUploadsSinceAsync(string connection, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Uploads.Count(upload => upload.Connection == connection && upload.UploadedUtc > sinceUtc));

        public Task<long> StoredBytesAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Uploads.GroupBy(upload => upload.Sha256).Sum(group => (long)group.Max(upload => upload.ByteCount)));

        public Task<IReadOnlyList<string>> DeleteUnusedUploadsAsync(DateTimeOffset uploadedBeforeUtc, CancellationToken cancellationToken = default)
        {
            List<CollectionUploadRecord> unused = this.Uploads
                .Where(upload => upload.UploadedUtc < uploadedBeforeUtc && !this.Collections.Values.Any(collection =>
                    collection.OwnerKeyHash == upload.OwnerKeyHash &&
                    !collection.IsDeleted &&
                    collection.Images.Any(image => image.Entry.Sha256 == upload.Sha256 || image.Entry.PinoutSha256 == upload.Sha256)))
                .ToList();

            foreach (CollectionUploadRecord upload in unused)
                this.Uploads.Remove(upload);

            IReadOnlyList<string> gone = unused
                .Select(upload => upload.Sha256)
                .Distinct(StringComparer.Ordinal)
                .Where(sha => !this.Uploads.Any(upload => upload.Sha256 == sha))
                .ToList();

            this.AfterDeleteUnusedUploads?.Invoke();

            return Task.FromResult(gone);
        }

        // Run once the unused uploads are gone, before the sweep deletes their bytes - a test sends the same picture again there,
        // as another key may at that very moment.
        public Action? AfterDeleteUnusedUploads { get; set; }

        public Task<bool> IsUploadedAsync(string sha256, CancellationToken cancellationToken = default) =>
            this.IsUploadedFails.Contains(sha256)
                ? throw new InvalidOperationException("The database could not be asked.")
                : Task.FromResult(this.Uploads.Any(upload => upload.Sha256 == sha256));

        // The pictures the database fails to be asked about - a connection dropped mid-sweep.
        public HashSet<string> IsUploadedFails { get; } = [];
    }
}
