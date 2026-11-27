using CRT.Server.Handlers.Collections;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // ICollectionStore in memory - the rules MySqlCollectionStore keeps in SQL, kept here in plain code: an id taken is no
    // add, a deleted collection takes no change, a date moved later asks for the reminder again, and an upload nothing shows is
    // deleted a day after it came.
    // ###########################################################################################
    public sealed class FakeCollectionStore : ICollectionStore
    {
        public Dictionary<string, CollectionRecord> Collections { get; } = new(StringComparer.Ordinal);

        public List<CollectionUploadRecord> Uploads { get; } = [];

        // Every comment, deleted ones too (their DeletedUtc set) - the fake joins its writer from `Accounts`.
        public List<(CollectionCommentRecord Comment, DateTimeOffset? DeletedUtc)> Comments { get; } = [];

        // The writers' names and whether each is an administrator - a test fills it as it makes accounts.
        public Dictionary<long, (string Name, bool IsAdministrator)> Accounts { get; } = [];

        // The accounts locked, as the database's join to accounts reads them at every ask - a test locking an account in its
        // FakeAccountStore puts it here too.
        public HashSet<long> LockedAccounts { get; } = [];

        private long thisNextCommentId = 1;

        public Task<CollectionRecord?> FindAsync(string id, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.TryGetValue(id, out CollectionRecord? collection) ? collection : null);

        public Task<CollectionRecord?> FindWithImageAsync(string id, string sha256, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(id, out CollectionRecord? collection))
                return Task.FromResult<CollectionRecord?>(null);

            List<CollectionImageRecord> images = collection.Images.Where(image => image.Entry.Sha256 == sha256).ToList();

            return Task.FromResult<CollectionRecord?>(collection with { Images = images, ImageCount = images.Count });
        }

        public Task<IReadOnlyList<CollectionRecord>> ListLiveForAccountAsync(long accountId, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionRecord>>(this.Collections.Values
                .Where(collection => collection.AccountId == accountId && !collection.IsDeleted)
                .OrderByDescending(collection => collection.CreatedUtc)
                .Select(this.Counted)
                .ToList());

        public Task<IReadOnlyList<CollectionRecord>> ListLiveCommentedOnAsync(long accountId, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionRecord>>(this.Collections.Values
                .Where(collection => collection.AccountId != accountId && !collection.IsDeleted)
                .Select(collection => (Collection: collection, Latest: this.Comments
                    .Where(kept => kept.DeletedUtc is null && kept.Comment.AccountId == accountId && kept.Comment.CollectionId == collection.Id &&
                                   this.IsOnACaptureShown(kept.Comment))
                    .Select(kept => kept.Comment.Id)
                    .DefaultIfEmpty(0)
                    .Max()))
                .Where(pair => pair.Latest > 0)
                .OrderByDescending(pair => pair.Latest)
                .Select(pair => this.Counted(pair.Collection))
                .ToList());

        public Task<IReadOnlyList<CollectionRecord>> ListAllLiveAsync(int count, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionRecord>>(this.Collections.Values
                .Where(collection => !collection.IsDeleted)
                .OrderByDescending(collection => collection.CreatedUtc)
                .Take(count)
                .Select(this.Counted)
                .ToList());

        public Task<IReadOnlyList<CollectionCommentRecord>> ListNewestCommentsAsync(int count, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionCommentRecord>>(this.Comments
                .Where(kept => kept.DeletedUtc is null &&
                               this.Collections.TryGetValue(kept.Comment.CollectionId, out CollectionRecord? collection) &&
                               !collection.IsDeleted && this.IsOnACaptureShown(kept.Comment))
                .Select(kept => this.AsRead(kept.Comment))
                .OrderByDescending(comment => comment.Id)
                .Take(count)
                .ToList());

        // As the database's OnACaptureShown: a comment counts only on a capture its collection shows now.
        private bool IsOnACaptureShown(CollectionCommentRecord comment) =>
            this.Collections.TryGetValue(comment.CollectionId, out CollectionRecord? collection) &&
            collection.Images.Any(image => image.Entry.Sha256 == comment.ImageSha256);

        // A comment as the join reads it now - its writer locked or not.
        private CollectionCommentRecord AsRead(CollectionCommentRecord comment) =>
            comment with { AuthorIsLocked = this.LockedAccounts.Contains(comment.AccountId) };

        // A collection as a list holds it: no captures, those counted, and the comments its page shows counted - none hidden,
        // none of a locked account.
        private CollectionRecord Counted(CollectionRecord collection)
        {
            List<CollectionCommentRecord> shown = this.Comments
                .Where(kept => kept.DeletedUtc is null && !kept.Comment.IsHidden && kept.Comment.CollectionId == collection.Id &&
                               !this.LockedAccounts.Contains(kept.Comment.AccountId) && this.IsOnACaptureShown(kept.Comment))
                .Select(kept => kept.Comment)
                .ToList();

            // The newest of others than the owner, as the database's.
            List<CollectionCommentRecord> others = shown.Where(comment => comment.AccountId != collection.AccountId).ToList();

            return collection with
            {
                Images = [],
                ImageCount = collection.Images.Count,
                CommentCount = shown.Count,
                LatestCommentId = others.Count == 0 ? 0 : others.Max(comment => comment.Id)
            };
        }

        // Run as a share counts the day's collections - another request coming at that very moment.
        public Action? WhileCountingCreated { get; set; }

        public Task<int> CountCreatedSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            Action? meanwhile = this.WhileCountingCreated;

            this.WhileCountingCreated = null;
            meanwhile?.Invoke();

            return Task.FromResult(this.Collections.Values.Count(collection => collection.AccountId == accountId && collection.CreatedUtc > sinceUtc));
        }

        public Task<bool> AddAsync(CollectionRecord collection, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.TryAdd(collection.Id, collection));

        public Task<CollectionRecord?> FindLiveByClientKeyAsync(long accountId, string clientKey, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.Values
                .Where(collection => collection.AccountId == accountId && collection.ClientKey == clientKey && !collection.IsDeleted)
                .OrderByDescending(collection => collection.CreatedUtc)
                .FirstOrDefault());

        public Task<bool> ReplaceContentAsync(CollectionRecord collection, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(collection.Id, out CollectionRecord? existing) || existing.IsDeleted)
                return Task.FromResult(false);

            // As the database's UPDATE: the key CRT first shared it under stays, and a share brings one whose sharing was stopped
            // back (migration 0028).
            this.Collections[collection.Id] = collection with
            {
                ClientKey = existing.ClientKey,
                AccountId = existing.AccountId,
                CreatedUtc = existing.CreatedUtc,
                RemindedUtc = collection.DeleteAtUtc != existing.DeleteAtUtc ? collection.RemindedUtc : existing.RemindedUtc,
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
                Images = [],
                ImageCount = 0
            };

            // Its comments go with it.
            this.Comments.RemoveAll(kept => kept.Comment.CollectionId == id);

            return Task.FromResult(true);
        }

        public Task<IReadOnlyList<CollectionCommentRecord>> ListCommentsAsync(string collectionId, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionCommentRecord>>(this.Comments
                .Where(kept => kept.DeletedUtc is null && kept.Comment.CollectionId == collectionId)
                .Select(kept => this.AsRead(kept.Comment))
                .OrderBy(comment => comment.CreatedUtc)
                .ThenBy(comment => comment.Id)
                .ToList());

        public Task<CollectionCommentRecord?> FindCommentAsync(long commentId, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Comments.Where(kept => kept.DeletedUtc is null && kept.Comment.Id == commentId).Select(kept => this.AsRead(kept.Comment)).FirstOrDefault());

        public Task<long> AddCommentAsync(string collectionId, string imageSha256, long accountId, string body, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            (string name, bool isAdministrator) = this.Accounts.TryGetValue(accountId, out var writer) ? writer : ($"Account {accountId}", false);
            long id = this.thisNextCommentId++;

            this.Comments.Add((new CollectionCommentRecord(id, collectionId, imageSha256, accountId, name, isAdministrator, body, whenUtc, null, null), null));

            return Task.FromResult(id);
        }

        public Task<bool> EditCommentAsync(long commentId, string body, DateTimeOffset whenUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Change(commentId, comment => (comment with { Body = body, EditedUtc = whenUtc }, null)));

        public Task<bool> DeleteCommentAsync(long commentId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Change(commentId, comment => (comment, whenUtc)));

        public Task<bool> SetCommentHiddenAsync(long commentId, bool hidden, DateTimeOffset whenUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Change(commentId, comment => (comment with { HiddenUtc = hidden ? whenUtc : null }, null)));

        public Task<int> CountCommentsSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Comments.Count(kept => kept.Comment.AccountId == accountId && kept.Comment.CreatedUtc > sinceUtc));

        // A comment not deleted, changed - false when there is none.
        private bool Change(long commentId, Func<CollectionCommentRecord, (CollectionCommentRecord Comment, DateTimeOffset? DeletedUtc)> change)
        {
            int at = this.Comments.FindIndex(kept => kept.DeletedUtc is null && kept.Comment.Id == commentId);

            if (at < 0)
                return false;

            this.Comments[at] = change(this.Comments[at].Comment);
            return true;
        }

        public Task<IReadOnlyList<string>> ListExpiredAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<string>>(this.Collections.Values
                .Where(collection => !collection.IsDeleted && collection.DeleteAtUtc <= nowUtc)
                .Select(collection => collection.Id)
                .ToList());

        public Task<IReadOnlyList<CollectionRecord>> ListToRemindAsync(DateTimeOffset byUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionRecord>>(this.Collections.Values
                .Where(collection => !collection.IsDeleted && collection.RemindedUtc is null && collection.DeleteAtUtc <= byUtc)
                .ToList());

        public Task MarkRemindedAsync(string id, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.Collections[id] = this.Collections[id] with { RemindedUtc = whenUtc };
            return Task.CompletedTask;
        }

        public Task<IReadOnlyDictionary<string, CollectionUploadRecord>> FindUploadsAsync(
            long accountId,
            IReadOnlyCollection<string> sha256s,
            CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyDictionary<string, CollectionUploadRecord>>(this.Uploads
                .Where(upload => upload.AccountId == accountId && sha256s.Contains(upload.Sha256))
                .ToDictionary(upload => upload.Sha256, StringComparer.Ordinal));

        public Task AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default)
        {
            int earlier = this.Uploads.FindIndex(row => row.AccountId == upload.AccountId && row.Sha256 == upload.Sha256);

            if (earlier < 0)
                this.Uploads.Add(upload);
            else if (upload.UploadedUtc > this.Uploads[earlier].UploadedUtc)
                this.Uploads[earlier] = this.Uploads[earlier] with { UploadedUtc = upload.UploadedUtc };

            return Task.CompletedTask;
        }

        public Task<int> CountUploadsSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Uploads.Count(upload => upload.AccountId == accountId && upload.UploadedUtc > sinceUtc));

        public Task<long> StoredBytesAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Uploads.GroupBy(upload => upload.Sha256).Sum(group => (long)group.Max(upload => upload.ByteCount)));

        public Task<IReadOnlyList<string>> DeleteUnusedUploadsAsync(DateTimeOffset uploadedBeforeUtc, CancellationToken cancellationToken = default)
        {
            List<CollectionUploadRecord> unused = this.Uploads
                .Where(upload => upload.UploadedUtc < uploadedBeforeUtc && !this.Collections.Values.Any(collection =>
                    collection.AccountId == upload.AccountId &&
                    !collection.IsDeleted &&
                    collection.Images.Any(image => image.Entry.Sha256 == upload.Sha256)))
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
        // as another account may at that very moment.
        public Action? AfterDeleteUnusedUploads { get; set; }

        public Task<bool> IsUploadedAsync(string sha256, CancellationToken cancellationToken = default) =>
            this.IsUploadedFails.Contains(sha256)
                ? throw new InvalidOperationException("The database could not be asked.")
                : Task.FromResult(this.Uploads.Any(upload => upload.Sha256 == sha256));

        // The pictures the database fails to be asked about - a connection dropped mid-sweep.
        public HashSet<string> IsUploadedFails { get; } = [];
    }
}
