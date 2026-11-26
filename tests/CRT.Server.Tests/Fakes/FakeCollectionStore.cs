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

        private long thisNextCommentId = 1;

        public Task<CollectionRecord?> FindAsync(string id, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.TryGetValue(id, out CollectionRecord? collection) ? collection : null);

        public Task<IReadOnlyList<CollectionRecord>> ListLiveForAccountAsync(long accountId, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<CollectionRecord>>(this.Collections.Values
                .Where(collection => collection.AccountId == accountId && !collection.IsDeleted)
                .OrderByDescending(collection => collection.CreatedUtc)
                .Select(collection =>
                {
                    List<CollectionCommentRecord> shown = this.Comments
                        .Where(kept => kept.DeletedUtc is null && !kept.Comment.IsHidden && kept.Comment.CollectionId == collection.Id)
                        .Select(kept => kept.Comment)
                        .ToList();

                    return collection with
                    {
                        Images = [],
                        ImageCount = collection.Images.Count,
                        CommentCount = shown.Count,
                        LatestCommentId = shown.Count == 0 ? 0 : shown.Max(comment => comment.Id)
                    };
                })
                .ToList());

        public Task<int> CountCreatedSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.Values.Count(collection => collection.AccountId == accountId && collection.CreatedUtc > sinceUtc));

        public Task<bool> AddAsync(CollectionRecord collection, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Collections.TryAdd(collection.Id, collection));

        public Task<bool> ReplaceContentAsync(CollectionRecord collection, CancellationToken cancellationToken = default)
        {
            if (!this.Collections.TryGetValue(collection.Id, out CollectionRecord? existing) || existing.IsDeleted)
                return Task.FromResult(false);

            this.Collections[collection.Id] = collection with
            {
                AccountId = existing.AccountId,
                CreatedUtc = existing.CreatedUtc,
                RemindedUtc = collection.DeleteAtUtc > existing.DeleteAtUtc ? null : existing.RemindedUtc
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
                .Select(kept => kept.Comment)
                .OrderBy(comment => comment.CreatedUtc)
                .ThenBy(comment => comment.Id)
                .ToList());

        public Task<CollectionCommentRecord?> FindCommentAsync(long commentId, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Comments.Where(kept => kept.DeletedUtc is null && kept.Comment.Id == commentId).Select(kept => kept.Comment).FirstOrDefault());

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

        public Task<bool> AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default)
        {
            if (this.Uploads.Any(earlier => earlier.AccountId == upload.AccountId && earlier.Sha256 == upload.Sha256))
                return Task.FromResult(false);

            this.Uploads.Add(upload);
            return Task.FromResult(true);
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

            return Task.FromResult<IReadOnlyList<string>>(unused
                .Select(upload => upload.Sha256)
                .Distinct(StringComparer.Ordinal)
                .Where(sha => !this.Uploads.Any(upload => upload.Sha256 == sha))
                .ToList());
        }
    }
}
