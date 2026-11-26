using System.Security.Cryptography;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // EVERY DECISION ABOUT A SHARED COLLECTION (2026-11-26; CRT.Data's CollectionRules says what one is, CollectionContract how
    // CRT sends it). Pure: the stores, the files and the mailer are handed in, so the tests run it on fakes.
    //
    //   - ONLY A SIGNED-IN ACCOUNT WITH A CONFIRMED ADDRESS shares (owner decision, 2026-11-25: accounts for everybody). An
    //     address not confirmed cannot sign in since server 14.0.0; a session from before is told to confirm it.
    //   - A collection shows only pictures ITS OWN ACCOUNT uploaded - a checksum read off somebody else's page is no way to
    //     show their picture under another collection.
    //   - A collection that is not the account's is "not found" - whether it exists is nobody else's business. Only the
    //     administrator may delete another account's (step 4's moderation).
    //   - Limits: CollectionRules' per account, the day's counted from the database; the space all captures take
    //     (CollectionsMaxStoredBytes) and the disk's reserve (MinimumFreeDiskBytes, the blob store's).
    // ###########################################################################################
    public static class CollectionFlows
    {
        public const string NotSetUp = "Sharing collections is not set up on this server yet.";

        public const string AddressNotConfirmed =
            "Confirm your address before you share a collection - the link is in the mail sent when the account was made. " +
            "No mail? \"I forgot my password\" in the account window mails a code that confirms it too.";

        public const string NoSuchCollection = "There is no collection of yours with that id.";

        public const string NoRoom = "The server has no room for more captures right now. Try again in a few days.";

        public const string ChecksumMismatch = "The capture's bytes do not match its checksum - it was damaged on the way. Share it again.";

        public static readonly TimeSpan UnusedUploadKept = TimeSpan.FromDays(1);

        // ###########################################################################################
        // PUT /api/collections/images/<sha256> - one picture.
        // ###########################################################################################
        public static async Task<CollectionOutcome> UploadImageAsync(
            AccountRecord? account,
            string? sha256,
            byte[] bytes,
            ICollectionStore store,
            CollectionFiles files,
            CollectionStorage storage,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.Refuse(account) is { } refused)
                return refused;

            if (!DataTreeLayout.IsSha256(sha256))
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, "The address does not name a capture by its checksum.");

            if (!string.Equals(Convert.ToHexStringLower(SHA256.HashData(bytes)), sha256, StringComparison.Ordinal))
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.ChecksumMismatch);

            CollectionPng.Result picture = CollectionPng.Check(bytes);

            if (!picture.IsCapture)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, picture.Problem ?? CollectionPng.NotAPicture);

            IReadOnlyDictionary<string, CollectionUploadRecord> earlier =
                await store.FindUploadsAsync(account!.Id, [sha256!], cancellationToken);

            bool isNew = !earlier.ContainsKey(sha256!);

            if (isNew &&
                await store.CountUploadsSinceAsync(account.Id, now - TimeSpan.FromDays(1), cancellationToken) >= CollectionRules.MaximumUploadsPerDay)
            {
                return CollectionOutcome.Refused(
                    CollectionVerdict.TooMany,
                    $"An account shares at most {CollectionRules.MaximumUploadsPerDay} captures a day - share the rest tomorrow.");
            }

            // Bytes already kept - this account's or another's - take no more room.
            if (!files.Has(sha256!))
            {
                if (!await storage.HasRoomForAsync(bytes.Length, store, cancellationToken))
                    return CollectionOutcome.Refused(CollectionVerdict.NoRoom, CollectionFlows.NoRoom);

                files.Write(sha256!, bytes);
            }

            if (isNew)
            {
                await store.AddUploadAsync(
                    new CollectionUploadRecord(account.Id, sha256!, bytes.Length, picture.Width, picture.Height, now),
                    cancellationToken);
            }

            return CollectionOutcome.Uploaded(new CollectionImageUploadAnswer(sha256!, picture.Width, picture.Height));
        }

        // ###########################################################################################
        // POST /api/collections - a collection shared for the first time.
        // ###########################################################################################
        public static async Task<CollectionOutcome> CreateAsync(
            AccountRecord? account,
            CollectionContentRequest? content,
            ICollectionStore store,
            CollectionFiles files,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.Refuse(account) is { } refused)
                return refused;

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, now);

            if (problems.Count > 0)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, problems);

            IReadOnlyList<CollectionRecord> live = await store.ListLiveForAccountAsync(account!.Id, cancellationToken);

            if (live.Count >= CollectionRules.MaximumLiveCollectionsPerAccount)
            {
                return CollectionOutcome.Refused(
                    CollectionVerdict.TooMany,
                    $"An account keeps at most {CollectionRules.MaximumLiveCollectionsPerAccount} shared collections - delete one it no longer needs first.");
            }

            if (await store.CountCreatedSinceAsync(account.Id, now - TimeSpan.FromDays(1), cancellationToken) >= CollectionRules.MaximumCollectionsPerDay)
            {
                return CollectionOutcome.Refused(
                    CollectionVerdict.TooMany,
                    $"An account shares at most {CollectionRules.MaximumCollectionsPerDay} new collections a day - share this one tomorrow.");
            }

            (IReadOnlyList<CollectionImageRecord>? images, CollectionOutcome? missing) =
                await CollectionFlows.ImagesOfAsync(account.Id, content!, store, files, cancellationToken);

            if (missing is not null)
                return missing;

            // An id taken is drawn again - 32^12 ids, so a second draw is all it can ever need; five is plenty.
            for (int attempt = 0; attempt < 5; attempt++)
            {
                var collection = CollectionFlows.RecordOf(CollectionRules.NewId(), account.Id, content!, images!, now, now);

                if (await store.AddAsync(collection, cancellationToken))
                    return CollectionOutcome.Created(CollectionFlows.AnswerOf(collection, siteBaseUrl));
            }

            throw new InvalidOperationException("No free collection id was found in five draws.");
        }

        // ###########################################################################################
        // PUT /api/collections/<id> - a collection shared before, as it is now. Its id, owner and creation stay.
        // ###########################################################################################
        public static async Task<CollectionOutcome> UpdateAsync(
            AccountRecord? account,
            string? id,
            CollectionContentRequest? content,
            ICollectionStore store,
            CollectionFiles files,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.Refuse(account) is { } refused)
                return refused;

            CollectionRecord? existing = CollectionRules.IsId(id) ? await store.FindAsync(id!, cancellationToken) : null;

            if (existing is null || existing.AccountId != account!.Id)
                return CollectionOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection);

            if (existing.IsDeleted)
                return CollectionOutcome.Refused(CollectionVerdict.Gone, CollectionFlows.DeletedLine(existing));

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, now);

            if (problems.Count > 0)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, problems);

            (IReadOnlyList<CollectionImageRecord>? images, CollectionOutcome? missing) =
                await CollectionFlows.ImagesOfAsync(account.Id, content!, store, files, cancellationToken);

            if (missing is not null)
                return missing;

            CollectionRecord changed = CollectionFlows.RecordOf(existing.Id, account.Id, content!, images!, existing.CreatedUtc, now) with
            {
                RemindedUtc = existing.RemindedUtc
            };

            if (!await store.ReplaceContentAsync(changed, cancellationToken))
                return CollectionOutcome.Refused(CollectionVerdict.Gone, "That collection was deleted meanwhile.");

            return CollectionOutcome.Done(CollectionFlows.AnswerOf(changed, siteBaseUrl));
        }

        // ###########################################################################################
        // DELETE /api/collections/<id> - by its owner, or by the administrator. Deleting what is deleted already is done.
        // ###########################################################################################
        public static async Task<CollectionOutcome> DeleteAsync(
            AccountRecord? account,
            string? id,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CollectionOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            CollectionRecord? existing = CollectionRules.IsId(id) ? await store.FindAsync(id!, cancellationToken) : null;

            bool isOwner = existing is not null && existing.AccountId == account.Id;

            if (existing is null || (!isOwner && !account.IsAdministrator))
                return CollectionOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection);

            if (!existing.IsDeleted)
            {
                await store.MarkDeletedAsync(
                    existing.Id,
                    isOwner ? CollectionDeletedReasons.Owner : CollectionDeletedReasons.Administrator,
                    now,
                    cancellationToken);
            }

            return CollectionOutcome.Removed();
        }

        // ###########################################################################################
        // GET /api/collections/mine.
        // ###########################################################################################
        public static async Task<CollectionOutcome> ListMineAsync(
            AccountRecord? account,
            ICollectionStore store,
            string siteBaseUrl,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CollectionOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            IReadOnlyList<CollectionRecord> live = await store.ListLiveForAccountAsync(account.Id, cancellationToken);

            return CollectionOutcome.Listed(new MyCollectionsAnswer(
                live.Select(collection => new CollectionSummaryAnswer(
                    collection.Id,
                    CollectionRules.PageUrl(siteBaseUrl, collection.Id),
                    collection.Title,
                    collection.HardwareName,
                    collection.BoardName,
                    collection.Region,
                    collection.ImageCount,
                    collection.CreatedUtc,
                    collection.UpdatedUtc,
                    collection.DeleteAtUtc,
                    collection.CommentCount,
                    collection.LatestCommentId)).ToList()));
        }

        // ###########################################################################################
        // THE HOURLY SWEEP (CollectionSweeper): collections whose time is up are deleted, their owners mailed 30 days before,
        // and a picture no collection of its account shows any more goes a day after its upload - a day, so a capture
        // uploaded for a collection about to be sent is not taken from under it.
        // ###########################################################################################
        public static async Task<CollectionSweepResult> SweepAsync(
            ICollectionStore store,
            CollectionFiles files,
            IAccountStore accounts,
            IEmailSender mailer,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            int expired = 0;

            foreach (string id in await store.ListExpiredAsync(now, cancellationToken))
            {
                if (await store.MarkDeletedAsync(id, CollectionDeletedReasons.Expired, now, cancellationToken))
                    expired++;
            }

            int reminded = 0;

            foreach (CollectionRecord collection in await store.ListToRemindAsync(now + CollectionRules.ReminderBefore, cancellationToken))
            {
                AccountRecord? owner = await accounts.FindByIdAsync(collection.AccountId, cancellationToken);

                // Marked either way: an owner gone or locked is reminded of nothing, and the next sweep does not ask again. Nor is
                // one kept a month or so (CollectionFlows.IsWorthReminding) - the reminder would come the day it was shared.
                if (owner is { IsVerified: true, IsLocked: false } && CollectionFlows.IsWorthReminding(collection))
                {
                    await mailer.SendAsync(
                        EmailTemplates.CollectionDeletingSoon(
                            owner.Email,
                            owner.DisplayName,
                            CollectionPages.NameOf(collection),
                            collection.DeleteAtUtc,
                            CollectionRules.PageUrl(siteBaseUrl, collection.Id)),
                        cancellationToken);

                    reminded++;
                }

                await store.MarkRemindedAsync(collection.Id, now, cancellationToken);
            }

            int filesDeleted = 0;

            foreach (string sha256 in await store.DeleteUnusedUploadsAsync(now - CollectionFlows.UnusedUploadKept, cancellationToken))
            {
                if (files.Delete(sha256))
                    filesDeleted++;
            }

            return new CollectionSweepResult(expired, reminded, filesDeleted);
        }

        // ###########################################################################################
        // A collection as anybody with its link may see it: none for an id of the wrong shape, none at all, or one whose owner's
        // account is locked - a locked account's collections go offline with it. A deleted one only when asked for (its page
        // says when it went). The page and the comments alike ask this.
        // ###########################################################################################
        public static async Task<CollectionRecord?> FindShownAsync(
            string? id,
            ICollectionStore store,
            IAccountStore accounts,
            CancellationToken cancellationToken,
            bool evenIfDeleted = false)
        {
            if (!CollectionRules.IsId(id))
                return null;

            CollectionRecord? collection = await store.FindAsync(id!, cancellationToken);

            if (collection is null || (collection.IsDeleted && !evenIfDeleted))
                return null;

            if (collection.IsDeleted)
                return collection;

            AccountRecord? owner = await accounts.FindByIdAsync(collection.AccountId, cancellationToken);

            return owner is { IsLocked: false } ? collection : null;
        }

        // A reminder only for a collection kept longer than twice the reminder's notice - "1 month" is told nothing.
        public static bool IsWorthReminding(CollectionRecord collection) =>
            collection.DeleteAtUtc - collection.CreatedUtc > CollectionRules.ReminderBefore * 2;

        // "Deleted on 2026-November-26" - said by the routes and the page.
        public static string DeletedLine(CollectionRecord collection) =>
            $"That collection was deleted on {CollectionRules.DateText(collection.DeletedUtc ?? collection.DeleteAtUtc)}.";

        private static CollectionOutcome? Refuse(AccountRecord? account)
        {
            if (account is null)
                return CollectionOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            if (!account.IsVerified)
                return CollectionOutcome.Refused(CollectionVerdict.Forbidden, CollectionFlows.AddressNotConfirmed);

            return null;
        }

        // The captures as kept - each with its picture's size from its upload - or the refusal naming the pictures the server
        // does not hold for this account (uploaded by it, the bytes still here).
        private static async Task<(IReadOnlyList<CollectionImageRecord>? Images, CollectionOutcome? Missing)> ImagesOfAsync(
            long accountId,
            CollectionContentRequest content,
            ICollectionStore store,
            CollectionFiles files,
            CancellationToken cancellationToken)
        {
            string[] wanted = content.Images.Select(image => image.Sha256).Distinct(StringComparer.Ordinal).ToArray();

            IReadOnlyDictionary<string, CollectionUploadRecord> uploads =
                await store.FindUploadsAsync(accountId, wanted, cancellationToken);

            List<string> missing = wanted.Where(sha => !uploads.ContainsKey(sha) || !files.Has(sha)).ToList();

            if (missing.Count > 0)
            {
                return (null, CollectionOutcome.MissingImages(
                    missing,
                    missing.Count == 1
                        ? "The server does not hold 1 of the captures yet - CRT sends it, then the collection again."
                        : $"The server does not hold {missing.Count} of the captures yet - CRT sends them, then the collection again."));
            }

            return (content.Images
                .Select(image => new CollectionImageRecord(CollectionFlows.Tidied(image), uploads[image.Sha256].Width, uploads[image.Sha256].Height))
                .ToList(), null);
        }

        private static CollectionRecord RecordOf(
            string id,
            long accountId,
            CollectionContentRequest content,
            IReadOnlyList<CollectionImageRecord> images,
            DateTimeOffset createdUtc,
            DateTimeOffset now) =>
            new(
                id,
                accountId,
                content.HardwareName.Trim(),
                content.BoardName.Trim(),
                CollectionFlows.Blank(content.Region),
                CollectionFlows.Blank(content.Title),
                CollectionFlows.Blank(content.Description),
                (content.ForumLinks ?? []).Select(link => link.Trim()).ToList(),
                createdUtc,
                now,
                content.DeleteAtUtc.ToUniversalTime(),
                RemindedUtc: null,
                DeletedUtc: null,
                DeletedReason: null,
                images,
                images.Count);

        private static CollectionImageEntry Tidied(CollectionImageEntry image) =>
            image with
            {
                ComponentLabel = image.ComponentLabel.Trim(),
                Pin = CollectionFlows.Blank(image.Pin),
                SignalName = CollectionFlows.Blank(image.SignalName),
                ExpectedReading = CollectionFlows.Blank(image.ExpectedReading),
                TimePerDivision = CollectionFlows.Blank(image.TimePerDivision),
                VoltsPerDivision = CollectionFlows.Blank(image.VoltsPerDivision),
                TriggerLevel = CollectionFlows.Blank(image.TriggerLevel),
                Caption = CollectionFlows.Blank(image.Caption),
                Baseline = image.Baseline is { } baseline
                    ? new CollectionBaseline(
                        CollectionFlows.Blank(baseline.Sha256),
                        CollectionFlows.Blank(baseline.FileName),
                        CollectionFlows.Blank(baseline.TimePerDivision),
                        CollectionFlows.Blank(baseline.VoltsPerDivision),
                        CollectionFlows.Blank(baseline.TriggerLevel))
                    : null
            };

        private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        private static CollectionAnswer AnswerOf(CollectionRecord collection, string siteBaseUrl) =>
            new(collection.Id, CollectionRules.PageUrl(siteBaseUrl, collection.Id), collection.CreatedUtc, collection.DeleteAtUtc, collection.Images.Count);
    }

    // ###########################################################################################
    // How much room the captures may take: the total (CollectionsMaxStoredBytes, 0 for none) and the disk's reserve.
    // ###########################################################################################
    public sealed class CollectionStorage(long maximumStoredBytes, Func<long?> freeBytes, long minimumFreeBytes)
    {
        public async Task<bool> HasRoomForAsync(long bytes, ICollectionStore store, CancellationToken cancellationToken)
        {
            // The probe answers null when it cannot measure - room, as the blob store takes it (BlobStore.HasRoomFor).
            if (freeBytes() is long free && free - bytes < minimumFreeBytes)
                return false;

            return maximumStoredBytes <= 0 || await store.StoredBytesAsync(cancellationToken) + bytes <= maximumStoredBytes;
        }
    }

    public enum CollectionVerdict
    {
        Done,
        Created,
        Removed,
        Unauthorized,
        Forbidden,
        BadRequest,
        NotFound,
        Gone,
        MissingImages,
        TooMany,
        NoRoom,
        Off
    }

    public sealed record CollectionOutcome(
        CollectionVerdict Verdict,
        IReadOnlyList<string> Errors,
        CollectionAnswer? Answer = null,
        IReadOnlyList<string>? Missing = null,
        CollectionImageUploadAnswer? Upload = null,
        MyCollectionsAnswer? Mine = null)
    {
        public static CollectionOutcome Refused(CollectionVerdict verdict, string error) => new(verdict, [error]);

        public static CollectionOutcome Refused(CollectionVerdict verdict, IReadOnlyList<string> errors) => new(verdict, errors);

        public static CollectionOutcome MissingImages(IReadOnlyList<string> missing, string error) =>
            new(CollectionVerdict.MissingImages, [error], Missing: missing);

        public static CollectionOutcome Created(CollectionAnswer answer) => new(CollectionVerdict.Created, [], Answer: answer);

        public static CollectionOutcome Done(CollectionAnswer answer) => new(CollectionVerdict.Done, [], Answer: answer);

        public static CollectionOutcome Uploaded(CollectionImageUploadAnswer upload) => new(CollectionVerdict.Done, [], Upload: upload);

        public static CollectionOutcome Listed(MyCollectionsAnswer mine) => new(CollectionVerdict.Done, [], Mine: mine);

        public static CollectionOutcome Removed() => new(CollectionVerdict.Removed, []);
    }

    public sealed record CollectionSweepResult(int Expired, int Reminded, int FilesDeleted);
}
