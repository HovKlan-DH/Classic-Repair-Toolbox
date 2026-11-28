using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // EVERY DECISION ABOUT A SHARED COLLECTION (2026-11-26; CRT.Data's CollectionRules says what one is, CollectionContract how
    // CRT sends it). Pure: the stores and the files are handed in, so the tests run it on fakes.
    //
    //   - *** ANYBODY SHARES, NO ACCOUNT (owner decision, 2026-11-28: "Entirely remove the new "Account" possibility ... I do
    //     not want the possibility for everyone to create an account"; asked, "Anybody, no account"; server 17.0.0). *** A
    //     collection is OWNED BY ITS KEY - the one CRT made for it as it first shared it, kept with the collection on that
    //     computer and sent with every request (CollectionRoutes.OwnerKeyHeader). Only its hash is kept (OwnerHashOf). A
    //     collection another key owns is "not found" - whether it exists is nobody else's business. A collection shared with an
    //     account before is owned by no key any more: it goes at its date, or by the administrator's delete on the website.
    //   - A collection shows only pictures ITS OWN KEY uploaded - a checksum read off somebody else's page is no way to show
    //     their picture under another collection.
    //   - Limits: CollectionRules' per INTERNET CONNECTION (an IPv4 address, an IPv6 /64), the day's counted from the database;
    //     the space all captures take (CollectionsMaxStoredBytes) and the disk's reserve (MinimumFreeDiskBytes, the blob store's).
    //   - The website's only signed-in reader is the ADMINISTRATOR (the web sign-in, WebSessions): it sees every collection, an
    //     offline one too, and deletes one. Nobody comments (the comments went with the accounts for everybody).
    // ###########################################################################################
    public static class CollectionFlows
    {
        public const string NotSetUp = "Sharing collections is not set up on this server yet.";

        public const string NoKey =
            "The request carries no collection key - it comes from a CRT that shares with an account. Update CRT, then share the " +
            "collection again.";

        public const string NoSuchCollection = "There is no collection with that id that this computer shared.";

        public const string NoRoom = "The server has no room for more captures right now. Try again in a few days.";

        public const string ChecksumMismatch = "The capture's bytes do not match its checksum - it was damaged on the way. Share it again.";

        public const string NotACaptureAddress = "The address does not name a capture by its checksum.";

        public const string OnlyTheAdministratorDeletesHere =
            "Only the administrator deletes a collection on the website - whoever shared it deletes it in CRT.";

        // When the connection a request came from is not known - counted as one, so an unknown sender is limited as a whole.
        public const string UnknownConnection = "unknown";

        public static readonly TimeSpan UnusedUploadKept = TimeSpan.FromDays(1);

        // The hash a key is kept by - lower-case hex SHA-256 of its characters - or null for no key of the right shape.
        public static string? OwnerHashOf(string? key) =>
            CollectionRules.IsOwnerKey(key) ? Convert.ToHexStringLower(SHA256.HashData(Encoding.ASCII.GetBytes(key!))) : null;

        // ###########################################################################################
        // PUT /api/collections/images/<sha256> - one picture, the key's.
        // ###########################################################################################
        public static async Task<CollectionOutcome> UploadImageAsync(
            string? ownerKey,
            string? connection,
            string? sha256,
            byte[] bytes,
            ICollectionStore store,
            CollectionFiles files,
            CollectionStorage storage,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.OwnerHashOf(ownerKey) is not { } owner)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.NoKey);

            if (!DataTreeLayout.IsSha256(sha256))
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.NotACaptureAddress);

            if (!string.Equals(Convert.ToHexStringLower(SHA256.HashData(bytes)), sha256, StringComparison.Ordinal))
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.ChecksumMismatch);

            CollectionPng.Result picture = CollectionPng.Check(bytes);

            if (!picture.IsCapture)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, picture.Problem ?? CollectionPng.NotAPicture);

            string from = connection ?? CollectionFlows.UnknownConnection;

            // The look, the day's count, the bytes and the row under ONE lock (code review, 2026-11-27: looked at first, a sweep
            // between the look and the lock took the old row and the bytes, the upload wrote the bytes and no row - a file nothing
            // named, never swept - and parallel uploads all passed the day's limit together).
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            try
            {
                IReadOnlyDictionary<string, CollectionUploadRecord> earlier =
                    await store.FindUploadsAsync(owner, [sha256!], cancellationToken);

                bool isNew = !earlier.ContainsKey(sha256!);

                if (isNew &&
                    await store.CountUploadsSinceAsync(from, now - TimeSpan.FromDays(1), cancellationToken) >= CollectionRules.MaximumUploadsPerDay)
                {
                    return CollectionOutcome.Refused(
                        CollectionVerdict.TooMany,
                        $"An internet connection shares at most {CollectionRules.MaximumUploadsPerDay} captures a day - share the rest tomorrow.");
                }

                // Bytes already kept - this key's or another's - take no more room.
                if (!files.Has(sha256!))
                {
                    if (!await storage.HasRoomForAsync(bytes.Length, store, cancellationToken))
                        return CollectionOutcome.Refused(CollectionVerdict.NoRoom, CollectionFlows.NoRoom);

                    files.Write(sha256!, bytes);
                }

                // A picture sent again has its upload's time moved to now, so the day the sweep leaves an unused upload counts
                // from its LAST send - an old row no collection shows would go at the next sweep, before the collection it was
                // sent again for is shared.
                await store.AddUploadAsync(
                    new CollectionUploadRecord(owner, sha256!, bytes.Length, picture.Width, picture.Height, now, from),
                    cancellationToken);
            }
            finally
            {
                CollectionFlows.PicturesLock.Release();
            }

            return CollectionOutcome.Uploaded(new CollectionImageUploadAnswer(sha256!, picture.Width, picture.Height));
        }

        // ###########################################################################################
        // *** A PICTURE'S BYTES AND ITS UPLOAD ROWS ARE CHANGED UNDER ONE LOCK (code review, 2026-11-27). *** An upload finding
        // the bytes here writes none and adds its row; the sweep, finding no row left, deletes the bytes. Between the sweep's
        // look and its delete, an upload of the same picture by another key lost its bytes for good - nothing asks for them
        // again, as its row says it has them. So the lock is held by an upload (its look, write and row), a share (finding its
        // pictures and keeping its collection - CreateAsync, UpdateAsync) and the sweep (letting go of unused uploads and
        // deleting their bytes). The service is one process (the test instance another, with folders of its own), so the lock
        // is the process's.
        // ###########################################################################################
        private static readonly SemaphoreSlim PicturesLock = new(1, 1);

        // ###########################################################################################
        // POST /api/collections - a collection shared for the first time, owned by the key it is sent with.
        //
        // *** A FIRST SHARE SENT AGAIN UNDER ITS KEY CHANGES THE PAGE IT MADE (code review, 2026-11-27). *** A share the server
        // kept whose answer was lost made a second page at the next "Share", the first one's id never reaching CRT - so the
        // key's collection still shown is the one changed, as PUT changes it, never a new one.
        // ###########################################################################################
        public static async Task<CollectionOutcome> CreateAsync(
            string? ownerKey,
            string? connection,
            CollectionContentRequest? content,
            ICollectionStore store,
            CollectionFiles files,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.OwnerHashOf(ownerKey) is not { } owner)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.NoKey);

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, now);

            if (problems.Count > 0)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, problems);

            string from = connection ?? CollectionFlows.UnknownConnection;

            // The key looked up, the connection's limit counted, the pictures found and the collection kept under the pictures'
            // lock, as ONE step: the sweep lets go of an upload no collection shows, and between a look and the write it took one
            // this collection shows; and looked at before the lock, shares sent at the same moment all passed the limit, and a
            // first share retried while the first was being written found no page under its key and made a second (code reviews,
            // 2026-11-27).
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            bool locked = true;

            try
            {
                // One whose time is up is no page to change - a new one is made. The page made before is changed as PUT changes
                // it, which takes the lock itself - so it is let go of first; the page is there, so a share retried meanwhile
                // finds it too.
                if (await store.FindLiveByOwnerAsync(owner, cancellationToken) is { } madeBefore && !CollectionFlows.IsGone(madeBefore, now))
                {
                    CollectionFlows.PicturesLock.Release();
                    locked = false;

                    return await CollectionFlows.UpdateAsync(ownerKey, madeBefore.Id, content, store, files, siteBaseUrl, now, cancellationToken);
                }

                if (await store.CountCreatedSinceAsync(from, now - TimeSpan.FromDays(1), cancellationToken) >= CollectionRules.MaximumCollectionsPerDay)
                {
                    return CollectionOutcome.Refused(
                        CollectionVerdict.TooMany,
                        $"An internet connection shares at most {CollectionRules.MaximumCollectionsPerDay} new collections a day - share this one tomorrow.");
                }

                (IReadOnlyList<CollectionImageRecord>? images, CollectionOutcome? missing) =
                    await CollectionFlows.ImagesOfAsync(owner, content!, store, files, cancellationToken);

                if (missing is not null)
                    return missing;

                // An id taken is drawn again - 32^12 ids, so a second draw is all it can ever need; five is plenty.
                for (int attempt = 0; attempt < 5; attempt++)
                {
                    var collection = CollectionFlows.RecordOf(CollectionRules.NewId(), owner, content!, images!, now, now) with { Connection = from };

                    if (await store.AddAsync(collection, cancellationToken))
                        return CollectionOutcome.Created(CollectionFlows.AnswerOf(collection, siteBaseUrl));
                }

                throw new InvalidOperationException("No free collection id was found in five draws.");
            }
            finally
            {
                if (locked)
                    CollectionFlows.PicturesLock.Release();
            }
        }

        // ###########################################################################################
        // PUT /api/collections/<id> - a collection shared before, as it is now. Its id, key, connection and creation stay.
        // ###########################################################################################
        public static async Task<CollectionOutcome> UpdateAsync(
            string? ownerKey,
            string? id,
            CollectionContentRequest? content,
            ICollectionStore store,
            CollectionFiles files,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            (CollectionRecord? existing, string? owner, CollectionOutcome? refused) = await CollectionFlows.OwnedAsync(ownerKey, id, store, now, cancellationToken);

            if (refused is not null)
                return refused;

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, now);

            if (problems.Count > 0)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, problems);

            // As CreateAsync: the pictures found and the collection kept under the pictures' lock.
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            try
            {
                (IReadOnlyList<CollectionImageRecord>? images, CollectionOutcome? missing) =
                    await CollectionFlows.ImagesOfAsync(owner!, content!, store, files, cancellationToken);

                if (missing is not null)
                    return missing;

                CollectionRecord changed = CollectionFlows.RecordOf(existing!.Id, owner!, content!, images!, existing.CreatedUtc, now) with
                {
                    Connection = existing.Connection
                };

                if (!await store.ReplaceContentAsync(changed, cancellationToken))
                    return CollectionOutcome.Refused(CollectionVerdict.Gone, "That collection was deleted meanwhile.");

                return CollectionOutcome.Done(CollectionFlows.AnswerOf(changed, siteBaseUrl));
            }
            finally
            {
                CollectionFlows.PicturesLock.Release();
            }
        }

        // ###########################################################################################
        // DELETE /api/collections/<id> - by its key. Deleting what is deleted already is done.
        // ###########################################################################################
        public static async Task<CollectionOutcome> DeleteAsync(
            string? ownerKey,
            string? id,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.OwnerHashOf(ownerKey) is not { } owner)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.NoKey);

            CollectionRecord? existing = CollectionRules.IsId(id) ? await store.FindAsync(id!, cancellationToken) : null;

            if (existing is null || !CollectionFlows.IsOwnedBy(existing, owner))
                return CollectionOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection);

            if (!existing.IsDeleted)
                await store.MarkDeletedAsync(existing.Id, CollectionDeletedReasons.Owner, now, cancellationToken);

            return CollectionOutcome.Removed();
        }

        // ###########################################################################################
        // POST /collections/<id>/delete - THE WEBSITE'S DELETE IS THE ADMINISTRATOR'S ALONE (step 4): whoever shared a collection
        // deletes it in CRT, whose window would otherwise not know of a delete made elsewhere until its next send. Anybody else
        // is refused (Forbidden) before anything is looked up. Deleting what is deleted already is done.
        // ###########################################################################################
        public static async Task<CollectionOutcome> DeleteFromWebAsync(
            AccountRecord? account,
            string? id,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CollectionOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            if (!account.IsAdministrator)
                return CollectionOutcome.Refused(CollectionVerdict.Forbidden, CollectionFlows.OnlyTheAdministratorDeletesHere);

            CollectionRecord? existing = CollectionRules.IsId(id) ? await store.FindAsync(id!, cancellationToken) : null;

            if (existing is null)
                return CollectionOutcome.Refused(CollectionVerdict.NotFound, "There is no collection with that id.");

            if (!existing.IsDeleted)
                await store.MarkDeletedAsync(existing.Id, CollectionDeletedReasons.Administrator, now, cancellationToken);

            return CollectionOutcome.Removed();
        }

        // ###########################################################################################
        // PUT /api/collections/<id>/sharing - ITS OWNER STOPS SHARING IT, OR SHARES IT AGAIN (owner decision, 2026-11-27: "Offline,
        // same link back"; migration 0028). Stopped, its page is offline - its link says it is not shared at the moment, its
        // captures are kept, its deletion date keeps counting; shared again, it is back under the same link. Its key's alone:
        // any other is "not found". Doing what is done already is done.
        // ###########################################################################################
        public static async Task<CollectionOutcome> SetSharedAsync(
            string? ownerKey,
            string? id,
            bool shared,
            ICollectionStore store,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            (CollectionRecord? existing, _, CollectionOutcome? refused) = await CollectionFlows.OwnedAsync(ownerKey, id, store, now, cancellationToken);

            if (refused is not null)
                return refused;

            if (!await store.SetUnsharedAsync(existing!.Id, shared ? null : now, cancellationToken))
                return CollectionOutcome.Refused(CollectionVerdict.Gone, "That collection was deleted meanwhile.");

            return CollectionOutcome.Done(CollectionFlows.AnswerOf(existing, siteBaseUrl));
        }

        // ###########################################################################################
        // DELETE /api/collections/<id>/images/<sha256> - ITS OWNER DELETES ONE CAPTURE, AND THE PAGE SHOWS IT NO MORE AT ONCE
        // (owner request, 2026-11-28: "The "Delete image" should also remove it right away from the shared page, and do not touch
        // the expiry date. This date should only change when the USER does something"; server 16.1.0). Nothing else changes:
        // the deletion date, whether it is shared, its words and its other captures stay as the last share left them - so a
        // change typed in CRT and not shared yet stays CRT's until "Share". The last capture may go too: the page then says it
        // has none. Deleting what is not there is done. Its key's alone - any other "not found".
        // ###########################################################################################
        public static async Task<CollectionOutcome> RemoveImageAsync(
            string? ownerKey,
            string? id,
            string? sha256,
            ICollectionStore store,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            (CollectionRecord? existing, _, CollectionOutcome? refused) = await CollectionFlows.OwnedAsync(ownerKey, id, store, now, cancellationToken);

            if (refused is not null)
                return refused;

            if (!DataTreeLayout.IsSha256(sha256))
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.NotACaptureAddress);

            if (!await store.RemoveImageAsync(existing!.Id, sha256!, now, cancellationToken))
                return CollectionOutcome.Refused(CollectionVerdict.Gone, "That collection was deleted meanwhile.");

            return CollectionOutcome.Done(CollectionFlows.AnswerOf(
                existing with { Images = existing.Images.Where(image => !string.Equals(image.Entry.Sha256, sha256, StringComparison.Ordinal)).ToList() },
                siteBaseUrl));
        }

        // ###########################################################################################
        // THE HOURLY SWEEP (CollectionSweeper): collections whose time is up are deleted, and a picture no collection of its key
        // shows any more goes a day after its upload - a day, so a capture uploaded for a collection about to be sent is not
        // taken from under it. (Nobody is reminded of the date any more: with no account there is nobody to mail.)
        // ###########################################################################################
        public static async Task<CollectionSweepResult> SweepAsync(
            ICollectionStore store,
            CollectionFiles files,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            int expired = 0;

            foreach (string id in await store.ListExpiredAsync(now, cancellationToken))
            {
                if (await store.MarkDeletedAsync(id, CollectionDeletedReasons.Expired, now, cancellationToken))
                    expired++;
            }

            int filesDeleted = 0;
            var filesNotDeleted = new List<string>();

            // *** THE UPLOADS LET GO OF AND THEIR BYTES DELETED UNDER THE PICTURES' LOCK, AS ONE STEP (code review, 2026-11-27). ***
            // A share finds its pictures and keeps its collection under the same lock (CreateAsync, UpdateAsync), an upload its
            // bytes and row: so none comes between the sweep finding an upload unused and its bytes going.
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            try
            {
                // The rows are gone once this answers, so nothing would ever name these files again: one that cannot be deleted
                // is said and the rest still go (code review, 2026-11-27: the first refusal stopped the loop, and every file
                // after it stayed on disk for good, counted by nothing against CollectionsMaxStoredBytes). The loop is FINISHED
                // whatever comes - the service stopping (no token: a few quick steps a picture), the database failing one ask
                // (that picture named, the rest still dealt with).
                foreach (string sha256 in await store.DeleteUnusedUploadsAsync(now - CollectionFlows.UnusedUploadKept, cancellationToken))
                {
                    try
                    {
                        // Another key may have an upload of the very same picture - its upload found the bytes here and wrote
                        // none.
                        if (await store.IsUploadedAsync(sha256, CancellationToken.None))
                            continue;

                        if (files.Delete(sha256))
                            filesDeleted++;
                    }
                    catch (Exception)
                    {
                        filesNotDeleted.Add(sha256);
                    }
                }
            }
            finally
            {
                CollectionFlows.PicturesLock.Release();
            }

            return new CollectionSweepResult(expired, filesDeleted, filesNotDeleted.Count == 0 ? null : filesNotDeleted);
        }

        // ###########################################################################################
        // A collection as a reader may see it: none for an id of the wrong shape or none at all; a deleted one only when asked for
        // (its page says when it went). One whose sharing was stopped is the ADMINISTRATOR'S to see still, said so (`Offline`) -
        // to everybody else only that it is not shared at the moment (`IsWithheld`). `viewer` is asked only then - a picture's
        // every request needs no session read for a collection shown to all.
        // ###########################################################################################
        public static async Task<ShownCollection?> FindShownToAsync(
            string? id,
            Func<CancellationToken, Task<AccountRecord?>>? viewer,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken,
            bool evenIfDeleted = false) =>
            CollectionRules.IsId(id)
                ? await CollectionFlows.ShownToAsync(await store.FindAsync(id!, cancellationToken), viewer, now, cancellationToken, evenIfDeleted)
                : null;

        // ###########################################################################################
        // ONE PICTURE OF A COLLECTION, BY THE SAME RULE (code review, 2026-11-27): the collection read with that capture's rows
        // alone (ICollectionStore.FindWithImageAsync) - a picture's every request read all of them, each view of a page of 300
        // captures some 90,000 rows. Null when the collection is not shown to the reader, or shows no such capture.
        // ###########################################################################################
        public static async Task<ShownCollection?> FindShownImageToAsync(
            string? id,
            string? sha256,
            Func<CancellationToken, Task<AccountRecord?>>? viewer,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (!CollectionRules.IsId(id) || !DataTreeLayout.IsSha256(sha256))
                return null;

            ShownCollection? shown = await CollectionFlows.ShownToAsync(
                await store.FindWithImageAsync(id!, sha256!, cancellationToken), viewer, now, cancellationToken, evenIfDeleted: false);

            // A capture, or a capture's pinout drawing (server 16.0.0).
            return shown is { IsWithheld: false } &&
                   shown.Collection.Images.Any(image => image.Entry.Sha256 == sha256 || image.Entry.PinoutSha256 == sha256)
                ? shown
                : null;
        }

        private static async Task<ShownCollection?> ShownToAsync(
            CollectionRecord? collection,
            Func<CancellationToken, Task<AccountRecord?>>? viewer,
            DateTimeOffset now,
            CancellationToken cancellationToken,
            bool evenIfDeleted)
        {
            // Its time up is DELETED from that moment - the hourly sweep only makes it so in the database.
            if (collection is not null && !collection.IsDeleted && CollectionFlows.IsGone(collection, now))
                collection = collection with { DeletedUtc = collection.DeleteAtUtc, DeletedReason = CollectionDeletedReasons.Expired };

            if (collection is null || (collection.IsDeleted && !evenIfDeleted))
                return null;

            if (collection.IsDeleted || !collection.IsSharingStopped)
                return new ShownCollection(collection, CollectionOffline.No);

            // ###########################################################################################
            // ITS OWNER STOPPED SHARING IT (migration 0028; owner decision, 2026-11-27: "Offline, same link back"): the page to the
            // administrator, said so - to everybody else only that it is not shared at the moment (`IsWithheld`), never its
            // captures or words. The reader is asked only here.
            // ###########################################################################################
            bool administrator = viewer is not null && await viewer(cancellationToken) is { IsAdministrator: true };

            return new ShownCollection(collection, CollectionOffline.SharingStopped, IsWithheld: !administrator);
        }

        // ###########################################################################################
        // *** A COLLECTION IS GONE ONCE DELETED - OR ONCE ITS TIME IS UP, before the hourly sweep makes it so (code review,
        // 2026-11-27). *** The one question the page, "Share", and the administrator's list ask, so none of them shows for an
        // hour what another says is gone.
        // ###########################################################################################
        public static bool IsGone(CollectionRecord collection, DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(collection);

            return collection.IsDeleted || collection.DeleteAtUtc <= now;
        }

        // "Deleted on 2026-November-26" - said by the routes and the page.
        public static string DeletedLine(CollectionRecord collection) =>
            $"That collection was deleted on {CollectionRules.DateText(collection.DeletedUtc ?? collection.DeleteAtUtc)}.";

        // Whether `ownerKeyHash` owns the collection - never one owned by no key (shared with an account before server 17.0.0).
        public static bool IsOwnedBy(CollectionRecord collection, string ownerKeyHash)
        {
            ArgumentNullException.ThrowIfNull(collection);

            return collection.OwnerKeyHash is { } kept &&
                   CryptographicOperations.FixedTimeEquals(Encoding.ASCII.GetBytes(kept), Encoding.ASCII.GetBytes(ownerKeyHash));
        }

        // The collection `id` that the key owns, still shown - or why not: no key, another key's (or none) "not found", gone.
        private static async Task<(CollectionRecord? Collection, string? Owner, CollectionOutcome? Refused)> OwnedAsync(
            string? ownerKey,
            string? id,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.OwnerHashOf(ownerKey) is not { } owner)
                return (null, null, CollectionOutcome.Refused(CollectionVerdict.BadRequest, CollectionFlows.NoKey));

            CollectionRecord? existing = CollectionRules.IsId(id) ? await store.FindAsync(id!, cancellationToken) : null;

            if (existing is null || !CollectionFlows.IsOwnedBy(existing, owner))
                return (null, null, CollectionOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection));

            // Gone - its time up too, as its page says: CRT then makes a new page, never brings the old one back.
            if (CollectionFlows.IsGone(existing, now))
                return (null, null, CollectionOutcome.Refused(CollectionVerdict.Gone, CollectionFlows.DeletedLine(existing)));

            return (existing, owner, null);
        }

        // The captures as kept - each with its picture's size from its upload - or the refusal naming the pictures the server
        // does not hold for this key (uploaded by it, the bytes still here): the captures, and their pinout drawings.
        private static async Task<(IReadOnlyList<CollectionImageRecord>? Images, CollectionOutcome? Missing)> ImagesOfAsync(
            string owner,
            CollectionContentRequest content,
            ICollectionStore store,
            CollectionFiles files,
            CancellationToken cancellationToken)
        {
            string[] wanted = content.Images
                .SelectMany(image => image.PinoutSha256 is { } pinout ? new[] { image.Sha256, pinout } : [image.Sha256])
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            IReadOnlyDictionary<string, CollectionUploadRecord> uploads =
                await store.FindUploadsAsync(owner, wanted, cancellationToken);

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
            string owner,
            CollectionContentRequest content,
            IReadOnlyList<CollectionImageRecord> images,
            DateTimeOffset createdUtc,
            DateTimeOffset now) =>
            new(
                id,
                owner,
                content.HardwareName.Trim(),
                content.BoardName.Trim(),
                CollectionFlows.Blank(content.Region),
                CollectionFlows.Blank(content.Title),
                CollectionFlows.Blank(content.Description),
                (content.ForumLinks ?? []).Select(link => link.Trim()).ToList(),
                createdUtc,
                now,
                CollectionFlows.WholeMilliseconds(content.DeleteAtUtc.ToUniversalTime()),
                DeletedUtc: null,
                DeletedReason: null,
                images,
                images.Count,
                Components: (content.Components ?? []).Select(CollectionFlows.Tidied).ToList(),
                Oscilloscopes: (content.Oscilloscopes ?? []).Select(oscilloscope => oscilloscope.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList());

        private static CollectionComponent Tidied(CollectionComponent component) =>
            new(
                component.Label.Trim(),
                CollectionFlows.Blank(component.Title),
                CollectionFlows.Blank(component.CategoryLine),
                CollectionFlows.Blank(component.Description),
                CollectionFlows.Blank(component.Purpose),
                (component.Pins ?? [])
                    .Select(pin => new CollectionPin(
                        pin.Pin.Trim(),
                        CollectionFlows.Blank(pin.Name),
                        CollectionFlows.Blank(pin.Direction),
                        CollectionFlows.Blank(pin.Net),
                        CollectionFlows.Blank(pin.Reading)))
                    .ToList());

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
                PinoutSha256 = CollectionFlows.Blank(image.PinoutSha256)
            };

        private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        // ###########################################################################################
        // The deletion date AS THE DATABASE KEEPS IT - whole milliseconds, its DATETIME(3) (code review, 2026-11-27). Kept to
        // the tick, the same date shared again compared as LATER than the one stored - which MariaDB had rounded.
        // ###########################################################################################
        internal static DateTimeOffset WholeMilliseconds(DateTimeOffset when) =>
            new(when.UtcTicks - (when.UtcTicks % TimeSpan.TicksPerMillisecond), TimeSpan.Zero);

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
        CollectionImageUploadAnswer? Upload = null)
    {
        public static CollectionOutcome Refused(CollectionVerdict verdict, string error) => new(verdict, [error]);

        public static CollectionOutcome Refused(CollectionVerdict verdict, IReadOnlyList<string> errors) => new(verdict, errors);

        public static CollectionOutcome MissingImages(IReadOnlyList<string> missing, string error) =>
            new(CollectionVerdict.MissingImages, [error], Missing: missing);

        public static CollectionOutcome Created(CollectionAnswer answer) => new(CollectionVerdict.Created, [], Answer: answer);

        public static CollectionOutcome Done(CollectionAnswer answer) => new(CollectionVerdict.Done, [], Answer: answer);

        public static CollectionOutcome Uploaded(CollectionImageUploadAnswer upload) => new(CollectionVerdict.Done, [], Upload: upload);

        public static CollectionOutcome Removed() => new(CollectionVerdict.Removed, []);
    }

    // `FilesNotDeleted`: the pictures nobody shows any more that could not be deleted - named in the log, to delete by hand.
    // Null when every one went.
    public sealed record CollectionSweepResult(int Expired, int FilesDeleted, IReadOnlyList<string>? FilesNotDeleted = null);

    // A collection a reader may see, and whether it is offline for everybody else (FindShownToAsync) - only ever so for the
    // administrator, its sharing stopped. `IsWithheld`: the reader may learn only that it is not shared at the moment - nothing
    // of it is shown.
    public sealed record ShownCollection(CollectionRecord Collection, CollectionOffline Offline, bool IsWithheld = false)
    {
        public bool IsOffline => this.Offline != CollectionOffline.No;
    }

    public enum CollectionOffline
    {
        No,

        // Its owner pressed "Stop sharing" in CRT (migration 0028).
        SharingStopped
    }
}
