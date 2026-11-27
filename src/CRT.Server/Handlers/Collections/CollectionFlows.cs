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

            // The look, the day's count, the bytes and the row under ONE lock (code review, 2026-11-27: looked at first, a sweep
            // between the look and the lock took the account's old row and the bytes, the upload wrote the bytes and no row - a
            // file nothing named, never swept - and parallel uploads all passed the day's limit together).
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            try
            {
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

                // A picture sent again has its upload's time moved to now, so the day the sweep leaves an unused upload counts
                // from its LAST send - an old row no collection shows would go at the next sweep, before the collection it was
                // sent again for is shared.
                await store.AddUploadAsync(
                    new CollectionUploadRecord(account.Id, sha256!, bytes.Length, picture.Width, picture.Height, now),
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
        // look and its delete, an upload of the same picture by another account lost its bytes for good - nothing asks for them
        // again, as its row says it has them. So the lock is held by an upload (its look, write and row), a share (finding its
        // pictures and keeping its collection - CreateAsync, UpdateAsync) and the sweep (letting go of unused uploads and
        // deleting their bytes). The service is one process (the test instance another, with folders of its own), so the lock
        // is the process's.
        // ###########################################################################################
        private static readonly SemaphoreSlim PicturesLock = new(1, 1);

        // ###########################################################################################
        // POST /api/collections - a collection shared for the first time.
        // ###########################################################################################
        // ###########################################################################################
        // *** A FIRST SHARE SENT AGAIN UNDER ITS KEY CHANGES THE PAGE IT MADE (code review, 2026-11-27; migration 0027). *** CRT
        // sends the collection's own local id (`clientKey`, the Idempotency-Key header). A share the server kept whose answer
        // was lost made a second page at the next "Share", the first one's id never reaching CRT - so a collection of the
        // account still shown under that key is the one changed, as PUT changes it, never a new one. A key of another shape
        // is no key.
        // ###########################################################################################
        public static async Task<CollectionOutcome> CreateAsync(
            AccountRecord? account,
            CollectionContentRequest? content,
            string? clientKey,
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

            string? key = CollectionRules.IsClientKey(clientKey) ? clientKey : null;

            // The key looked up, the account's limits counted, the pictures found and the collection kept under the pictures'
            // lock, as ONE step: the sweep lets go of an upload no collection shows, and between a look and the write it took one
            // this collection shows; and looked at before the lock, shares sent at the same moment all passed the limits, and a
            // first share retried while the first was being written found no page under its key and made a second (code reviews,
            // 2026-11-27).
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            bool locked = true;

            try
            {
                // One whose time is up is no page to change - a new one is made. The page made before is changed as PUT changes
                // it, which takes the lock itself - so it is let go of first; the page is there, so a share retried meanwhile
                // finds it too.
                if (key is not null && await store.FindLiveByClientKeyAsync(account!.Id, key, cancellationToken) is { } madeBefore &&
                    !CollectionFlows.IsGone(madeBefore, now))
                {
                    CollectionFlows.PicturesLock.Release();
                    locked = false;

                    return await CollectionFlows.UpdateAsync(account, madeBefore.Id, content, store, files, siteBaseUrl, now, cancellationToken);
                }

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
                    var collection = CollectionFlows.RecordOf(CollectionRules.NewId(), account.Id, content!, images!, now, now) with { ClientKey = key };

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

            // Gone - its time up too, as its page says: CRT then makes a new page, never brings the old one back.
            if (CollectionFlows.IsGone(existing, now))
                return CollectionOutcome.Refused(CollectionVerdict.Gone, CollectionFlows.DeletedLine(existing));

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, now);

            if (problems.Count > 0)
                return CollectionOutcome.Refused(CollectionVerdict.BadRequest, problems);

            // As CreateAsync: the pictures found and the collection kept under the pictures' lock.
            await CollectionFlows.PicturesLock.WaitAsync(cancellationToken);

            try
            {
                (IReadOnlyList<CollectionImageRecord>? images, CollectionOutcome? missing) =
                    await CollectionFlows.ImagesOfAsync(account.Id, content!, store, files, cancellationToken);

                if (missing is not null)
                    return missing;

                // Its RemindedUtc is what a new date would make of it - the store keeps the one decided before while the date stays.
                CollectionRecord changed = CollectionFlows.RecordOf(existing.Id, account.Id, content!, images!, existing.CreatedUtc, now);

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
        // PUT /api/collections/<id>/sharing - ITS OWNER STOPS SHARING IT, OR SHARES IT AGAIN (owner decision, 2026-11-27: "Offline,
        // same link back"; migration 0028). Stopped, its page is offline - its link says it is not shared at the moment, its
        // captures and comments are kept, its deletion date keeps counting; shared again, it is back under the same link. The
        // owner's alone, as sharing it is: anybody else's is "not found". Doing what is done already is done.
        // ###########################################################################################
        public static async Task<CollectionOutcome> SetSharedAsync(
            AccountRecord? account,
            string? id,
            bool shared,
            ICollectionStore store,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionFlows.Refuse(account) is { } refused)
                return refused;

            CollectionRecord? existing = CollectionRules.IsId(id) ? await store.FindAsync(id!, cancellationToken) : null;

            if (existing is null || existing.AccountId != account!.Id)
                return CollectionOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection);

            if (CollectionFlows.IsGone(existing, now))
                return CollectionOutcome.Refused(CollectionVerdict.Gone, CollectionFlows.DeletedLine(existing));

            if (!await store.SetUnsharedAsync(existing.Id, shared ? null : now, cancellationToken))
                return CollectionOutcome.Refused(CollectionVerdict.Gone, "That collection was deleted meanwhile.");

            return CollectionOutcome.Done(CollectionFlows.AnswerOf(existing, siteBaseUrl));
        }

        public const string OnlyTheAdministratorDeletesHere =
            "Only the administrator deletes a collection on the website - its owner deletes it in CRT.";

        // ###########################################################################################
        // POST /collections/<id>/delete - THE WEBSITE'S DELETE IS THE ADMINISTRATOR'S ALONE (step 4): the owner deletes their own
        // in CRT, whose window would otherwise not know of a delete made elsewhere until its next send. Anybody else is refused
        // (Forbidden) before anything is looked up; the administrator's is DeleteAsync's.
        // ###########################################################################################
        public static Task<CollectionOutcome> DeleteFromWebAsync(
            AccountRecord? account,
            string? id,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return Task.FromResult(CollectionOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first."));

            return account.IsAdministrator
                ? CollectionFlows.DeleteAsync(account, id, store, now, cancellationToken)
                : Task.FromResult(CollectionOutcome.Refused(CollectionVerdict.Forbidden, CollectionFlows.OnlyTheAdministratorDeletesHere));
        }

        // ###########################################################################################
        // GET /api/collections/mine.
        // ###########################################################################################
        public static async Task<CollectionOutcome> ListMineAsync(
            AccountRecord? account,
            ICollectionStore store,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CollectionOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            IReadOnlyList<CollectionRecord> live = await store.ListLiveForAccountAsync(account.Id, cancellationToken);

            return CollectionOutcome.Listed(new MyCollectionsAnswer(
                live.Where(collection => !CollectionFlows.IsGone(collection, now)).Select(collection => new CollectionSummaryAnswer(
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

                // Marked either way: an owner gone or locked is reminded of nothing, and the next sweep does not ask again. One
                // kept a month or so is never listed here - it was marked as its date was set (ReminderSettledAt).
                if (owner is { IsVerified: true, IsLocked: false })
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
                        // Another account may have an upload of the very same picture - its upload found the bytes here and
                        // wrote none.
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

            return new CollectionSweepResult(expired, reminded, filesDeleted, filesNotDeleted.Count == 0 ? null : filesNotDeleted);
        }

        // ###########################################################################################
        // A collection as anybody with its link may see it: none for an id of the wrong shape, none at all, or one whose owner's
        // account is locked or gone - a locked account's collections go offline with it. A deleted one only when asked for (its
        // page says when it went). The comments ask this; the page asks FindShownToAsync, which is this question with the reader,
        // and a picture FindShownImageToAsync, the same with one capture's rows read.
        // ###########################################################################################
        public static async Task<CollectionRecord?> FindShownAsync(
            string? id,
            ICollectionStore store,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken,
            bool evenIfDeleted = false) =>
            (await CollectionFlows.FindShownToAsync(id, viewer: null, store, accounts, now, cancellationToken, evenIfDeleted)) is { IsWithheld: false } shown
                ? shown.Collection
                : null;

        // ###########################################################################################
        // THE SAME QUESTION, ASKED BY A READER (code review, 2026-11-27: the administrator's "Every collection" linked a locked
        // owner's collection, and its page said "There is no such collection" - no "Delete this collection", no hiding its
        // comments): a collection offline for everybody else is the ADMINISTRATOR'S to see still, said so (`Offline`). `viewer`
        // is asked only then - a picture's every request needs no session read for a collection shown to all.
        // ###########################################################################################
        public static async Task<ShownCollection?> FindShownToAsync(
            string? id,
            Func<CancellationToken, Task<AccountRecord?>>? viewer,
            ICollectionStore store,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken,
            bool evenIfDeleted = false) =>
            CollectionRules.IsId(id)
                ? await CollectionFlows.ShownToAsync(await store.FindAsync(id!, cancellationToken), viewer, accounts, now, cancellationToken, evenIfDeleted)
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
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (!CollectionRules.IsId(id) || !DataTreeLayout.IsSha256(sha256))
                return null;

            ShownCollection? shown = await CollectionFlows.ShownToAsync(
                await store.FindWithImageAsync(id!, sha256!, cancellationToken), viewer, accounts, now, cancellationToken, evenIfDeleted: false);

            return shown is { IsWithheld: false } && shown.Collection.Images.Any(image => image.Entry.Sha256 == sha256) ? shown : null;
        }

        private static async Task<ShownCollection?> ShownToAsync(
            CollectionRecord? collection,
            Func<CancellationToken, Task<AccountRecord?>>? viewer,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken,
            bool evenIfDeleted)
        {
            // Its time up is DELETED from that moment - the hourly sweep only makes it so in the database (code review,
            // 2026-11-27: for up to an hour the page took comments under "Deleted on" a date passed).
            if (collection is not null && !collection.IsDeleted && CollectionFlows.IsGone(collection, now))
                collection = collection with { DeletedUtc = collection.DeleteAtUtc, DeletedReason = CollectionDeletedReasons.Expired };

            if (collection is null || (collection.IsDeleted && !evenIfDeleted))
                return null;

            if (collection.IsDeleted)
                return new ShownCollection(collection, CollectionOffline.No);

            AccountRecord? owner = await accounts.FindByIdAsync(collection.AccountId, cancellationToken);

            if (owner is { IsLocked: false })
            {
                if (!collection.IsSharingStopped)
                    return new ShownCollection(collection, CollectionOffline.No);

                // ###########################################################################################
                // ITS OWNER STOPPED SHARING IT (migration 0028; owner decision, 2026-11-27: "Offline, same link back"): its own
                // page to its owner and the administrator, said so - to everybody else only that it is not shared at the moment
                // (`IsWithheld`), never its captures, words or comments. The reader is asked only here.
                // ###########################################################################################
                AccountRecord? reader = viewer is null ? null : await viewer(cancellationToken);
                bool mayRead = reader is not null && (reader.Id == collection.AccountId || reader.IsAdministrator);

                return new ShownCollection(collection, CollectionOffline.SharingStopped, IsWithheld: !mayRead);
            }

            CollectionOffline offline = owner is null ? CollectionOffline.OwnerGone : CollectionOffline.OwnerLocked;

            return viewer is not null && await viewer(cancellationToken) is { IsAdministrator: true }
                ? new ShownCollection(collection, offline)
                : null;
        }

        // ###########################################################################################
        // *** A COLLECTION IS GONE ONCE DELETED - OR ONCE ITS TIME IS UP, before the hourly sweep makes it so (code review,
        // 2026-11-27). *** The one question the page, CRT's comment list and "Share", and every list of collections ask, so none
        // of them shows for an hour what another says is gone.
        // ###########################################################################################
        public static bool IsGone(CollectionRecord collection, DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(collection);

            return collection.IsDeleted || collection.DeleteAtUtc <= now;
        }

        // ###########################################################################################
        // *** WHETHER A REMINDER IS DUE IS DECIDED AS THE DELETION DATE IS SET (code review, 2026-11-27). *** A date set no more
        // than twice the reminder's notice ahead - "1 month" - is told nothing: it is marked reminded at once (`now`), so the
        // sweep passes it by. A date further ahead waits for its reminder (null). The store takes this only when the date
        // CHANGES (ReplaceContentAsync); a share keeping its date keeps what was decided - measured from the last share instead,
        // a caption changed 45 days before its date lost the reminder, and measured from the making, a year-old collection
        // given one month more was reminded the day after.
        // ###########################################################################################
        public static DateTimeOffset? ReminderSettledAt(DateTimeOffset deleteAtUtc, DateTimeOffset now) =>
            deleteAtUtc - now > CollectionRules.ReminderBefore * 2 ? null : now;

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
                CollectionFlows.WholeMilliseconds(content.DeleteAtUtc.ToUniversalTime()),
                RemindedUtc: CollectionFlows.ReminderSettledAt(CollectionFlows.WholeMilliseconds(content.DeleteAtUtc.ToUniversalTime()), now),
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

        // ###########################################################################################
        // The deletion date AS THE DATABASE KEEPS IT - whole milliseconds, its DATETIME(3) (code review, 2026-11-27). Kept to
        // the tick, the same date shared again compared as LATER than the one stored - which MariaDB had rounded - and the
        // month's reminder was mailed a second time.
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

    // `FilesNotDeleted`: the pictures nobody shows any more that could not be deleted - named in the log, to delete by hand.
    // Null when every one went.
    public sealed record CollectionSweepResult(int Expired, int Reminded, int FilesDeleted, IReadOnlyList<string>? FilesNotDeleted = null);

    // A collection a reader may see, and whether it is offline for everybody else (FindShownToAsync) - only ever so for the
    // administrator, and for its owner when they stopped sharing it. `IsWithheld`: the reader may learn only that it is not
    // shared at the moment - nothing of it is shown (its sharing stopped, and the reader neither its owner nor the
    // administrator).
    public sealed record ShownCollection(CollectionRecord Collection, CollectionOffline Offline, bool IsWithheld = false)
    {
        public bool IsOffline => this.Offline != CollectionOffline.No;
    }

    public enum CollectionOffline
    {
        No,
        OwnerLocked,
        OwnerGone,

        // Its owner pressed "Stop sharing" in CRT (migration 0028).
        SharingStopped
    }
}
