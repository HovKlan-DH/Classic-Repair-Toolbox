using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // EVERY DECISION ABOUT A COMMENT ON A SHARED COLLECTION (step 3, 2026-11-26; CRT.Data's CollectionComments says the rules).
    // Pure as CollectionFlows is: stores and mailer handed in.
    //
    //   - Writing: any account with a confirmed address, on a collection shown (CollectionFlows.FindShownAsync), on one of
    //     its captures, within the day's limit. The owner is MAILED for every comment but their own (owner decision,
    //     2026-11-26: "A mail per comment"); nobody else is.
    //   - Editing: its writer's alone, marked edited. Deleting: its writer's, or an administrator's.
    //   - Hiding and showing again: the collection's owner's, or an administrator's - a hidden comment is gone from the page
    //     for everybody else. A locked account's comments are the administrator's alone to see (ShownTo).
    //   - A comment not the account's to touch is "not found", as a collection is.
    // ###########################################################################################
    public static class CollectionCommentFlows
    {
        public const string NoSuchComment = "There is no comment of yours with that id.";

        public const string NoSuchCapture = "That capture is not in the collection.";

        // A comment's count of the day and its write, one at a time (AddAsync).
        private static readonly SemaphoreSlim WritingLock = new(1, 1);

        public static async Task<CommentOutcome> AddAsync(
            AccountRecord? account,
            string? collectionId,
            string? imageSha256,
            string? body,
            ICollectionStore store,
            IAccountStore accounts,
            IEmailSender mailer,
            string siteBaseUrl,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionCommentFlows.Refuse(account) is { } refused)
                return refused;

            if (await CollectionFlows.FindShownAsync(collectionId, store, accounts, now, cancellationToken) is not { } collection)
                return CommentOutcome.Refused(CollectionVerdict.NotFound, "There is no such collection.");

            if (!collection.Images.Any(image => image.Entry.Sha256 == imageSha256))
                return CommentOutcome.Refused(CollectionVerdict.BadRequest, CollectionCommentFlows.NoSuchCapture);

            if (CollectionComments.Problem(body) is { } problem)
                return CommentOutcome.Refused(CollectionVerdict.BadRequest, problem);

            string text = CollectionComments.Tidied(body!);
            long id;

            // The day's count and the comment kept as ONE step (code review, 2026-11-27: counted apart, comments posted at the same
            // moment all passed the limit, each mailing the owner). The service is one process, so the lock is the process's.
            await CollectionCommentFlows.WritingLock.WaitAsync(cancellationToken);

            try
            {
                if (await store.CountCommentsSinceAsync(account!.Id, now - TimeSpan.FromDays(1), cancellationToken) >= CollectionComments.MaximumPerDay)
                {
                    return CommentOutcome.Refused(
                        CollectionVerdict.TooMany,
                        $"An account writes at most {CollectionComments.MaximumPerDay} comments a day - write the rest tomorrow.");
                }

                id = await store.AddCommentAsync(collection.Id, imageSha256!, account.Id, text, now, cancellationToken);
            }
            finally
            {
                CollectionCommentFlows.WritingLock.Release();
            }

            // The owner hears of every comment but their own.
            if (account.Id != collection.AccountId &&
                await accounts.FindByIdAsync(collection.AccountId, cancellationToken) is { IsVerified: true, IsLocked: false } owner)
            {
                CollectionImageEntry capture = collection.Images.First(image => image.Entry.Sha256 == imageSha256).Entry;

                await mailer.SendAsync(
                    EmailTemplates.CollectionCommented(
                        owner.Email,
                        owner.DisplayName,
                        CollectionPages.NameOf(collection),
                        account.DisplayName,
                        CollectionRules.WhereOf(capture),
                        text,
                        CollectionCommentFlows.CommentUrl(siteBaseUrl, collection.Id, id)),
                    cancellationToken);
            }

            return CommentOutcome.Written(id);
        }

        // A comment of the collection the route names, or none (code review, 2026-11-27: "/collections/<A>/comments/<c>" acted on
        // a comment of collection B) - only an id of the right shape names one.
        private static async Task<CollectionCommentRecord?> FindOfAsync(string? collectionId, long commentId, ICollectionStore store, CancellationToken cancellationToken) =>
            CollectionRules.IsId(collectionId) && await store.FindCommentAsync(commentId, cancellationToken) is { } comment &&
            string.Equals(comment.CollectionId, collectionId, StringComparison.Ordinal)
                ? comment
                : null;

        public static async Task<CommentOutcome> EditAsync(
            AccountRecord? account,
            string? collectionId,
            long commentId,
            string? body,
            ICollectionStore store,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionCommentFlows.Refuse(account) is { } refused)
                return refused;

            // A hidden comment is gone for its writer too (ShownTo) - and not to be rewritten into something harmless to be shown
            // again, then back (code review, 2026-11-27).
            if (await CollectionCommentFlows.FindOfAsync(collectionId, commentId, store, cancellationToken) is not { } comment || comment.AccountId != account!.Id ||
                comment.IsHidden ||
                await CollectionFlows.FindShownAsync(comment.CollectionId, store, accounts, now, cancellationToken) is null)
            {
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionCommentFlows.NoSuchComment);
            }

            if (CollectionComments.Problem(body) is { } problem)
                return CommentOutcome.Refused(CollectionVerdict.BadRequest, problem);

            await store.EditCommentAsync(commentId, CollectionComments.Tidied(body!), now, cancellationToken);

            return CommentOutcome.Written(commentId);
        }

        // Its writer's, or an administrator's. Deleting what is gone already is "not found" - there is nothing of it left.
        public static async Task<CommentOutcome> DeleteAsync(
            AccountRecord? account,
            string? collectionId,
            long commentId,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            // Of a collection still there - one whose time is up is gone, its comments with it (CollectionFlows.IsGone).
            if (await CollectionCommentFlows.FindOfAsync(collectionId, commentId, store, cancellationToken) is not { } comment ||
                (comment.AccountId != account.Id && !account.IsAdministrator) ||
                await store.FindAsync(comment.CollectionId, cancellationToken) is not { } collection || CollectionFlows.IsGone(collection, now))
            {
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionCommentFlows.NoSuchComment);
            }

            await store.DeleteCommentAsync(commentId, now, cancellationToken);

            return CommentOutcome.Written(commentId);
        }

        // The collection's owner's, or an administrator's.
        public static async Task<CommentOutcome> SetHiddenAsync(
            AccountRecord? account,
            string? collectionId,
            long commentId,
            bool hidden,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            if (await CollectionCommentFlows.FindOfAsync(collectionId, commentId, store, cancellationToken) is not { } comment ||
                await store.FindAsync(comment.CollectionId, cancellationToken) is not { } collection || CollectionFlows.IsGone(collection, now) ||
                (collection.AccountId != account.Id && !account.IsAdministrator))
            {
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionCommentFlows.NoSuchComment);
            }

            await store.SetCommentHiddenAsync(commentId, hidden, now, cancellationToken);

            return CommentOutcome.Written(commentId);
        }

        // ###########################################################################################
        // A collection's comments for its owner's CRT - hidden ones too, each saying whose it is; a locked account's to an
        // administrator alone (AccountLockFlows), as on the page, and said so (AuthorLocked), as the page flags it.
        // ###########################################################################################
        public static async Task<CommentOutcome> ListForOwnerAsync(
            AccountRecord? account,
            string? collectionId,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            CollectionRecord? collection = CollectionRules.IsId(collectionId) ? await store.FindAsync(collectionId!, cancellationToken) : null;

            if (collection is null || (collection.AccountId != account.Id && !account.IsAdministrator))
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection);

            // Taken down - by its time, or by the administrator on the website: said to its owner as gone, and when, so CRT
            // forgets the page (code review, 2026-11-27). Its time up is gone from that moment, as on the page (CollectionFlows
            // .IsGone).
            if (CollectionFlows.IsGone(collection, now))
                return CommentOutcome.Refused(CollectionVerdict.Gone, CollectionFlows.DeletedLine(collection));

            IReadOnlyList<CollectionCommentRecord> comments = await store.ListCommentsAsync(collection.Id, cancellationToken);

            return CommentOutcome.Listed(new CollectionCommentsAnswer(comments
                .Where(comment => account.IsAdministrator || !comment.AuthorIsLocked)
                .Select(comment => new CollectionCommentAnswer(
                    comment.Id,
                    comment.ImageSha256,
                    comment.AuthorName,
                    ByOwner: comment.AccountId == collection.AccountId,
                    ByAdministrator: comment.AuthorIsAdministrator,
                    Mine: comment.AccountId == account.Id,
                    comment.Body,
                    comment.CreatedUtc,
                    comment.EditedUtc,
                    comment.IsHidden,
                    comment.AuthorIsLocked))
                .ToList()));
        }

        // ###########################################################################################
        // The comments a viewer of the page sees: every one not hidden - and hidden ones too for the collection's owner and an
        // administrator, who may show them again. A LOCKED account's are the administrator's alone (owner decision, 2026-11-27,
        // locking from the website - AccountLockFlows): locking a spammer takes the spam down with it, and unlocking puts it back.
        // ###########################################################################################
        public static IReadOnlyList<CollectionCommentRecord> ShownTo(IReadOnlyList<CollectionCommentRecord> comments, CollectionRecord collection, AccountRecord? viewer) =>
            viewer is { IsAdministrator: true }
                ? comments
                : comments
                    .Where(comment => !comment.AuthorIsLocked && (!comment.IsHidden || (viewer is not null && viewer.Id == collection.AccountId)))
                    .ToList();

        public static string CommentUrl(string siteBaseUrl, string collectionId, long commentId) =>
            $"{CollectionRules.PageUrl(siteBaseUrl, collectionId)}#comment-{commentId}";

        private static CommentOutcome? Refuse(AccountRecord? account)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            return account.IsVerified ? null : CommentOutcome.Refused(CollectionVerdict.Forbidden, CollectionFlows.AddressNotConfirmed);
        }
    }

    public sealed record CommentOutcome(
        CollectionVerdict Verdict,
        IReadOnlyList<string> Errors,
        long? CommentId = null,
        CollectionCommentsAnswer? Comments = null)
    {
        public bool IsDone => this.Verdict == CollectionVerdict.Done;

        public static CommentOutcome Refused(CollectionVerdict verdict, string error) => new(verdict, [error]);

        public static CommentOutcome Written(long id) => new(CollectionVerdict.Done, [], CommentId: id);

        public static CommentOutcome Listed(CollectionCommentsAnswer comments) => new(CollectionVerdict.Done, [], Comments: comments);
    }
}
