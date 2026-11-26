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
    //     for everybody else.
    //   - A comment not the account's to touch is "not found", as a collection is.
    // ###########################################################################################
    public static class CollectionCommentFlows
    {
        public const string NoSuchComment = "There is no comment of yours with that id.";

        public const string NoSuchCapture = "That capture is not in the collection.";

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

            if (await CollectionFlows.FindShownAsync(collectionId, store, accounts, cancellationToken) is not { } collection)
                return CommentOutcome.Refused(CollectionVerdict.NotFound, "There is no such collection.");

            if (!collection.Images.Any(image => image.Entry.Sha256 == imageSha256))
                return CommentOutcome.Refused(CollectionVerdict.BadRequest, CollectionCommentFlows.NoSuchCapture);

            if (CollectionComments.Problem(body) is { } problem)
                return CommentOutcome.Refused(CollectionVerdict.BadRequest, problem);

            if (await store.CountCommentsSinceAsync(account!.Id, now - TimeSpan.FromDays(1), cancellationToken) >= CollectionComments.MaximumPerDay)
            {
                return CommentOutcome.Refused(
                    CollectionVerdict.TooMany,
                    $"An account writes at most {CollectionComments.MaximumPerDay} comments a day - write the rest tomorrow.");
            }

            string text = CollectionComments.Tidied(body!);
            long id = await store.AddCommentAsync(collection.Id, imageSha256!, account.Id, text, now, cancellationToken);

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

        public static async Task<CommentOutcome> EditAsync(
            AccountRecord? account,
            long commentId,
            string? body,
            ICollectionStore store,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (CollectionCommentFlows.Refuse(account) is { } refused)
                return refused;

            if (await store.FindCommentAsync(commentId, cancellationToken) is not { } comment || comment.AccountId != account!.Id ||
                await CollectionFlows.FindShownAsync(comment.CollectionId, store, accounts, cancellationToken) is null)
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
            long commentId,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            if (await store.FindCommentAsync(commentId, cancellationToken) is not { } comment ||
                (comment.AccountId != account.Id && !account.IsAdministrator))
            {
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionCommentFlows.NoSuchComment);
            }

            await store.DeleteCommentAsync(commentId, now, cancellationToken);

            return CommentOutcome.Written(commentId);
        }

        // The collection's owner's, or an administrator's.
        public static async Task<CommentOutcome> SetHiddenAsync(
            AccountRecord? account,
            long commentId,
            bool hidden,
            ICollectionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            if (await store.FindCommentAsync(commentId, cancellationToken) is not { } comment ||
                await store.FindAsync(comment.CollectionId, cancellationToken) is not { IsDeleted: false } collection ||
                (collection.AccountId != account.Id && !account.IsAdministrator))
            {
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionCommentFlows.NoSuchComment);
            }

            await store.SetCommentHiddenAsync(commentId, hidden, now, cancellationToken);

            return CommentOutcome.Written(commentId);
        }

        // ###########################################################################################
        // A collection's comments for its owner's CRT - hidden ones too, each saying whose it is.
        // ###########################################################################################
        public static async Task<CommentOutcome> ListForOwnerAsync(
            AccountRecord? account,
            string? collectionId,
            ICollectionStore store,
            CancellationToken cancellationToken)
        {
            if (account is null)
                return CommentOutcome.Refused(CollectionVerdict.Unauthorized, "Sign in first.");

            CollectionRecord? collection = CollectionRules.IsId(collectionId) ? await store.FindAsync(collectionId!, cancellationToken) : null;

            if (collection is null || collection.IsDeleted || (collection.AccountId != account.Id && !account.IsAdministrator))
                return CommentOutcome.Refused(CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection);

            IReadOnlyList<CollectionCommentRecord> comments = await store.ListCommentsAsync(collection.Id, cancellationToken);

            return CommentOutcome.Listed(new CollectionCommentsAnswer(comments
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
                    comment.IsHidden))
                .ToList()));
        }

        // ###########################################################################################
        // The comments a viewer of the page sees: every one not hidden - and hidden ones too for the collection's owner and an
        // administrator, who may show them again.
        // ###########################################################################################
        public static IReadOnlyList<CollectionCommentRecord> ShownTo(IReadOnlyList<CollectionCommentRecord> comments, CollectionRecord collection, AccountRecord? viewer) =>
            viewer is not null && (viewer.Id == collection.AccountId || viewer.IsAdministrator)
                ? comments
                : comments.Where(comment => !comment.IsHidden).ToList();

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
