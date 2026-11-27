using Handlers.DataHandling;

namespace Handlers.Collections
{
    // ###########################################################################################
    // ONE COMMENT AS THE COLLECTION WINDOW LISTS IT (step 3): who and when - "You" for the owner's own - its text, and what the
    // owner may do with it: hide another's, show a hidden one again, delete their own. Editing is the page's. Pure - the
    // window's rows show it (CollectionWindow.axaml), so it lives here, beside CollectionSharing (code review, 2026-11-27).
    //
    // *** THE NAME, ITS BADGES AND ITS MARKS ARE SHOWN APART (code review, 2026-11-27, Security). *** Written as one line, a
    // display name such as "Dennis (Administrators)" or "Bob - NEW" read as a badge or a mark of CRT's own; the window draws
    // "You", "Administrator" and "NEW" as pills beside the name, as the web page draws its badges, so typed words never pass
    // for them.
    // ###########################################################################################
    public sealed class CollectionCommentRow
    {
        public long Id { get; init; }

        // The writer's display name as typed - or "You", shown as a pill, for the owner's own (IsMine).
        public string Author { get; init; } = string.Empty;

        public bool IsMine { get; init; }

        // "Administrator" for an administrator's comment, never the owner's own; null otherwise - a pill beside the name.
        public string? Badge { get; init; }

        public bool HasBadge => this.Badge is not null;

        // When, and what is said of it: edited, hidden from the page, its writer's account locked.
        public string Details { get; init; } = string.Empty;

        public string Body { get; init; } = string.Empty;

        public bool CanHide { get; init; }

        public bool CanShow { get; init; }

        public bool CanDelete { get; init; }

        // Newer than the owner had read as the window opened - and not their own. A pill too.
        public bool IsNew { get; init; }

        // Its writer's account is locked: the administrator, the one owner answered such a comment, is the one who sees it.
        internal const string AuthorLockedNote = " - account locked, only you see it";

        internal static CollectionCommentRow Of(CollectionCommentAnswer comment, long seenBefore) =>
            new()
            {
                Id = comment.Id,
                Author = comment.Mine ? DisplayNames.You : comment.AuthorName,
                IsMine = comment.Mine,
                Badge = comment.ByAdministrator && !comment.Mine ? CollectionComments.AdministratorBadge : null,
                Details =
                    $"{CollectionComments.WhenText(comment.CreatedUtc)}" +
                    (comment.EditedUtc is null ? string.Empty : " - edited") +
                    (comment.Hidden ? " - hidden from the page" : string.Empty) +
                    (comment.AuthorLocked ? CollectionCommentRow.AuthorLockedNote : string.Empty),
                Body = comment.Body,
                CanHide = !comment.Mine && !comment.Hidden,
                CanShow = comment.Hidden,
                CanDelete = comment.Mine,
                IsNew = comment.Id > seenBefore && !comment.Mine
            };
    }
}
