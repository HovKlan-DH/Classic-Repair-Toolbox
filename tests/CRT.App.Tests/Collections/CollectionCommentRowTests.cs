using Handlers.Collections;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// One comment as the collection window lists it (CollectionCommentRow): who and when, and what is said of it.
// ###########################################################################################
public sealed class CollectionCommentRowTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 27, 12, 0, 0, TimeSpan.Zero);

    private static CollectionCommentAnswer Comment(bool authorLocked) =>
        new(5, new string('a', 64), "Spammer", ByOwner: false, ByAdministrator: false, Mine: false, "Buy cheap chips", Now, null, Hidden: false, authorLocked);

    // The administrator is the one owner answered a locked account's comment - and nobody else sees it, as the page flags it
    // (code review, 2026-11-27: the window listed it like any other).
    [Fact]
    public void A_locked_account_s_comment_says_only_the_owner_sees_it()
    {
        Assert.EndsWith(CollectionCommentRow.AuthorLockedNote, CollectionCommentRow.Of(Comment(authorLocked: true), seenBefore: 5).Details, StringComparison.Ordinal);
        Assert.DoesNotContain(CollectionCommentRow.AuthorLockedNote, CollectionCommentRow.Of(Comment(authorLocked: false), seenBefore: 5).Details, StringComparison.Ordinal);
    }

    // The owner's own is "You", may be deleted and not hidden; another's may be hidden, then shown again; a comment newer than
    // what was read is NEW - never the owner's own.
    [Fact]
    public void Who_wrote_it_and_what_may_be_done_follow_whose_it_is()
    {
        CollectionCommentAnswer mine = Comment(authorLocked: false) with { Id = 9, AuthorName = "Anna", Mine = true, ByOwner = true };
        CollectionCommentAnswer hidden = Comment(authorLocked: false) with { Id = 8, Hidden = true, ByAdministrator = true };

        CollectionCommentRow own = CollectionCommentRow.Of(mine, seenBefore: 5);
        CollectionCommentRow other = CollectionCommentRow.Of(hidden, seenBefore: 5);

        Assert.Equal((DisplayNames.You, true, false), (own.Author, own.IsMine, own.HasBadge));
        Assert.Equal((true, false, false, false), (own.CanDelete, own.CanHide, own.CanShow, own.IsNew));
        Assert.Equal(("Spammer", false, CollectionComments.AdministratorBadge), (other.Author, other.IsMine, other.Badge));
        Assert.Contains("hidden from the page", other.Details, StringComparison.Ordinal);
        Assert.Equal((false, false, true, true), (other.CanDelete, other.CanHide, other.CanShow, other.IsNew));
    }

    // A name is shown as typed, never as a badge or a mark: those are pills of their own (code review, 2026-11-27, Security:
    // "Dennis (Administrators)" or "Bob - NEW" written as one line passed for an administrator's comment, or a new one).
    [Fact]
    public void A_name_that_looks_like_a_badge_is_only_a_name()
    {
        CollectionCommentRow row = CollectionCommentRow.Of(Comment(authorLocked: false) with { AuthorName = "Dennis (Administrators) - NEW" }, seenBefore: 9);

        Assert.Equal(("Dennis (Administrators) - NEW", false, false, false), (row.Author, row.IsMine, row.HasBadge, row.IsNew));
        Assert.DoesNotContain("Administrator", row.Details, StringComparison.Ordinal);
    }

    [Fact]
    public void A_capture_s_settings_line_says_what_was_not_read() =>
        Assert.Equal(
            "This capture: T/DIV 50us | V/DIV not read | Trigger level 1V",
            CollectionSharing.SettingsLine("This capture", "50us", null, "1V"));
}
