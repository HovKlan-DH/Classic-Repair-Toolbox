using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using CRT.Server.Handlers.Email;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Comments on a shared collection's captures (CollectionCommentFlows; owner decisions, 2026-11-26): anybody with a
    // confirmed address comments on a capture, the owner mailed for every comment but their own; the writer edits and
    // deletes their own; the owner hides any - all on the fake stores.
    // ###########################################################################################
    public sealed class CollectionCommentFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private const string Site = "https://classic-repair-toolbox.dk";

        private const string Id = "k7mq3x9vbn2p";

        private static readonly string Sha = new('a', 64);

        private readonly FakeCollectionStore thisStore = new();

        private readonly FakeAccountStore thisAccounts = new();

        private readonly FakeEmailSender thisMailer = new();

        private static readonly AccountRecord Owner = Account(1, "owner@example.com", "Anna");

        private static readonly AccountRecord Helper = Account(2, "helper@example.com", "Bo");

        private static readonly AccountRecord Administrator = Account(9, "admin@example.com", "Dennis@CRT", administrator: true);

        public CollectionCommentFlowsTests()
        {
            foreach (AccountRecord account in new[] { Owner, Helper, Administrator })
            {
                this.thisAccounts.Accounts[account.Id] = account;
                this.thisStore.Accounts[account.Id] = (account.DisplayName, account.IsAdministrator);
            }

            this.thisStore.Collections[Id] = new CollectionRecord(
                Id, Owner.Id, "Commodore 64", "250407", "PAL", "No picture", null, [], Now, Now, Now.AddYears(1), null, null, null,
                [new CollectionImageRecord(new CollectionImageEntry(Sha, "U7", "12", "/RAS", null, null, null, null, null, null, false, Now), 800, 480)],
                1);
        }

        private static AccountRecord Account(long id, string email, string name, bool verified = true, bool administrator = false) =>
            new(id, email, email, "hash", name, verified, administrator, IsLocked: false, Now.AddDays(-10), null);

        private Task<CommentOutcome> AddAsync(AccountRecord? account, string body = "Pin 12 should pulse - check R7.", string? image = null, string id = Id) =>
            CollectionCommentFlows.AddAsync(account, id, image ?? Sha, body, this.thisStore, this.thisAccounts, this.thisMailer, Site, Now, CancellationToken.None);

        [Fact]
        public async Task A_comment_is_kept_on_its_capture_and_the_owner_is_mailed_with_it()
        {
            CommentOutcome outcome = await this.AddAsync(Helper, "  Pin 12 should pulse - check R7.\r\nThen U8.  ");

            Assert.True(outcome.IsDone);

            CollectionCommentRecord kept = Assert.Single(await this.thisStore.ListCommentsAsync(Id));

            Assert.Equal((Sha, Helper.Id, "Pin 12 should pulse - check R7.\nThen U8."), (kept.ImageSha256, kept.AccountId, kept.Body));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);

            Assert.Equal("owner@example.com", mail.ToAddress);
            Assert.Contains("Bo", mail.Body, StringComparison.Ordinal);
            Assert.Contains("U7, pin 12", mail.Body, StringComparison.Ordinal);
            Assert.Contains("check R7.", mail.Body, StringComparison.Ordinal);
            Assert.Contains($"{Site}/collections/{Id}#comment-{outcome.CommentId}", mail.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task The_owner_s_own_comment_mails_nobody()
        {
            Assert.True((await this.AddAsync(Owner, "Thanks - R7 it was.")).IsDone);
            Assert.Empty(this.thisMailer.Sent);
        }

        [Fact]
        public async Task Nobody_signed_in_or_unconfirmed_comments()
        {
            Assert.Equal(CollectionVerdict.Unauthorized, (await this.AddAsync(null)).Verdict);
            Assert.Equal(CollectionVerdict.Forbidden, (await this.AddAsync(Account(3, "new@example.com", "New", verified: false))).Verdict);
            Assert.Empty(await this.thisStore.ListCommentsAsync(Id));
        }

        [Fact]
        public async Task A_comment_on_no_collection_or_on_a_capture_it_does_not_show_is_refused()
        {
            Assert.Equal(CollectionVerdict.NotFound, (await this.AddAsync(Helper, id: "nonexistent12")).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await this.AddAsync(Helper, id: "../../etc")).Verdict);

            CommentOutcome outcome = await this.AddAsync(Helper, image: new string('b', 64));

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal([CollectionCommentFlows.NoSuchCapture], outcome.Errors);
        }

        // A locked owner's collection is offline - nobody comments on it.
        [Fact]
        public async Task A_collection_of_a_locked_owner_takes_no_comment()
        {
            this.thisAccounts.Accounts[Owner.Id] = Owner with { IsLocked = true };

            Assert.Equal(CollectionVerdict.NotFound, (await this.AddAsync(Helper)).Verdict);
        }

        // Locking a spammer takes the spam down with it (owner decision, 2026-11-27 - AccountLockFlows): a locked account's
        // comments are the administrator's alone to see, on the page and in the owner's CRT, and unlocking puts them back.
        [Fact]
        public async Task A_locked_account_s_comments_are_shown_to_the_administrator_alone()
        {
            await this.AddAsync(Helper, "Buy cheap chips here");
            this.thisAccounts.Accounts[Helper.Id] = Helper with { IsLocked = true };
            this.thisStore.LockedAccounts.Add(Helper.Id);

            CollectionRecord collection = this.thisStore.Collections[Id];
            IReadOnlyList<CollectionCommentRecord> kept = await this.thisStore.ListCommentsAsync(Id);

            Assert.Empty(CollectionCommentFlows.ShownTo(kept, collection, null));
            Assert.Empty(CollectionCommentFlows.ShownTo(kept, collection, Owner));
            Assert.True(Assert.Single(CollectionCommentFlows.ShownTo(kept, collection, Administrator)).AuthorIsLocked);

            Assert.Empty((await CollectionCommentFlows.ListForOwnerAsync(Owner, Id, this.thisStore, Now, CancellationToken.None)).Comments!.Comments);
            // And said so, as the page flags it - CRT's window marks it "only you see it" (code review, 2026-11-27).
            Assert.True(Assert.Single((await CollectionCommentFlows.ListForOwnerAsync(Administrator, Id, this.thisStore, Now, CancellationToken.None)).Comments!.Comments).AuthorLocked);

            // Nor counted as new in the owner's CRT.
            CollectionRecord mine = Assert.Single(await this.thisStore.ListLiveForAccountAsync(Owner.Id));
            Assert.Equal((0, 0L), (mine.CommentCount, mine.LatestCommentId));

            this.thisStore.LockedAccounts.Remove(Helper.Id);
            Assert.Single(CollectionCommentFlows.ShownTo(await this.thisStore.ListCommentsAsync(Id), collection, null));
        }

        [Theory]
        [InlineData("   ", CollectionComments.Missing)]
        [InlineData(null, CollectionComments.Missing)]
        public async Task An_empty_comment_is_refused(string? body, string expected)
        {
            CommentOutcome outcome = await CollectionCommentFlows.AddAsync(Helper, Id, Sha, body, this.thisStore, this.thisAccounts, this.thisMailer, Site, Now, CancellationToken.None);

            Assert.Equal([expected], outcome.Errors);
        }

        [Fact]
        public async Task A_comment_too_long_is_refused_saying_how_long_one_may_be()
        {
            CommentOutcome outcome = await this.AddAsync(Helper, new string('x', CollectionComments.MaximumLength + 1));

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Contains($"at most {CollectionComments.MaximumLength}", outcome.Errors[0], StringComparison.Ordinal);
        }

        [Fact]
        public async Task An_account_writes_its_day_s_comments_and_no_more()
        {
            for (int i = 0; i < CollectionComments.MaximumPerDay; i++)
                await this.thisStore.AddCommentAsync(Id, Sha, Helper.Id, "x", Now.AddHours(-1));

            Assert.Equal(CollectionVerdict.TooMany, (await this.AddAsync(Helper)).Verdict);
            Assert.True((await this.AddAsync(Owner)).IsDone);
        }

        [Fact]
        public async Task Only_its_writer_edits_a_comment_and_it_is_marked_edited()
        {
            long id = (await this.AddAsync(Helper)).CommentId!.Value;

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.EditAsync(Owner, Id, id, "Changed", this.thisStore, this.thisAccounts, Now, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.EditAsync(Administrator, Id, id, "Changed", this.thisStore, this.thisAccounts, Now, CancellationToken.None)).Verdict);

            Assert.True((await CollectionCommentFlows.EditAsync(Helper, Id, id, " Check R7 and R8. ", this.thisStore, this.thisAccounts, Now.AddMinutes(5), CancellationToken.None)).IsDone);

            CollectionCommentRecord kept = (await this.thisStore.FindCommentAsync(id))!;

            Assert.Equal("Check R7 and R8.", kept.Body);
            Assert.Equal(Now.AddMinutes(5), kept.EditedUtc);
        }

        [Fact]
        public async Task Its_writer_or_an_administrator_deletes_a_comment_and_nobody_else()
        {
            long first = (await this.AddAsync(Helper)).CommentId!.Value;
            long second = (await this.AddAsync(Helper)).CommentId!.Value;

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Owner, Id, first, this.thisStore, Now, CancellationToken.None)).Verdict);

            Assert.True((await CollectionCommentFlows.DeleteAsync(Helper, Id, first, this.thisStore, Now, CancellationToken.None)).IsDone);
            Assert.True((await CollectionCommentFlows.DeleteAsync(Administrator, Id, second, this.thisStore, Now, CancellationToken.None)).IsDone);
            Assert.Empty(await this.thisStore.ListCommentsAsync(Id));

            // Gone: there is nothing of it to delete again.
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Helper, Id, first, this.thisStore, Now, CancellationToken.None)).Verdict);
        }

        [Fact]
        public async Task The_owner_hides_a_comment_from_everybody_else_and_shows_it_again()
        {
            long id = (await this.AddAsync(Helper)).CommentId!.Value;

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.SetHiddenAsync(Helper, Id, id, true, this.thisStore, Now, CancellationToken.None)).Verdict);
            Assert.True((await CollectionCommentFlows.SetHiddenAsync(Owner, Id, id, true, this.thisStore, Now, CancellationToken.None)).IsDone);

            IReadOnlyList<CollectionCommentRecord> all = await this.thisStore.ListCommentsAsync(Id);
            CollectionRecord collection = this.thisStore.Collections[Id];

            Assert.Empty(CollectionCommentFlows.ShownTo(all, collection, null));
            Assert.Empty(CollectionCommentFlows.ShownTo(all, collection, Helper));
            Assert.Single(CollectionCommentFlows.ShownTo(all, collection, Owner));
            Assert.Single(CollectionCommentFlows.ShownTo(all, collection, Administrator));

            Assert.True((await CollectionCommentFlows.SetHiddenAsync(Owner, Id, id, false, this.thisStore, Now, CancellationToken.None)).IsDone);
            Assert.Single(CollectionCommentFlows.ShownTo(await this.thisStore.ListCommentsAsync(Id), collection, null));
        }

        [Fact]
        public async Task The_owner_s_CRT_lists_every_comment_saying_whose_it_is()
        {
            long helpers = (await this.AddAsync(Helper)).CommentId!.Value;
            await this.AddAsync(Owner, "Thanks");
            await CollectionCommentFlows.SetHiddenAsync(Owner, Id, helpers, true, this.thisStore, Now, CancellationToken.None);

            CommentOutcome outcome = await CollectionCommentFlows.ListForOwnerAsync(Owner, Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(
                [("Bo", false, false, true), ("Anna", true, true, false)],
                outcome.Comments!.Comments.Select(comment => (comment.AuthorName, comment.ByOwner, comment.Mine, comment.Hidden)));

            // Not the owner's: not found.
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.ListForOwnerAsync(Helper, Id, this.thisStore, Now, CancellationToken.None)).Verdict);
        }

        // A route names its collection: a comment of another one is "not found" there, whatever the caller may do to it in its own
        // (code review, 2026-11-27: "/collections/<A>/comments/<c>" acted on a comment of collection B).
        [Fact]
        public async Task A_comment_is_acted_on_only_under_its_own_collection()
        {
            const string other = "otherset2222";

            this.thisStore.Collections[other] = this.thisStore.Collections[Id] with { Id = other };
            long id = (await this.AddAsync(Helper)).CommentId!.Value;

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.SetHiddenAsync(Owner, other, id, true, this.thisStore, Now, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.EditAsync(Helper, other, id, "Changed", this.thisStore, this.thisAccounts, Now, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Helper, other, id, this.thisStore, Now, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Helper, "../../x", id, this.thisStore, Now, CancellationToken.None)).Verdict);

            CollectionCommentRecord kept = (await this.thisStore.FindCommentAsync(id))!;

            Assert.False(kept.IsHidden);
            Assert.Null(kept.EditedUtc);
        }

        // A hidden comment is gone for its writer - not to be rewritten into something harmless, shown again, and changed back
        // (code review, 2026-11-27).
        [Fact]
        public async Task A_hidden_comment_is_not_edited_by_its_writer()
        {
            long id = (await this.AddAsync(Helper, "Spam")).CommentId!.Value;
            await CollectionCommentFlows.SetHiddenAsync(Owner, Id, id, true, this.thisStore, Now, CancellationToken.None);

            CommentOutcome edit = await CollectionCommentFlows.EditAsync(Helper, Id, id, "Harmless", this.thisStore, this.thisAccounts, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.NotFound, edit.Verdict);
            Assert.Equal("Spam", (await this.thisStore.FindCommentAsync(id))!.Body);
        }

        // Its time up, a collection's comments are gone with it - no hiding or deleting them before the sweep (code review,
        // 2026-11-27).
        [Fact]
        public async Task A_collection_whose_time_is_up_takes_no_hide_or_delete()
        {
            long id = (await this.AddAsync(Helper)).CommentId!.Value;
            DateTimeOffset after = this.thisStore.Collections[Id].DeleteAtUtc.AddMinutes(1);

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.SetHiddenAsync(Owner, Id, id, true, this.thisStore, after, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Helper, Id, id, this.thisStore, after, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.Gone, (await CollectionCommentFlows.ListForOwnerAsync(Owner, Id, this.thisStore, after, CancellationToken.None)).Verdict);
        }

        // The owner's own comment - an answer written on the website - is nothing new to them: CRT's "new comments" counts others'
        // alone (code review, 2026-11-27).
        [Fact]
        public async Task The_owner_s_own_comment_is_no_new_comment_to_them()
        {
            long helpers = (await this.AddAsync(Helper)).CommentId!.Value;
            await this.AddAsync(Owner, "Thanks - it was R7.");

            CollectionSummaryAnswer summary = Assert.Single((await CollectionFlows.ListMineAsync(Owner, this.thisStore, Site, Now, CancellationToken.None)).Mine!.Collections);

            Assert.Equal((2, helpers), (summary.CommentCount, summary.LatestCommentId));
        }

        // A comment on a capture the owner took out and shared again without is on no page and in no CRT window: it counts for
        // nothing - no "1 comment", no "new comments", not "You commented on" - and comes back with the capture (code review,
        // 2026-11-27: nobody could see, hide or delete what was counted).
        [Fact]
        public async Task A_comment_on_a_capture_taken_out_counts_for_nothing()
        {
            await this.AddAsync(Helper);
            IReadOnlyList<CollectionImageRecord> images = this.thisStore.Collections[Id].Images;
            this.thisStore.Collections[Id] = this.thisStore.Collections[Id] with { Images = [] };

            CollectionSummaryAnswer summary = Assert.Single((await CollectionFlows.ListMineAsync(Owner, this.thisStore, Site, Now, CancellationToken.None)).Mine!.Collections);

            Assert.Equal((0, 0L), (summary.CommentCount, summary.LatestCommentId));
            Assert.Empty(await this.thisStore.ListLiveCommentedOnAsync(Helper.Id));
            Assert.Empty(await this.thisStore.ListNewestCommentsAsync(10));

            this.thisStore.Collections[Id] = this.thisStore.Collections[Id] with { Images = images };

            Assert.Equal(1, Assert.Single((await CollectionFlows.ListMineAsync(Owner, this.thisStore, Site, Now, CancellationToken.None)).Mine!.Collections).CommentCount);
        }

        // The owner's CRT asking for a collection the administrator took down on the website hears it went, and when (code
        // review, 2026-11-27: "There is no collection of yours with that id" - and CRT went on showing it shared).
        [Fact]
        public async Task The_owner_asking_for_the_comments_of_a_collection_taken_down_hears_it_went()
        {
            await CollectionFlows.DeleteAsync(Administrator, Id, this.thisStore, Now, CancellationToken.None);

            CommentOutcome outcome = await CollectionCommentFlows.ListForOwnerAsync(Owner, Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Gone, outcome.Verdict);
            Assert.Equal(CollectionFlows.DeletedLine(this.thisStore.Collections[Id]), Assert.Single(outcome.Errors));

            // Somebody else is told nothing of it.
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.ListForOwnerAsync(Helper, Id, this.thisStore, Now, CancellationToken.None)).Verdict);
        }

        [Fact]
        public async Task A_collection_deleted_takes_its_comments_with_it()
        {
            await this.AddAsync(Helper);

            await CollectionFlows.DeleteAsync(Owner, Id, this.thisStore, Now, CancellationToken.None);

            Assert.Empty(await this.thisStore.ListCommentsAsync(Id));
        }

        // The owner's list in CRT says how many comments the page shows and the newest's id - so CRT marks what is new.
        [Fact]
        public async Task The_owner_s_list_counts_the_comments_shown()
        {
            await this.AddAsync(Helper);
            long newest = (await this.AddAsync(Helper)).CommentId!.Value;
            long hidden = (await this.AddAsync(Helper)).CommentId!.Value;
            await CollectionCommentFlows.SetHiddenAsync(Owner, Id, hidden, true, this.thisStore, Now, CancellationToken.None);

            CollectionSummaryAnswer summary = Assert.Single((await CollectionFlows.ListMineAsync(Owner, this.thisStore, Site, Now, CancellationToken.None)).Mine!.Collections);

            Assert.Equal((2, newest), (summary.CommentCount, summary.LatestCommentId));
        }

        [Fact]
        public void A_name_in_a_subject_line_is_one_line()
        {
            Assert.Equal("No picture  Bcc: x@y", EmailTemplates.OneLine("No picture\r\nBcc: x@y"));
            Assert.Equal(80, EmailTemplates.OneLine(new string('x', 200)).Length);

            // Never half of an emoji (code review, 2026-11-27).
            string cut = EmailTemplates.OneLine(new string('x', 76) + "😀" + new string('y', 20));
            Assert.Equal(new string('x', 76) + "...", cut);
        }

        [Fact]
        public void A_subject_cut_after_a_whole_emoji_keeps_it()
        {
            string cut = EmailTemplates.OneLine(new string('x', 75) + "😀" + new string('y', 20));
            Assert.Equal(new string('x', 75) + "😀...", cut);
        }
    }
}
