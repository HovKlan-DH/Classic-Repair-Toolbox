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

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.EditAsync(Owner, id, "Changed", this.thisStore, this.thisAccounts, Now, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.EditAsync(Administrator, id, "Changed", this.thisStore, this.thisAccounts, Now, CancellationToken.None)).Verdict);

            Assert.True((await CollectionCommentFlows.EditAsync(Helper, id, " Check R7 and R8. ", this.thisStore, this.thisAccounts, Now.AddMinutes(5), CancellationToken.None)).IsDone);

            CollectionCommentRecord kept = (await this.thisStore.FindCommentAsync(id))!;

            Assert.Equal("Check R7 and R8.", kept.Body);
            Assert.Equal(Now.AddMinutes(5), kept.EditedUtc);
        }

        [Fact]
        public async Task Its_writer_or_an_administrator_deletes_a_comment_and_nobody_else()
        {
            long first = (await this.AddAsync(Helper)).CommentId!.Value;
            long second = (await this.AddAsync(Helper)).CommentId!.Value;

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Owner, first, this.thisStore, Now, CancellationToken.None)).Verdict);

            Assert.True((await CollectionCommentFlows.DeleteAsync(Helper, first, this.thisStore, Now, CancellationToken.None)).IsDone);
            Assert.True((await CollectionCommentFlows.DeleteAsync(Administrator, second, this.thisStore, Now, CancellationToken.None)).IsDone);
            Assert.Empty(await this.thisStore.ListCommentsAsync(Id));

            // Gone: there is nothing of it to delete again.
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.DeleteAsync(Helper, first, this.thisStore, Now, CancellationToken.None)).Verdict);
        }

        [Fact]
        public async Task The_owner_hides_a_comment_from_everybody_else_and_shows_it_again()
        {
            long id = (await this.AddAsync(Helper)).CommentId!.Value;

            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.SetHiddenAsync(Helper, id, true, this.thisStore, Now, CancellationToken.None)).Verdict);
            Assert.True((await CollectionCommentFlows.SetHiddenAsync(Owner, id, true, this.thisStore, Now, CancellationToken.None)).IsDone);

            IReadOnlyList<CollectionCommentRecord> all = await this.thisStore.ListCommentsAsync(Id);
            CollectionRecord collection = this.thisStore.Collections[Id];

            Assert.Empty(CollectionCommentFlows.ShownTo(all, collection, null));
            Assert.Empty(CollectionCommentFlows.ShownTo(all, collection, Helper));
            Assert.Single(CollectionCommentFlows.ShownTo(all, collection, Owner));
            Assert.Single(CollectionCommentFlows.ShownTo(all, collection, Administrator));

            Assert.True((await CollectionCommentFlows.SetHiddenAsync(Owner, id, false, this.thisStore, Now, CancellationToken.None)).IsDone);
            Assert.Single(CollectionCommentFlows.ShownTo(await this.thisStore.ListCommentsAsync(Id), collection, null));
        }

        [Fact]
        public async Task The_owner_s_CRT_lists_every_comment_saying_whose_it_is()
        {
            long helpers = (await this.AddAsync(Helper)).CommentId!.Value;
            await this.AddAsync(Owner, "Thanks");
            await CollectionCommentFlows.SetHiddenAsync(Owner, helpers, true, this.thisStore, Now, CancellationToken.None);

            CommentOutcome outcome = await CollectionCommentFlows.ListForOwnerAsync(Owner, Id, this.thisStore, CancellationToken.None);

            Assert.Equal(
                [("Bo", false, false, true), ("Anna", true, true, false)],
                outcome.Comments!.Comments.Select(comment => (comment.AuthorName, comment.ByOwner, comment.Mine, comment.Hidden)));

            // Not the owner's: not found.
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionCommentFlows.ListForOwnerAsync(Helper, Id, this.thisStore, CancellationToken.None)).Verdict);
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
            await CollectionCommentFlows.SetHiddenAsync(Owner, hidden, true, this.thisStore, Now, CancellationToken.None);

            CollectionSummaryAnswer summary = Assert.Single((await CollectionFlows.ListMineAsync(Owner, this.thisStore, Site, CancellationToken.None)).Mine!.Collections);

            Assert.Equal((2, newest), (summary.CommentCount, summary.LatestCommentId));
        }

        [Fact]
        public void A_name_in_a_subject_line_is_one_line()
        {
            Assert.Equal("No picture  Bcc: x@y", EmailTemplates.OneLine("No picture\r\nBcc: x@y"));
            Assert.Equal(80, EmailTemplates.OneLine(new string('x', 200)).Length);
        }
    }
}
