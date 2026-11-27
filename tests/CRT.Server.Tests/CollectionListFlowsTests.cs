using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The website's "Your collections" (CollectionListFlows; owner decisions, 2026-11-25): a reader's own collections and those
    // they commented on - and every collection with the newest comments for the administrator alone.
    // ###########################################################################################
    public sealed class CollectionListFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private static readonly string Sha = new('a', 64);

        private readonly FakeCollectionStore thisStore = new();

        private readonly FakeAccountStore thisAccounts = new();

        private static readonly AccountRecord Anna = Account(1, "anna@example.com", "Anna");

        private static readonly AccountRecord Bo = Account(2, "bo@example.com", "Bo");

        private static readonly AccountRecord Cy = Account(3, "cy@example.com", "Cy");

        private static readonly AccountRecord Administrator = Account(9, "admin@example.com", "Dennis@CRT", administrator: true);

        public CollectionListFlowsTests()
        {
            foreach (AccountRecord account in new[] { Anna, Bo, Cy, Administrator })
            {
                this.thisAccounts.Accounts[account.Id] = account;
                this.thisStore.Accounts[account.Id] = (account.DisplayName, account.IsAdministrator);
            }

            this.Collection("annafirst222", Anna, Now.AddDays(-3));
            this.Collection("annanext2222", Anna, Now.AddDays(-1));
            this.Collection("bertfirst222", Bo, Now.AddDays(-2));
            this.Collection("cyfirst22222", Cy, Now.AddDays(-4));
        }

        private static AccountRecord Account(long id, string email, string name, bool administrator = false) =>
            new(id, email, email, "hash", name, true, administrator, IsLocked: false, Now.AddDays(-10), null);

        private void Collection(string id, AccountRecord owner, DateTimeOffset created) =>
            this.thisStore.Collections[id] = new CollectionRecord(
                id, owner.Id, "Commodore 64", "250407", "PAL", $"{owner.DisplayName}'s", null, [], created, created, created.AddYears(1), null, null, null,
                [new CollectionImageRecord(new CollectionImageEntry(Sha, "U7", "12", "/RAS", null, null, null, null, null, null, created), 800, 480)],
                1);

        private Task<long> CommentAsync(AccountRecord writer, string collection) =>
            this.thisStore.AddCommentAsync(collection, Sha, writer.Id, "Check R7.", Now);

        private Task<CollectionListView> ListAsync(AccountRecord reader) =>
            CollectionListFlows.ListForAsync(reader, this.thisStore, this.thisAccounts, Now, CancellationToken.None);

        [Fact]
        public async Task A_reader_sees_their_own_newest_first_and_those_they_commented_on_last_commented_first()
        {
            await this.CommentAsync(Anna, "cyfirst22222");
            await this.CommentAsync(Anna, "bertfirst222");
            await this.CommentAsync(Anna, "annafirst222");     // their own: listed as theirs, not as commented on
            await this.CommentAsync(Bo, "cyfirst22222");       // somebody else's comment says nothing of Anna

            CollectionListView view = await this.ListAsync(Anna);

            Assert.Equal(["annanext2222", "annafirst222"], view.Own.Select(collection => collection.Id));
            Assert.Equal(["bertfirst222", "cyfirst22222"], view.CommentedOn.Select(collection => collection.Id));
            Assert.Null(view.All);
            Assert.Null(view.NewestComments);
        }

        [Fact]
        public async Task A_deleted_comment_or_collection_is_no_reason_to_list_it()
        {
            long comment = await this.CommentAsync(Anna, "bertfirst222");
            await this.CommentAsync(Anna, "cyfirst22222");

            await this.thisStore.DeleteCommentAsync(comment, Now);
            await this.thisStore.MarkDeletedAsync("cyfirst22222", CollectionDeletedReasons.Owner, Now);
            await this.thisStore.MarkDeletedAsync("annanext2222", CollectionDeletedReasons.Owner, Now);

            CollectionListView view = await this.ListAsync(Anna);

            Assert.Equal(["annafirst222"], view.Own.Select(collection => collection.Id));
            Assert.Empty(view.CommentedOn);
        }

        // A locked owner's collections are offline for everybody, so a reader's list leaves them out too.
        [Fact]
        public async Task A_locked_owner_s_collection_is_not_listed_as_commented_on()
        {
            await this.CommentAsync(Anna, "bertfirst222");
            this.thisAccounts.Accounts[Bo.Id] = Bo with { IsLocked = true };

            Assert.Empty((await this.ListAsync(Anna)).CommentedOn);
        }

        [Fact]
        public async Task The_administrator_sees_every_collection_with_its_owner_and_the_newest_comments()
        {
            long first = await this.CommentAsync(Bo, "annafirst222");
            long second = await this.CommentAsync(Cy, "bertfirst222");
            await this.thisStore.SetCommentHiddenAsync(second, hidden: true, Now);
            this.thisAccounts.Accounts[Cy.Id] = Cy with { IsLocked = true };

            CollectionListView view = await this.ListAsync(Administrator);

            Assert.Equal(
                [("annanext2222", "Anna", false), ("bertfirst222", "Bo", false), ("annafirst222", "Anna", false), ("cyfirst22222", "Cy", true)],
                view.All!.Select(listed => (listed.Collection.Id, listed.OwnerName, listed.OwnerLocked)));
            Assert.Equal("anna@example.com", view.All![0].OwnerEmail);

            // Hidden ones too - moderating starts there.
            Assert.Equal([second, first], view.NewestComments!.Select(comment => comment.Id));
        }

        [Fact]
        public async Task A_collection_whose_owner_has_no_account_is_listed_to_the_administrator_as_offline()
        {
            this.thisAccounts.Accounts.Remove(Cy.Id);

            CollectionListed listed = (await this.ListAsync(Administrator)).All!.Single(item => item.Collection.Id == "cyfirst22222");

            // Offline, but not "locked" - nobody locked an owner that is not there (code review, 2026-11-27).
            Assert.Equal(
                (CollectionListFlows.NoOwner, (string?)null, false, true, true),
                (listed.OwnerName, listed.OwnerEmail, listed.OwnerLocked, listed.OwnerGone, listed.IsOffline));
        }

        // Locking from the website (AccountLockFlows): the accounts locked listed to unlock, and an administrator's collection
        // marked so that no "Lock account" is offered for it.
        [Fact]
        public async Task The_administrator_sees_the_accounts_locked_and_whose_collections_are_an_administrator_s()
        {
            this.thisAccounts.Accounts[Bo.Id] = Bo with { IsLocked = true };
            this.Collection("adminset2222", Administrator, Now);

            CollectionListView view = await this.ListAsync(Administrator);

            Assert.Equal([new LockedAccount(Bo.Id, "Bo", "bo@example.com")], view.Locked);
            Assert.True(view.All!.Single(listed => listed.Collection.Id == "adminset2222").OwnerIsAdministrator);
            Assert.False(view.All!.Single(listed => listed.Collection.Id == "annafirst222").OwnerIsAdministrator);

            Assert.Null((await this.ListAsync(Anna)).Locked);
        }

        [Fact]
        public async Task The_newest_comments_leave_out_those_of_a_deleted_collection()
        {
            await this.CommentAsync(Bo, "annafirst222");
            await this.thisStore.MarkDeletedAsync("annafirst222", CollectionDeletedReasons.Administrator, Now);

            Assert.Empty((await this.ListAsync(Administrator)).NewestComments!);
        }

        // A collection whose time is up is gone at once, before the hourly sweep marks it - out of every list, and its comments out
        // of the newest (code review, 2026-11-27: listed until the sweep, linking to a page saying it was deleted).
        [Fact]
        public async Task A_collection_whose_time_is_up_is_listed_nowhere()
        {
            this.thisStore.Collections["annafirst222"] = this.thisStore.Collections["annafirst222"] with { DeleteAtUtc = Now.AddMinutes(-1) };
            this.thisStore.Collections["bertfirst222"] = this.thisStore.Collections["bertfirst222"] with { DeleteAtUtc = Now };
            await this.CommentAsync(Anna, "bertfirst222");
            await this.CommentAsync(Bo, "annafirst222");

            CollectionListView anna = await this.ListAsync(Anna);
            CollectionListView administrator = await this.ListAsync(Administrator);

            Assert.Equal(["annanext2222"], anna.Own.Select(collection => collection.Id));
            Assert.Empty(anna.CommentedOn);
            Assert.DoesNotContain(administrator.All!, listed => listed.Collection.Id is "annafirst222" or "bertfirst222");
            Assert.Empty(administrator.NewestComments!);
        }
    }
}
