using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The website's list of collections (CollectionListFlows): the administrator's alone - every collection, so the pages can be
    // checked (owner request, 2026-11-28: "I (as web admin) still want to be able to see all collections") - and nobody else's.
    // ###########################################################################################
    public sealed class CollectionListFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private static readonly string Sha = new('a', 64);

        private readonly FakeCollectionStore thisStore = new();

        private static readonly AccountRecord AMaintainer = Account(2, "bo@example.com", "Bo");

        private static readonly AccountRecord Administrator = Account(9, "admin@example.com", "Dennis@CRT", administrator: true);

        public CollectionListFlowsTests()
        {
            this.Collection("first2222222", Now.AddDays(-3));
            this.Collection("newest222222", Now.AddDays(-1));
            this.Collection("middle222222", Now.AddDays(-2));
        }

        private static AccountRecord Account(long id, string email, string name, bool administrator = false) =>
            new(id, email, email, "hash", name, true, administrator, IsLocked: false, Now.AddDays(-10), null);

        private void Collection(string id, DateTimeOffset created, string? ownerKeyHash = "0f") =>
            this.thisStore.Collections[id] = new CollectionRecord(
                id, ownerKeyHash, "Commodore 64", "250407", "PAL", id, null, [], created, created, created.AddYears(1), null, null,
                [new CollectionImageRecord(new CollectionImageEntry(Sha, "U7", "12", "/RAS", null, null, null, null, null, created), 800, 480)],
                1);

        private Task<IReadOnlyList<CollectionRecord>?> ListAsync(AccountRecord reader) =>
            CollectionListFlows.ListForAsync(reader, this.thisStore, Now, CancellationToken.None);

        // Newest first, each counted - a collection shared with an account before server 17.0.0 (owned by no key) too.
        [Fact]
        public async Task The_administrator_sees_every_collection_newest_first()
        {
            this.Collection("fromanaccnt2", Now.AddHours(-1), ownerKeyHash: null);

            IReadOnlyList<CollectionRecord>? all = await this.ListAsync(Administrator);

            Assert.Equal(["fromanaccnt2", "newest222222", "middle222222", "first2222222"], all!.Select(collection => collection.Id));
            Assert.All(all!, collection => Assert.Equal(1, collection.ImageCount));
        }

        // Nobody else signs in on the website - a maintainer's session there gets no list either.
        [Fact]
        public async Task Nobody_but_the_administrator_gets_a_list()
        {
            Assert.Null(await this.ListAsync(AMaintainer));
        }

        // A collection deleted, or whose time is up - gone at once, before the hourly sweep marks it - is listed nowhere (code
        // review, 2026-11-27: listed until the sweep, linking to a page saying it was deleted). One whose sharing stopped is
        // still the administrator's to check.
        [Fact]
        public async Task A_collection_deleted_or_whose_time_is_up_is_not_listed()
        {
            await this.thisStore.MarkDeletedAsync("first2222222", CollectionDeletedReasons.Owner, Now);
            this.thisStore.Collections["newest222222"] = this.thisStore.Collections["newest222222"] with { DeleteAtUtc = Now };
            await this.thisStore.SetUnsharedAsync("middle222222", Now);

            Assert.Equal(["middle222222"], (await this.ListAsync(Administrator))!.Select(collection => collection.Id));
        }
    }
}
