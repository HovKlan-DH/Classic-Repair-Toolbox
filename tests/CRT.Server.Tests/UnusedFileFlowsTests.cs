using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // UnusedFileFlows - the administrator's "Unused files" screen (2026-09-25): the list per tree,
    // and removing the files the administrator chose, with the manifest rewritten and an audit
    // row naming every file. Only an administrator may remove; the negative test is first.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class UnusedFileFlowsTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-unused-flows", Guid.NewGuid().ToString("N"));
        private readonly string thisData;
        private readonly string thisManifest;

        // A CRT 4.0 tree (2026-10-25): the board lists "sheet1.png"; a stored file no board lists
        // lies beside it.
        private static readonly string Cited = DataTreeBuilder.Stored("sheet", "sheet1.png");
        private static readonly string Orphan = DataTreeBuilder.Stored("7408 picture", "7408.jpg");

        public UnusedFileFlowsTests()
        {
            this.thisData = Path.Combine(this.thisRoot, "app-data-v4-BETA", "Data");
            this.thisManifest = Path.Combine(this.thisRoot, "app-data-v4-BETA", "dataChecksums.json");

            DataTreeBuilder.Master(this.thisData, DataTreeBuilder.BoardId);
            DataTreeBuilder.Board(this.thisData, DataTreeBuilder.BoardId, ("sheet1.png", "sheet"));
            DataTreeBuilder.StoreOnly(this.thisData, "7408.jpg", "7408 picture");
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        private ServerOptions Options() => new()
        {
            DataTreeRoot = this.thisData,
            ManifestPath = this.thisManifest,
            PublicDataBaseUrl = "https://example.com/app-data-BETA/Data"
        };

        private UnusedFileFlows.DataTree Beta()
        {
            Assert.True(UnusedFileFlows.TryResolve("beta", this.Options(), out UnusedFileFlows.DataTree? tree, out string error), error);
            return tree!;
        }

        private static ReviewAccess Access(bool administrator) =>
            ReviewAccess.For(
                new AccountRecord(1, "a@example.com", "a@example.com", "hash", "A",
                    IsVerified: true, IsAdministrator: administrator, IsLocked: false, UnusedFileFlowsTests.Now, null),
                administrator ? [] : [DataTreeBuilder.BoardId]);

        [Fact]
        public async Task A_maintainer_cannot_remove_unused_files_and_nothing_is_touched()
        {
            var accounts = new FakeAccountStore();

            UnusedFileRemoval removal = await UnusedFileFlows.RemoveAsync(
                UnusedFileFlowsTests.Access(administrator: false), this.Beta(), [UnusedFileFlowsTests.Orphan],
                new PublishLock(), accounts, NullLogger.Instance, UnusedFileFlowsTests.Now);

            Assert.Empty(removal.Removed);
            Assert.Equal(UnusedFileFlows.AdministratorOnly, removal.NotDoneBecause);
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisData, UnusedFileFlowsTests.Orphan)));
            Assert.Empty(accounts.Audit);
        }

        [Fact]
        public void The_list_names_every_unused_file_with_its_size()
        {
            UnusedFileListing listing = UnusedFileFlows.List(this.Beta());

            Assert.True(listing.IsComplete, string.Join(" ", listing.Problems));
            UnusedFileEntry entry = Assert.Single(listing.Files);
            Assert.Equal(UnusedFileFlowsTests.Orphan, entry.Path);
            Assert.Equal(new FileInfo(DataTreeBuilder.Full(this.thisData, UnusedFileFlowsTests.Orphan)).Length, entry.SizeBytes);
            Assert.Equal("beta", listing.Tree);

            // Where the tree is published, so the Maintainer tab's tree can show and open the file
            // (2026-10-04).
            Assert.Equal(this.Beta().PublicBaseUrl, listing.PublicDataUrl);
        }

        // Clients are told by the manifest: until it is rewritten they keep being offered a file
        // that is gone. And the audit row names what went.
        [Fact]
        public async Task The_administrator_removes_the_chosen_files_the_manifest_forgets_them_and_the_audit_names_them()
        {
            var accounts = new FakeAccountStore();

            UnusedFileRemoval removal = await UnusedFileFlows.RemoveAsync(
                UnusedFileFlowsTests.Access(administrator: true), this.Beta(), [UnusedFileFlowsTests.Orphan],
                new PublishLock(), accounts, NullLogger.Instance, UnusedFileFlowsTests.Now);

            Assert.Equal([UnusedFileFlowsTests.Orphan], removal.Removed);
            Assert.False(File.Exists(DataTreeBuilder.Full(this.thisData, UnusedFileFlowsTests.Orphan)));

            string manifest = File.ReadAllText(this.thisManifest);
            Assert.Contains(UnusedFileFlowsTests.Cited, manifest, StringComparison.Ordinal);
            Assert.DoesNotContain(UnusedFileFlowsTests.Orphan, manifest, StringComparison.Ordinal);

            AuditEntry audit = Assert.Single(accounts.Audit);
            Assert.Equal(UnusedFileFlows.RemovedAction, audit.Action);
            Assert.Equal("beta", audit.Subject);
            Assert.Contains(UnusedFileFlowsTests.Orphan, audit.Detail, StringComparison.Ordinal);
        }

        [Fact]
        public void Production_is_refused_while_publishing_to_it_is_not_switched_on()
        {
            Assert.False(UnusedFileFlows.TryResolve("production", this.Options(), out _, out string error));
            Assert.Equal(ProductionPromotionFlow.NotConfiguredMessage, error);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("../etc")]
        [InlineData("prod")]
        public void Anything_but_beta_or_production_is_refused(string? name)
        {
            Assert.False(UnusedFileFlows.TryResolve(name, this.Options(), out UnusedFileFlows.DataTree? tree, out _));
            Assert.Null(tree);
        }
    }
}
