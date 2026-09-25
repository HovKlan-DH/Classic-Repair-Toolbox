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

        private const string Cited = "Commodore/C64/250407/sheet1.png";
        private const string Orphan = "Generic shared files/Component images/7408.jpg";

        public UnusedFileFlowsTests()
        {
            this.thisData = Path.Combine(this.thisRoot, "app-data-BETA", "Data");
            this.thisManifest = Path.Combine(this.thisRoot, "app-data-BETA", "dataChecksums.json");

            DataTreeBuilder.Master(this.thisData, DataTreeBuilder.Workbook);
            DataTreeBuilder.Board(this.thisData, DataTreeBuilder.Workbook, UnusedFileFlowsTests.Cited);
            DataTreeBuilder.Files(this.thisData, UnusedFileFlowsTests.Cited, UnusedFileFlowsTests.Orphan);
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
                administrator ? [] : [DataTreeBuilder.BoardFolder]);

        [Fact]
        public async Task A_reviewer_cannot_remove_unused_files_and_nothing_is_touched()
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
            Assert.Contains("sheet1.png", manifest, StringComparison.Ordinal);
            Assert.DoesNotContain("7408.jpg", manifest, StringComparison.Ordinal);

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
