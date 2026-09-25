using System.Security.Cryptography;
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
    // Covers publishing a system from BETA to PRODUCTION, end to end (owner request,
    // 2026-09-25): a real submission is published into a temp BETA tree through ApprovePublishFlow,
    // exactly as the Approve button does, and then promoted into a temp production tree.
    //
    // Asserted on the DISK, like ApprovePublishFlowTests: production is what every user downloads,
    // and the only honest way to test a refusal is to check nothing arrived there.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class ProductionPromotionFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private const string SystemId = "Commodore/C64/250407";

        private readonly string thisRoot;
        private readonly string thisBeta;
        private readonly string thisProduction;
        private readonly string thisBlobRoot;

        public ProductionPromotionFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-promote", Guid.NewGuid().ToString("N"));
            this.thisBeta = Path.Combine(this.thisRoot, "app-data-BETA", "Data");
            this.thisProduction = Path.Combine(this.thisRoot, "app-data", "Data");
            this.thisBlobRoot = Path.Combine(this.thisRoot, "blobs");

            Directory.CreateDirectory(this.thisBeta);
            Directory.CreateDirectory(this.thisProduction);
            Directory.CreateDirectory(this.thisBlobRoot);

            // The generation the BETA publish writes into - see ApprovePublishFlowTests.
            File.WriteAllText(Path.Combine(this.thisBeta, "Classic-Repair-Toolbox.v2.0.0.xlsx"), "master v2");
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private ServerOptions Options(bool configured = true) => new()
        {
            DataTreeRoot = this.thisBeta,
            ProductionTreeRoot = Path.Combine(this.thisRoot, "app-data"),
            ProductionDataTreeRoot = configured ? this.thisProduction : null,
            ProductionManifestPath = configured ? Path.Combine(this.thisRoot, "app-data", "dataChecksums.json") : null,
            ProductionPublicDataBaseUrl = configured ? "https://example.com/app-data/Data" : null
        };

        private static ReviewAccess Admin() =>
            ReviewAccess.For(ProductionPromotionFlowTests.Account(1, administrator: true));

        private static ReviewAccess MaintainerOf(params string[] systems) =>
            ReviewAccess.For(ProductionPromotionFlowTests.Account(2, administrator: false), systems);

        private static AccountRecord Account(long id, bool administrator) =>
            new(id, $"a{id}@example.com", $"a{id}@example.com", "hash", $"A{id}",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false,
                ProductionPromotionFlowTests.Now, null);

        private BlobStore Blobs() => new(this.thisBlobRoot, NullLogger<BlobStore>.Instance);

        private static byte[] Png(string marker) => PublishExecutorTests.Png(marker);

        private async Task<string> PutBlobAsync(byte[] bytes)
        {
            string hash = Convert.ToHexStringLower(SHA256.HashData(bytes));

            BlobStore blobs = this.Blobs();
            using var content = new MemoryStream(bytes);
            await blobs.AppendChunkAsync(99, hash, 0, content, CancellationToken.None);
            await blobs.TryCompleteAsync(99, hash, CancellationToken.None);

            return hash;
        }

        // ###########################################################################################
        // Publishes a one-image submission into BETA, the way the Approve button does, and returns
        // the store holding the system's row with its new BETA content hash.
        // ###########################################################################################
        private async Task<FakeSubmissionStore> PublishToBetaAsync(
            FakeSubmissionStore? store = null,
            string image = "Commodore/C64/250407/Images/sheet1.png",
            string marker = "SHEET1",
            DateTimeOffset? when = null)
        {
            store ??= new FakeSubmissionStore();

            byte[] bytes = ProductionPromotionFlowTests.Png(marker);
            string hash = await this.PutBlobAsync(bytes);

            var manifest = new SubmissionManifest
            {
                SystemId = ProductionPromotionFlowTests.SystemId,
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Files = [new SubmissionFile { Path = image, Sha256 = hash, SizeBytes = bytes.LongLength }]
            };

            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = image });
            manifest.Rows.RevisionDate = "2026-September-25";
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA" });

            long id = await store.CreateAsync(
                new NewSubmission(
                    manifest.SystemId, manifest.Manufacturer, manifest.Hardware, manifest.Board,
                    null, "contributor@example.com", "192.0.2.1", "hash", "r0", "A change.", 1,
                    [.. manifest.Files], ProductionPromotionFlowTests.Now, ProductionPromotionFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SavePayloadAsync(id, manifest, CancellationToken.None);
            await store.SetStateAsync(id, SubmissionState.Pending, ProductionPromotionFlowTests.Now, CancellationToken.None);

            var approvals = new ApprovePublishFlow(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                new FakeAccountStore(),
                NullLogger<ApprovePublishFlow>.Instance);

            ApproveOutcome outcome = await approvals.ApproveAsync(
                id, ProductionPromotionFlowTests.Admin(), this.thisBeta, when ?? ProductionPromotionFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            return store;
        }

        private ProductionPromotionFlow Flow(FakeSubmissionStore store, FakeAccountStore? accounts = null) =>
            new(
                store,
                accounts ?? new FakeAccountStore(),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                new PublishLock(),
                NullLogger<ProductionPromotionFlow>.Instance);

        private static string BetaHashOf(FakeSubmissionStore store) =>
            store.PublishedSystems[ProductionPromotionFlowTests.SystemId].ContentHash;

        private string InProduction(string relative) =>
            Path.Combine(this.thisProduction, relative.Replace('/', Path.DirectorySeparatorChar));

        // -----------------------------------------------------------------------------------
        // The happy path
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_maintainer_of_the_system_publishes_it_to_production_and_the_board_arrives_whole()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsPublished, outcome.Error);

            // The image, the workbook and the sidecar - the whole board. No system.json: retired.
            Assert.True(File.Exists(this.InProduction("Commodore/C64/250407/Images/sheet1.png")));
            Assert.True(File.Exists(this.InProduction("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx")));
            Assert.True(File.Exists(this.InProduction("Commodore/C64/250407/Data C64 250407 v2.0.0.json")));
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/system.json")));

            // And the board READS BACK from production through the ordinary reader.
            BoardData? board = await BoardDataReader.LoadAsync(
                this.InProduction("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"), "promote-" + Guid.NewGuid().ToString("N"));

            Assert.Equal("U8", Assert.Single(board!.Components).BoardLabel);

            // Recorded, so it is no longer waiting.
            Assert.Equal(ProductionPromotionFlowTests.BetaHashOf(store), store.ProductionSystems[ProductionPromotionFlowTests.SystemId].ContentHash);
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(await store.FindSystemAsync(ProductionPromotionFlowTests.SystemId)));
        }

        [Fact]
        public async Task A_system_json_already_in_production_is_REMOVED_by_the_promotion()
        {
            // Retired (2026-09-25): one carried across by hand before promotions existed must not
            // stay where every user downloads it.
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            string leftover = this.InProduction("Commodore/C64/250407/system.json");
            Directory.CreateDirectory(Path.GetDirectoryName(leftover)!);
            await File.WriteAllTextAsync(leftover, "{}");

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.False(File.Exists(leftover));
        }

        [Fact]
        public async Task A_promotion_is_written_to_the_audit_trail()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            var accounts = new FakeAccountStore();

            await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now);

            AuditEntry entry = Assert.Single(accounts.Audit);
            Assert.Equal(ProductionPromotionFlow.PublishedAction, entry.Action);
            Assert.Equal(ProductionPromotionFlowTests.SystemId, entry.Subject);
        }

        [Fact]
        public async Task A_second_promotion_copies_ONLY_what_changed_in_BETA_since()
        {
            // The ordinary life of a board: a typo fix in BETA must not rewrite every image in
            // production.
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now);

            // A second submission to the same board adds one image.
            await this.PublishToBetaAsync(store, "Commodore/C64/250407/Images/sheet2.png", "SHEET2", ProductionPromotionFlowTests.Now.AddHours(1));

            PromotionPlanOutcome plan = await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options());

            Assert.Null(plan.Refusal);
            Assert.Contains(plan.Plan!.Files, file => file.Path.EndsWith("sheet2.png", StringComparison.Ordinal) && file.Change == PromotionChange.Added);
            Assert.DoesNotContain(plan.Plan.Files, file => file.Path.EndsWith("sheet1.png", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // "Only after he has checked it in BETA"
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task If_BETA_changed_since_the_maintainer_looked_NOTHING_is_copied()
        {
            // *** THE CHECKED-IN-BETA INTERLOCK. *** The maintainer checked one state; another
            // contribution was published into BETA after that. Promoting now would send out work
            // nobody looked at.
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            string checkedHash = ProductionPromotionFlowTests.BetaHashOf(store);

            await this.PublishToBetaAsync(store, "Commodore/C64/250407/Images/sheet2.png", "SHEET2", ProductionPromotionFlowTests.Now.AddHours(1));

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, checkedHash,
                this.Options(), ProductionPromotionFlowTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsConflict);
            Assert.Contains("changed in BETA", outcome.Error, StringComparison.Ordinal);
            Assert.False(Directory.Exists(this.InProduction("Commodore")));
        }

        [Fact]
        public async Task A_system_production_already_has_is_a_conflict_not_a_second_copy()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            string hash = ProductionPromotionFlowTests.BetaHashOf(store);

            await this.Flow(store).PromoteAsync(ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, hash, this.Options(), ProductionPromotionFlowTests.Now);

            PromotionOutcome again = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, hash, this.Options(), ProductionPromotionFlowTests.Now.AddMinutes(1));

            Assert.True(again.IsConflict);
        }

        // -----------------------------------------------------------------------------------
        // Who may
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_maintainer_of_ANOTHER_system_cannot_publish_it_to_production_and_nothing_arrives()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf("Commodore/C128/310378"),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsForbidden);
            Assert.False(Directory.Exists(this.InProduction("Commodore")));
        }

        // An account store in which the board HAS a maintainer (account 2, MaintainerOf's id).
        private static FakeAccountStore AccountsWithAMaintainer()
        {
            // The pool row AND its account: the store drops a row whose account is missing, as
            // the real join does.
            var accounts = new FakeAccountStore();
            accounts.Accounts[2] = ProductionPromotionFlowTests.Account(2, administrator: false);
            accounts.Maintainers.Add((ProductionPromotionFlowTests.SystemId, 2));
            return accounts;
        }

        [Fact]
        public async Task A_promotion_changing_a_SHARED_file_needs_the_maintainer_AND_the_administrator()
        {
            // The project owner's rule for production as for BETA (2026-09-25). The first approval is
            // recorded and NOTHING is copied; the second publishes.
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");
            FakeAccountStore accounts = ProductionPromotionFlowTests.AccountsWithAMaintainer();
            string hash = ProductionPromotionFlowTests.BetaHashOf(store);

            PromotionOutcome byMaintainer = await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, hash, this.Options(), ProductionPromotionFlowTests.Now);

            Assert.False(byMaintainer.IsPublished);
            Assert.True(byMaintainer.IsAwaitingApproval);
            Assert.Equal([ApproverRole.Administrator], byMaintainer.WaitingFor);
            Assert.False(File.Exists(this.InProduction("Commodore/Shared files/74LS08.png")));

            PromotionOutcome byAdmin = await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, hash, this.Options(), ProductionPromotionFlowTests.Now);

            Assert.True(byAdmin.IsPublished, byAdmin.Error);
            Assert.True(File.Exists(this.InProduction("Commodore/Shared files/74LS08.png")));
        }

        [Fact]
        public async Task A_production_approval_does_NOT_carry_over_once_BETA_has_changed()
        {
            // The maintainer approved one BETA state; another publish landed in BETA since. The
            // administrator's approval now starts again - nobody approved the new state twice.
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");
            FakeAccountStore accounts = ProductionPromotionFlowTests.AccountsWithAMaintainer();

            await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now);

            await this.PublishToBetaAsync(store, "Commodore/Shared files/74LS08.png", "SHARED-2", ProductionPromotionFlowTests.Now.AddHours(1));

            PromotionOutcome byAdmin = await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now.AddHours(2));

            Assert.True(byAdmin.IsAwaitingApproval);
            Assert.Equal([ApproverRole.Maintainer], byAdmin.WaitingFor);
            Assert.False(File.Exists(this.InProduction("Commodore/Shared files/74LS08.png")));
        }

        [Fact]
        public async Task A_promotion_with_NO_shared_change_is_published_by_the_maintainer_alone()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionOutcome outcome = await this.Flow(store, ProductionPromotionFlowTests.AccountsWithAMaintainer()).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsPublished, outcome.Error);
        }

        [Fact]
        public async Task The_list_offers_a_maintainer_only_their_own_systems()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            Assert.Single(await this.Flow(store).ListAwaitingAsync(ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId)));
            Assert.Empty(await this.Flow(store).ListAwaitingAsync(ProductionPromotionFlowTests.MaintainerOf("Commodore/C128/310378")));
            Assert.Single(await this.Flow(store).ListAwaitingAsync(ProductionPromotionFlowTests.Admin()));
        }

        // -----------------------------------------------------------------------------------
        // Switched off
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task With_the_production_settings_unset_nothing_can_be_promoted()
        {
            // A server configured before this existed must be exactly as unable to write
            // production as it always was.
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store), this.Options(configured: false), ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsNotConfigured);
            Assert.False(Directory.Exists(this.InProduction("Commodore")));

            Assert.True((await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options(configured: false))).IsNotConfigured);
        }
        // ###########################################################################################
        // WHAT A PROMOTION REMOVES FROM PRODUCTION (owner decisions, 2026-09-25: no orphan
        // files, and the list shown BEFORE anyone approves). Production holds the older board,
        // which cites a manual the BETA board no longer does.
        // ###########################################################################################
        private const string OldManual = "Commodore/C64/250407/old-manual.pdf";
        private const string Sheet = "Commodore/C64/250407/Images/sheet1.png";

        private void ProductionHasTheOlderBoard()
        {
            DataTreeBuilder.Master(this.thisProduction, DataTreeBuilder.Workbook);
            DataTreeBuilder.Board(this.thisProduction, DataTreeBuilder.Workbook, ProductionPromotionFlowTests.Sheet, ProductionPromotionFlowTests.OldManual);
            DataTreeBuilder.Files(this.thisProduction, ProductionPromotionFlowTests.OldManual);
        }

        [Fact]
        public async Task The_plan_shows_what_publishing_to_production_would_remove()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            this.ProductionHasTheOlderBoard();

            PromotionPlanOutcome plan = await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options());

            Assert.False(plan.Removals.IsBlocked, plan.Removals.BlockedBecause);
            Assert.Equal([ProductionPromotionFlowTests.OldManual], plan.Removals.Files);
        }

        [Fact]
        public async Task Publishing_to_production_with_the_list_shown_removes_exactly_that_list()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            this.ProductionHasTheOlderBoard();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now,
                [ProductionPromotionFlowTests.OldManual],
                CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal([ProductionPromotionFlowTests.OldManual], outcome.RemovedFiles);
            Assert.False(File.Exists(this.InProduction(ProductionPromotionFlowTests.OldManual)));
            Assert.True(File.Exists(this.InProduction(ProductionPromotionFlowTests.Sheet)));
        }

        // Nothing is copied and nothing removed when the list the maintainer saw is not the list the
        // promotion would now remove. An EMPTY list is a list: the maintainer was shown nothing.
        [Fact]
        public async Task Publishing_to_production_with_a_different_list_is_refused_and_nothing_is_copied()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            this.ProductionHasTheOlderBoard();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now,
                shownRemovals: [],
                CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsConflict);
            Assert.Equal(ProductionPromotionFlow.RemovalsChangedMessage, outcome.Error);
            Assert.True(File.Exists(this.InProduction(ProductionPromotionFlowTests.OldManual)));
            Assert.False(File.Exists(this.InProduction(ProductionPromotionFlowTests.Sheet)));
        }

        // ###########################################################################################
        // *** NO LIST AT ALL IS AN OLDER MAINTAINER APPLICATION, NOT A CHANGED ONE (code review,
        // 2026-09-25). *** The current one always sends its list, empty or not; one built before
        // the list existed sends none. It used to be told the list had "changed since you opened
        // it" - false, and reopening changed nothing, so it could never publish and never learn why.
        // ###########################################################################################
        [Fact]
        public async Task A_review_application_that_sends_no_list_is_told_to_update_and_nothing_is_copied()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            this.ProductionHasTheOlderBoard();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now,
                shownRemovals: null,
                CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(outcome.IsConflict);
            Assert.Equal(ApprovePublishFlow.RemovalsNotSentMessage, outcome.Error);
            Assert.True(File.Exists(this.InProduction(ProductionPromotionFlowTests.OldManual)));
            Assert.False(File.Exists(this.InProduction(ProductionPromotionFlowTests.Sheet)));
        }
    }
}
