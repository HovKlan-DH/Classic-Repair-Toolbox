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

            // The generation the BETA publish writes into - see ApprovePublishFlowTests. A real
            // workbook: the fixture's system is new, and its publish adds its row (2026-09-27).
            DataTreeBuilder.ListingMaster(this.thisBeta, ApprovePublishFlowTests.OtherListedBoard);
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

        // `administratorsOnly` is OFF here unless a test asks for it, although the server defaults
        // it ON: the tests below describe a maintainer publishing their own systems, as designed
        // (2026-09-25). The "for now, only the administrator" rule (2026-10-05) has its own tests at
        // the end of the class.
        private ServerOptions Options(bool configured = true, bool administratorsOnly = false) => new()
        {
            DataTreeRoot = this.thisBeta,
            ProductionTreeRoot = Path.Combine(this.thisRoot, "app-data"),
            ProductionDataTreeRoot = configured ? this.thisProduction : null,
            ProductionManifestPath = configured ? Path.Combine(this.thisRoot, "app-data", "dataChecksums.json") : null,
            ProductionPublicDataBaseUrl = configured ? "https://example.com/app-data/Data" : null,
            ProductionPublishingAdministratorsOnly = administratorsOnly
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

            long id = await this.PendingAsync(store, image, marker);

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

        // A pending one-image submission of the fixture's system, placed in the drop-down lists.
        private async Task<long> PendingAsync(
            FakeSubmissionStore store,
            string image = "Commodore/C64/250407/Images/sheet1.png",
            string marker = "SHEET1")
        {
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

            // A new system is placed in the drop-down lists before it can reach BETA (2026-09-27).
            await store.SetPlacementAsync(manifest.SystemId, ApprovePublishFlowTests.Placement(), 1, ProductionPromotionFlowTests.Now, CancellationToken.None);

            return id;
        }

        // What the project owner does by hand as root: the whole BETA data tree copied into production.
        private void CopyBetaToProductionByHand()
        {
            foreach (string file in Directory.EnumerateFiles(this.thisBeta, "*", SearchOption.AllDirectories))
            {
                string target = Path.Combine(this.thisProduction, Path.GetRelativePath(this.thisBeta, file));
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                File.Copy(file, target, overwrite: true);
            }
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
        public async Task A_promotion_REPLACING_a_SHARED_file_needs_the_maintainer_AND_the_administrator()
        {
            // The project owner's rule for production as for BETA (2026-09-25) - since 2026-09-27 for
            // a REPLACEMENT only: production already has the shared file, differently. The first
            // approval is recorded and NOTHING is copied; the second publishes.
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");
            this.WriteInProduction("Commodore/Shared files/74LS08.png", "OLD");
            FakeAccountStore accounts = ProductionPromotionFlowTests.AccountsWithAMaintainer();
            string hash = ProductionPromotionFlowTests.BetaHashOf(store);

            PromotionOutcome byMaintainer = await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, hash, this.Options(), ProductionPromotionFlowTests.Now);

            Assert.False(byMaintainer.IsPublished);
            Assert.True(byMaintainer.IsAwaitingApproval);
            Assert.Equal([ApproverRole.Administrator], byMaintainer.WaitingFor);
            Assert.Equal("OLD", File.ReadAllText(this.InProduction("Commodore/Shared files/74LS08.png")));

            PromotionOutcome byAdmin = await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, hash, this.Options(), ProductionPromotionFlowTests.Now);

            Assert.True(byAdmin.IsPublished, byAdmin.Error);
            Assert.NotEqual("OLD", File.ReadAllText(this.InProduction("Commodore/Shared files/74LS08.png")));
        }

        // A NEW shared file reaches no board that did not ask for it - the maintainer alone promotes
        // it (owner decision, 2026-09-27).
        [Fact]
        public async Task A_promotion_ADDING_a_new_shared_file_is_published_by_the_maintainer_alone()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");

            PromotionOutcome outcome = await this.Flow(store, ProductionPromotionFlowTests.AccountsWithAMaintainer()).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, ProductionPromotionFlowTests.BetaHashOf(store), this.Options(), ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.True(File.Exists(this.InProduction("Commodore/Shared files/74LS08.png")));
        }

        // Writes a file into the production tree, folders and all.
        private void WriteInProduction(string relative, string text)
        {
            string full = this.InProduction(relative);
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllText(full, text);
        }

        [Fact]
        public async Task A_production_approval_does_NOT_carry_over_once_BETA_has_changed()
        {
            // The maintainer approved one BETA state; another publish landed in BETA since. The
            // administrator's approval now starts again - nobody approved the new state twice. (A
            // REPLACEMENT of production's shared file, the kind that needs two.)
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");
            this.WriteInProduction("Commodore/Shared files/74LS08.png", "OLD");
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
            Assert.Equal("OLD", File.ReadAllText(this.InProduction("Commodore/Shared files/74LS08.png")));
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

        // ###########################################################################################
        // *** A BOARD COPIED TO PRODUCTION BY HAND IS NOT WAITING FOR PRODUCTION (code review,
        // 2026-09-29). *** Its record said it waited for ever - it was never PROMOTED - so it sat in
        // Beta > Prod and blocked every new approval of the system. The list now asks the trees:
        // production already holding exactly BETA's state is recorded as in production, audited,
        // and left out.
        // ###########################################################################################
        [Fact]
        public async Task A_board_copied_to_production_by_hand_is_not_waiting_and_is_recorded_as_in_production()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            var accounts = new FakeAccountStore();

            this.CopyBetaToProductionByHand();

            IReadOnlyList<SystemRecord> waiting = await this.Flow(store, accounts).ListAwaitingAsync(
                ProductionPromotionFlowTests.Admin(), this.Options(), ProductionPromotionFlowTests.Now.AddHours(1));

            Assert.Empty(waiting);
            Assert.Equal(ProductionPromotionFlowTests.BetaHashOf(store), store.ProductionSystems[ProductionPromotionFlowTests.SystemId].ContentHash);
            Assert.Contains(accounts.Audit, entry =>
                entry.Action == SystemHistoryEvents.FoundInProduction && entry.Subject == ProductionPromotionFlowTests.SystemId);
        }

        // The other half, which matters as much: a system production genuinely lacks keeps waiting,
        // and nothing is recorded for it.
        [Fact]
        public async Task A_board_production_does_not_have_keeps_waiting_and_nothing_is_recorded()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            var accounts = new FakeAccountStore();

            IReadOnlyList<SystemRecord> waiting = await this.Flow(store, accounts).ListAwaitingAsync(
                ProductionPromotionFlowTests.Admin(), this.Options(), ProductionPromotionFlowTests.Now.AddHours(1));

            Assert.Single(waiting);
            Assert.False(store.ProductionSystems.ContainsKey(ProductionPromotionFlowTests.SystemId));
            Assert.DoesNotContain(accounts.Audit, entry => entry.Action == SystemHistoryEvents.FoundInProduction);
        }

        // ###########################################################################################
        // The same check behind the one-in-BETA rule: a second submission of a system whose BETA
        // state was copied to production by hand is approved, where it was refused for ever.
        // ###########################################################################################
        [Fact]
        public async Task A_second_approval_goes_through_once_production_holds_BETA_by_hand()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            var accounts = new FakeAccountStore();

            var approvals = new ApprovePublishFlow(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                accounts,
                NullLogger<ApprovePublishFlow>.Instance,
                publishLock: null,
                options: this.Options(),
                production: this.Flow(store, accounts));

            long second = await this.PendingAsync(store, "Commodore/C64/250407/Images/sheet2.png", "SHEET2");

            // Still genuinely waiting: refused, as the rule says.
            ApproveOutcome refused = await approvals.ApproveAsync(
                second, ProductionPromotionFlowTests.Admin(), this.thisBeta, ProductionPromotionFlowTests.Now.AddMinutes(1), CancellationToken.None);

            Assert.False(refused.IsPublished);
            Assert.Equal(OneSubmissionInBeta.BusyMessage(ProductionPromotionFlowTests.SystemId), refused.Error);

            // Copied to production by hand: the same approval goes through.
            this.CopyBetaToProductionByHand();

            ApproveOutcome published = await approvals.ApproveAsync(
                second, ProductionPromotionFlowTests.Admin(), this.thisBeta, ProductionPromotionFlowTests.Now.AddMinutes(2), CancellationToken.None);

            Assert.True(published.IsPublished, published.Error);
        }

        // ###########################################################################################
        // *** THE LIST ASKS A FIXED NUMBER OF QUERIES, HOWEVER MANY SYSTEMS WAIT (code review,
        // 2026-09-29). *** Every open Maintainer tab reads it every minute, and it used to ask three
        // queries per waiting system to set two booleans. Four systems here: one this account has
        // already approved for production, one carrying a submission whose contributor discarded
        // their draft - both flags still right, from one approvals read and one discards read.
        // ###########################################################################################
        [Fact]
        public async Task The_list_reads_its_two_facts_once_for_every_waiting_system_together()
        {
            var store = new FakeSubmissionStore();
            string[] ids = ["Commodore/C64/250407", "Commodore/C64/250425", "Commodore/C128/310378", "Commodore/VIC20/250403"];

            foreach (string id in ids)
            {
                string[] parts = id.Split('/');
                await store.EnsureSystemAsync(id, parts[0], parts[1], parts[2], "pipeline", ProductionPromotionFlowTests.Now);
                store.PublishedSystems[id] = new PublishedSystemRow("2026-September-25", "beta-" + parts[2], ProductionPromotionFlowTests.Now);
            }

            // The administrator (account 1) has approved the second one's BETA state for production.
            store.ProductionApprovals[(ids[1], "beta-250425")] =
                [new GivenApproval(ApproverRole.Administrator, "A1", ProductionPromotionFlowTests.Now, 1)];

            // A merged submission of the third, whose contributor has since discarded their draft.
            long merged = await store.CreateAsync(
                new NewSubmission(
                    ids[2], "Commodore", "C128", "310378", null, "c@example.com", "192.0.2.1", "hash", "r0",
                    "A fix.", 1, [], ProductionPromotionFlowTests.Now.AddDays(-2), ProductionPromotionFlowTests.Now.AddDays(-1)),
                CancellationToken.None);
            await store.SetStateAsync(merged, SubmissionState.Merged, ProductionPromotionFlowTests.Now.AddDays(-1), CancellationToken.None);
            await store.RecordDraftDiscardedAsync(merged, ProductionPromotionFlowTests.Now.AddHours(-1), CancellationToken.None);

            IReadOnlyList<ProductionListEntry> entries = await this.Flow(store).ListEntriesAsync(
                ProductionPromotionFlowTests.Admin(), this.Options(), ProductionPromotionFlowTests.Now);

            Assert.Equal(4, entries.Count);
            Assert.False(entries.Single(entry => entry.SystemId == ids[1]).AwaitsYou);
            Assert.True(entries.Single(entry => entry.SystemId == ids[0]).AwaitsYou);
            Assert.True(entries.Single(entry => entry.SystemId == ids[2]).CarriesDiscardedDraft);
            Assert.False(entries.Single(entry => entry.SystemId == ids[3]).CarriesDiscardedDraft);

            // One read of each fact for the whole list - never one per system.
            Assert.Equal(1, store.ProductionApprovalReads);
            Assert.Equal(1, store.DraftDiscardReads);
            Assert.Equal(0, store.MergedSubmissionReads);
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

        // ###########################################################################################
        // *** PRODUCTION TOO REMOVES ONLY INSIDE THE SYSTEM'S OWN FOLDER (owner decision, 2026-09-27).
        // *** The older production board also cited a shared datasheet nothing else uses; the BETA
        // board does not. It stays in production, unused, for Account > Unused files - and is not on
        // the list the maintainer is shown.
        // ###########################################################################################
        [Fact]
        public async Task A_shared_file_the_board_stops_using_is_not_removed_from_production()
        {
            const string shared = "Commodore/Shared files/Datasheets/old-chip.pdf";

            FakeSubmissionStore store = await this.PublishToBetaAsync();
            DataTreeBuilder.Master(this.thisProduction, DataTreeBuilder.Workbook);
            DataTreeBuilder.Board(this.thisProduction, DataTreeBuilder.Workbook, ProductionPromotionFlowTests.Sheet, ProductionPromotionFlowTests.OldManual, shared);
            DataTreeBuilder.Files(this.thisProduction, ProductionPromotionFlowTests.OldManual, shared);

            PromotionPlanOutcome plan = await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options());

            Assert.Equal([ProductionPromotionFlowTests.OldManual], plan.Removals.Files);

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now,
                [ProductionPromotionFlowTests.OldManual],
                CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.False(File.Exists(this.InProduction(ProductionPromotionFlowTests.OldManual)));
            Assert.True(File.Exists(this.InProduction(shared)));
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

        // -----------------------------------------------------------------------------------
        // A NEW system's row in production's drop-down lists (owner decision, 2026-09-27: "insert
        // at the same place").
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Promoting_a_NEW_system_adds_its_row_to_productions_list_after_the_same_neighbour()
        {
            MasterListingRow first = new("Commodore VIC-20", "324003", "Commodore/VIC-20/324003/Data VIC20 324003 v2.0.0.xlsx", string.Empty);
            DataTreeBuilder.ListingMaster(this.thisProduction, first, ApprovePublishFlowTests.OtherListedBoard);

            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now);

            Assert.True(outcome.IsPublished, outcome.Error);

            // After the Amstrad board, as in BETA - and BETA's own row, names and notes included.
            Assert.Equal(
                [first, ApprovePublishFlowTests.OtherListedBoard, DataTreeBuilder.ListedIn(this.thisBeta)[1]],
                DataTreeBuilder.ListedIn(this.thisProduction));
            Assert.Equal("Commodore 64", DataTreeBuilder.ListedIn(this.thisProduction)[2].HardwareName);
        }

        // Production can list another system under the names BETA gave this one - refused on the
        // plan, with nothing copied.
        [Fact]
        public async Task A_system_production_lists_another_system_under_the_names_of_is_refused_before_anything_is_copied()
        {
            MasterListingRow sameNames = new("Commodore 64", "250407", "Commodore/C64/326298/Data C64 326298 v2.0.0.xlsx", string.Empty);
            DataTreeBuilder.ListingMaster(this.thisProduction, ApprovePublishFlowTests.OtherListedBoard, sameNames);

            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionPlanOutcome plan = await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options());

            Assert.Contains("is already in the drop-down lists", plan.Refusal);

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx")));
        }

        // ###########################################################################################
        // A system neither list carries would reach production and be seen by nobody. Refused - and
        // shown on the plan before anybody presses the button - with nothing copied.
        // ###########################################################################################
        [Fact]
        public async Task A_system_neither_list_carries_is_refused_before_anything_is_copied()
        {
            DataTreeBuilder.ListingMaster(this.thisProduction, ApprovePublishFlowTests.OtherListedBoard);
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            // BETA's list loses the row (edited by hand, say) after the publish added it.
            DataTreeBuilder.ListingMaster(this.thisBeta, ApprovePublishFlowTests.OtherListedBoard);

            PromotionPlanOutcome plan = await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options());

            Assert.Contains("not in the drop-down lists", plan.Refusal);

            PromotionOutcome outcome = await this.Flow(store).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId,
                ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(),
                ProductionPromotionFlowTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx")));
        }

        // ###########################################################################################
        // *** FOR NOW, ONLY THE ADMINISTRATOR PUBLISHES TO STABLE (owner request, 2026-10-05: "I do
        // not want to pollute the stable yet. Only me, as admin, should be able to publish to
        // stable"). *** ServerOptions.ProductionPublishingAdministratorsOnly, ON by default. A
        // maintainer still sees the system and its plan, and may push it back or reject it - only
        // the publish is closed to them.
        // ###########################################################################################

        // The default is what makes "for now" true on a server nobody re-configured: forgetting the
        // setting must leave stable closed to maintainers, not open.
        [Fact]
        public void A_server_not_told_otherwise_lets_only_administrators_publish_to_stable()
        {
            Assert.True(new ServerOptions().ProductionPublishingAdministratorsOnly);
        }

        [Fact]
        public async Task While_only_administrators_publish_a_maintainer_is_refused_and_nothing_arrives_in_stable()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            string hash = ProductionPromotionFlowTests.BetaHashOf(store);

            PromotionOutcome outcome = await this.Flow(store, ProductionPromotionFlowTests.AccountsWithAMaintainer()).PromoteAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, hash, this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now,
                [], CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsForbidden);
            Assert.Equal(StablePublishing.AdministratorsOnlyMessage, outcome.Error);
            Assert.False(Directory.Exists(this.InProduction("Commodore")));

            // Nor is half an approval left behind for the administrator's to complete.
            Assert.False(store.ProductionApprovals.ContainsKey((ProductionPromotionFlowTests.SystemId, hash)));
            Assert.False(store.ProductionSystems.ContainsKey(ProductionPromotionFlowTests.SystemId));
        }

        [Fact]
        public async Task While_only_administrators_publish_the_administrator_still_publishes_to_stable()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();

            PromotionOutcome outcome = await this.Flow(store, ProductionPromotionFlowTests.AccountsWithAMaintainer()).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId, ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now, [], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.True(File.Exists(this.InProduction(ProductionPromotionFlowTests.Sheet)));
        }

        // ###########################################################################################
        // A REPLACED shared file normally needs the board's maintainer AND the administrator. With
        // the maintainer shut out of the publish, asking for their half would leave the
        // administrator's approval waiting for ever - so the administrator's alone publishes.
        // ###########################################################################################
        [Fact]
        public async Task While_only_administrators_publish_a_shared_file_replacement_needs_the_administrator_alone()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");
            this.WriteInProduction("Commodore/Shared files/74LS08.png", "OLD");
            FakeAccountStore accounts = ProductionPromotionFlowTests.AccountsWithAMaintainer();

            PromotionPlanOutcome plan = await this.Flow(store, accounts).PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options(administratorsOnly: true));

            Assert.True(plan.Plan!.TouchesSharedFiles);
            Assert.True(plan.Approval!.ApprovalPublishes);

            // Still SAID to be a shared-file replacement: Required is what CRT's "Replaces a shared
            // file other boards may use" line is written from, and a stable publish cannot be undone.
            // Collapsing it to "any one approval" took that warning off the administrator's screen
            // (code review, 2026-10-05).
            Assert.Equal([ApproverRole.Administrator], plan.Approval.Required);

            // And the maintainer is not offered a publish they cannot make: the administrator's
            // approval is the one needed, theirs is not asked for.
            PromotionPlanOutcome maintainers = await this.Flow(store, accounts).PlanAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, this.Options(administratorsOnly: true));

            Assert.False(maintainers.Approval!.CanApprove);
            Assert.Equal([ApproverRole.Administrator], maintainers.Approval.WaitingFor);

            PromotionOutcome outcome = await this.Flow(store, accounts).PromoteAsync(
                ProductionPromotionFlowTests.Admin(),
                ProductionPromotionFlowTests.SystemId, ProductionPromotionFlowTests.BetaHashOf(store),
                this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now, [], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.NotEqual("OLD", File.ReadAllText(this.InProduction("Commodore/Shared files/74LS08.png")));
        }

        // ###########################################################################################
        // *** AN APPROVAL GIVEN BEFORE THE SETTING WAS SWITCHED ON (code review, 2026-10-05). *** The
        // administrator gave the first of two approvals on a shared-file replacement; with only
        // administrators publishing, nothing else is asked for, so their next press publishes. The
        // list must say so: "you approved, so it is with the other approver" left the system dimmed
        // and off the administrator's badge, with nobody prompted to finish it.
        // ###########################################################################################
        [Fact]
        public async Task An_administrators_approval_given_before_only_administrators_published_still_leaves_the_system_awaiting_them()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync(image: "Commodore/Shared files/74LS08.png", marker: "SHARED");
            this.WriteInProduction("Commodore/Shared files/74LS08.png", "OLD");
            FakeAccountStore accounts = ProductionPromotionFlowTests.AccountsWithAMaintainer();
            ProductionPromotionFlow flow = this.Flow(store, accounts);
            string hash = ProductionPromotionFlowTests.BetaHashOf(store);

            PromotionOutcome first = await flow.PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, hash,
                this.Options(administratorsOnly: false), ProductionPromotionFlowTests.Now, [], CancellationToken.None);

            Assert.True(first.IsAwaitingApproval);

            ProductionListEntry entry = Assert.Single(await flow.ListEntriesAsync(
                ProductionPromotionFlowTests.Admin(), this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now));

            Assert.True(entry.AwaitsYou);

            PromotionOutcome second = await flow.PromoteAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, hash,
                this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now, [], CancellationToken.None);

            Assert.True(second.IsPublished, second.Error);
        }

        // ###########################################################################################
        // The plan is what CRT's button follows (ProductionPlanAnswer.CanPublish needs a null
        // refusal): a maintainer's carries the reason, the administrator's does not. Everything
        // else - the files, the removals - is still shown to the maintainer, who may push it back.
        // ###########################################################################################
        [Fact]
        public async Task While_only_administrators_publish_a_maintainers_plan_carries_the_reason_and_the_administrators_does_not()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            this.ProductionHasTheOlderBoard();
            ProductionPromotionFlow flow = this.Flow(store, ProductionPromotionFlowTests.AccountsWithAMaintainer());

            PromotionPlanOutcome maintainers = await flow.PlanAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, this.Options(administratorsOnly: true));

            Assert.False(maintainers.IsForbidden);
            Assert.Equal(StablePublishing.AdministratorsOnlyMessage, maintainers.Refusal);
            Assert.Equal([ProductionPromotionFlowTests.OldManual], maintainers.Removals.Files);

            PromotionPlanOutcome administrators = await flow.PlanAsync(
                ProductionPromotionFlowTests.Admin(), ProductionPromotionFlowTests.SystemId, this.Options(administratorsOnly: true));

            Assert.Null(administrators.Refusal);

            // Switched off, the maintainer's plan is as it always was.
            PromotionPlanOutcome open = await flow.PlanAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, this.Options(administratorsOnly: false));

            Assert.Null(open.Refusal);
        }

        // A plan refusal the maintainer CAN act on (placing the system) is still said, first.
        [Fact]
        public async Task While_only_administrators_publish_a_maintainer_is_still_told_the_plans_own_refusal_first()
        {
            DataTreeBuilder.ListingMaster(this.thisProduction, ApprovePublishFlowTests.OtherListedBoard);
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            DataTreeBuilder.ListingMaster(this.thisBeta, ApprovePublishFlowTests.OtherListedBoard);

            PromotionPlanOutcome plan = await this.Flow(store).PlanAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                ProductionPromotionFlowTests.SystemId, this.Options(administratorsOnly: true));

            Assert.StartsWith("This system is not in the drop-down lists", plan.Refusal, StringComparison.Ordinal);
            Assert.EndsWith(StablePublishing.AdministratorsOnlyMessage, plan.Refusal, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The maintainer still SEES the system in the BETA queue - they may push it back or reject
        // it - but it never awaits them, so the tab's badge does not count work they cannot do.
        // ###########################################################################################
        [Fact]
        public async Task While_only_administrators_publish_the_list_shows_a_maintainer_the_system_but_never_as_awaiting_them()
        {
            FakeSubmissionStore store = await this.PublishToBetaAsync();
            ProductionPromotionFlow flow = this.Flow(store);

            ProductionListEntry maintainers = Assert.Single(await flow.ListEntriesAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now));

            Assert.False(maintainers.AwaitsYou);

            // ...and says it waits for the administrator, so CRT does not call it "with the other
            // approver" - which the maintainer never was (code review, 2026-10-05).
            Assert.True(maintainers.WaitsForAdministrator);

            ProductionListEntry administrators = Assert.Single(await flow.ListEntriesAsync(
                ProductionPromotionFlowTests.Admin(), this.Options(administratorsOnly: true), ProductionPromotionFlowTests.Now));

            Assert.True(administrators.AwaitsYou);
            Assert.False(administrators.WaitsForAdministrator);

            ProductionListEntry open = Assert.Single(await flow.ListEntriesAsync(
                ProductionPromotionFlowTests.MaintainerOf(ProductionPromotionFlowTests.SystemId),
                this.Options(administratorsOnly: false), ProductionPromotionFlowTests.Now));

            Assert.True(open.AwaitsYou);
            Assert.False(open.WaitsForAdministrator);
        }
    }
}
