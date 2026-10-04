using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // SystemOrderFlow - the administrator putting CRT's drop-down lists in order (owner request,
    // 2026-10-04: "a possibility to be able to sort the list of systems, which then gets saved to both
    // sources (BETA + stable) after my save"). Two real trees, each with a real main Excel data file;
    // the manifest writer is a fake that records which trees it was asked to rebuild.
    // ###########################################################################################
    public sealed class SystemOrderFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

        private static readonly MasterListingRow C64 =
            new("Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", string.Empty);

        private static readonly MasterListingRow C128 =
            new("Commodore 128", "310378 (C128 & C128D)", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx", "Has 6581 SID.");

        private static readonly MasterListingRow Spectrum =
            new("ZX Spectrum 16K/48K", "Issue 4B", "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Issue 4B v2.0.0.xlsx", string.Empty);

        // In BETA only - not promoted yet.
        private static readonly MasterListingRow Open128 =
            new("Commodore 128", "310378 Open128", "Commodore/C128/310378 Open128/Data C128 310378 Open128 v2.0.0.xlsx", string.Empty);

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-order", Guid.NewGuid().ToString("N"));

        private string Beta => Path.Combine(this.thisRoot, "beta");

        private string Stable => Path.Combine(this.thisRoot, "stable");

        private readonly List<string> thisRebuilt = [];

        private readonly FakeAccountStore thisAccounts = new();

        public SystemOrderFlowTests()
        {
            Directory.CreateDirectory(this.Beta);
            Directory.CreateDirectory(this.Stable);

            DataTreeBuilder.ListingMaster(this.Beta, C64, C128, Open128, Spectrum);
            DataTreeBuilder.ListingMaster(this.Stable, C64, C128, Spectrum);
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

        private ServerOptions Options(bool stable = true)
        {
            var options = new ServerOptions
            {
                DataTreeRoot = this.Beta,
                ManifestPath = Path.Combine(this.thisRoot, "beta.json"),
                PublicDataBaseUrl = "https://example.invalid/app-data-BETA/",
            };

            if (stable)
            {
                options.ProductionDataTreeRoot = this.Stable;
                options.ProductionManifestPath = Path.Combine(this.thisRoot, "stable.json");
                options.ProductionPublicDataBaseUrl = "https://example.invalid/app-data/";
            }

            return options;
        }

        private static ReviewAccess Administrator() =>
            ReviewAccess.For(new AccountRecord(
                7, "dennis@example.com", "dennis@example.com", "hash", "Dennis",
                IsVerified: true, IsAdministrator: true, IsLocked: false, SystemOrderFlowTests.Now, null));

        private Task<SystemOrderOutcome> Order(ServerOptions options, params MasterListingRow[] rows) =>
            SystemOrderFlow.SetAsync(
                SystemOrderFlowTests.Administrator(),
                new SystemOrderRequest(rows.Select(row => row.SystemId).ToList()),
                options,
                tree =>
                {
                    this.thisRebuilt.Add(tree.Name);
                    return 100;
                },
                new PublishLock(),
                this.thisAccounts,
                SystemOrderFlowTests.Now);

        private static IReadOnlyList<string> Boards(string root) =>
            DataTreeBuilder.ListedIn(root).Select(row => row.BoardName).ToList();

        // ###########################################################################################
        // BETA's list takes the order sent; the stable source's follows it - its own systems in
        // BETA's order, the one it lacks (Open128) simply not there - and each tree's manifest is
        // rebuilt, so CRT downloads the new order.
        // ###########################################################################################
        [Fact]
        public async Task The_order_is_written_into_BETA_and_the_stable_source_and_both_manifests_are_rebuilt()
        {
            SystemOrderOutcome outcome = await this.Order(this.Options(), Spectrum, Open128, C128, C64);

            Assert.NotNull(outcome.Answer);
            Assert.Equal(new SystemOrderAnswer(true, true, null), outcome.Answer);

            Assert.Equal(["Issue 4B", "310378 Open128", "310378 (C128 & C128D)", "250407 (long board)"], SystemOrderFlowTests.Boards(this.Beta));
            Assert.Equal(["Issue 4B", "310378 (C128 & C128D)", "250407 (long board)"], SystemOrderFlowTests.Boards(this.Stable));

            Assert.Equal(["beta", "production"], this.thisRebuilt);

            AuditEntry audit = Assert.Single(this.thisAccounts.Audit);
            Assert.Equal(SystemOrderFlow.OrderedAction, audit.Action);
        }

        // ###########################################################################################
        // *** THE WHOLE LIST OR NOTHING. *** An order naming fewer systems than BETA lists - one was
        // placed while the list was on screen - is refused, and neither file is touched.
        // ###########################################################################################
        [Fact]
        public async Task An_order_that_does_not_name_exactly_BETAs_systems_is_refused_and_nothing_is_written()
        {
            byte[] beta = File.ReadAllBytes(Path.Combine(this.Beta, "Classic-Repair-Toolbox.v2.0.0.xlsx"));
            byte[] stable = File.ReadAllBytes(Path.Combine(this.Stable, "Classic-Repair-Toolbox.v2.0.0.xlsx"));

            SystemOrderOutcome outcome = await this.Order(this.Options(), Spectrum, C128, C64);

            Assert.Null(outcome.Answer);
            Assert.True(outcome.IsConflict);
            Assert.Equal(SystemOrderFlow.ChangedMessage, outcome.Error);

            Assert.Equal(beta, File.ReadAllBytes(Path.Combine(this.Beta, "Classic-Repair-Toolbox.v2.0.0.xlsx")));
            Assert.Equal(stable, File.ReadAllBytes(Path.Combine(this.Stable, "Classic-Repair-Toolbox.v2.0.0.xlsx")));
            Assert.Empty(this.thisRebuilt);
            Assert.Empty(this.thisAccounts.Audit);
        }

        [Theory]
        [InlineData(new string[0], "No systems")]
        [InlineData(new[] { "Commodore/C64/250407", "commodore/c64/250407" }, "more than once")]
        [InlineData(new[] { "Commodore/C64/250407", " " }, "no id")]
        public async Task A_list_that_is_not_one_is_refused(string[] ids, string expected)
        {
            SystemOrderOutcome outcome = await SystemOrderFlow.SetAsync(
                SystemOrderFlowTests.Administrator(),
                new SystemOrderRequest(ids),
                this.Options(),
                _ => 1,
                new PublishLock(),
                this.thisAccounts,
                SystemOrderFlowTests.Now);

            Assert.False(outcome.IsConflict);
            Assert.Contains(expected, outcome.Error, StringComparison.Ordinal);
        }

        // An order the lists already have writes nothing, rebuilds nothing and records nothing.
        [Fact]
        public async Task The_order_the_lists_already_have_changes_nothing()
        {
            SystemOrderOutcome outcome = await this.Order(this.Options(), C64, C128, Open128, Spectrum);

            Assert.Equal(new SystemOrderAnswer(false, false, null), outcome.Answer);
            Assert.Empty(this.thisRebuilt);
            Assert.Empty(this.thisAccounts.Audit);
        }

        // A server with no stable source: BETA alone, and the answer says there is no stable list.
        [Fact]
        public async Task Without_a_stable_source_only_BETA_is_written()
        {
            SystemOrderOutcome outcome = await this.Order(this.Options(stable: false), Spectrum, Open128, C128, C64);

            Assert.Equal(new SystemOrderAnswer(true, null, null), outcome.Answer);
            Assert.Equal(["beta"], this.thisRebuilt);
            Assert.Equal(["250407 (long board)", "310378 (C128 & C128D)", "Issue 4B"], SystemOrderFlowTests.Boards(this.Stable));
        }

        // ###########################################################################################
        // *** BETA FIRST, AND A LATER FAILURE IS SAID, NOT HIDDEN. *** A stable source with no main
        // Excel data file to write: BETA is saved all the same, and the answer says what was not.
        // ###########################################################################################
        [Fact]
        public async Task A_stable_list_that_cannot_be_written_is_said_after_BETA_was_saved()
        {
            File.Delete(Path.Combine(this.Stable, "Classic-Repair-Toolbox.v2.0.0.xlsx"));

            SystemOrderOutcome outcome = await this.Order(this.Options(), Spectrum, Open128, C128, C64);

            Assert.True(outcome.Answer!.BetaChanged);
            Assert.False(outcome.Answer.StableChanged);
            Assert.Contains("stable source", outcome.Answer.Problem, StringComparison.Ordinal);
            Assert.Equal("Issue 4B", SystemOrderFlowTests.Boards(this.Beta)[0]);
        }

        // A manifest that could not be rebuilt is said too - CRT would not see the new order.
        [Fact]
        public async Task A_manifest_that_could_not_be_rebuilt_is_said()
        {
            SystemOrderOutcome outcome = await SystemOrderFlow.SetAsync(
                SystemOrderFlowTests.Administrator(),
                new SystemOrderRequest([Spectrum.SystemId, Open128.SystemId, C128.SystemId, C64.SystemId]),
                this.Options(),
                tree => tree.Name == "beta" ? -1 : 10,
                new PublishLock(),
                this.thisAccounts,
                SystemOrderFlowTests.Now);

            Assert.True(outcome.Answer!.BetaChanged);
            Assert.Contains("checksum manifest of BETA", outcome.Answer.Problem, StringComparison.Ordinal);

            // And where to rebuild it, by the Maintainer tab's own name for the screen (2026-10-04).
            Assert.Contains("use \"Account\" > \"Rebuild checksum manifests\"", outcome.Answer.Problem, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** ONCE BETA's LIST IS WRITTEN, A CLIENT GIVING UP STOPS NOTHING (code review,
        // 2026-10-04). *** The request is cancelled while BETA's manifest is being hashed - CRT hit
        // its wait limit, or was closed. The stable list, its manifest and the audit row still
        // follow: stopping there left a stable manifest naming the old file's checksum, so every
        // CRT threw the download away on each sync until somebody rebuilt the manifests by hand.
        // ###########################################################################################
        [Fact]
        public async Task A_request_cancelled_after_BETAs_list_is_written_still_finishes_the_stable_source()
        {
            using var cancellation = new CancellationTokenSource();

            SystemOrderOutcome outcome = await SystemOrderFlow.SetAsync(
                SystemOrderFlowTests.Administrator(),
                new SystemOrderRequest([Spectrum.SystemId, Open128.SystemId, C128.SystemId, C64.SystemId]),
                this.Options(),
                tree =>
                {
                    this.thisRebuilt.Add(tree.Name);

                    if (tree.Name == "beta")
                        cancellation.Cancel();

                    return 100;
                },
                new PublishLock(),
                this.thisAccounts,
                SystemOrderFlowTests.Now,
                cancellation.Token);

            Assert.Equal(new SystemOrderAnswer(true, true, null), outcome.Answer);
            Assert.Equal(["Issue 4B", "310378 (C128 & C128D)", "250407 (long board)"], SystemOrderFlowTests.Boards(this.Stable));
            Assert.Equal(["beta", "production"], this.thisRebuilt);
            Assert.Single(this.thisAccounts.Audit);
        }
    }
}
