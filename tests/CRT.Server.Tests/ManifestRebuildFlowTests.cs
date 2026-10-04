using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;

namespace CRT.Server.Tests;

// ###########################################################################################
// REBUILDING dataChecksums.json BY HAND (owner request, 2026-10-01) - ManifestRebuildFlow.
//
// The button exists because the project owner edits files in the trees as root, which the service
// never sees. So what matters here is not that a manifest gets written - DataChecksumManifest's
// own tests cover that - but that the flow decides the right things around it: which trees to
// touch, that one tree's failure cannot cost the other its rebuild, and that what it SAYS matches
// what happened. An administrator who is told "rebuilt" when nothing was is worse off than one
// told nothing at all: they will go and look for a bug in the sync.
// ###########################################################################################
public class ManifestRebuildFlowTests
{
    private static readonly DateTimeOffset Now = new(2026, 10, 1, 9, 0, 0, TimeSpan.Zero);

    // ###########################################################################################
    // A server with both trees configured rebuilds both, and says each one's count.
    // ###########################################################################################
    [Fact]
    public async Task Both_configured_trees_are_rebuilt_and_each_count_is_reported()
    {
        var store = new FakeAccountStore();

        var written = new List<string>();

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> outcomes = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree =>
            {
                written.Add(tree.Name);
                return tree.Name == "beta" ? 1180 : 1175;
            },
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        // BETA first - it is the tree being edited by hand while something is tested.
        Assert.Equal(["beta", "production"], written);

        Assert.Equal(2, outcomes.Count);
        Assert.Equal((1180, false), (outcomes[0].Entries, outcomes[0].Skipped));
        Assert.Equal((1175, false), (outcomes[1].Entries, outcomes[1].Skipped));

        Assert.Contains("1180", outcomes[0].Message, StringComparison.Ordinal);
        Assert.Equal("2 checksum manifests were rebuilt.", ManifestRebuildFlow.Headline(outcomes));
    }

    // ###########################################################################################
    // *** A SERVER WITHOUT PRODUCTION PUBLISHING IS A NORMAL SERVER, NOT A BROKEN ONE. *** Until
    // DEPLOYMENT.md step 13 is done, nothing leaves BETA (ServerOptions.IsProductionPublishingConfigured),
    // and the service must never write a production tree it has not been told about. The flow must
    // therefore SKIP that tree rather than "fail" to rebuild it - a red line on the Account screen
    // would send the administrator looking for a fault that is not there.
    //
    // *** SKIPPED IS REPORTED, NOT LEFT OUT (code review, 2026-10-01). *** The tree used to be
    // missing from the answer altogether while the API contract and the flow's header both promised
    // a Skipped line - so the Maintainer tab's handling of one could never run.
    // ###########################################################################################
    [Fact]
    public async Task A_server_with_no_production_tree_rebuilds_only_BETA_and_calls_it_a_success()
    {
        var store = new FakeAccountStore();

        var written = new List<string>();

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> outcomes = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: false),
            tree =>
            {
                written.Add(tree.Name);
                return 900;
            },
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        Assert.Equal(["beta"], written);
        Assert.Equal(["beta", "production"], outcomes.Select(outcome => outcome.Tree));
        Assert.False(outcomes[0].Skipped);
        Assert.True(outcomes[1].Skipped);
        Assert.Contains("not configured", outcomes[1].Message, StringComparison.Ordinal);
        Assert.Equal("1 checksum manifest was rebuilt.", ManifestRebuildFlow.Headline(outcomes));

        // The audit names the tree that was rebuilt, not the one skipped.
        AuditEntry entry = Assert.Single(store.Audit);
        Assert.Equal("beta=900", entry.Detail);
    }

    // ###########################################################################################
    // *** A CLIENT GONE PART-WAY STILL LEAVES A RECORD (code review, 2026-10-01). *** The BETA
    // manifest was already rewritten when the disconnect cancelled the production scan, and the
    // audit after the loop was skipped with it - a manifest every client downloads changed with
    // nothing to say so.
    // ###########################################################################################
    [Fact]
    public async Task A_rebuild_cancelled_after_the_first_tree_is_still_audited()
    {
        var store = new FakeAccountStore();
        using var cancel = new CancellationTokenSource();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree =>
            {
                // The request goes away while BETA is being written.
                cancel.Cancel();
                return 1180;
            },
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now,
            cancel.Token));

        AuditEntry entry = Assert.Single(store.Audit);
        Assert.Contains("beta=1180", entry.Detail!, StringComparison.Ordinal);
        Assert.Contains("cancelled", entry.Detail!, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** A FAILED AUDIT DOES NOT TURN A DONE REBUILD INTO A FAILURE (code review, 2026-10-01). ***
    // Both manifests were rewritten; a throw from the audit write made the request a 500, and the
    // administrator would press again for something that had worked.
    // ###########################################################################################
    [Fact]
    public async Task A_failed_audit_write_still_answers_what_was_rebuilt()
    {
        var store = new FakeAccountStore { FailAuditWrites = true };

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> outcomes = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree => 1180,
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        Assert.Equal("2 checksum manifests were rebuilt.", ManifestRebuildFlow.Headline(outcomes));
    }

    // ###########################################################################################
    // *** ONE TREE'S FAILURE MUST NOT COST THE OTHER ITS REBUILD. *** The likely failure is exactly
    // the situation this button is for: a manifest or its folder owned by root, which the service
    // cannot replace. If that aborted the whole call, a production folder the owner had copied in
    // by hand would permanently block rebuilding BETA - the tree that is actually being tested.
    //
    // Both an exception and DataChecksumManifest's own -1 are covered, because Write answers -1
    // for most failures but a hostile filesystem can still throw.
    // ###########################################################################################
    [Fact]
    public async Task One_trees_failure_leaves_the_other_rebuilt_and_the_headline_says_so()
    {
        var store = new FakeAccountStore();

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> thrown = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree => tree.Name == "beta" ? 1180 : throw new UnauthorizedAccessException("root owns it"),
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        Assert.Equal(1180, thrown[0].Entries);
        Assert.Equal(-1, thrown[1].Entries);
        Assert.Equal("1 of 2 checksum manifests were rebuilt.", ManifestRebuildFlow.Headline(thrown));

        // The failing line says so in words, not only by its number - the Account screen prints it.
        Assert.Contains("could NOT be rebuilt", thrown[1].Message, StringComparison.Ordinal);

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> refused = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree => tree.Name == "beta" ? 1180 : -1,
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        Assert.Equal(-1, refused[1].Entries);
        Assert.Equal("1 of 2 checksum manifests were rebuilt.", ManifestRebuildFlow.Headline(refused));
    }

    // ###########################################################################################
    // *** ANY FAILURE IS A FAILED TREE, NOT A FAILED CALL (code review, 2026-10-01). *** Only an IO or
    // access exception used to be caught, so a malformed configured path or a serialising failure
    // escaped the loop: the other tree was never tried and the audit was never written.
    // ###########################################################################################
    [Fact]
    public async Task An_unexpected_failure_in_one_tree_still_rebuilds_the_other_and_is_audited()
    {
        var store = new FakeAccountStore();

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> outcomes = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree => tree.Name == "beta" ? throw new ArgumentException("not a path") : 1175,
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        Assert.Equal(-1, outcomes[0].Entries);
        Assert.Equal(1175, outcomes[1].Entries);
        Assert.Equal("1 of 2 checksum manifests were rebuilt.", ManifestRebuildFlow.Headline(outcomes));

        Assert.Single(store.Audit);
    }

    // Neither rebuilt: the headline must not read as a partial success.
    [Fact]
    public async Task Nothing_rebuilt_is_said_plainly()
    {
        var store = new FakeAccountStore();

        IReadOnlyList<ManifestRebuildFlow.TreeOutcome> outcomes = await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            _ => -1,
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        Assert.Equal("No checksum manifest could be rebuilt.", ManifestRebuildFlow.Headline(outcomes));

        // And a server with no tree at all says that, rather than showing an empty list.
        Assert.Equal(
            "No data tree is configured on this server, so there is no manifest to rebuild.",
            ManifestRebuildFlow.Headline([]));
    }

    // ###########################################################################################
    // *** AUDITED WHATEVER HAPPENED, INCLUDING A FAILURE. *** A manifest rebuilt by hand changes
    // what every CRT client downloads, and the failed attempt is the more interesting of the two
    // to find afterwards - it is the one that explains why a hand-copied board never arrived.
    // ###########################################################################################
    [Fact]
    public async Task Every_rebuild_is_audited_with_what_each_tree_answered()
    {
        var store = new FakeAccountStore();

        await ManifestRebuildFlow.RebuildAsync(
            ManifestRebuildFlowTests.Administrator(),
            ManifestRebuildFlowTests.Options(production: true),
            tree => tree.Name == "beta" ? 1180 : -1,
            new PublishLock(),
            store,
            ManifestRebuildFlowTests.Now);

        AuditEntry entry = Assert.Single(store.Audit);

        Assert.Equal(ManifestRebuildFlow.RebuiltAction, entry.Action);
        Assert.Contains("beta=1180", entry.Detail!, StringComparison.Ordinal);
        Assert.Contains("production=-1", entry.Detail!, StringComparison.Ordinal);
    }

    // A tree whose root is set but whose manifest path is not cannot be written to, so it is not
    // offered at all - there is no file to replace.
    [Fact]
    public void A_tree_with_no_manifest_path_is_not_offered()
    {
        var options = new ServerOptions
        {
            DataTreeRoot = "/srv/beta/Data",
            ManifestPath = string.Empty,
            PublicDataBaseUrl = "https://example.invalid/app-data-BETA/",
        };

        Assert.Empty(ManifestRebuildFlow.TreesToRebuild(options));
    }

    private static ReviewAccess Administrator() =>
        ReviewAccess.For(new AccountRecord(
            7, "anna@example.com", "anna@example.com", "hash", "Anna",
            IsVerified: true, IsAdministrator: true, IsLocked: false, ManifestRebuildFlowTests.Now, null));

    private static ServerOptions Options(bool production)
    {
        var options = new ServerOptions
        {
            DataTreeRoot = "/srv/beta/Data",
            ManifestPath = "/srv/beta/dataChecksums.json",
            PublicDataBaseUrl = "https://example.invalid/app-data-BETA/",
        };

        if (production)
        {
            options.ProductionDataTreeRoot = "/srv/prod/Data";
            options.ProductionManifestPath = "/srv/prod/dataChecksums.json";
            options.ProductionPublicDataBaseUrl = "https://example.invalid/app-data/";
        }

        return options;
    }
}
