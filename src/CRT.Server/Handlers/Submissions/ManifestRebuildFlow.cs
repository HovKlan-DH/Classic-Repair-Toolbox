using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // REBUILDING dataChecksums.json BY HAND (owner request, 2026-10-01: "I need it to update the
    // checksum again ... Still here in the beginning, when testing a lot, I need to do some manual
    // editing of files from time to time").
    //
    // *** WHY THIS EXISTS AT ALL. *** Every path that WRITES a data tree through the service already
    // rebuilds the manifest beside it - an approval (ReviewEndpoints), a promotion
    // (ProductionEndpoints), a rollback (BetaRollbackFlow) and an unused-file removal
    // (UnusedFileFlows). But the project owner also copies files into the trees BY HAND as root
    // (DEPLOYMENT.md step 3, and the "served file is replaced, never opened" note in CLAUDE.md).
    // Nothing on the server sees those edits, so the manifest keeps describing the tree as it was
    // and every CRT client stays on the old bytes - the file looks published and reaches nobody.
    //
    // *** BOTH TREES IN ONE PRESS (owner decision, 2026-10-01). *** One button, BETA and stable
    // together, rather than one per tree. A hand edit is as likely in one as the other, and
    // rebuilding a tree that did not change costs only the scan - the manifest it writes is
    // byte-identical. A tree that is not configured is reported as skipped (Skipped: true, a line
    // of its own) rather than as an error or not at all: a server with no production publishing
    // set up (ServerOptions.IsProductionPublishingConfigured) is a normal, supported state, not
    // a fault - and the answer still names both trees, so the administrator is not left
    // wondering whether the stable one was looked at. (It was LEFT OUT until the code review of
    // 2026-10-01, while this header and the API contract both said "skipped".)
    //
    // *** IT NEVER FAILS AS A WHOLE. *** Each tree is rebuilt independently and its own outcome
    // reported, so an unreadable production folder cannot cost the BETA rebuild that would
    // otherwise have succeeded. The administrator sees per-tree counts and acts on what failed.
    // That includes the AUDIT: a manifest already rewritten is not turned into a 500 by a
    // failed audit write (code review, 2026-10-01) - that would send the administrator to press
    // again for something that worked. The failure is logged instead.
    //
    // *** AUTHORITY IS THE ROUTE'S, AND SAYS SO (code review, 2026-10-01). *** This flow takes the
    // ReviewAccess only to NAME the actor in the audit; it does not check that the account is an
    // administrator - AdminEndpoints' group does, for every route in it
    // (AdminEndpoints.AuthoriseAsync). A refusal constant that nothing read used to sit here and
    // advertised a guard the class does not have. Anything that reaches RebuildAsync some other
    // way must do that check itself.
    //
    // Pure but for the two calls it is given (the manifest writer and the audit write), so the
    // decisions - which trees to touch, what counts as skipped, what the answer says - are tested
    // without a data tree or a database.
    // ###########################################################################################
    public static class ManifestRebuildFlow
    {
        public const string RebuiltAction = "data.manifest_rebuilt";

        // One tree's outcome. `Entries` is what DataChecksumManifest.Write returned: the number of
        // files listed, or -1 for a failure. `Skipped` says the tree is not configured on this
        // server, which is not a failure.
        public sealed record TreeOutcome(string Tree, bool Skipped, int Entries, string Message);

        // ###########################################################################################
        // The trees to rebuild, in the order the answer lists them: BETA first, because it is the
        // one that is edited by hand while something is being tested.
        //
        // A tree with no root or no manifest path configured is left out entirely - there is
        // nothing to write and no file to write it to. UnusedFileFlows.TryResolve answers the same
        // question for its own screen; this does not route through it because that one refuses a
        // tree with an error message, and here an absent tree is simply skipped.
        // ###########################################################################################
        public static IReadOnlyList<UnusedFileFlows.DataTree> TreesToRebuild(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);

            var trees = new List<UnusedFileFlows.DataTree>();

            if (!string.IsNullOrWhiteSpace(options.DataTreeRoot) && !string.IsNullOrWhiteSpace(options.ManifestPath))
            {
                trees.Add(new UnusedFileFlows.DataTree(
                    UnusedFileFlows.Beta,
                    options.DataTreeRoot,
                    options.ManifestPath,
                    options.PublicDataBaseUrl ?? string.Empty));
            }

            if (options.IsProductionPublishingConfigured &&
                !string.IsNullOrWhiteSpace(options.ProductionDataTreeRoot) &&
                !string.IsNullOrWhiteSpace(options.ProductionManifestPath))
            {
                trees.Add(new UnusedFileFlows.DataTree(
                    UnusedFileFlows.Production,
                    options.ProductionDataTreeRoot,
                    options.ProductionManifestPath,
                    options.ProductionPublicDataBaseUrl ?? string.Empty));
            }

            return trees;
        }

        // ###########################################################################################
        // What one tree's result says, in the words the Maintainer tab shows unchanged.
        //
        // *** A FAILURE IS NAMED, NOT COUNTED. *** DataChecksumManifest.Write answers -1 for every
        // kind of failure it has - an empty scan, an unwritable path - and distinguishing them here
        // would mean guessing. The message says what the administrator can check instead, since the
        // overwhelmingly likely cause is the one this button exists for: a folder or a manifest file
        // owned by root that the service cannot replace.
        // ###########################################################################################
        public static string DescribeSkipped(string tree) =>
            $"The {tree} data tree is not configured on this server, so its manifest was skipped.";

        public static string Describe(string tree, int entries) =>
            entries < 0
                ? $"The {tree} manifest could NOT be rebuilt - check that the service may write it, and the server log for why."
                : entries == 1
                    ? $"The {tree} manifest was rebuilt: 1 file listed."
                    : $"The {tree} manifest was rebuilt: {entries} files listed.";

        // The line above the per-tree results. Nothing configured at all is said plainly rather
        // than shown as an empty list.
        public static string Headline(IReadOnlyList<TreeOutcome> outcomes)
        {
            ArgumentNullException.ThrowIfNull(outcomes);

            int rebuilt = outcomes.Count(outcome => !outcome.Skipped && outcome.Entries >= 0);
            int failed = outcomes.Count(outcome => !outcome.Skipped && outcome.Entries < 0);

            if (rebuilt == 0 && failed == 0)
                return "No data tree is configured on this server, so there is no manifest to rebuild.";

            if (failed == 0)
                return rebuilt == 1
                    ? "1 checksum manifest was rebuilt."
                    : $"{rebuilt} checksum manifests were rebuilt.";

            return rebuilt == 0
                ? "No checksum manifest could be rebuilt."
                : $"{rebuilt} of {rebuilt + failed} checksum manifests were rebuilt.";
        }

        // ###########################################################################################
        // Rebuilds every configured tree's manifest and audits what happened. Both trees are in the
        // answer, BETA first; one not configured is Skipped.
        //
        // `write` is DataChecksumManifest.Write in the service and a fake in the tests. The scan
        // walks the whole tree and hashes every file, so it runs on the pool rather than on the
        // request thread.
        //
        // *** IT TAKES THE PUBLISH LOCK. *** A rebuild during an approval or a promotion would scan
        // a tree that is being written and record a manifest describing neither the old state nor
        // the new one. DataChecksumManifest.Write has a write gate of its own, but that only
        // serialises the WRITES - it cannot keep a scan away from a publish that is mid-copy.
        //
        // *** AUDITED EVEN WHEN CANCELLED PART-WAY (code review, 2026-10-01). *** A client that
        // disconnects after the BETA manifest was rewritten cancels the production scan - and the
        // audit used to follow the loop, so a manifest every client downloads changed with no
        // record. The audit now runs in a finally, with no token (the request's is the one that was
        // cancelled), whenever at least one tree was attempted, and says when the run stopped early.
        // ###########################################################################################
        public static async Task<IReadOnlyList<TreeOutcome>> RebuildAsync(
            ReviewAccess access,
            ServerOptions options,
            Func<UnusedFileFlows.DataTree, int> write,
            PublishLock publishLock,
            IAccountStore accounts,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default,
            ILogger? logger = null)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(write);
            ArgumentNullException.ThrowIfNull(publishLock);
            ArgumentNullException.ThrowIfNull(accounts);

            IReadOnlyList<UnusedFileFlows.DataTree> configured = ManifestRebuildFlow.TreesToRebuild(options);

            var outcomes = new List<TreeOutcome>();
            bool finished = false;

            try
            {
                using (await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false))
                {
                    foreach (string name in new[] { UnusedFileFlows.Beta, UnusedFileFlows.Production })
                    {
                        if (configured.FirstOrDefault(candidate => candidate.Name == name) is not { } tree)
                        {
                            outcomes.Add(new TreeOutcome(name, Skipped: true, 0, ManifestRebuildFlow.DescribeSkipped(name)));
                            continue;
                        }

                        // One tree's failure must not cost the other its rebuild - see the header.
                        int entries;

                        try
                        {
                            // On the pool: the scan walks the tree and hashes every file, which on a
                            // board-sized tree is seconds, and this is called straight from a request.
                            entries = await Task.Run(() => write(tree), cancellationToken).ConfigureAwait(false);
                        }
                        catch (Exception ex) when (ex is not OperationCanceledException)
                        {
                            // *** ANY FAILURE OF ONE TREE IS -1 (code review, 2026-10-01). *** Only an IO or
                            // access failure used to be caught here; a malformed configured path
                            // (ArgumentException), a serialising failure or a security exception escaped
                            // the loop, so the second tree was never tried AND the audit was never
                            // written - both of this class's stated guarantees broken at once, with the
                            // administrator seeing a bare 500. DataChecksumManifest.Write answers -1 for
                            // every failure it knows; this is the same contract for the ones it does not.
                            // Cancellation is the caller's own decision and still propagates.
                            entries = -1;
                        }

                        outcomes.Add(new TreeOutcome(
                            tree.Name,
                            Skipped: false,
                            entries,
                            ManifestRebuildFlow.Describe(tree.Name, entries)));
                    }
                }

                finished = true;
            }
            finally
            {
                // Audited whatever the result: a manifest rebuilt by hand changes what every client
                // downloads, and a rebuild that FAILED is the more interesting of the two to find
                // later. Nothing attempted (cancelled waiting for the lock, or no tree configured)
                // changed nothing, and is not audited.
                if (outcomes.Any(outcome => !outcome.Skipped))
                    await ManifestRebuildFlow.AuditAsync(access, outcomes, finished, accounts, nowUtc, logger).ConfigureAwait(false);
            }

            return outcomes;
        }

        // The audit row, never thrown out of: the manifests are already written - see the header.
        private static async Task AuditAsync(
            ReviewAccess access,
            IReadOnlyList<TreeOutcome> outcomes,
            bool finished,
            IAccountStore accounts,
            DateTimeOffset nowUtc,
            ILogger? logger)
        {
            List<TreeOutcome> attempted = [.. outcomes.Where(outcome => !outcome.Skipped)];

            string detail = string.Join(" | ", attempted.Select(outcome => $"{outcome.Tree}={outcome.Entries}"));

            if (!finished)
                detail += " | cancelled before the rest";

            try
            {
                await accounts.WriteAuditAsync(
                    new AuditEntry(
                        access.Account.Id,
                        access.Account.Email,
                        ManifestRebuildFlow.RebuiltAction,
                        string.Join("+", attempted.Select(outcome => outcome.Tree)),
                        detail,
                        nowUtc),
                    CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                logger?.LogError(ex, "The checksum manifest rebuild could not be audited ({Detail}).", detail);
            }
        }
    }
}
