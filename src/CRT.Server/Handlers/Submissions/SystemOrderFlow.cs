using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE ORDER OF CRT's DROP-DOWN LISTS, SET BY THE ADMINISTRATOR (owner request, 2026-10-04: "In the
    // 'Admin' top menu I would like a possibility to be able to sort the list of systems, which then
    // gets saved to both sources (BETA + stable) after my save").
    //
    // The administrator drags BETA's list into order in the Maintainer tab and sends every system in
    // it, by id. This writes that order into BETA's newest main Excel data file and the stable
    // source's, and rebuilds the checksum manifest of each tree it changed, so CRT downloads the new
    // order on its next launch.
    //
    // *** THE WHOLE LIST OR NOTHING. *** The request must name exactly the systems BETA's list holds -
    // a system placed or removed since the list was read makes it a 409, so the administrator looks
    // again rather than having a list applied that was not the one on screen.
    //
    // *** THE STABLE LIST FOLLOWS BETA's ORDER. *** It lists fewer systems (those not promoted yet are
    // only in BETA) and may list one BETA lacks; MasterListing.ArrangeAs puts each it shares with BETA
    // in BETA's order and keeps any other straight after the row it followed. Writing the stable
    // source directly - not through a promotion - is the owner's request: the order is a fact of the
    // lists, not of any one system's data waiting in BETA.
    //
    // *** BETA FIRST, AND A LATER FAILURE IS REPORTED, NOT ROLLED BACK. *** Once BETA's list is
    // written, a stable list or a manifest that could not be written is said in the answer
    // (SystemOrderAnswer.Problem); pressing Save again writes what is still out of order and changes
    // nothing that already is.
    //
    // Authority is the route's: AdminEndpoints' group refuses anybody who is not an administrator.
    // Every write happens under the one PublishLock, so no publish or promotion reads a list half
    // written. The manifest writer is handed in, so the order of the steps is tested on a real tree
    // without a real manifest.
    // ###########################################################################################
    public static class SystemOrderFlow
    {
        public const string OrderedAction = "system.list_ordered";

        public const string ChangedMessage =
            "The list in BETA has changed since you opened it - a system was added to it or taken out. " +
            "Open \"Order of systems\" again and put it in order from there.";

        public static async Task<SystemOrderOutcome> SetAsync(
            ReviewAccess access,
            SystemOrderRequest? request,
            ServerOptions options,
            Func<UnusedFileFlows.DataTree, int> writeManifest,
            PublishLock publishLock,
            IAccountStore accounts,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default,
            ILogger? logger = null)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(writeManifest);
            ArgumentNullException.ThrowIfNull(publishLock);
            ArgumentNullException.ThrowIfNull(accounts);

            if (!SystemOrderFlow.TryNormalise(request, out IReadOnlyList<string> order, out string problem))
                return SystemOrderOutcome.BadRequest(problem);

            IReadOnlyList<UnusedFileFlows.DataTree> trees = ManifestRebuildFlow.TreesToRebuild(options);
            UnusedFileFlows.DataTree? beta = trees.FirstOrDefault(tree => tree.Name == UnusedFileFlows.Beta);
            UnusedFileFlows.DataTree? stable = trees.FirstOrDefault(tree => tree.Name == UnusedFileFlows.Production);

            // Only the BETA root is needed to write its list; without a manifest path it is still
            // written, and simply has no manifest to rebuild.
            string? betaRoot = beta?.Root ?? options.DataTreeRoot;

            if (string.IsNullOrWhiteSpace(betaRoot))
                return SystemOrderOutcome.Conflict("The server has no data tree configured.");

            using IDisposable gate = await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            string? betaMaster = MasterListing.NewestMasterPath(betaRoot);

            if (betaMaster is null)
                return SystemOrderOutcome.Conflict(SystemListingRules.NoMasterMessage);

            if (!MasterListing.TryRead(betaMaster, out IReadOnlyList<MasterListingRow> listed, out string why))
                return SystemOrderOutcome.Conflict($"BETA's main Excel data file could not be read: {why}");

            if (!SystemOrderFlow.NamesExactly(listed, order))
                return SystemOrderOutcome.Conflict(SystemOrderFlow.ChangedMessage);

            MasterListingEdit betaEdit = MasterListing.Reorder(betaMaster, order, nowUtc);

            if (!betaEdit.IsDone)
                return SystemOrderOutcome.Conflict(betaEdit.Failure ?? "BETA's main Excel data file could not be written.");

            // *** FROM HERE ON, NOTHING IS CANCELLED (code review, 2026-10-04). *** BETA's list is
            // written: a client giving up now - CRT's wait limit, or CRT closed - must still leave
            // the stable list in BETA's order and each changed tree's manifest naming its new
            // checksum, or every CRT throws the download away on each sync. SystemDeletionFlow
            // rebuilds its manifests uncancellably for the same reason.
            var problems = new List<string>();

            if (betaEdit.Changed && beta is not null)
                await SystemOrderFlow.RebuildAsync(beta, writeManifest, problems).ConfigureAwait(false);

            bool? stableChanged = null;

            if (stable is not null)
            {
                stableChanged = false;
                string? stableMaster = MasterListing.NewestMasterPath(stable.Root);

                if (stableMaster is null)
                {
                    problems.Add("The stable source has no versioned main Excel data file, so its list was not changed.");
                }
                else
                {
                    MasterListingEdit stableEdit = MasterListing.Reorder(stableMaster, order, nowUtc);

                    if (!stableEdit.IsDone)
                    {
                        problems.Add($"The stable source's list could not be changed: {stableEdit.Failure}");
                    }
                    else if (stableEdit.Changed)
                    {
                        stableChanged = true;
                        await SystemOrderFlow.RebuildAsync(stable, writeManifest, problems).ConfigureAwait(false);
                    }
                }
            }

            if (betaEdit.Changed || stableChanged == true)
            {
                try
                {
                    await accounts.WriteAuditAsync(
                        new AuditEntry(
                            access.Account.Id,
                            access.Account.Email,
                            SystemOrderFlow.OrderedAction,
                            MasterWorkbookSchema.SheetName,
                            $"beta={(betaEdit.Changed ? "changed" : "unchanged")} | stable={(stableChanged switch { true => "changed", false => "unchanged", null => "none" })} | {order.Count} systems",
                            nowUtc),
                        CancellationToken.None).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    // The lists are written - an audit that failed must not turn that into a 500.
                    logger?.LogError(ex, "The new order of the drop-down lists could not be audited.");
                }
            }

            foreach (string line in problems)
                logger?.LogWarning("Ordering the drop-down lists: {Problem}", line);

            return SystemOrderOutcome.Saved(new SystemOrderAnswer(
                betaEdit.Changed,
                stableChanged,
                problems.Count == 0 ? null : string.Join(" ", problems)));
        }

        // ###########################################################################################
        // The request made into an order: every id trimmed, none blank, none twice (any case). An
        // empty list is refused - there is no list with nothing in it to put in order.
        // ###########################################################################################
        public static bool TryNormalise(SystemOrderRequest? request, out IReadOnlyList<string> order, out string problem)
        {
            List<string> ids = (request?.SystemIds ?? []).Select(id => id?.Trim() ?? string.Empty).ToList();
            order = ids;

            if (ids.Count == 0)
            {
                problem = "No systems were sent to put in order.";
                return false;
            }

            if (ids.Any(id => id.Length == 0))
            {
                problem = "A system in the list has no id.";
                return false;
            }

            if (ids.Distinct(StringComparer.OrdinalIgnoreCase).Count() != ids.Count)
            {
                problem = "A system is in the list more than once.";
                return false;
            }

            problem = string.Empty;
            return true;
        }

        // Whether `order` names exactly the systems `listed` holds - no more, no fewer (any case).
        public static bool NamesExactly(IReadOnlyList<MasterListingRow> listed, IReadOnlyList<string> order)
        {
            ArgumentNullException.ThrowIfNull(listed);
            ArgumentNullException.ThrowIfNull(order);

            var held = new HashSet<string>(listed.Select(row => row.SystemId), StringComparer.OrdinalIgnoreCase);
            var sent = new HashSet<string>(order, StringComparer.OrdinalIgnoreCase);

            return held.SetEquals(sent);
        }

        // ###########################################################################################
        // One tree's manifest, on the pool (it hashes the whole tree); a failure is a problem line.
        // Takes no token: it only ever runs after a list was written (see SetAsync).
        // ###########################################################################################
        private static async Task RebuildAsync(
            UnusedFileFlows.DataTree tree,
            Func<UnusedFileFlows.DataTree, int> writeManifest,
            List<string> problems)
        {
            int entries;

            try
            {
                entries = await Task.Run(() => writeManifest(tree), CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception)
            {
                entries = -1;
            }

            if (entries < 0)
            {
                string source = tree.Name == UnusedFileFlows.Beta ? "BETA" : "the stable source";

                problems.Add(
                    $"The checksum manifest of {source} could not be rebuilt, so CRT will not download the new order from there yet - " +
                    "use " + MaintainerScreenWording.AccountQuoted + " > \"Rebuild checksum manifests\".");
            }
        }
    }

    public sealed record SystemOrderOutcome(SystemOrderAnswer? Answer, string? Error, bool IsConflict = false)
    {
        public static SystemOrderOutcome Saved(SystemOrderAnswer answer) => new(answer, null);

        public static SystemOrderOutcome Conflict(string error) => new(null, error, IsConflict: true);

        public static SystemOrderOutcome BadRequest(string error) => new(null, error);
    }
}
