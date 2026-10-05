using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE ADMINISTRATOR'S "UNUSED FILES" SCREEN (owner decision, 2026-09-25: "there must be
    // no orphan files").
    //
    // A publish removes only what IT made unused. Files that were orphans before any of this
    // existed - the project owner reviewed and approved a list of 50 in the shipped data - are shown
    // here, per tree, and removed only when the administrator presses the button, so the first
    // clean-up is looked at before anything goes.
    //
    // Administrator-only (AdminEndpoints refuses everyone else). The removal takes the
    // PublishLock, removes only files the administrator sent that the tree - read again - still
    // does not use (UnusedFileRemover), regenerates that tree's dataChecksums.json so clients stop
    // being offered them, and writes an audit row.
    // ###########################################################################################
    public static class UnusedFileFlows
    {
        public const string RemovedAction = "data.unused_removed";

        public const string AdministratorOnly = "Only an administrator can remove unused files.";

        public const string Beta = "beta";
        public const string Production = "production";

        // One tree: where it is, and its manifest.
        public sealed record DataTree(string Name, string Root, string ManifestPath, string PublicBaseUrl);

        // ###########################################################################################
        // "beta" or "production" to the tree it names. Production only when publishing to it is
        // configured - otherwise the service has no business writing there (INSTALLING.md ("Folders and permissions")).
        // ###########################################################################################
        public static bool TryResolve(string? name, ServerOptions options, out DataTree? tree, out string error)
        {
            ArgumentNullException.ThrowIfNull(options);
            tree = null;
            error = string.Empty;

            if (string.Equals(name, UnusedFileFlows.Beta, StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(options.DataTreeRoot))
                {
                    error = "The server has no BETA data configured.";
                    return false;
                }

                tree = new DataTree(UnusedFileFlows.Beta, options.DataTreeRoot, options.ManifestPath ?? string.Empty, options.PublicDataBaseUrl ?? string.Empty);
                return true;
            }

            if (string.Equals(name, UnusedFileFlows.Production, StringComparison.OrdinalIgnoreCase))
            {
                if (!options.IsProductionPublishingConfigured)
                {
                    error = ProductionPromotionFlow.NotConfiguredMessage;
                    return false;
                }

                tree = new DataTree(
                    UnusedFileFlows.Production,
                    options.ProductionDataTreeRoot!,
                    options.ProductionManifestPath ?? string.Empty,
                    options.ProductionPublicDataBaseUrl ?? string.Empty);
                return true;
            }

            error = "Name the data as \"beta\" or \"production\".";
            return false;
        }

        // ###########################################################################################
        // The list: every file the tree does not use, with its size - or why none can be named.
        // ###########################################################################################
        public static UnusedFileListing List(DataTree tree)
        {
            ArgumentNullException.ThrowIfNull(tree);

            // Shown, not acted on - RemoveAsync reads the tree afresh - so the preview cache is
            // safe here, and a second look at the list costs a directory walk, not a re-parse.
            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root, cache: ApprovePublishFlow.PreviewReads);
            string fullRoot = Path.GetFullPath(tree.Root);

            List<UnusedFileEntry> files = usage.UnusedFiles
                .Select(path => new UnusedFileEntry(path, UnusedFileFlows.SizeOf(fullRoot, path)))
                .ToList();

            return new UnusedFileListing(
                tree.Name,
                usage.IsComplete,
                usage.Problems,
                usage.MasterCount,
                usage.BoardWorkbookCount,
                usage.Files.Count,
                files,

                // Where the tree is published, so the Maintainer tab's tree can show and open a file
                // as CRT downloads it (2026-10-04).
                string.IsNullOrWhiteSpace(tree.PublicBaseUrl) ? null : tree.PublicBaseUrl);
        }

        // ###########################################################################################
        // Removes the files the administrator chose, under the publish lock.
        // ###########################################################################################
        public static async Task<UnusedFileRemoval> RemoveAsync(
            ReviewAccess? access,
            DataTree tree,
            IReadOnlyCollection<string> files,
            PublishLock publishLock,
            IAccountStore accounts,
            ILogger logger,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(tree);
            ArgumentNullException.ThrowIfNull(files);
            ArgumentNullException.ThrowIfNull(publishLock);
            ArgumentNullException.ThrowIfNull(accounts);

            // The route refuses everyone else first; this is the same rule where a test reaches it.
            if (!ReviewAuthority.CanAdminister(access))
                return new UnusedFileRemoval([], files.ToList(), UnusedFileFlows.AdministratorOnly);

            UnusedFileRemoval removal;

            using (await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false))
            {
                removal = UnusedFileRemover.Remove(tree.Root, files, logger);

                // Clients are told by the manifest; until it is rewritten they keep being offered
                // files that are no longer there. It cannot fail the request (see its header).
                if (removal.Removed.Count > 0)
                {
                    int written = DataChecksumManifest.Write(tree.Root, tree.PublicBaseUrl, tree.ManifestPath);

                    if (written < 0)
                    {
                        logger.LogWarning(
                            "Unused files were removed from {Tree} but its checksum manifest at [{Path}] could not be regenerated.",
                            tree.Name, tree.ManifestPath);
                    }
                }
            }

            if (removal.Removed.Count > 0)
            {
                await accounts.WriteAuditAsync(
                    new AuditEntry(
                        access!.Account.Id,
                        access.Account.Email,
                        UnusedFileFlows.RemovedAction,
                        tree.Name,
                        $"{removal.Removed.Count} file(s): {string.Join(" | ", removal.Removed)}",
                        nowUtc),
                    cancellationToken).ConfigureAwait(false);
            }

            return removal;
        }

        private static long SizeOf(string fullRoot, string relativePath)
        {
            try
            {
                return new FileInfo(Path.Combine(fullRoot, relativePath.Replace('/', Path.DirectorySeparatorChar))).Length;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return 0;
            }
        }
    }
}
