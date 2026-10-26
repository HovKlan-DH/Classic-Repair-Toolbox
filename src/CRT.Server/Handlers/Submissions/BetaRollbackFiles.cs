using System.Collections.Concurrent;
using System.Security.Cryptography;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The DISK half of a BETA rollback (owner decision, 2026-09-27): listing each tree's copy of
    // one board - its board file and the stored files it lists (the CRT 4.0 tree, 2026-10-25) - and
    // comparing them, so BetaRollbackPlan can stay pure.
    //
    // *** CHEAP COMPARISONS FIRST (code review, 2026-09-27). *** The first version SHA-256'd both
    // trees' copy of EVERY production file, synchronously, and did it twice per rollback - once for
    // the confirmation, once again under the PublishLock - so a board with 1 GB of scans read 4 GB
    // and blocked every approval while it did. Now:
    //
    //   - two files of different LENGTH are different, and neither is read;
    //   - a hash is computed once per file VERSION (path + length + last-write time) and kept in
    //     FileHashes, so the pass under the lock reuses what the confirmation already read - the
    //     WorkbookReadCache idea, applied to bytes. A file written in between has a new version and
    //     is read again;
    //   - all of it is async, off the request thread.
    //
    // A cached PRODUCTION hash that is stale can only cause a restore to be REFUSED: the writer
    // copies through VerifiedFileCopy, which hashes the bytes it actually copies and replaces the
    // target only on a match. The BETA side is kept honest by the writer forgetting every file it
    // writes - see Forget.
    // ###########################################################################################
    public static class BetaRollbackFiles
    {
        // ###########################################################################################
        // The plan for one board, with the production hash of every path it restores - the value
        // the writer's verified copy must see, so the bytes copied are the bytes compared.
        // ###########################################################################################
        public static async Task<BetaRollbackFilePlan> PlanAsync(
            string betaRoot,
            string productionRoot,
            BoardRecord board,
            IReadOnlyList<CarriedSubmission> returning,
            CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(betaRoot);
            ArgumentException.ThrowIfNullOrWhiteSpace(productionRoot);
            ArgumentNullException.ThrowIfNull(board);

            // ###########################################################################################
            // *** A BOARD IN THE CRT 4.0 TREE (2026-10-25) is its board file and the stored files it
            // lists *** ("Files/<ab>/<sha256>.<ext>", DataTreeLayout) - so those are each tree's
            // "own files" here. A board file is compared by its bytes; a stored file by being there,
            // since its name is its checksum. There are no shared files to put back any more.
            // ###########################################################################################
            IReadOnlyList<string> beta = BetaRollbackFiles.FilesOfBoard(betaRoot, board);
            IReadOnlyList<string> production = BetaRollbackFiles.FilesOfBoard(productionRoot, board);

            // Worked out up front, asynchronously: BetaRollbackPlan asks through a plain delegate.
            var same = new Dictionary<string, bool>(StringComparer.Ordinal);

            foreach (string path in production)
            {
                same[path] = AutomaticRemovalScope.IsStoredFile(path)
                    ? PublishedBoardFiles.IsStored(betaRoot, path)
                    : await BetaRollbackFiles.SameBytesAsync(betaRoot, productionRoot, path, cancellationToken);
            }

            // "Never promoted" is what the board's record says, never what an empty listing
            // suggests - a missing or unreadable production folder lists nothing too.
            bool promotedBefore =
                board.ProductionPublishedUtc is not null ||
                !string.IsNullOrWhiteSpace(board.ProductionContentHash);

            BetaRollbackPlanResult plan = BetaRollbackPlan.Build(
                beta,
                production,
                path => same.TryGetValue(path, out bool isSame) && isSame,
                returning,
                sharedFiles: null,
                promotedBefore);

            plan = BetaRollbackFiles.InWriteOrder(BetaRollbackFiles.KeepingWhatOthersUse(plan, betaRoot, board, production));

            // The production hash of each path the writer will copy - its verified copy's target: a
            // stored file's checksum is its name; the board file is hashed.
            var hashes = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (string path in plan.AllRestored)
            {
                if (BetaRollbackFiles.ChecksumInName(path) is string named)
                {
                    hashes[path] = named;
                }
                else if (SubmissionPathRules.TryResolve(productionRoot, path, out string full, out _) &&
                         await BetaRollbackFiles.HashOfAsync(full, cancellationToken) is string hash)
                {
                    hashes[path] = hash;
                }
            }

            return new BetaRollbackFilePlan(plan, hashes);
        }

        // ###########################################################################################
        // The board's files in one tree: its board file, and every stored file it lists that the
        // tree holds - data-root-relative. None when the tree has no board file of it (a board never
        // promoted, in production); just the board file when it cannot be read - the plan then
        // restores it whole, and the stored files it names are restored by the next push-back.
        // ###########################################################################################
        public static IReadOnlyList<string> FilesOfBoard(string root, BoardRecord board)
        {
            ArgumentNullException.ThrowIfNull(board);

            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(root, board.BoardId);

            if (!location.Exists)
                return [];

            var files = new List<string> { DataTreeLayout.BoardFilePath(board.BoardId) };

            foreach (BoardDocument.StoredFileRow file in BoardFiles.ReadDocument(location.WorkbookPath)?.Files ?? [])
            {
                string stored = PublishedBoardFiles.StoredRelative(file);

                if (stored.Length > 0 && !files.Contains(stored, StringComparer.Ordinal) && PublishedBoardFiles.IsStored(root, stored))
                    files.Add(stored);
            }

            return files;
        }

        // ###########################################################################################
        // *** A STORED FILE ANOTHER BETA BOARD LISTS STAYS. *** The plan removes every stored file
        // the rolled-back board lists that production's does not - but in the 4.0 tree bytes are
        // stored once for every board naming them, so one may be another board's too. Only those no
        // BETA board lists once the board is back as production has it (or gone, for a board never
        // promoted) go - DataTreeUsage, failing closed: a tree that cannot be read completely keeps
        // every stored file, and only the board file is removed.
        // ###########################################################################################
        private static BetaRollbackPlanResult KeepingWhatOthersUse(
            BetaRollbackPlanResult plan,
            string betaRoot,
            BoardRecord board,
            IReadOnlyList<string> production)
        {
            List<string> stored = plan.Removed.Where(AutomaticRemovalScope.IsStoredFile).ToList();

            if (stored.Count == 0)
                return plan;

            List<string> after = plan.Kind == BetaRollbackKind.RemoveFromBeta
                ? []
                : production.Where(AutomaticRemovalScope.IsStoredFile).ToList();

            IReadOnlyList<string> removable = DataTreeUsage
                .Compute(betaRoot, new Dictionary<string, IReadOnlyCollection<string>> { [DataTreeLayout.BoardFilePath(board.BoardId)] = after })
                .RemovableFrom(stored);

            var keep = new HashSet<string>(removable, StringComparer.OrdinalIgnoreCase);

            return plan with
            {
                Removed = plan.Removed.Where(path => !AutomaticRemovalScope.IsStoredFile(path) || keep.Contains(path)).ToList()
            };
        }

        // ###########################################################################################
        // The publish's order, for the same reason: the stored files a restored board file names go
        // in BEFORE it, so BETA never holds a board file naming bytes that are not there yet; and a
        // removed board file goes before the stored files nothing lists after it.
        // ###########################################################################################
        private static BetaRollbackPlanResult InWriteOrder(BetaRollbackPlanResult plan) => plan with
        {
            Restored = plan.Restored
                .OrderBy(path => AutomaticRemovalScope.IsStoredFile(path) ? 0 : 1)
                .ThenBy(path => path, StringComparer.Ordinal)
                .ToList(),
            Removed = plan.Removed
                .OrderBy(path => AutomaticRemovalScope.IsStoredFile(path) ? 1 : 0)
                .ThenBy(path => path, StringComparer.Ordinal)
                .ToList()
        };

        // The SHA-256 a stored file's name is ("Files/3f/3f7a...e1.png"), or null for any other path.
        private static string? ChecksumInName(string path)
        {
            if (!AutomaticRemovalScope.IsStoredFile(path))
                return null;

            string name = Path.GetFileName(path);
            int dot = name.IndexOf('.');
            string sha = dot < 0 ? name : name[..dot];

            return DataTreeLayout.IsSha256(sha) ? sha : null;
        }

        // The board's folder in a tree ("Commodore/C64/250407") - where its board file lies, and what
        // a board never promoted leaves behind once it is taken out of BETA.
        public static string BoardFolder(BoardRecord board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return $"{board.Manufacturer.Trim()}/{board.Hardware.Trim()}/{board.Board.Trim()}";
        }

        // ###########################################################################################
        // Do both trees hold identical bytes at this path? LENGTH FIRST - two files of different
        // length can never be the same, and deciding that reads nothing - then the hash, from
        // FileHashes when this version of the file has been read before.
        //
        // An unreadable file answers FALSE - "not known to be the same" - so it is restored. Failing
        // towards writing production's known-good copy is the safe direction here.
        // ###########################################################################################
        public static async Task<bool> SameBytesAsync(
            string betaRoot,
            string productionRoot,
            string path,
            CancellationToken cancellationToken = default)
        {
            if (!SubmissionPathRules.TryResolve(betaRoot, path, out string beta, out _) ||
                !SubmissionPathRules.TryResolve(productionRoot, path, out string production, out _))
            {
                return false;
            }

            try
            {
                var betaInfo = new FileInfo(beta);
                var productionInfo = new FileInfo(production);

                if (!betaInfo.Exists || !productionInfo.Exists)
                    return false;

                if (betaInfo.Length != productionInfo.Length)
                    return false;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }

            string? betaHash = await BetaRollbackFiles.HashOfAsync(beta, cancellationToken);
            string? productionHash = await BetaRollbackFiles.HashOfAsync(production, cancellationToken);

            return betaHash is not null && string.Equals(betaHash, productionHash, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // A file's SHA-256, lower-case hex as VerifiedFileCopy compares it, once per file version.
        // Null when it cannot be read.
        // ###########################################################################################
        public static async Task<string?> HashOfAsync(string fullPath, CancellationToken cancellationToken = default)
        {
            try
            {
                var info = new FileInfo(fullPath);

                if (!info.Exists)
                    return null;

                var version = new FileVersion(info.Length, info.LastWriteTimeUtc);

                if (FileHashes.TryGetValue(fullPath, out (FileVersion Version, string Hash) cached) && cached.Version == version)
                    return cached.Hash;

                string hash;

                await using (FileStream stream = new(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
                {
                    hash = Convert.ToHexStringLower(await SHA256.HashDataAsync(stream, cancellationToken));
                }

                // A bound, so a long-running service cannot grow this without limit. Clearing costs
                // one re-read of whatever is asked about next - nothing is ever wrong because of it.
                if (FileHashes.Count >= BetaRollbackFiles.MaximumCachedHashes)
                    FileHashes.Clear();

                FileHashes[fullPath] = (version, hash);

                return hash;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }

        // ###########################################################################################
        // *** THE WRITER FORGETS WHAT IT WRITES. *** A version is length + last-write time, and a
        // file time only moves in ticks - on Windows commonly ~16 ms. A file rewritten with the same
        // length inside one tick of its previous write therefore looks unchanged, and the cache
        // answers the OLD hash. Nothing else writes these trees that quickly after a write; the
        // rollback's own restore does, when it replaces a file the same pass has just hashed. The
        // full suite caught it: a second push-back re-restored a file the first had already put
        // back, and the same staleness the other way would SKIP a restore that is needed.
        // ###########################################################################################
        public static void Forget(string fullPath) => FileHashes.TryRemove(fullPath, out _);

        private const int MaximumCachedHashes = 20_000;

        private static readonly ConcurrentDictionary<string, (FileVersion Version, string Hash)> FileHashes =
            new(StringComparer.Ordinal);

        private readonly record struct FileVersion(long Length, DateTime LastWriteUtc);
    }

    // A rollback plan plus the production hash of every path it restores.
    public sealed record BetaRollbackFilePlan(
        BetaRollbackPlanResult Plan,
        IReadOnlyDictionary<string, string> ProductionHashes);
}
