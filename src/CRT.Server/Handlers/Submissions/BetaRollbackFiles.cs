using System.Collections.Concurrent;
using System.Security.Cryptography;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The DISK half of a BETA rollback (owner decision, 2026-09-27): listing each tree's copy of
    // one board's folder, comparing files, and finding what the returning submissions did to the
    // shared folders - so BetaRollbackPlan can stay pure.
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
            IReadOnlyList<SubmissionFileRecord> returningFiles,
            CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(betaRoot);
            ArgumentException.ThrowIfNullOrWhiteSpace(productionRoot);
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(returningFiles);

            string folder = BetaRollbackFiles.BoardFolder(board);

            IReadOnlyList<string> beta = BetaRollbackFiles.FilesUnder(betaRoot, folder);
            IReadOnlyList<string> production = BetaRollbackFiles.FilesUnder(productionRoot, folder);

            // Worked out up front, asynchronously: BetaRollbackPlan asks through a plain delegate.
            var same = new Dictionary<string, bool>(StringComparer.Ordinal);

            foreach (string path in production)
                same[path] = await BetaRollbackFiles.SameBytesAsync(betaRoot, productionRoot, path, cancellationToken);

            IReadOnlyList<BetaRollbackSharedFile> shared = await BetaRollbackFiles.SharedFilesAsync(
                betaRoot, productionRoot, board, returningFiles, cancellationToken);

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
                shared,
                promotedBefore);

            // The production hash of each path the writer will copy - its verified copy's target.
            var hashes = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (string path in plan.AllRestored)
            {
                if (SubmissionPathRules.TryResolve(productionRoot, path, out string full, out _) &&
                    await BetaRollbackFiles.HashOfAsync(full, cancellationToken) is string hash)
                {
                    hashes[path] = hash;
                }
            }

            return new BetaRollbackFilePlan(plan, hashes);
        }

        // ###########################################################################################
        // *** ONLY THE BOARD'S OWN FOLDER is listed here. *** The shared folders are handled one
        // file at a time, and only for files a returning submission carried - see SharedFilesAsync.
        // Restoring a whole shared folder from production would revert every other board's change.
        // ###########################################################################################
        public static string BoardFolder(BoardRecord board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return $"{board.Manufacturer.Trim()}/{board.Hardware.Trim()}/{board.Board.Trim()}";
        }

        // Every file under one folder of a tree, as data-root-relative forward-slashed paths.
        // A folder that does not exist is no files, which is a real answer: production has none for
        // a board never promoted.
        public static IReadOnlyList<string> FilesUnder(string root, string folder)
        {
            if (!SubmissionPathRules.TryResolve(root, folder, out string resolved, out _))
                return [];

            if (!Directory.Exists(resolved))
                return [];

            var files = new List<string>();

            foreach (string path in Directory.EnumerateFiles(resolved, "*", SearchOption.AllDirectories))
            {
                // A hidden temporary file beside a copy in progress (VerifiedFileCopy's dot-name) is
                // not part of the board - the checksum manifest skips it for the same reason.
                if (Path.GetFileName(path).StartsWith('.'))
                    continue;

                string relative = Path.GetRelativePath(root, path).Replace('\\', '/');

                // A path that escapes the root cannot be acted on - the same fail-closed rule
                // SubmissionPathRules exists for.
                if (!relative.StartsWith("..", StringComparison.Ordinal))
                    files.Add(relative);
            }

            return files;
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
        // The shared files the returning submissions carried, and what each tree holds for them.
        //
        // "Carried" is the submission's own file list (submission_files), whose hashes are the
        // bytes the submission put there. BETA still holding THOSE bytes is the proof the change is
        // this rollback's to undo; anything else was written again since and belongs to somebody
        // else - see BetaRollbackPlan's header.
        // ###########################################################################################
        public static async Task<IReadOnlyList<BetaRollbackSharedFile>> SharedFilesAsync(
            string betaRoot,
            string productionRoot,
            BoardRecord board,
            IReadOnlyList<SubmissionFileRecord> returningFiles,
            CancellationToken cancellationToken = default)
        {
            var result = new List<BetaRollbackSharedFile>();

            // One entry per path. Two returning submissions naming one shared file both count as
            // "ours" if BETA holds EITHER one's bytes - both are being rolled back.
            foreach (IGrouping<string, SubmissionFileRecord> byPath in returningFiles
                .Where(file => BetaRollbackFiles.IsShared(board, file.Path))
                .GroupBy(file => file.Path.Trim(), StringComparer.Ordinal))
            {
                string path = byPath.Key;

                if (!SubmissionPathRules.TryResolve(betaRoot, path, out string beta, out _) ||
                    !SubmissionPathRules.TryResolve(productionRoot, path, out string production, out _))
                {
                    continue;
                }

                string? betaHash = File.Exists(beta) ? await BetaRollbackFiles.HashOfAsync(beta, cancellationToken) : null;

                bool ours = betaHash is not null && byPath.Any(file =>
                    string.Equals(file.Sha256, betaHash, StringComparison.OrdinalIgnoreCase));

                bool inProduction = File.Exists(production);

                bool same = inProduction && ours &&
                    await BetaRollbackFiles.SameBytesAsync(betaRoot, productionRoot, path, cancellationToken);

                result.Add(new BetaRollbackSharedFile(path, ours, inProduction, same));
            }

            return result;
        }

        // Is this path in one of the two shared folders, for this board's manufacturer?
        private static bool IsShared(BoardRecord board, string? path) =>
            SubmissionFileScopes.Classify(board.Manufacturer, board.Hardware, board.Board, path) is
                SubmissionFileScope.ManufacturerShared or SubmissionFileScope.GenericShared;

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
