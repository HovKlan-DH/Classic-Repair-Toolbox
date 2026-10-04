using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // HOW BIG EACH FILE IN A FILE TREE IS (owner request, 2026-10-04: "Ideally the 'Files' actually
    // also states its size (everywhere)") - for the Systems screen's Files view, a submission's
    // Files view and the production plan's tree.
    //
    // A file's size is the size of what OPENS from its row: BETA's or production's file on disk, or
    // the submission's upload (its manifest says). A file not written yet has none.
    //
    // *** READ THROUGH THE ONE CONTAINMENT RULE. *** The paths are the server's own - a tree walk, a
    // board's citations, a validated manifest - but a size read is still a read under a data root,
    // so it goes through SubmissionPathRules like every other. A path it refuses, or a file not
    // there, has no size - never a wrong one, and never an exception that fails the tree.
    // ###########################################################################################
    public static class TreeFileSizes
    {
        // The size of `relativePath` under `treeRoot`, or null when it is not there.
        public static long? Of(string? treeRoot, string? relativePath)
        {
            if (string.IsNullOrWhiteSpace(treeRoot) || string.IsNullOrWhiteSpace(relativePath) ||
                !SubmissionPathRules.TryResolve(treeRoot, relativePath, out string resolved, out _))
            {
                return null;
            }

            try
            {
                var info = new FileInfo(resolved);
                return info.Exists ? info.Length : null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
            {
                return null;
            }
        }

        // ###########################################################################################
        // Each entry's size, where its bytes are: BETA's tree, production's, or the submission's
        // upload (`submitted`, path -> bytes, from its manifest).
        // ###########################################################################################
        public static IReadOnlyList<SystemFileEntry> Attach(
            IReadOnlyList<SystemFileEntry> entries,
            string betaRoot,
            string? productionRoot = null,
            IReadOnlyDictionary<string, long>? submitted = null) =>
            SystemFileEntries.WithSizes(entries, entry => entry.OpenFrom switch
            {
                SystemFileSource.Beta => TreeFileSizes.Of(betaRoot, entry.Path),
                SystemFileSource.Production => TreeFileSizes.Of(productionRoot, entry.Path),
                SystemFileSource.Submission => submitted is not null && submitted.TryGetValue(entry.Path, out long size) ? size : null,
                _ => null
            });

        // A submission's uploads by path - the keys spelled as the tree spells them.
        public static IReadOnlyDictionary<string, long> OfSubmission(SubmissionManifest manifest)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            var sizes = new Dictionary<string, long>(StringComparer.Ordinal);

            foreach (SubmissionFile file in manifest.Files)
            {
                string path = (file.Path ?? string.Empty).Replace('\\', '/').Trim('/');

                if (path.Length > 0)
                    sizes[path] = file.SizeBytes;
            }

            return sizes;
        }

        // ###########################################################################################
        // The production plan's sizes (ProductionPlanAnswer.FileSizes): what it copies and what is
        // already the same as BETA holds them - what production will have - and what it removes as
        // production holds it now. A file not there is left out.
        // ###########################################################################################
        public static IReadOnlyDictionary<string, long> ForPromotion(
            string betaRoot,
            string? productionRoot,
            IEnumerable<string>? copied,
            IEnumerable<string>? unchanged,
            IEnumerable<string>? removed)
        {
            var sizes = new Dictionary<string, long>(StringComparer.Ordinal);

            void Put(string? root, IEnumerable<string>? paths)
            {
                foreach (string path in paths ?? [])
                {
                    if (TreeFileSizes.Of(root, path) is long size)
                        sizes[path] = size;
                }
            }

            Put(betaRoot, unchanged);
            Put(betaRoot, copied);
            Put(productionRoot, removed);

            return sizes;
        }
    }
}
