using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The REAL PublishedTreeView: the two questions the submission rules ask about the published
    // tree, answered from the data tree on this server (security review, 2026-09-25).
    //
    // An I/O boundary, thin on purpose - the rules that USE these answers live in CRT.Data and are
    // tested against a dictionary. What is here is a directory listing and a cached hash.
    //
    // Every relative path goes through SubmissionPathRules before it touches the disk, the one
    // containment rule this service has; a hostile path cannot make the probe list or hash
    // anything outside the tree.
    // ###########################################################################################
    public static class PublishedTreeProbe
    {
        // ###########################################################################################
        // A view of the tree at `dataTreeRoot`, or NULL when there is no usable root. Null is what
        // the rules treat as "the tree could not be consulted", which refuses every foreign file -
        // the safe answer for a misconfigured service.
        // ###########################################################################################
        public static PublishedTreeView? For(string? dataTreeRoot)
        {
            if (string.IsNullOrWhiteSpace(dataTreeRoot) || !Directory.Exists(dataTreeRoot))
                return null;

            return new PublishedTreeView(
                relativePath => PublishedFileHashes.TryHash(dataTreeRoot, relativePath),
                relativeFolder => PublishedTreeProbe.List(dataTreeRoot, relativeFolder));
        }

        private static IReadOnlyCollection<string>? List(string dataTreeRoot, string relativeFolder)
        {
            string folder;

            if (relativeFolder.Length == 0)
            {
                folder = dataTreeRoot;
            }
            else if (!SubmissionPathRules.TryResolve(dataTreeRoot, relativeFolder, out folder, out _))
            {
                return null;
            }

            try
            {
                if (!Directory.Exists(folder))
                    return null;

                return Directory
                    .EnumerateFileSystemEntries(folder)
                    .Select(Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Select(name => name!)
                    .ToList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }
    }
}
