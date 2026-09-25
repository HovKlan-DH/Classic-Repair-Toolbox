namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Refuses to publish THROUGH a symbolic link (security review, 2026-09-25).
    //
    // SubmissionPathRules checks containment LEXICALLY: it resolves "..", but it cannot see that a
    // folder inside the data tree is really a link to somewhere else. A link placed in the BETA
    // tree - by hand, by a restore, by a misjudged convenience - would let a perfectly contained
    // submitted path write wherever the link points. None exists today; this is what keeps it that
    // way rather than relying on nobody ever making one. ExternalTargetLauncher in CRT.App makes
    // the same real-path check on the reading side.
    //
    // The walk takes the "is this a link" question as a delegate, so the rule is tested without
    // creating links (which Windows only permits with elevated rights).
    // ###########################################################################################
    public static class PublishPathSafety
    {
        // ###########################################################################################
        // The first path from just below `root` down to and including `target` that is a link, or
        // null when there is none. The root itself is the operator's choice and is not questioned
        // - a data tree mounted from elsewhere is a legitimate deployment.
        //
        // A target outside the root answers with the target itself: nothing should ever ask, and
        // "refuse" is the safe answer if something does.
        // ###########################################################################################
        public static string? FindLinkOnPath(string root, string target, Func<string, bool> isLink)
        {
            ArgumentNullException.ThrowIfNull(isLink);

            string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar);
            string fullTarget = Path.GetFullPath(target);

            if (!fullTarget.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.Ordinal))
                return fullTarget;

            string current = fullRoot;

            foreach (string segment in fullTarget[(fullRoot.Length + 1)..].Split(Path.DirectorySeparatorChar))
            {
                current = Path.Combine(current, segment);

                if (isLink(current))
                    return current;
            }

            return null;
        }

        // ###########################################################################################
        // The real question, for the running service. A path that does not exist yet is not a
        // link. A path whose nature cannot be determined IS treated as one - failing closed.
        // ###########################################################################################
        public static bool IsLink(string path)
        {
            try
            {
                return new FileInfo(path).LinkTarget is not null
                    || new DirectoryInfo(path).LinkTarget is not null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return true;
            }
        }
    }
}
