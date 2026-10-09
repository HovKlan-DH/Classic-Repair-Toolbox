using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // system.json IS RETIRED (owner decision, 2026-09-25): "I do not want this file visible in
    // the source ... it should not be something downloaded by all users, as this file is not
    // relevant for them."
    //
    // It was written into every published board's folder (Phase 4 task 7) and so synced to every
    // user, while nothing in CRT showed anything from it. Every fact it held is in the database
    // already - the revision and content hash on `boards`, the maintainers in `maintainers`, the
    // origin on `boards.origin` - so it is not moved anywhere: it is simply no longer written.
    //
    // Builds before this one DID write it, so one can still be sitting in a board folder in BETA,
    // and in production if BETA was copied across by hand. This removes such a leftover: the BETA
    // publish does it for the board it publishes, the production promotion for the board it
    // promotes. Once removed it drops out of that tree's dataChecksums.json, and CRT deletes its
    // own copy on the next sync when the user allows deletion of unused files.
    // ###########################################################################################
    public static class RetiredBoardDescriptor
    {
        // ###########################################################################################
        // Deletes "<boardFolder>/system.json" if it is there. True when there is none afterwards.
        //
        // The folder has already been resolved inside `dataRoot` by the caller; the link check is
        // repeated here because this is a DELETE, and a link on the way would delete elsewhere.
        // Never throws: a leftover that could not be removed is a stray file, not a failed publish.
        // ###########################################################################################
        public static bool TryRemove(string dataRoot, string boardFolder, ILogger? logger = null)
        {
            if (string.IsNullOrWhiteSpace(dataRoot) || string.IsNullOrWhiteSpace(boardFolder))
                return false;

            string path = Path.Combine(boardFolder, BoardDescriptorStore.FileName);

            try
            {
                if (!File.Exists(path))
                    return true;

                if (PublishPathSafety.FindLinkOnPath(dataRoot, path, PublishPathSafety.IsLink) is not null)
                {
                    logger?.LogWarning("Did not remove the retired [{Path}]: a symbolic link is on the way to it.", path);
                    return false;
                }

                File.Delete(path);
                logger?.LogInformation("Removed the retired [{Path}].", path);

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                logger?.LogWarning(ex, "Could not remove the retired [{Path}].", path);
                return false;
            }
        }
    }
}
