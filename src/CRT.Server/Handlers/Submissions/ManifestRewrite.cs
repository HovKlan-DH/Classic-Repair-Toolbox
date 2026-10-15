using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A DATA TREE'S dataChecksums.json, REBUILT AFTER SOMETHING WROTE THE TREE OUTSIDE THE LOCK
    // (code review, 2026-10-13).
    //
    // Four callers rebuilt a manifest after the publish lock had been let go - after a publish to
    // BETA, a change published from the Boards screen, a new board placed in BETA's lists, and a
    // promotion to the stable source - on the request thread, and scanning a tree another publish
    // could be writing at that moment. The manifest then described neither state. Now each takes
    // the one PublishLock for the scan and the write, and runs them on the pool: the scan walks
    // every file of the tree (hashing only those that changed - DataChecksumManifest's cache).
    //
    // *** NEVER CANCELLED. *** It runs after a write that has already happened; a client giving up
    // must not leave the tree's manifest older than the tree. Never throws either: -1 is a failure,
    // which DataChecksumManifest.Write has already logged, and the caller says what it means.
    //
    // NOT for a caller already holding the lock - the lock is not re-entrant. The flows that rebuild
    // INSIDE their own hold of it (a push-back, a deletion, the unused files, the drop-down order,
    // the administrator's rebuild) call DataChecksumManifest.Write directly.
    // ###########################################################################################
    public static class ManifestRewrite
    {
        public static async Task<int> RewriteAsync(
            PublishLock publishLock,
            string? dataTreeRoot,
            string? publicBaseUrl,
            string? manifestPath,
            Func<string, string, string, int>? write = null)
        {
            ArgumentNullException.ThrowIfNull(publishLock);

            Func<string, string, string, int> writer = write ?? DataChecksumManifest.Write;

            try
            {
                using IDisposable gate = await publishLock.EnterAsync(CancellationToken.None).ConfigureAwait(false);

                return await Task.Run(() => writer(
                    dataTreeRoot ?? string.Empty,
                    publicBaseUrl ?? string.Empty,
                    manifestPath ?? string.Empty)).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Could not rebuild the checksum manifest [{manifestPath}] - [{ex}]");
                return -1;
            }
        }
    }
}
