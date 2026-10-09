using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Reads the board a board is PUBLISHED at, so a submission can be compared against it
    // (NewContributeStrategy.md Phase 5, task 3).
    //
    // *** A MISSING BOARD IS AN ANSWER, NOT AN ERROR. *** It means the board has never been
    // published - a NEW BOARD - which is the highest-risk submission there is (Phase 6 task 3)
    // and must reach a maintainer rather than failing to open. `ReviewSummary.Compare` takes a null
    // published board and reports "New board, N rows" for exactly this case, so the null travels
    // all the way through rather than being turned into an empty board here. An empty board would
    // report every row as an addition, which is the same information stripped of the one fact
    // that matters: that nobody has ever vetted this board.
    //
    // *** IT READS THE NEWEST GENERATION, using the same rule publishing writes with. *** If this
    // read an older, frozen generation the maintainer would be shown a diff against a board no
    // current build uses - every change would look real and most of them would be noise from the
    // generation gap. DataGenerationRules owns that decision for both directions.
    //
    // *** THIS IS WHERE EPPlus RUNS ON THE SERVER. *** Open question 8 predicted Phase 4 task 3;
    // it turned out to be here and the publish writer. The project owner has decided to proceed on
    // the current licence and settle it separately.
    //
    // An I/O boundary, so it is thin on purpose: resolving WHICH file and deciding what a missing
    // one means are the parts worth testing, and they live in PublishedBoardLocator.
    // ###########################################################################################
    public sealed class PublishedBoardReader
    {
        private readonly ILogger<PublishedBoardReader> thisLogger;

        public PublishedBoardReader(ILogger<PublishedBoardReader> logger)
        {
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // The published board for a board, or null when it has never been published.
        //
        // *** THE CACHE IS DELIBERATELY BYPASSED, AND THIS WAS A REAL BUG. *** BoardDataReader
        // keeps a static cache keyed by whatever string the caller supplies, and the first version
        // here passed the workbook's own PATH - which looks like the obviously right key and is
        // exactly wrong for this caller. The file at that path is REWRITTEN by every publish, so
        // after a merge the next maintainer to open a submission for that board would be served the
        // PRE-PUBLISH board out of cache and shown a diff against data that no longer exists. In
        // the app the same key is safe because CRT never rewrites a published board; on the server
        // it is the one thing that does.
        //
        // Caught by PublishExecutorTests.Re_running_the_same_publish_is_safe failing
        // intermittently once these tests shared an assembly - the second publish read the first
        // one's cached board.
        //
        // A fresh key per read means every call parses. That is the correct trade here: a review
        // opens one board at a time and a parse is milliseconds, whereas serving a stale board to
        // somebody deciding what to publish is unbounded damage. The cache entry is cleared
        // immediately afterwards so the unbounded key space cannot grow either.
        // ###########################################################################################
        public async Task<BoardData?> TryReadAsync(
            string dataTreeRoot,
            SubmissionManifest manifest,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return await this.ReadAsync(PublishedBoardLocator.Locate(dataTreeRoot, manifest), manifest.BoardId, cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // The same, for a board named by its id rather than by a submission - the Boards screen's
        // Board data and Files views (2026-10-03). Located by PublishedBoardLocator.LocateBoard, the
        // rule the review queue tells a new board by.
        // ###########################################################################################
        public Task<BoardData?> TryReadBoardAsync(
            string dataTreeRoot,
            string boardId,
            CancellationToken cancellationToken = default) =>
            this.ReadAsync(PublishedBoardLocator.LocateBoard(dataTreeRoot, boardId), boardId, cancellationToken);

        private async Task<BoardData?> ReadAsync(
            PublishedBoardLocation location,
            string? boardId,
            CancellationToken cancellationToken)
        {
            if (!location.Exists)
            {
                // Logged at INFORMATION rather than WARNING: a new board is an ordinary and
                // expected state, not a fault, and logging it as one trains people to ignore
                // warnings.
                this.thisLogger.LogInformation(
                    "No published board for {BoardId}; treating as a new board.",
                    boardId);

                return null;
            }

            cancellationToken.ThrowIfCancellationRequested();

            string cacheKey = $"review:{Guid.NewGuid():N}";

            try
            {
                return await BoardDataReader
                    .LoadAsync(location.WorkbookPath, cacheKey)
                    .ConfigureAwait(false);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                // A board that cannot be READ is NOT the same as one that does not exist, and the
                // difference is the whole point: returning null here would tell the maintainer this
                // is a brand-new board, and they would approve it on that basis while a
                // published board sat on disk unreadable. So this fails loudly instead.
                this.thisLogger.LogError(
                    exception,
                    "The published board for {BoardId} at {Path} could not be read.",
                    boardId,
                    location.WorkbookPath);

                throw;
            }
            finally
            {
                // The key is unique per call, so leaving it behind would grow the static cache by
                // one full board per review opened and never hit again.
                BoardDataReader.ClearCache(cacheKey);
            }
        }
    }
}
