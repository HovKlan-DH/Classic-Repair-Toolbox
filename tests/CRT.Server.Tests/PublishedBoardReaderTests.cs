using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedBoardReader - reading the board a submission is compared against.
    //
    // *** THE TEST THAT MATTERS MOST HERE IS THE FRESHNESS ONE. *** BoardDataReader keeps a static
    // cache keyed by whatever string the caller supplies, and the first version of this class
    // passed the workbook's own PATH - which reads as obviously correct and is exactly wrong on a
    // server. The file at that path is REWRITTEN by every publish, so a cached read would serve
    // the PRE-PUBLISH board to the next maintainer opening a submission for that board, and show
    // them a diff against data that no longer exists.
    //
    // It surfaced as PublishExecutorTests.Re_running_the_same_publish_is_safe failing
    // intermittently, which is the sort of flake that gets rerun until it passes. The fix is a
    // fresh cache key per read; this file is what stops it coming back.
    //
    // Uses a real temp folder and a real workbook, because what is under test is reading a file
    // that changes underneath - which a fake filesystem could not express.
    // ###########################################################################################
    // Shares the "BoardFiles" collection - see PublishExecutorTests for why these must not run
    // in parallel.
    [Collection("BoardFiles")]
    public sealed class PublishedBoardReaderTests : IDisposable
    {
        private readonly string thisRoot;

        public PublishedBoardReaderTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-published-reader", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(this.thisRoot);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(this.thisRoot))
                    Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless; failing a test over cleanup is not.
            }
        }

        private static PublishedBoardReader Reader() => new(NullLogger<PublishedBoardReader>.Instance);

        private static SubmissionManifest Manifest() => new()
        {
            BoardId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407"
        };

        private string WriteBoard(string partNumber)
        {
            string folder = Path.Combine(this.thisRoot, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);

            string path = Path.Combine(folder, "Data C64 250407 v2.0.0.xlsx");

            BoardWorkbookWriter.Write(path, new BoardData
            {
                RevisionDate = "2026-August-21",
                Components =
                [
                    new ComponentEntry
                    {
                        BoardLabel = "U8",
                        FriendlyName = "PLA",
                        TechnicalNameOrValue = "906114-01",
                        PartNumber = partNumber
                    }
                ]
            });

            return path;
        }

        [Fact]
        public async Task A_published_board_is_read()
        {
            this.WriteBoard("906114");

            BoardData? board = await PublishedBoardReaderTests.Reader()
                .TryReadAsync(this.thisRoot, PublishedBoardReaderTests.Manifest());

            Assert.NotNull(board);
            Assert.Equal("906114", Assert.Single(board!.Components).PartNumber);
        }

        [Fact]
        public async Task A_REWRITTEN_board_is_read_afresh_rather_than_from_cache()
        {
            // *** THE REGRESSION TEST FOR THE STALE-CACHE BUG. *** Publishing rewrites the board
            // in place. A reader caching by path would hand the second call the FIRST board, and
            // the maintainer would be comparing a submission against a board that has already been
            // replaced - with no indication anything was wrong.
            //
            // This fails against the version that passed location.WorkbookPath as the cache key.
            this.WriteBoard("906114");

            BoardData? before = await PublishedBoardReaderTests.Reader()
                .TryReadAsync(this.thisRoot, PublishedBoardReaderTests.Manifest());

            Assert.Equal("906114", Assert.Single(before!.Components).PartNumber);

            // The same path, different contents - exactly what a publish does.
            this.WriteBoard("251715-01");

            BoardData? after = await PublishedBoardReaderTests.Reader()
                .TryReadAsync(this.thisRoot, PublishedBoardReaderTests.Manifest());

            Assert.Equal("251715-01", Assert.Single(after!.Components).PartNumber);
        }

        [Fact]
        public async Task A_board_that_was_never_published_reads_as_NULL()
        {
            // The new-board case, and a first-class answer rather than a failure: a new board is
            // the highest-risk submission there is and must reach a maintainer.
            BoardData? board = await PublishedBoardReaderTests.Reader()
                .TryReadAsync(this.thisRoot, PublishedBoardReaderTests.Manifest());

            Assert.Null(board);
        }

        [Fact]
        public async Task A_null_published_board_produces_a_NEW_BOARD_summary()
        {
            // The whole point of returning null rather than an empty board: an empty board would
            // report every row as an addition, which is the same information with the one fact
            // that matters stripped out - that nobody has ever vetted this board.
            BoardData? published = await PublishedBoardReaderTests.Reader()
                .TryReadAsync(this.thisRoot, PublishedBoardReaderTests.Manifest());

            ReviewChangeSummary summary = ReviewSummary.Compare(
                published,
                new BoardData
                {
                    Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "A", TechnicalNameOrValue = "B" }]
                });

            Assert.True(summary.IsNewBoard);
            Assert.Contains("New board", summary.Describe());
        }

        [Fact]
        public async Task A_null_manifest_is_refused()
        {
            await Assert.ThrowsAsync<ArgumentNullException>(
                () => PublishedBoardReaderTests.Reader().TryReadAsync(this.thisRoot, null!));
        }
    }
}
