using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SubmissionFileTreeFlow - a submission's files as the BETA data will hold them after
    // approving it (owner request, 2026-09-28), against a real temp tree: the published board, the
    // publish plan and the removal preview are all real. In the CRT 4.0 tree (2026-10-25) the tree
    // is the board's NAMES - its board file and every file it lists.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class SubmissionFileTreeFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 28, 12, 0, 0, TimeSpan.Zero);

        private const string Board = "Commodore/C64/250407";

        private static readonly byte[] ManualBytes = System.Text.Encoding.UTF8.GetBytes("%PDF-1.4 the manual");

        private readonly string thisRoot;
        private readonly string thisDataTree;

        public SubmissionFileTreeFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-file-tree", Guid.NewGuid().ToString("N"));
            this.thisDataTree = Path.Combine(this.thisRoot, "app-data-BETA");
            Directory.CreateDirectory(this.thisDataTree);

            // The main file a new board's publish adds its row to - see ApprovePublishFlowTests.
            DataTreeBuilder.ListingMaster(this.thisDataTree, ApprovePublishFlowTests.OtherListedBoard);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private static SubmissionManifest Manifest(string summary, params SubmissionFile[] files)
        {
            var manifest = new SubmissionManifest
            {
                BoardId = SubmissionFileTreeFlowTests.Board,
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Files = [.. files]
            };

            foreach (SubmissionFile file in files)
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", Description = summary });
            manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1",
                BoardLabel = "U8",
                X = "100",
                Y = "200",
                Width = "40",
                Height = "20"
            });

            return manifest;
        }

        private static SubmissionFile Manual() => new()
        {
            Path = "manual.pdf",
            Sha256 = DataTreeBuilder.Sha(SubmissionFileTreeFlowTests.ManualBytes),
            SizeBytes = SubmissionFileTreeFlowTests.ManualBytes.Length
        };

        // BETA's board as `before` publishes it, listing the manual and `more`; the main file lists
        // only it, so the tree reads whole and a removal can be shown.
        private void Published(SubmissionManifest before, params (string Name, byte[] Bytes)[] more)
        {
            DataTreeBuilder.Master(this.thisDataTree, SubmissionFileTreeFlowTests.Board);
            DataTreeBuilder.Board(
                this.thisDataTree,
                SubmissionFileTreeFlowTests.Board,
                PublishMerge.Build(before, null),
                [("manual.pdf", SubmissionFileTreeFlowTests.ManualBytes), .. more]);
        }

        private Task<IReadOnlyList<BoardFileEntry>> BuildAsync(SubmissionManifest manifest) =>
            SubmissionFileTreeFlow.BuildAsync(
                this.thisDataTree, manifest, new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance), SubmissionFileTreeFlowTests.Now);

        private static BoardFileEntry Of(IReadOnlyList<BoardFileEntry> entries, string path) =>
            Assert.Single(entries, entry => entry.Path == path);

        // ###########################################################################################
        // A NEW BOARD "should just show how the new file-system will look like": its file new, and
        // the board file the approval writes new too, with nothing to open yet.
        // ###########################################################################################
        [Fact]
        public async Task A_new_boards_tree_is_all_new()
        {
            IReadOnlyList<BoardFileEntry> entries = await this.BuildAsync(SubmissionFileTreeFlowTests.Manifest(
                "b",
                new SubmissionFile { Path = "manual.pdf", Sha256 = new string('a', 64), SizeBytes = 10 }));

            Assert.All(entries, entry => Assert.Equal(BoardFileChange.Added, entry.Change));
            Assert.Equal(BoardFileSource.Submission, Of(entries, "manual.pdf").OpenFrom);

            BoardFileEntry boardFile = Of(entries, DataTreeLayout.BoardFileName);
            Assert.True(boardFile.WrittenOnApproval);
            Assert.Equal(BoardFileSource.NotWrittenYet, boardFile.OpenFrom);
            Assert.Equal(DataTreeLayout.BoardFilePath(SubmissionFileTreeFlowTests.Board), boardFile.PlaceInTree);

            // Sizes (owner request, 2026-10-04): the upload's, from the manifest; none for a file
            // with nothing to open yet.
            Assert.Equal(10, Of(entries, "manual.pdf").SizeBytes);
            Assert.Null(boardFile.SizeBytes);
        }

        // ###########################################################################################
        // *** THE REPORTED CASE (2026-09-28): only a one-liner changed. *** The board file is written
        // again - it carries the rows - and the manual the board lists, carried unchanged, is there
        // beside it as BETA holds it: opened from where its bytes lie, sized from there.
        // ###########################################################################################
        [Fact]
        public async Task A_rows_only_change_shows_only_the_board_file_changing()
        {
            this.Published(SubmissionFileTreeFlowTests.Manifest("a", SubmissionFileTreeFlowTests.Manual()));

            IReadOnlyList<BoardFileEntry> entries = await this.BuildAsync(
                SubmissionFileTreeFlowTests.Manifest("b", SubmissionFileTreeFlowTests.Manual()));

            Assert.Equal(
                [DataTreeLayout.BoardFileName],
                entries.Where(entry => entry.Change != BoardFileChange.Unchanged).Select(entry => entry.Path));

            BoardFileEntry manual = Of(entries, "manual.pdf");
            Assert.Equal(BoardFileChange.Unchanged, manual.Change);
            Assert.Equal(BoardFileSource.Beta, manual.OpenFrom);
            Assert.Equal(DataTreeBuilder.Stored("%PDF-1.4 the manual", "manual.pdf"), manual.PlaceInTree);
            Assert.Equal(SubmissionFileTreeFlowTests.ManualBytes.Length, manual.SizeBytes);
        }

        // ###########################################################################################
        // A file the board stops listing is shown REMOVED, by the name the board knew it by - the
        // same list the approval is sent back - and opened from where its bytes still lie.
        // ###########################################################################################
        [Fact]
        public async Task A_file_the_board_stops_listing_is_shown_removed_by_its_name()
        {
            byte[] old = System.Text.Encoding.UTF8.GetBytes("%PDF-1.4 the old pinout");
            this.Published(SubmissionFileTreeFlowTests.Manifest("a", SubmissionFileTreeFlowTests.Manual()), ("Pinouts/old.pdf", old));

            IReadOnlyList<BoardFileEntry> entries = await this.BuildAsync(
                SubmissionFileTreeFlowTests.Manifest("a", SubmissionFileTreeFlowTests.Manual()));

            BoardFileEntry removed = Of(entries, "Pinouts/old.pdf");
            Assert.Equal(BoardFileChange.Removed, removed.Change);
            Assert.Equal(DataTreeLayout.StoredFilePath(DataTreeBuilder.Sha(old), "Pinouts/old.pdf"), removed.PlaceInTree);
        }

        // ###########################################################################################
        // A file of the board carried with other bytes is CHANGED under its name - and, the file
        // being the board's own, it names no other board: in the 4.0 tree nothing is shared
        // (BoardFileEntry.AlsoUsedBy stays empty).
        // ###########################################################################################
        [Fact]
        public async Task A_replaced_file_is_changed_under_its_name_and_names_no_other_board()
        {
            this.Published(SubmissionFileTreeFlowTests.Manifest("a", SubmissionFileTreeFlowTests.Manual()));
            DataTreeBuilder.BoardWithBytes(this.thisDataTree, "Commodore/C128/310378", ("manual.pdf", SubmissionFileTreeFlowTests.ManualBytes));

            IReadOnlyList<BoardFileEntry> entries = await this.BuildAsync(SubmissionFileTreeFlowTests.Manifest(
                "a better manual",
                new SubmissionFile { Path = "manual.pdf", Sha256 = new string('b', 64), SizeBytes = 20 }));

            BoardFileEntry replaced = Of(entries, "manual.pdf");

            Assert.Equal(BoardFileChange.Changed, replaced.Change);
            Assert.Null(replaced.AlsoUsedBy);
        }
    }
}
