using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedTreeProbe - the server's real answers to "what does the published board list
    // under this name" and "what is in this folder", which the submission rules ask (security
    // review, 2026-09-25). In the CRT 4.0 tree (2026-10-25) a view is of ONE board: its board
    // file's list of files by name - nothing is hashed - and the tree's own folders.
    // ###########################################################################################
    public sealed class PublishedTreeProbeTests : IDisposable
    {
        private readonly string thisRoot;
        private readonly string thisTree;

        public PublishedTreeProbeTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-probe-tests", Guid.NewGuid().ToString("N"));
            this.thisTree = Path.Combine(this.thisRoot, "Data");

            DataTreeBuilder.Board(this.thisTree, DataTreeBuilder.BoardId, ("Images/Sheet1.png", "scan"));
            File.WriteAllText(Path.Combine(this.thisRoot, "secret.txt"), "outside the tree");
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // Harmless.
            }
        }

        [Fact]
        public void A_name_the_board_lists_answers_its_checksum_and_size()
        {
            PublishedTreeView board = PublishedTreeProbe.For(this.thisTree, DataTreeBuilder.BoardId)!;

            Assert.Equal(DataTreeBuilder.Sha("scan"), board.HashOf("Images/Sheet1.png"));
            Assert.Equal(4, board.SizeOf("Images/Sheet1.png"));

            // Names compare without regard to capitals, as a board's names do.
            Assert.Equal(DataTreeBuilder.Sha("scan"), board.HashOf("images/sheet1.PNG"));
            Assert.Null(board.HashOf("Images/other.png"));
        }

        [Fact]
        public void A_case_variant_of_a_real_board_folder_is_found_on_disk()
        {
            PublishedTreeView board = PublishedTreeProbe.For(this.thisTree, "commodore/C64/250407")!;

            // On a case-insensitive filesystem the probe still reports the name AS IT EXISTS, so the
            // check works the same on the developer's Windows machine and on the Linux server.
            Assert.Equal("Commodore", board.FindCaseVariant("commodore/C64/250407"));
            Assert.Null(board.FindCaseVariant("Commodore/C64/250407"));
        }

        // A hostile path must not make the probe list anything outside the tree.
        [Fact]
        public void Nothing_outside_the_tree_is_listed_or_answered()
        {
            PublishedTreeView board = PublishedTreeProbe.For(this.thisTree, DataTreeBuilder.BoardId)!;

            Assert.Null(board.HashOf("../secret.txt"));
            Assert.Null(board.FindCaseVariant("../secret.txt"));
        }

        // A board the tree does not hold is an EMPTY list - a new board has no files yet.
        [Fact]
        public void A_board_the_tree_does_not_hold_lists_nothing()
        {
            PublishedTreeView board = PublishedTreeProbe.For(this.thisTree, "Commodore/C64/250425")!;

            Assert.Null(board.HashOf("Images/Sheet1.png"));
        }

        // No usable root, or a board file that cannot be read, is "what is published could not be
        // consulted" - the safe answer: nothing is admitted for being published already.
        [Fact]
        public void No_usable_root_or_an_unreadable_board_file_gives_no_view_at_all()
        {
            Assert.Null(PublishedTreeProbe.For(null, DataTreeBuilder.BoardId));
            Assert.Null(PublishedTreeProbe.For(Path.Combine(this.thisRoot, "does-not-exist"), DataTreeBuilder.BoardId));

            File.WriteAllText(Path.Combine(this.thisTree, "Commodore", "C64", "250407", "board.json"), "{ not json");

            Assert.Null(PublishedTreeProbe.For(this.thisTree, DataTreeBuilder.BoardId));
        }
    }
}
