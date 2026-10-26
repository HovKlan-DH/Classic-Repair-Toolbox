using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedBoardLister - the boards in the data tree a maintainer can be assigned to.
    //
    // A real temp folder, because the thing under test IS the folder walk: which folders count as
    // a board, and which are the tree's own folders that must never get a maintainer. In the CRT
    // 4.0 tree (2026-10-25) a board is a folder holding its board file ("board.json").
    // ###########################################################################################
    public sealed class PublishedBoardListerTests : IDisposable
    {
        private readonly string thisRoot;

        public PublishedBoardListerTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-lister-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(this.thisRoot);
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

        private void Board(string manufacturer, string hardware, string board, string file = DataTreeLayout.BoardFileName)
        {
            string folder = Path.Combine(this.thisRoot, manufacturer, hardware, board);
            Directory.CreateDirectory(folder);

            if (file.Length > 0)
                File.WriteAllText(Path.Combine(folder, file), "x");
        }

        [Fact]
        public void Every_board_folder_holding_a_board_file_is_listed_with_its_three_parts()
        {
            this.Board("Commodore", "C64", "250407");
            this.Board("Commodore", "C128", "310378");
            this.Board("Amstrad", "CPC", "464");

            IReadOnlyList<PublishedBoardLister.KnownBoard> boards = PublishedBoardLister.List(this.thisRoot);

            Assert.Equal(
                ["Amstrad/CPC/464", "Commodore/C128/310378", "Commodore/C64/250407"],
                boards.Select(board => board.BoardId));

            PublishedBoardLister.KnownBoard c64 = boards.Single(board => board.BoardId == "Commodore/C64/250407");
            Assert.Equal(("Commodore", "C64", "250407"), (c64.Manufacturer, c64.Hardware, c64.Board));
        }

        // ###########################################################################################
        // The tree's own folders - the stored files ("Files/<ab>/...") and the MiniPro tests - are
        // not boards, and a maintainer assigned to one would review nothing. Both carry sub-folders
        // deep enough to look like a board to a naive walk.
        // ###########################################################################################
        [Fact]
        public void The_trees_own_folders_are_never_boards()
        {
            this.Board("Commodore", "C64", "250407");
            this.Board(DataTreeLayout.FilesFolder, "ab", "cd");
            this.Board("MiniPro", "IC tests", "74");

            IReadOnlyList<PublishedBoardLister.KnownBoard> boards = PublishedBoardLister.List(this.thisRoot);

            Assert.Equal("Commodore/C64/250407", Assert.Single(boards).BoardId);
        }

        // A folder with no board file is not a board - an Excel workbook alone included, which is a
        // generation of its own tree.
        [Fact]
        public void A_folder_with_no_board_file_is_not_a_board()
        {
            this.Board("Commodore", "C64", "250407", file: string.Empty);
            this.Board("Commodore", "C64", "250466", file: "Data C64 250466 v2.0.0.xlsx");
            this.Board("Commodore", "C64", "250425");

            Assert.Equal("Commodore/C64/250425", Assert.Single(PublishedBoardLister.List(this.thisRoot)).BoardId);
        }

        [Fact]
        public void A_missing_or_blank_root_lists_nothing()
        {
            Assert.Empty(PublishedBoardLister.List(null));
            Assert.Empty(PublishedBoardLister.List(string.Empty));
            Assert.Empty(PublishedBoardLister.List(Path.Combine(this.thisRoot, "does-not-exist")));
        }
    }
}
