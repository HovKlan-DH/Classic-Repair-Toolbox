using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedBoardLocator - which file holds a board's published board, and what
    // "there is none" means.
    //
    // TWO THINGS HERE ARE SECURITY-SHAPED, not merely correctness-shaped:
    //
    //   1. The identity comes from a SUBMISSION and is untrusted. On a WRITE path a traversal
    //      would corrupt the tree; on this READ path it would disclose an arbitrary file's
    //      contents to a maintainer, turning a review screen into a file-disclosure hole.
    //   2. A missing board must read as "new board", never as an error - because a new board is
    //      the highest-risk submission there is and must reach a maintainer rather than failing to
    //      open. But a board that exists and cannot be READ must NOT read as "new board", or a
    //      maintainer approves a replacement for a board they were told did not exist.
    //
    // Uses a real temp folder: what is under test is which file on disk is chosen, which a fake
    // filesystem would not exercise.
    // ###########################################################################################
    public sealed class PublishedBoardLocatorTests : IDisposable
    {
        private readonly string thisRoot;

        public PublishedBoardLocatorTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-locator-tests", Guid.NewGuid().ToString("N"));
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

        private static SubmissionManifest Manifest(
            string manufacturer = "Commodore",
            string hardware = "C64",
            string board = "250407") => new()
            {
                BoardId = $"{manufacturer}/{hardware}/{board}",
                Manufacturer = manufacturer,
                Hardware = hardware,
                Board = board
            };

        private string BoardFolder(params string[] files)
        {
            string folder = Path.Combine(this.thisRoot, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);

            foreach (string file in files)
                File.WriteAllText(Path.Combine(folder, file), "not a real board");

            return folder;
        }

        // -----------------------------------------------------------------------------------
        // Finding the board
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** IN THE CRT 4.0 TREE A BOARD IS ITS BOARD FILE (2026-10-25). *** "<board id>/board.json",
        // always that name - no generations to choose between. An Excel workbook left beside it is
        // not the board, and is never chosen.
        // ###########################################################################################
        [Fact]
        public void The_board_file_is_chosen_and_never_a_workbook_beside_it()
        {
            this.BoardFolder("Data C64 250407 v2.0.0.xlsx", DataTreeLayout.BoardFileName);

            PublishedBoardLocation location =
                PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest());

            Assert.True(location.Exists);
            Assert.Equal(Path.Combine(this.thisRoot, "Commodore", "C64", "250407", DataTreeLayout.BoardFileName), location.WorkbookPath);
            Assert.Null(location.Generation);
        }

        // ###########################################################################################
        // From a board ID - what the review QUEUE holds, which loads no manifest (2026-09-26: the
        // "New board" badge). The same board file as from the manifest, and nothing at all for an id
        // that is not three well-formed parts - a traversal included.
        // ###########################################################################################
        [Fact]
        public void A_board_id_finds_the_same_board_as_its_manifest()
        {
            this.BoardFolder(DataTreeLayout.BoardFileName);

            PublishedBoardLocation byId = PublishedBoardLocator.LocateBoard(this.thisRoot, "Commodore/C64/250407");

            Assert.Equal(PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest()), byId);
            Assert.True(byId.Exists);
        }

        [Theory]
        [InlineData("Commodore/C64")]
        [InlineData("Commodore/C64/250407/extra")]
        [InlineData("../C64/250407")]
        [InlineData("")]
        [InlineData(null)]
        public void A_malformed_board_id_locates_nothing(string? boardId)
        {
            this.BoardFolder(DataTreeLayout.BoardFileName);

            Assert.False(PublishedBoardLocator.LocateBoard(this.thisRoot, boardId).Exists);
        }

        // A folder holding only an Excel generation's workbook is no board of this tree - a tree
        // the 4.0 converter has not been run over, or one copied wrongly by hand.
        [Fact]
        public void A_folder_holding_only_an_Excel_workbook_is_no_board()
        {
            this.BoardFolder("Data C64 250407.xlsx", "Data C64 250407 v2.0.0.xlsx");

            Assert.False(PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest()).Exists);
        }

        // The tree's own folders - the stored files, the MiniPro tests - are never a board's,
        // whatever lies in them (DataTreeLayout.ReservedTopFolders).
        [Theory]
        [InlineData("Files")]
        [InlineData("MiniPro")]
        [InlineData("files")]
        public void A_folder_of_the_tree_itself_is_never_a_board(string top)
        {
            string folder = Path.Combine(this.thisRoot, top, "ab", "cd");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, DataTreeLayout.BoardFileName), "{}");

            Assert.False(PublishedBoardLocator.LocateBoard(this.thisRoot, $"{top}/ab/cd").Exists);
        }

        // -----------------------------------------------------------------------------------
        // A new board - an answer, not a failure
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_board_with_no_folder_at_all_is_reported_as_having_no_board()
        {
            // The new-board case. It must not throw: a new board is the highest-risk submission
            // there is and has to reach a maintainer.
            PublishedBoardLocation location =
                PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest());

            Assert.False(location.Exists);
        }

        [Fact]
        public void A_folder_holding_no_board_file_is_reported_as_having_no_board()
        {
            // A board folder can exist carrying only images - a partially published board, or
            // one whose files arrived before its board file.
            string folder = this.BoardFolder();
            File.WriteAllText(Path.Combine(folder, "sheet1.png"), "x");

            Assert.False(PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest()).Exists);
        }

        // -----------------------------------------------------------------------------------
        // The identity is untrusted
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("..")]
        [InlineData("../..")]
        [InlineData("/etc")]
        public void A_TRAVERSAL_in_the_identity_finds_nothing_rather_than_escaping_the_tree(string manufacturer)
        {
            // *** ON A READ PATH, A TRAVERSAL IS FILE DISCLOSURE. *** The identity arrives in a
            // submission. Without containment this endpoint would hand a maintainer the contents of
            // any file the service can read, which is a far worse outcome than the write-side
            // corruption the same check prevents elsewhere.
            //
            // NOTE: on its own this test is WEAK - see the one below. It passes even with
            // containment removed entirely, because these paths point at folders that do not
            // exist. It is kept because it costs nothing and covers the shapes, but the test that
            // actually proves the check is A_traversal_onto_a_REAL_folder_outside_the_tree_is_refused.
            PublishedBoardLocation location = PublishedBoardLocator.Locate(
                this.thisRoot,
                PublishedBoardLocatorTests.Manifest(manufacturer, "x", "y"));

            Assert.False(location.Exists);
        }

        [Fact]
        public void A_traversal_onto_a_REAL_folder_outside_the_tree_is_refused()
        {
            // *** THE TEST THAT ACTUALLY PROVES CONTAINMENT, and the reason it exists. ***
            //
            // The theory above was written first and was VACUOUS: removing SubmissionPathRules
            // from the locator entirely left all twelve tests passing, because every traversal it
            // tried pointed at a folder that did not exist, and the "does this folder exist"
            // check refused them for the wrong reason. A traversal is only dangerous when it
            // lands somewhere REAL, so that is what this builds.
            //
            // The layout mirrors a genuine deployment: the data tree is one folder among several
            // under a parent, and a sibling holds a board the maintainer must never be shown as if
            // it were this board's history.
            string parent = Path.Combine(this.thisRoot, "parent");
            string dataTree = Path.Combine(parent, "beta");
            string sibling = Path.Combine(parent, "Commodore", "C64", "250407");

            Directory.CreateDirectory(dataTree);
            Directory.CreateDirectory(sibling);
            File.WriteAllText(Path.Combine(sibling, DataTreeLayout.BoardFileName), "outside the tree");

            // "../Commodore" resolves to a folder that exists and holds a board file.
            PublishedBoardLocation escaped = PublishedBoardLocator.Locate(
                dataTree,
                PublishedBoardLocatorTests.Manifest("../Commodore", "C64", "250407"));

            Assert.False(escaped.Exists);

            // Anti-vacuity: the same folder INSIDE the tree is found, so the refusal above is
            // containment doing its job rather than the locator failing to find anything at all.
            string inside = Path.Combine(dataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(inside);
            File.WriteAllText(Path.Combine(inside, DataTreeLayout.BoardFileName), "inside the tree");

            PublishedBoardLocation found = PublishedBoardLocator.Locate(
                dataTree,
                PublishedBoardLocatorTests.Manifest("Commodore", "C64", "250407"));

            Assert.True(found.Exists);
            Assert.StartsWith(Path.GetFullPath(dataTree), Path.GetFullPath(found.WorkbookPath), StringComparison.Ordinal);
        }

        [Fact]
        public void A_blank_identity_finds_nothing_rather_than_the_data_root_itself()
        {
            // Falling back to the root would compare the submission against whatever board file
            // happened to sit at the top of the data tree - a diff against an unrelated board,
            // presented to the maintainer as this board's history.
            PublishedBoardLocation location = PublishedBoardLocator.Locate(
                this.thisRoot,
                PublishedBoardLocatorTests.Manifest("", "", ""));

            Assert.False(location.Exists);
        }

        [Fact]
        public void A_blank_data_root_finds_nothing()
        {
            Assert.False(PublishedBoardLocator.Locate("", PublishedBoardLocatorTests.Manifest()).Exists);
            Assert.False(PublishedBoardLocator.Locate(null, PublishedBoardLocatorTests.Manifest()).Exists);
        }

        [Fact]
        public void A_null_manifest_is_refused()
        {
            Assert.Throws<ArgumentNullException>(() => PublishedBoardLocator.Locate(this.thisRoot, null!));
        }

        // -----------------------------------------------------------------------------------
        // Containment, proved positively
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_located_board_is_always_INSIDE_the_data_tree()
        {
            // The anti-vacuity partner to the traversal tests above: those pass trivially if
            // nothing is ever found. This proves the ordinary case really does resolve, and
            // resolves inside the root.
            this.BoardFolder(DataTreeLayout.BoardFileName);

            PublishedBoardLocation location =
                PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest());

            Assert.True(location.Exists);
            Assert.StartsWith(
                Path.GetFullPath(this.thisRoot),
                Path.GetFullPath(location.WorkbookPath),
                StringComparison.Ordinal);
        }
    }
}
