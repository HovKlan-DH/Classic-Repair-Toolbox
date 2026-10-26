using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedBoardFiles - a published board's files in the CRT 4.0 tree (2026-10-25): the
    // names its board file lists, and where each one's bytes are stored. Against a real temp tree,
    // since what is under test is which file on disk answers.
    //
    // The two answers that must never blur: a board the tree does not hold has NO files (a new
    // board, every file an addition), while a board file that cannot be read is NOT KNOWN - an
    // empty list there would turn every file of a real board into an addition.
    // ###########################################################################################
    public sealed class PublishedBoardFilesTests : IDisposable
    {
        private const string BoardId = "Commodore/C64/250407";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-published-files", Guid.NewGuid().ToString("N"));

        public PublishedBoardFilesTests()
        {
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

        [Fact]
        public void A_boards_files_are_found_by_name_whatever_the_capitals()
        {
            DataTreeBuilder.Board(this.thisRoot, PublishedBoardFilesTests.BoardId, ("Images/Sheet1.png", "sheet"), ("manual.pdf", "manual"));

            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? files = PublishedBoardFiles.ByName(this.thisRoot, PublishedBoardFilesTests.BoardId);

            Assert.NotNull(files);
            Assert.Equal(2, files!.Count);
            Assert.Equal(DataTreeBuilder.Sha("sheet"), files["images/sheet1.PNG"].Sha256);
        }

        [Fact]
        public void A_board_the_tree_does_not_hold_has_no_files_and_an_unreadable_one_is_not_known()
        {
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? absent = PublishedBoardFiles.ByName(this.thisRoot, PublishedBoardFilesTests.BoardId);

            Assert.NotNull(absent);
            Assert.Empty(absent!);

            File.WriteAllText(DataTreeBuilder.Full(this.thisRoot, DataTreeLayout.BoardFilePath(PublishedBoardFilesTests.BoardId)), "{ not a board file");

            Assert.Null(PublishedBoardFiles.ByName(this.thisRoot, PublishedBoardFilesTests.BoardId));
        }

        // A listed file's stored place, from its checksum - lower-cased, as the tree stores it - and
        // none for a checksum that is not one.
        [Fact]
        public void A_listed_files_stored_place_comes_from_its_checksum()
        {
            string sha = DataTreeBuilder.Sha("manual");

            Assert.Equal(
                DataTreeLayout.StoredFilePath(sha, "manual.pdf"),
                PublishedBoardFiles.StoredRelative(new BoardDocument.StoredFileRow { Name = "manual.pdf", Sha256 = sha.ToUpperInvariant() }));

            Assert.Equal(string.Empty, PublishedBoardFiles.StoredRelative(new BoardDocument.StoredFileRow { Name = "manual.pdf", Sha256 = "../../etc" }));
        }

        // Whether the bytes are there - and never a place outside the tree.
        [Fact]
        public void Whether_a_stored_file_is_there_is_asked_inside_the_tree_only()
        {
            string stored = DataTreeBuilder.StoreOnly(this.thisRoot, "manual.pdf", "manual");

            Assert.True(PublishedBoardFiles.IsStored(this.thisRoot, stored));
            Assert.False(PublishedBoardFiles.IsStored(this.thisRoot, DataTreeBuilder.Stored("other", "other.pdf")));
            Assert.False(PublishedBoardFiles.IsStored(this.thisRoot, "../outside.pdf"));
            Assert.False(PublishedBoardFiles.IsStored(null, stored));
        }

        // ###########################################################################################
        // Bytes are found by checksum whatever name they were stored under: the stored file of the
        // name asked first, then the same bytes stored for a name of another extension. Nothing for
        // bytes the tree lacks or a checksum that is not one.
        // ###########################################################################################
        [Fact]
        public void Bytes_are_found_by_their_checksum_under_any_extension()
        {
            string png = DataTreeBuilder.StoreOnly(this.thisRoot, "sheet.png", "the same bytes");
            string sha = DataTreeBuilder.Sha("the same bytes");

            Assert.Equal(DataTreeBuilder.Full(this.thisRoot, png), PublishedBoardFiles.FindStored(this.thisRoot, sha, "sheet.png"));
            Assert.Equal(DataTreeBuilder.Full(this.thisRoot, png), PublishedBoardFiles.FindStored(this.thisRoot, sha, "Sheet copy.PNG"));
            Assert.Equal(DataTreeBuilder.Full(this.thisRoot, png), PublishedBoardFiles.FindStored(this.thisRoot, sha, "sheet.jpg"));

            Assert.Equal(string.Empty, PublishedBoardFiles.FindStored(this.thisRoot, DataTreeBuilder.Sha("not stored"), "x.png"));
            Assert.Equal(string.Empty, PublishedBoardFiles.FindStored(this.thisRoot, "not a checksum", "x.png"));
        }

        // ###########################################################################################
        // The names a promotion plan answer carries (ProductionEndpoints.NamesOf): the plan's own,
        // and every removal's - the plan's winning where both name one place, since it describes
        // what stable holds afterwards.
        // ###########################################################################################
        [Fact]
        public void A_production_plans_names_are_the_plans_and_the_removals_together()
        {
            var plan = new ProductionPromotionResult(
                [], 0, [], TouchesSharedFiles: false,
                Names: new Dictionary<string, string> { ["Files/3f/3f7a.png"] = "Images/u8.png", ["Files/aa/aa11.pdf"] = "manual.pdf" });

            var removals = new FileRemovalPreview(
                ["Files/9c/9c01.pdf", "Files/aa/aa11.pdf"],
                null,
                new Dictionary<string, string> { ["Files/9c/9c01.pdf"] = "old.pdf", ["Files/aa/aa11.pdf"] = "another name.pdf" });

            IReadOnlyDictionary<string, string> names = ProductionEndpoints.NamesOf(plan, removals);

            Assert.Equal(3, names.Count);
            Assert.Equal("Images/u8.png", names["Files/3f/3f7a.png"]);
            Assert.Equal("old.pdf", names["Files/9c/9c01.pdf"]);
            Assert.Equal("manual.pdf", names["Files/aa/aa11.pdf"]);

            Assert.Empty(ProductionEndpoints.NamesOf(new ProductionPromotionResult([], 0, [], false), null));
        }
    }
}
