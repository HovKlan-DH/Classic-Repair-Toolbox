using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedBoardLocator - which file holds a system's published board, and what
    // "there is none" means.
    //
    // TWO THINGS HERE ARE SECURITY-SHAPED, not merely correctness-shaped:
    //
    //   1. The identity comes from a SUBMISSION and is untrusted. On a WRITE path a traversal
    //      would corrupt the tree; on this READ path it would disclose an arbitrary file's
    //      contents to a maintainer, turning a review screen into a file-disclosure hole.
    //   2. A missing board must read as "new system", never as an error - because a new system is
    //      the highest-risk submission there is and must reach a maintainer rather than failing to
    //      open. But a board that exists and cannot be READ must NOT read as "new system", or a
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
                SystemId = $"{manufacturer}/{hardware}/{board}",
                Manufacturer = manufacturer,
                Hardware = hardware,
                Board = board
            };

        private string SystemFolder(params string[] workbooks)
        {
            string folder = Path.Combine(this.thisRoot, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);

            foreach (string workbook in workbooks)
                File.WriteAllText(Path.Combine(folder, workbook), "not a real workbook");

            return folder;
        }

        // -----------------------------------------------------------------------------------
        // Finding the board
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_NEWEST_generation_is_chosen()
        {
            // *** THE SAME RULE PUBLISHING WRITES WITH. *** Reading an older, frozen generation
            // would show the maintainer a diff against a board no current build uses, and every
            // difference between the generations would appear as a change the contributor made.
            this.SystemFolder("Data C64 250407.xlsx", "Data C64 250407 v2.0.0.xlsx");

            PublishedBoardLocation location =
                PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest());

            Assert.True(location.Exists);
            Assert.EndsWith("Data C64 250407 v2.0.0.xlsx", location.WorkbookPath);
            Assert.Equal(Version.Parse("2.0.0"), location.Generation);
        }

        // ###########################################################################################
        // From a system ID - what the review QUEUE holds, which loads no manifest (2026-09-26: the
        // "New system" badge). The same workbook as from the manifest, and nothing at all for an id
        // that is not three well-formed parts - a traversal included.
        // ###########################################################################################
        [Fact]
        public void A_system_id_finds_the_same_board_as_its_manifest()
        {
            this.SystemFolder("Data C64 250407.xlsx", "Data C64 250407 v2.0.0.xlsx");

            PublishedBoardLocation byId = PublishedBoardLocator.LocateSystem(this.thisRoot, "Commodore/C64/250407");

            Assert.Equal(PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest()), byId);
            Assert.True(byId.Exists);
        }

        [Theory]
        [InlineData("Commodore/C64")]
        [InlineData("Commodore/C64/250407/extra")]
        [InlineData("../C64/250407")]
        [InlineData("")]
        [InlineData(null)]
        public void A_malformed_system_id_locates_nothing(string? systemId)
        {
            this.SystemFolder("Data C64 250407.xlsx");

            Assert.False(PublishedBoardLocator.LocateSystem(this.thisRoot, systemId).Exists);
        }

        [Fact]
        public void A_tree_with_only_the_unversioned_original_finds_it()
        {
            this.SystemFolder("Data C64 250407.xlsx");

            PublishedBoardLocation location =
                PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest());

            Assert.True(location.Exists);
            Assert.Null(location.Generation);
        }

        [Fact]
        public void The_board_stem_is_READ_from_disk_rather_than_rebuilt_from_the_identity()
        {
            // Board file names do not follow the folder names mechanically: "Data C128DCR 250477"
            // lives under C128/250477, and "Data VIC20 250403" under VIC-20/250403. Rebuilding a
            // name from the identity would miss every board whose file is not named after its
            // folder - which is most of them.
            string folder = Path.Combine(this.thisRoot, "Commodore", "C128", "250477");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "Data C128DCR 250477 v2.0.0.xlsx"), "x");

            PublishedBoardLocation location = PublishedBoardLocator.Locate(
                this.thisRoot,
                PublishedBoardLocatorTests.Manifest("Commodore", "C128", "250477"));

            Assert.True(location.Exists);
            Assert.EndsWith("Data C128DCR 250477 v2.0.0.xlsx", location.WorkbookPath);
        }

        // -----------------------------------------------------------------------------------
        // A new system - an answer, not a failure
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_system_with_no_folder_at_all_is_reported_as_having_no_board()
        {
            // The new-system case. It must not throw: a new system is the highest-risk submission
            // there is and has to reach a maintainer.
            PublishedBoardLocation location =
                PublishedBoardLocator.Locate(this.thisRoot, PublishedBoardLocatorTests.Manifest());

            Assert.False(location.Exists);
        }

        [Fact]
        public void A_folder_holding_no_workbook_is_reported_as_having_no_board()
        {
            // A system folder can exist carrying only images - a partially published system, or
            // one whose files arrived before its workbook.
            string folder = this.SystemFolder();
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
            // it were this system's history.
            string parent = Path.Combine(this.thisRoot, "parent");
            string dataTree = Path.Combine(parent, "beta");
            string sibling = Path.Combine(parent, "Commodore", "C64", "250407");

            Directory.CreateDirectory(dataTree);
            Directory.CreateDirectory(sibling);
            File.WriteAllText(Path.Combine(sibling, "Data C64 250407 v2.0.0.xlsx"), "outside the tree");

            // "../Commodore" resolves to a folder that exists and holds a real workbook.
            PublishedBoardLocation escaped = PublishedBoardLocator.Locate(
                dataTree,
                PublishedBoardLocatorTests.Manifest("../Commodore", "C64", "250407"));

            Assert.False(escaped.Exists);

            // Anti-vacuity: the same folder INSIDE the tree is found, so the refusal above is
            // containment doing its job rather than the locator failing to find anything at all.
            string inside = Path.Combine(dataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(inside);
            File.WriteAllText(Path.Combine(inside, "Data C64 250407 v2.0.0.xlsx"), "inside the tree");

            PublishedBoardLocation found = PublishedBoardLocator.Locate(
                dataTree,
                PublishedBoardLocatorTests.Manifest("Commodore", "C64", "250407"));

            Assert.True(found.Exists);
            Assert.StartsWith(Path.GetFullPath(dataTree), Path.GetFullPath(found.WorkbookPath), StringComparison.Ordinal);
        }

        [Fact]
        public void A_blank_identity_finds_nothing_rather_than_the_data_root_itself()
        {
            // Falling back to the root would compare the submission against whatever workbook
            // happened to sit at the top of the data tree - a diff against an unrelated board,
            // presented to the maintainer as this system's history.
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
            this.SystemFolder("Data C64 250407 v2.0.0.xlsx");

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
