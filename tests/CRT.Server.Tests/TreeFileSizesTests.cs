using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // TreeFileSizes - how big each file in a file tree is (owner request, 2026-10-04: "Ideally the
    // 'Files' actually also states its size (everywhere)"): the size of what opens from its row, read
    // through the one containment rule, and never a wrong size or an exception for a file not there.
    // ###########################################################################################
    public sealed class TreeFileSizesTests : IDisposable
    {
        private readonly string thisBeta;
        private readonly string thisProduction;

        public TreeFileSizesTests()
        {
            string root = Path.Combine(Path.GetTempPath(), "crt-tree-sizes", Guid.NewGuid().ToString("N"));
            this.thisBeta = Path.Combine(root, "beta");
            this.thisProduction = Path.Combine(root, "production");

            TreeFileSizesTests.Write(this.thisBeta, "Commodore/C64/250407/a.png", 10);
            TreeFileSizesTests.Write(this.thisBeta, "Commodore/Shared files/b.pdf", 20);
            TreeFileSizesTests.Write(this.thisProduction, "Commodore/C64/250407/old.png", 30);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(Path.GetDirectoryName(this.thisBeta)!, recursive: true);
            }
            catch (IOException)
            {
                // A temp folder left behind is not a failure.
            }
        }

        private static void Write(string root, string relative, int bytes)
        {
            string full = Path.Combine([root, .. relative.Split('/')]);
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllBytes(full, new byte[bytes]);
        }

        [Fact]
        public void A_file_in_the_tree_has_its_size_and_one_not_there_has_none()
        {
            Assert.Equal(10, TreeFileSizes.Of(this.thisBeta, "Commodore/C64/250407/a.png"));
            Assert.Null(TreeFileSizes.Of(this.thisBeta, "Commodore/C64/250407/missing.png"));
            Assert.Null(TreeFileSizes.Of(null, "Commodore/C64/250407/a.png"));
            Assert.Null(TreeFileSizes.Of(this.thisBeta, null));
        }

        // A path that climbs out of the tree is refused by the containment rule - no size, no read.
        [Fact]
        public void A_path_out_of_the_tree_has_no_size()
        {
            Assert.Null(TreeFileSizes.Of(this.thisBeta, "../production/Commodore/C64/250407/old.png"));
        }

        // The size of what OPENS from the row: BETA's file, production's, or the upload.
        [Fact]
        public void Each_entry_has_the_size_of_where_its_bytes_are()
        {
            IReadOnlyList<BoardFileEntry> entries = TreeFileSizes.Attach(
                [
                    new BoardFileEntry("Commodore/C64/250407/a.png", BoardFileChange.Unchanged, BoardFileSource.Beta),
                    new BoardFileEntry("Commodore/C64/250407/old.png", BoardFileChange.Removed, BoardFileSource.Production),
                    new BoardFileEntry("Commodore/C64/250407/new.png", BoardFileChange.Added, BoardFileSource.Submission, new string('a', 64)),
                    new BoardFileEntry("Commodore/C64/250407/Data.xlsx", BoardFileChange.Added, BoardFileSource.NotWrittenYet, WrittenOnApproval: true)
                ],
                this.thisBeta,
                this.thisProduction,
                new Dictionary<string, long> { ["Commodore/C64/250407/new.png"] = 40 });

            Assert.Equal([10L, 30L, 40L, null], entries.Select(entry => entry.SizeBytes));
        }

        [Fact]
        public void A_submissions_uploads_are_keyed_as_the_tree_spells_its_paths()
        {
            var manifest = new SubmissionManifest();
            manifest.Files.Add(new SubmissionFile { Path = "/Commodore\\C64/250407/a.png", Sha256 = new string('a', 64), SizeBytes = 7 });

            Assert.Equal(7, TreeFileSizes.OfSubmission(manifest)["Commodore/C64/250407/a.png"]);
        }

        // The production plan's sizes: what it copies and keeps as BETA holds them - what production
        // will have - and what it removes as production holds it now. A file not there is left out.
        [Fact]
        public void The_production_plans_sizes_come_from_where_each_file_will_be()
        {
            IReadOnlyDictionary<string, long> sizes = TreeFileSizes.ForPromotion(
                this.thisBeta,
                this.thisProduction,
                copied: ["Commodore/C64/250407/a.png"],
                unchanged: ["Commodore/Shared files/b.pdf", "Commodore/C64/250407/gone.png"],
                removed: ["Commodore/C64/250407/old.png"]);

            Assert.Equal(10, sizes["Commodore/C64/250407/a.png"]);
            Assert.Equal(20, sizes["Commodore/Shared files/b.pdf"]);
            Assert.Equal(30, sizes["Commodore/C64/250407/old.png"]);
            Assert.False(sizes.ContainsKey("Commodore/C64/250407/gone.png"));
        }
    }
}
