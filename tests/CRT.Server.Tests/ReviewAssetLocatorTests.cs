using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewAssetLocator - which bytes a maintainer is allowed to fetch (Phase 5, task 4).
    //
    // *** THIS IS THE FILE-DISCLOSURE BOUNDARY, so read the anti-vacuity note before adding a
    // test here. *** Two traversal tests in this project have already been caught proving nothing
    // (PublishedBoardLocator's, and the board round-trip): a traversal aimed at a folder that does
    // not exist is refused by File.Exists rather than by containment, so the test passes with the
    // guard deleted. Every containment test below therefore points at a REAL file outside the
    // tree, and carries an anti-vacuity half proving the same shape INSIDE the tree is still
    // found - otherwise "refused" could just mean "found nothing anywhere".
    //
    // The submitted side and the published side are guarded by DIFFERENT mechanisms on purpose,
    // because they take different input:
    //
    //   SUBMITTED - addressed by HASH. A hash cannot carry a traversal, so containment is not the
    //               risk; the risk is a maintainer fetching a blob belonging to some OTHER
    //               submission. Guarded by requiring the manifest to reference that hash.
    //   PUBLISHED - addressed by the board's NAME for a file (the CRT 4.0 tree, 2026-10-25).
    //               Guarded by the published board's own list - only a name it lists is served -
    //               and by serving the stored file of its checksum, never a place built from the
    //               name.
    // ###########################################################################################
    // Shares the "BoardFiles" collection - see PublishExecutorTests for why these must not run
    // in parallel.
    [Collection("BoardFiles")]
    public sealed class ReviewAssetLocatorTests : IDisposable
    {
        private readonly string thisRoot;
        private readonly string thisDataTree;

        public ReviewAssetLocatorTests()
        {
            this.thisRoot = Path.Combine(
                Path.GetTempPath(), "crt-review-assets", Guid.NewGuid().ToString("N"));

            this.thisDataTree = Path.Combine(this.thisRoot, "app-data-BETA");

            Directory.CreateDirectory(this.thisDataTree);
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

        // ###########################################################################################
        // The submission under review.
        //
        // *** `referencedFiles` MATTERS SINCE 2026-09-23: the published-file route now serves only
        // a path the submission own BOARD references. *** Naming the file in Files alone is not
        // enough - that list is the uploaded blobs, whereas the scope check reads the rows. A test
        // that wants a published file to be found must name it here.
        // ###########################################################################################
        private static SubmissionManifest Manifest(params string[] hashes)
        {
            return ReviewAssetLocatorTests.ManifestReferencing(
                ["Commodore/C64/250407/Schematics/board.png"],
                hashes);
        }

        private static SubmissionManifest ManifestReferencing(
            string[] referencedFiles,
            params string[] hashes)
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407"
            };

            foreach (string file in referencedFiles)
            {
                manifest.Rows.Schematics.Add(new BoardSchematicEntry
                {
                    SchematicName = file,
                    SchematicImageFile = file
                });
            }

            foreach (string hash in hashes)
            {
                manifest.Files.Add(new SubmissionFile
                {
                    Path = "Schematics/board.png",
                    Sha256 = hash,
                    SizeBytes = 1024
                });
            }

            return manifest;
        }

        private static string Hash(char fill) => new(fill, 64);

        // -----------------------------------------------------------------------------------
        // The SUBMITTED side: addressed by hash, scoped to the manifest.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_hash_the_manifest_references_is_allowed()
        {
            string hash = ReviewAssetLocatorTests.Hash('a');

            Assert.True(ReviewAssetLocator.IsSubmittedBlobAllowed(
                ReviewAssetLocatorTests.Manifest(hash), hash));
        }

        [Fact]
        public void A_hash_the_manifest_does_NOT_reference_is_refused()
        {
            // *** THE CONFINEMENT THAT MATTERS ON THIS SIDE. *** The blob store is shared across
            // every submission and is content-addressed, so without this a maintainer holding one
            // submission id could fetch any blob in the store whose hash they could obtain - and
            // hashes travel in manifests, which maintainers routinely see. Scoping the fetch to the
            // submission being reviewed keeps "may review submission 7" from meaning "may read
            // every file anyone has ever uploaded".
            Assert.False(ReviewAssetLocator.IsSubmittedBlobAllowed(
                ReviewAssetLocatorTests.Manifest(ReviewAssetLocatorTests.Hash('a')),
                ReviewAssetLocatorTests.Hash('b')));
        }

        [Fact]
        public void A_MALFORMED_hash_is_refused_before_it_reaches_the_store()
        {
            // BlobStorePaths re-validates too, and that belt-and-braces is deliberate. Refusing
            // here means the answer does not depend on the store having remembered.
            SubmissionManifest manifest = ReviewAssetLocatorTests.Manifest();

            manifest.Files.Add(new SubmissionFile { Path = "x.png", Sha256 = "../../etc/passwd" });

            Assert.False(ReviewAssetLocator.IsSubmittedBlobAllowed(manifest, "../../etc/passwd"));
        }

        [Fact]
        public void The_hash_comparison_is_case_sensitive()
        {
            // SubmissionPathRules.IsValidHash requires lowercase, and the store keys on the exact
            // string. Accepting an uppercase spelling here would allow a fetch the store would
            // then miss, which reads as "the file is gone" rather than "you asked wrongly".
            string hash = ReviewAssetLocatorTests.Hash('a');

            Assert.False(ReviewAssetLocator.IsSubmittedBlobAllowed(
                ReviewAssetLocatorTests.Manifest(hash), hash.ToUpperInvariant()));
        }

        [Fact]
        public void A_null_manifest_allows_NOTHING()
        {
            // A submission whose payload could not be loaded must not fall open. The maintainer sees
            // the findings explaining why instead.
            Assert.False(ReviewAssetLocator.IsSubmittedBlobAllowed(
                null, ReviewAssetLocatorTests.Hash('a')));
        }

        // -----------------------------------------------------------------------------------
        // The PUBLISHED side: asked for by the board's NAME, scoped to the published board's list.
        //
        // In the CRT 4.0 tree (2026-10-25) a request names a file as the board does, and only a name
        // the published board of this submission lists is answered - from the stored file of its
        // checksum ("Files/<ab>/<sha256>.<ext>"), so no name is ever turned into a place in the tree.
        // -----------------------------------------------------------------------------------

        // The published board's files by name, as PublishedBoardFiles reads them.
        private static IReadOnlyDictionary<string, BoardDocument.StoredFileRow> Listed(params (string Name, string Sha256)[] files) =>
            files.ToDictionary(
                file => file.Name,
                file => new BoardDocument.StoredFileRow { Name = file.Name, Sha256 = file.Sha256, Size = 5 },
                StringComparer.OrdinalIgnoreCase);

        // Stores `content` in the tree as the file of `name` would be, and returns its checksum.
        private string Store(string name, string content)
        {
            string sha = DataTreeBuilder.Sha(content);
            File.WriteAllText(DataTreeBuilder.Full(this.thisDataTree, DataTreeLayout.StoredFilePath(sha, name)), content);
            return sha;
        }

        [Fact]
        public void A_name_the_published_board_lists_is_found_at_its_stored_bytes()
        {
            string sha = this.Store("Schematics/board.png", "bytes");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Schematics/board.png",
                out string resolved,
                ReviewAssetLocatorTests.Listed(("Schematics/board.png", sha))));

            Assert.Equal("bytes", File.ReadAllText(resolved));
        }

        // ###########################################################################################
        // *** THE OLD PICTURE OF A CHANGED OR DELETED ROW (2026-09-26). *** The submission no longer
        // cites it - only the published board does - and it is the "before" side of the comparison.
        // ###########################################################################################
        [Fact]
        public void A_name_only_the_PUBLISHED_board_lists_is_found()
        {
            string sha = this.Store("old board.png", "old bytes");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.ManifestReferencing(["new board.png"]),
                "old board.png",
                out string resolved,
                ReviewAssetLocatorTests.Listed(("old board.png", sha))));

            Assert.Equal("old bytes", File.ReadAllText(resolved));
        }

        // A file lying in the board's folder under the name, but not in its list, is not the board's.
        [Fact]
        public void A_name_the_published_board_does_NOT_list_is_refused_whatever_lies_on_disk()
        {
            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "private-notes.png"), "not listed");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "private-notes.png",
                out _,
                ReviewAssetLocatorTests.Listed(("Schematics/board.png", this.Store("Schematics/board.png", "bytes")))));
        }

        // ###########################################################################################
        // *** A NAME NEVER BECOMES A PLACE. *** Even a list naming a traversal - the scope check then
        // passes - reaches only the stored file of its checksum, so the secret next to the tree,
        // a REAL file in a REAL folder, is never served. The anti-vacuity half: the stored bytes of
        // the same name ARE found.
        // ###########################################################################################
        [Fact]
        public void A_traversal_the_published_list_itself_names_reaches_only_the_store()
        {
            string secrets = Path.Combine(this.thisRoot, "secrets");
            Directory.CreateDirectory(secrets);
            File.WriteAllText(Path.Combine(secrets, "appsettings.Production.json"), "connection string");

            const string Escape = "../secrets/appsettings.Production.json";
            string sha = DataTreeBuilder.Sha("connection string");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree, ReviewAssetLocatorTests.Manifest(), Escape, out _, ReviewAssetLocatorTests.Listed((Escape, sha))));

            this.Store(Escape, "connection string");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree, ReviewAssetLocatorTests.Manifest(), Escape, out string resolved, ReviewAssetLocatorTests.Listed((Escape, sha))));

            Assert.StartsWith(Path.GetFullPath(Path.Combine(this.thisDataTree, "Files")), resolved, StringComparison.Ordinal);
        }

        [Fact]
        public void A_listed_name_without_a_usable_checksum_is_refused()
        {
            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Schematics/board.png",
                out _,
                ReviewAssetLocatorTests.Listed(("Schematics/board.png", "../../etc/passwd"))));
        }

        [Fact]
        public void A_listed_file_whose_bytes_are_not_in_the_tree_is_refused_WITHOUT_throwing()
        {
            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Schematics/board.png",
                out _,
                ReviewAssetLocatorTests.Listed(("Schematics/board.png", ReviewAssetLocatorTests.Hash('c')))));
        }

        [Fact]
        public void A_BLANK_data_tree_root_is_refused_rather_than_resolving_against_the_working_directory()
        {
            // A missing DataTreeRoot setting must not silently make the service's own working
            // directory the data tree - which is where its configuration and binaries live.
            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                string.Empty,
                ReviewAssetLocatorTests.Manifest(),
                "Schematics/board.png",
                out _,
                ReviewAssetLocatorTests.Listed(("Schematics/board.png", ReviewAssetLocatorTests.Hash('a')))));
        }

        // No manifest, or a published board that could not be read, locates nothing.
        [Fact]
        public void A_null_manifest_or_an_unread_published_board_locates_NOTHING()
        {
            string sha = this.Store("Schematics/board.png", "bytes");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree, null, "Schematics/board.png", out _, ReviewAssetLocatorTests.Listed(("Schematics/board.png", sha))));

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree, ReviewAssetLocatorTests.Manifest(), "Schematics/board.png", out _, null));
        }

        // -----------------------------------------------------------------------------------
        // What a fetched asset is served AS.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("board.png", "image/png")]
        [InlineData("board.PNG", "image/png")]
        [InlineData("photo.jpg", "image/jpeg")]
        [InlineData("photo.jpeg", "image/jpeg")]
        [InlineData("scan.gif", "image/gif")]
        [InlineData("scan.bmp", "image/bmp")]
        [InlineData("scan.webp", "image/webp")]
        public void An_IMAGE_is_served_as_its_own_type(string name, string expected)
        {
            Assert.Equal(expected, ReviewAssetLocator.ContentTypeFor(name));
        }

        [Theory]
        [InlineData("datasheet.pdf")]
        [InlineData("notes.txt")]
        [InlineData("board.xlsx")]
        [InlineData("script.html")]
        [InlineData("anything.exe")]
        [InlineData("")]
        public void ANYTHING_that_is_not_a_known_image_is_served_as_OCTET_STREAM(string name)
        {
            // *** NEVER text/html, AND NEVER A TYPE GUESSED FROM THE FILE NAME. *** These bytes
            // are contributor-supplied and are served from the server's own origin. A file served
            // as text/html would run script in that origin against a maintainer's session; an
            // allowlist of image types with octet-stream underneath means an attacker cannot
            // choose the type by choosing the extension.
            //
            // The Maintainer tab renders images. Everything else downloads, which is also the correct
            // behaviour for the datasheet a maintainer wants to open.
            Assert.Equal("application/octet-stream", ReviewAssetLocator.ContentTypeFor(name));
        }

        [Fact]
        public void A_DOUBLE_EXTENSION_is_read_from_the_LAST_one()
        {
            // "evil.png.html" is HTML. Reading the first extension would serve it as an image and
            // let the browser sniff it back to HTML; reading the last is what the filesystem and
            // every browser agree on.
            Assert.Equal("application/octet-stream", ReviewAssetLocator.ContentTypeFor("evil.png.html"));

            // ...and the reverse is a genuine image, so this is not simply refusing dots.
            Assert.Equal("image/png", ReviewAssetLocator.ContentTypeFor("evil.html.png"));
        }
    }
}
