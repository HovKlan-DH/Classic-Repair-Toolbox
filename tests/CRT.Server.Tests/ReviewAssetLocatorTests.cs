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
    //   PUBLISHED - addressed by a caller-supplied PATH into the data tree. Guarded by
    //               SubmissionPathRules, the same containment the write paths use.
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
                SystemId = "Commodore/C64/250407",
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
        // The PUBLISHED side: addressed by path, contained to the system folder.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE PATH IS RELATIVE TO THE DATA TREE ROOT, NOT THE SYSTEM FOLDER (corrected
        // 2026-09-23). ***
        //
        // This test used to pass "Schematics/board.png" and expect it resolved under the system
        // folder. That was wrong about the real data: a board stores its references relative to
        // the DATA ROOT ("Commodore/C64/250407/Schematics/board.png"), which is how the desktop
        // app resolves them. The old base doubled the system segments, so every published image
        // answered 404 and the maintainer was told "No published file at this path" about files that
        // are in fact published.
        //
        // The expectation is corrected rather than the code bent to it - see the locator header
        // for why the system folder could never have been right.
        // ###########################################################################################
        [Fact]
        public void A_published_file_is_found_by_its_DATA_TREE_relative_path()
        {
            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407", "Schematics");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "board.png"), "bytes");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Commodore/C64/250407/Schematics/board.png",
                out string resolved));

            Assert.True(File.Exists(resolved));
        }

        // ###########################################################################################
        // *** THE CASE THE OLD SYSTEM-FOLDER BASE COULD NEVER HAVE SERVED. ***
        //
        // A shared component image is stored as "Commodore/Shared files/Component images/6526.png"
        // - outside the system folder by design, and cited by boards across a manufacturer. Under
        // the old base this was unreachable, so an entire legitimate category of board reference
        // could not be shown to a maintainer at all.
        // ###########################################################################################
        [Fact]
        public void A_SHARED_file_outside_the_system_folder_is_found()
        {
            const string Shared = "Commodore/Shared files/Component images/6526.png";

            string folder = Path.Combine(
                this.thisDataTree, "Commodore", "Shared files", "Component images");

            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "6526.png"), "bytes");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.ManifestReferencing([Shared]),
                Shared,
                out string resolved));

            Assert.True(File.Exists(resolved));
        }

        // ###########################################################################################
        // *** THE OLD PICTURE OF A CHANGED OR DELETED ROW (2026-09-26). *** The submission no
        // longer cites it - only the published board does - and it is the "before" side of the
        // comparison. It used to answer 404, so every removed image read "No published file at this
        // path" in the change summary.
        // ###########################################################################################
        [Fact]
        public void A_file_only_the_PUBLISHED_board_cites_is_found()
        {
            const string Old = "Commodore/C64/250407/old board.png";

            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "old board.png"), "bytes");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.ManifestReferencing(["Commodore/C64/250407/new board.png"]),
                Old,
                out string resolved,
                publishedBoardFiles: () => [Old]));

            Assert.True(File.Exists(resolved));
        }

        // Neither board cites it: still refused, whatever sits on disk beside the board.
        [Fact]
        public void A_file_NEITHER_board_cites_is_still_refused()
        {
            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "unrelated.png"), "bytes");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.ManifestReferencing(["Commodore/C64/250407/new board.png"]),
                "Commodore/C64/250407/unrelated.png",
                out _,
                publishedBoardFiles: () => ["Commodore/C64/250407/old board.png"]));
        }

        // The published board is read only when the submission does not cite the path itself - the
        // ordinary request (a file the submission names) never pays for a workbook read.
        [Fact]
        public void The_published_board_is_not_read_for_a_file_the_submission_cites()
        {
            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407", "Schematics");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "board.png"), "bytes");

            bool asked = false;

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Commodore/C64/250407/Schematics/board.png",
                out _,
                publishedBoardFiles: () =>
                {
                    asked = true;
                    return [];
                }));

            Assert.False(asked);
        }

        // ###########################################################################################
        // *** THE SCOPE LIMIT THAT REPLACED THE SYSTEM FOLDER, and it is STRICTER. ***
        //
        // The old base confined reads to the system own folder. The new one confines them to the
        // files the submission board actually NAMES - so a file sitting inside the system folder,
        // which the old rule would have served, is refused unless the board references it.
        // ###########################################################################################
        // ###########################################################################################
        // *** THE CONTAINMENT TEST THAT CANNOT PASS VACUOUSLY, and it is here because the obvious
        // version DID. ***
        //
        // The traversal tests below hand the locator a path the board does not reference, so since
        // 2026-09-23 the SCOPE check refuses them before SubmissionPathRules is ever reached -
        // meaning they would keep passing with containment deleted outright. Verified by deleting
        // it: all thirty tests still went green.
        //
        // This one closes that hole by having the BOARD ITSELF reference the traversal, which is
        // the real threat model - the manifest is contributor-supplied, so an attacker controls
        // the row values and can name "../../../secrets/..." as a schematic image. The scope check
        // then passes and only containment stands between the request and the file.
        //
        // The secret is a REAL file in a REAL folder, so the refusal cannot come from File.Exists.
        // Delete SubmissionPathRules from the locator and this goes red.
        // ###########################################################################################
        [Fact]
        public void A_traversal_the_BOARD_ITSELF_REFERENCES_is_still_refused()
        {
            string secrets = Path.Combine(this.thisRoot, "secrets");
            Directory.CreateDirectory(secrets);
            File.WriteAllText(Path.Combine(secrets, "appsettings.Production.json"), "connection string");

            // The data tree is <root>/data, so one level up reaches thisRoot.
            const string Escape = "../secrets/appsettings.Production.json";

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.ManifestReferencing([Escape]),
                Escape,
                out _));
        }

        [Fact]
        public void A_file_the_board_does_NOT_reference_is_refused_even_inside_the_system_folder()
        {
            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "private-notes.png"), "not referenced");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Commodore/C64/250407/private-notes.png",
                out _));
        }

        [Fact]
        public void A_TRAVERSAL_ONTO_A_REAL_FILE_OUTSIDE_THE_TREE_IS_REFUSED()
        {
            // *** THE TEST THIS FILE EXISTS FOR, and the one shape that has already been written
            // vacuously twice in this project. *** The secret is placed in a REAL sibling folder
            // holding a REAL file, so the refusal cannot come from File.Exists returning false.
            // Delete SubmissionPathRules from the locator and this goes red.
            string secrets = Path.Combine(this.thisRoot, "secrets");
            Directory.CreateDirectory(secrets);
            File.WriteAllText(Path.Combine(secrets, "appsettings.Production.json"), "connection string");

            // The system folder is <tree>/Commodore/C64/250407, so four levels up reaches thisRoot.
            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "../../../../secrets/appsettings.Production.json",
                out _));
        }

        [Fact]
        public void The_same_shape_INSIDE_the_tree_IS_found_so_the_refusal_is_containment()
        {
            // The anti-vacuity half of the test above. Same nesting depth, same file name, the
            // only difference being that it stays inside the tree AND is referenced by the board -
            // so a locator that simply finds nothing would fail HERE.
            const string Inside = "Commodore/C64/250407/a/b/c/secrets/appsettings.Production.json";

            string folder = Path.Combine(
                this.thisDataTree, "Commodore", "C64", "250407", "a", "b", "c", "secrets");

            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "appsettings.Production.json"), "harmless");

            Assert.True(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.ManifestReferencing([Inside]),
                Inside,
                out string resolved));

            Assert.Equal("harmless", File.ReadAllText(resolved));
        }

        [Fact]
        public void A_file_in_ANOTHER_systems_folder_is_refused()
        {
            // Containment is to THIS system, not merely to the data tree. A maintainer opening a
            // C64 submission has no business reading an Amstrad board's files through it, and
            // "inside the tree" would permit exactly that.
            string other = Path.Combine(this.thisDataTree, "Amstrad", "CPC464", "Z70200");
            Directory.CreateDirectory(other);
            File.WriteAllText(Path.Combine(other, "board.png"), "other system");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "../../../Amstrad/CPC464/Z70200/board.png",
                out _));
        }

        [Fact]
        public void An_ABSOLUTE_path_is_refused()
        {
            string absolute = Path.Combine(this.thisRoot, "secrets.txt");
            File.WriteAllText(absolute, "secret");

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                absolute,
                out _));
        }

        [Fact]
        public void A_file_that_simply_does_not_exist_is_refused_WITHOUT_throwing()
        {
            // An ordinary miss: the published board references an image that is not on disk. The
            // maintainer gets "not found" and the rest of the screen still draws.
            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree,
                ReviewAssetLocatorTests.Manifest(),
                "Schematics/missing.png",
                out _));
        }

        [Fact]
        public void An_UNSAFE_IDENTITY_in_the_manifest_cannot_relocate_the_system_folder()
        {
            // The identity is untrusted too - it arrives inside the submission. A manufacturer of
            // ".." would move the system folder up the tree and make every subsequent containment
            // check contain the WRONG folder.
            string secrets = Path.Combine(this.thisRoot, "secrets");
            Directory.CreateDirectory(secrets);
            File.WriteAllText(Path.Combine(secrets, "key.txt"), "secret");

            var manifest = new SubmissionManifest
            {
                Manufacturer = "..",
                Hardware = "..",
                Board = ".."
            };

            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree, manifest, "secrets/key.txt", out _));
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
                out _));
        }

        [Fact]
        public void A_null_manifest_locates_NOTHING()
        {
            Assert.False(ReviewAssetLocator.TryLocatePublishedFile(
                this.thisDataTree, null, "Schematics/board.png", out _));
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
            // The maintainer app renders images. Everything else downloads, which is also the correct
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
