using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedFileHashes - the SHA-256 of each published file a submission also names,
    // which the Maintainer tab uses to drop byte-identical pairs from the image comparison.
    //
    // *** THIS EXISTS BECAUSE OF A REPORTED SCREEN (2026-09-23). *** A submission that changed
    // one component's short description was shown to the maintainer as "1178 images to compare",
    // every one of them identical on both sides, because nothing told the Maintainer tab the
    // published hashes. ReviewImageComparison.Plan already knew how to drop a matching pair; it
    // was never handed anything to match against.
    //
    // Real files in a real temp folder throughout - the same seam SubmissionFlowTests uses. No
    // network, no display, no spawned process.
    // ###########################################################################################
    public sealed class PublishedFileHashesTests : IDisposable
    {
        private readonly string thisRoot;
        private readonly string thisDataTree;

        public PublishedFileHashesTests()
        {
            this.thisRoot = Path.Combine(
                Path.GetTempPath(),
                "crt-hashes-" + Guid.NewGuid().ToString("N"));

            this.thisDataTree = Path.Combine(this.thisRoot, "app-data-BETA");
            Directory.CreateDirectory(this.thisDataTree);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A stranded temp folder is untidy, not a failed test.
            }
        }

        private string Write(string relativePath, string content)
        {
            string full = Path.Combine(
                this.thisDataTree,
                relativePath.Replace('/', Path.DirectorySeparatorChar));

            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllText(full, content);

            return full;
        }

        private static string Sha256Of(string content)
        {
            return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(content)));
        }

        private static SubmissionManifest ManifestNaming(params string[] paths)
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407"
            };

            foreach (string path in paths)
            {
                manifest.Files.Add(new SubmissionFile
                {
                    Path = path,
                    Sha256 = new string('a', 64),
                    SizeBytes = 1
                });
            }

            return manifest;
        }

        // ---------------------------------------------------------------- the hash itself

        [Fact]
        public async Task The_hash_is_the_lowercase_hex_SHA256_of_the_bytes_on_disk()
        {
            // Lowercase, because the client's manifest and dataChecksums.json both use lowercase
            // and ReviewImageComparison compares ORDINALLY - an uppercase hash here would compare
            // unequal to an identical file and read as a replacement.
            const string Path1 = "Commodore/C64/250407/Board Layout 250407 NTSC.png";
            this.Write(Path1, "board bytes");

            IReadOnlyDictionary<string, string> hashes = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree, [Path1], CancellationToken.None);

            Assert.Equal(PublishedFileHashesTests.Sha256Of("board bytes"), hashes[Path1]);
            Assert.Equal(hashes[Path1], hashes[Path1].ToLowerInvariant());
        }

        [Fact]
        public async Task The_result_is_keyed_by_the_PATH_AS_GIVEN_not_by_the_resolved_location()
        {
            // The Maintainer tab looks the hash up by the path in its published-file list, which is
            // the data-root-relative string with forward slashes. A key rewritten to the OS
            // separator, or to an absolute path, would match nothing and every pair would be
            // shown as replaced - silently.
            const string Path1 = "Commodore/C64/250407/Scope baseline/U1_1_PAL.png";
            this.Write(Path1, "trace");

            IReadOnlyDictionary<string, string> hashes = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree, [Path1], CancellationToken.None);

            Assert.True(hashes.ContainsKey(Path1));
        }

        // ---------------------------------------------------------------- what is left out

        [Fact]
        public async Task A_file_that_does_not_exist_is_simply_absent_and_nothing_throws()
        {
            // Absent rather than an error: the planner then treats the file as replaced and shows
            // it, which is the safe direction. A throw here would take down the whole detail
            // response over one missing scan.
            this.Write("Commodore/C64/250407/real.png", "real");

            IReadOnlyDictionary<string, string> hashes = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree,
                ["Commodore/C64/250407/real.png", "Commodore/C64/250407/gone.png"],
                CancellationToken.None);

            Assert.Single(hashes);
            Assert.False(hashes.ContainsKey("Commodore/C64/250407/gone.png"));
        }

        [Fact]
        public async Task A_traversal_onto_a_REAL_file_outside_the_tree_is_refused()
        {
            // The secret is a REAL file in a REAL sibling folder, so the refusal cannot come from
            // the file merely being absent. Delete SubmissionPathRules from the hasher and this
            // goes red.
            string secrets = Path.Combine(this.thisRoot, "secrets");
            Directory.CreateDirectory(secrets);
            File.WriteAllText(Path.Combine(secrets, "appsettings.Production.json"), "connection string");

            IReadOnlyDictionary<string, string> hashes = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree,
                ["../secrets/appsettings.Production.json"],
                CancellationToken.None);

            Assert.Empty(hashes);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public async Task No_data_tree_root_yields_no_hashes_rather_than_hashing_relative_to_nowhere(string? root)
        {
            IReadOnlyDictionary<string, string> hashes = await PublishedFileHashes.ComputeAsync(
                root, ["Commodore/C64/250407/x.png"], CancellationToken.None);

            Assert.Empty(hashes);
        }

        // ---------------------------------------------------------------- the cache

        [Fact]
        public async Task A_file_rewritten_with_new_bytes_gets_a_NEW_hash()
        {
            // The anti-staleness half of the cache. A publish replaces bytes in place, and a cache
            // that kept serving the old hash would tell the maintainer the new scan is identical to
            // the old one - hiding the change, which is the one direction this must never be wrong
            // in. The write time is set explicitly so the test does not depend on timer resolution.
            const string Path1 = "Commodore/C64/250407/scan.png";
            string full = this.Write(Path1, "first version");
            File.SetLastWriteTimeUtc(full, new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc));

            IReadOnlyDictionary<string, string> before = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree, [Path1], CancellationToken.None);

            File.WriteAllText(full, "second version - different length too");
            File.SetLastWriteTimeUtc(full, new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc));

            IReadOnlyDictionary<string, string> after = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree, [Path1], CancellationToken.None);

            Assert.NotEqual(before[Path1], after[Path1]);
            Assert.Equal(PublishedFileHashesTests.Sha256Of("second version - different length too"), after[Path1]);
        }

        [Fact]
        public async Task Hashing_the_same_unchanged_file_twice_gives_the_same_answer()
        {
            // The cache must be invisible: whatever it does internally, the second call answers
            // exactly what the first did.
            const string Path1 = "Commodore/C64/250407/stable.png";
            this.Write(Path1, "stable");

            IReadOnlyDictionary<string, string> first = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree, [Path1], CancellationToken.None);

            IReadOnlyDictionary<string, string> second = await PublishedFileHashes.ComputeAsync(
                this.thisDataTree, [Path1], CancellationToken.None);

            Assert.Equal(first[Path1], second[Path1]);
        }

        // -------------------------------------------------------- one file, for the tree view

        // ###########################################################################################
        // *** "WHICH PATHS ARE HASHED" IS NO LONGER A SELECTION (security review, 2026-09-25). ***
        // Only paths on BOTH the old board's list and the submission's used to be hashed, so a
        // submitted file OVERWRITING something the old board never cited was reported as "added".
        // The review endpoint now hashes every submitted path; what is left to pin is the
        // single-file hash PublishedTreeProbe answers "is this published unchanged?" with.
        // ###########################################################################################
        [Fact]
        public void A_single_published_file_hashes_to_the_same_value_as_the_batch()
        {
            const string Path1 = "Commodore/C128/310378/Scope baseline/notes.txt";
            this.Write(Path1, "published text");

            Assert.Equal(PublishedFileHashesTests.Sha256Of("published text"), PublishedFileHashes.TryHash(this.thisDataTree, Path1));
        }

        // ###########################################################################################
        // The single-file hash is taken on the calling thread now, not by blocking on the async one
        // (code review, 2026-09-25) - and it must see a rewrite exactly as the batch does, or the
        // tree would answer "unchanged" about a replaced file.
        // ###########################################################################################
        [Fact]
        public void A_single_file_rewritten_with_new_bytes_gets_a_NEW_hash()
        {
            const string Path1 = "Commodore/Shared files/74LS08.png";
            string full = this.Write(Path1, "first");
            File.SetLastWriteTimeUtc(full, new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc));

            Assert.Equal(PublishedFileHashesTests.Sha256Of("first"), PublishedFileHashes.TryHash(this.thisDataTree, Path1));

            File.WriteAllText(full, "second, and longer");
            File.SetLastWriteTimeUtc(full, new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc));

            Assert.Equal(PublishedFileHashesTests.Sha256Of("second, and longer"), PublishedFileHashes.TryHash(this.thisDataTree, Path1));
        }

        [Fact]
        public void A_path_with_nothing_published_at_it_hashes_to_null()
        {
            Assert.Null(PublishedFileHashes.TryHash(this.thisDataTree, "Commodore/C64/250407/absent.png"));
        }

        // The containment rule applies to a question as much as to a write: a hostile path must not
        // be able to make the probe read - and so fingerprint - a file outside the tree.
        [Theory]
        [InlineData("../outside.txt")]
        [InlineData("/etc/passwd")]
        [InlineData("")]
        public void A_path_that_leaves_the_tree_is_never_hashed(string path)
        {
            File.WriteAllText(Path.Combine(this.thisRoot, "outside.txt"), "secret");

            Assert.Null(PublishedFileHashes.TryHash(this.thisDataTree, path));
        }

        [Fact]
        public void No_data_tree_means_no_hash()
        {
            Assert.Null(PublishedFileHashes.TryHash(null, "Commodore/C64/250407/a.png"));
        }
    }
}
