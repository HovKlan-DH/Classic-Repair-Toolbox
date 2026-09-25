using System.Collections.Generic;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionSharedFiles - whether a submission changes a shared file, which is what
    // sends it to the administrator rather than to the system's maintainers (Phase 6 roles).
    //
    // The property that matters most is the NEGATIVE one: a board that merely USES a shared image
    // unchanged must not count, or nearly every submission would bypass its maintainers.
    // ###########################################################################################
    public sealed class SubmissionSharedFilesTests
    {
        private const string Hash = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        private const string OtherHash = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";

        private static SubmissionManifest Manifest(params (string Path, string Sha256)[] files)
        {
            var manifest = new SubmissionManifest
            {
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407"
            };

            foreach ((string path, string sha) in files)
                manifest.Files.Add(new SubmissionFile { Path = path, Sha256 = sha, SizeBytes = 1 });

            return manifest;
        }

        private static PublishedTreeView Tree(params (string Path, string Sha256)[] published)
        {
            var hashes = new Dictionary<string, string>(System.StringComparer.Ordinal);

            foreach ((string path, string sha) in published)
                hashes[path] = sha;

            return new PublishedTreeView(path => hashes.TryGetValue(path, out string? sha) ? sha : null, _ => null);
        }

        [Fact]
        public void A_shared_file_cited_UNCHANGED_does_not_count()
        {
            // The ordinary shape of a board: it uses a shared component image as published.
            SubmissionManifest manifest = SubmissionSharedFilesTests.Manifest(
                ("Commodore/Shared files/74LS08.png", SubmissionSharedFilesTests.Hash));

            Assert.False(SubmissionSharedFiles.TouchesSharedFiles(
                manifest, SubmissionSharedFilesTests.Tree(("Commodore/Shared files/74LS08.png", SubmissionSharedFilesTests.Hash))));
        }

        [Fact]
        public void A_shared_file_with_DIFFERENT_bytes_counts()
        {
            SubmissionManifest manifest = SubmissionSharedFilesTests.Manifest(
                ("Commodore/Shared files/74LS08.png", SubmissionSharedFilesTests.OtherHash));

            Assert.True(SubmissionSharedFiles.TouchesSharedFiles(
                manifest, SubmissionSharedFilesTests.Tree(("Commodore/Shared files/74LS08.png", SubmissionSharedFilesTests.Hash))));
        }

        [Fact]
        public void A_shared_file_that_is_NEW_counts()
        {
            SubmissionManifest manifest = SubmissionSharedFilesTests.Manifest(
                ("Generic shared files/new-chip.png", SubmissionSharedFilesTests.Hash));

            Assert.True(SubmissionSharedFiles.TouchesSharedFiles(manifest, SubmissionSharedFilesTests.Tree()));
        }

        [Fact]
        public void The_boards_OWN_files_never_count_however_much_they_change()
        {
            SubmissionManifest manifest = SubmissionSharedFilesTests.Manifest(
                ("Commodore/C64/250407/main.png", SubmissionSharedFilesTests.OtherHash),
                ("Commodore/C64/250407/new.png", SubmissionSharedFilesTests.Hash));

            Assert.False(SubmissionSharedFiles.TouchesSharedFiles(manifest, SubmissionSharedFilesTests.Tree()));
        }

        [Fact]
        public void Without_a_view_of_the_tree_a_shared_file_counts_as_changed()
        {
            // "Could not check" must fall on the administrator-only side, never the maintainer side.
            SubmissionManifest manifest = SubmissionSharedFilesTests.Manifest(
                ("Commodore/Shared files/74LS08.png", SubmissionSharedFilesTests.Hash));

            Assert.True(SubmissionSharedFiles.TouchesSharedFiles(manifest, null));
        }

        [Fact]
        public void A_submission_with_no_files_touches_nothing()
        {
            Assert.False(SubmissionSharedFiles.TouchesSharedFiles(SubmissionSharedFilesTests.Manifest(), null));
        }
    }
}
