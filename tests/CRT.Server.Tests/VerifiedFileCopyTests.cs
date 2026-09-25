using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers VerifiedFileCopy - the one verified copy a publish, a production promotion and a blob
    // store import all make (code review, 2026-09-25: the import used to carry its own copy of the
    // loop). What must hold whichever caller it is: only bytes that hash to what was expected ever
    // reach the destination, and nothing - not even a temporary file - is left behind otherwise.
    // ###########################################################################################
    public sealed class VerifiedFileCopyTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-verified-copy", Guid.NewGuid().ToString("N"));

        public VerifiedFileCopyTests() => Directory.CreateDirectory(this.thisRoot);

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless; failing a test over cleanup is not.
            }
        }

        private string Source(string text)
        {
            string path = Path.Combine(this.thisRoot, "source.bin");
            File.WriteAllText(path, text);
            return path;
        }

        private string Destination => Path.Combine(this.thisRoot, "tree", "board", "a.png");

        private static string HashOf(string text) => Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(text)));

        // Everything in the destination's folder - so a temporary file left behind shows up.
        private string[] FolderContents() =>
            Directory.Exists(Path.GetDirectoryName(this.Destination))
                ? Directory.GetFiles(Path.GetDirectoryName(this.Destination)!).Select(Path.GetFileName).Select(name => name!).ToArray()
                : [];

        [Fact]
        public async Task A_copy_that_hashes_as_expected_replaces_the_destination_and_leaves_nothing_else()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(this.Destination)!);
            File.WriteAllText(this.Destination, "the old image");

            VerifiedCopyResult result = await VerifiedFileCopy.CopyAsync(
                this.Source("the new image"), VerifiedFileCopyTests.HashOf("the new image"), this.Destination);

            Assert.Equal(VerifiedCopyResult.Copied, result);
            Assert.Equal("the new image", File.ReadAllText(this.Destination));
            Assert.Equal(["a.png"], this.FolderContents());
        }

        [Fact]
        public async Task A_copy_that_does_not_hash_as_expected_leaves_the_destination_as_it_was()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(this.Destination)!);
            File.WriteAllText(this.Destination, "the old image");

            VerifiedCopyResult result = await VerifiedFileCopy.CopyAsync(
                this.Source("changed since it was checked"), VerifiedFileCopyTests.HashOf("what was checked"), this.Destination);

            Assert.Equal(VerifiedCopyResult.HashMismatch, result);
            Assert.Equal("the old image", File.ReadAllText(this.Destination));
            Assert.Equal(["a.png"], this.FolderContents());
        }

        // The import's rule: past the cap nothing is written, not even a partial temporary file.
        [Fact]
        public async Task A_source_past_the_size_cap_writes_nothing()
        {
            VerifiedCopyResult result = await VerifiedFileCopy.CopyAsync(
                this.Source("0123456789"), VerifiedFileCopyTests.HashOf("0123456789"), this.Destination, maximumBytes: 9);

            Assert.Equal(VerifiedCopyResult.TooLarge, result);
            Assert.False(File.Exists(this.Destination));
            Assert.Empty(this.FolderContents());
        }

        // The cap is a MAXIMUM: a file exactly that size is allowed, as an upload of it would be.
        [Fact]
        public async Task A_source_exactly_at_the_size_cap_is_copied()
        {
            VerifiedCopyResult result = await VerifiedFileCopy.CopyAsync(
                this.Source("0123456789"), VerifiedFileCopyTests.HashOf("0123456789"), this.Destination, maximumBytes: 10);

            Assert.Equal(VerifiedCopyResult.Copied, result);
        }

        [Fact]
        public async Task A_missing_source_is_reported_and_creates_nothing()
        {
            VerifiedCopyResult result = await VerifiedFileCopy.CopyAsync(
                Path.Combine(this.thisRoot, "gone.bin"), new string('a', 64), this.Destination);

            Assert.Equal(VerifiedCopyResult.Missing, result);
            Assert.False(File.Exists(this.Destination));
        }
    }
}
