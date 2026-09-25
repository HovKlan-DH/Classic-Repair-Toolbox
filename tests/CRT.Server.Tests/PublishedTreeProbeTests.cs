using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedTreeProbe - the server's real answers to "what is published at this path" and
    // "what is in this folder", which the submission rules ask (security review, 2026-09-25).
    // ###########################################################################################
    public sealed class PublishedTreeProbeTests : IDisposable
    {
        private readonly string thisRoot;
        private readonly string thisTree;

        public PublishedTreeProbeTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-probe-tests", Guid.NewGuid().ToString("N"));
            this.thisTree = Path.Combine(this.thisRoot, "Data");

            Directory.CreateDirectory(Path.Combine(this.thisTree, "Commodore", "C64", "250407"));
            File.WriteAllText(Path.Combine(this.thisTree, "Commodore", "C64", "250407", "Sheet1.png"), "scan");
            File.WriteAllText(Path.Combine(this.thisRoot, "secret.txt"), "outside the tree");
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // Harmless.
            }
        }

        [Fact]
        public void A_published_file_hashes_to_its_real_bytes()
        {
            PublishedTreeView tree = PublishedTreeProbe.For(this.thisTree)!;

            Assert.Equal(
                Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes("scan"))),
                tree.HashOf("Commodore/C64/250407/Sheet1.png"));
        }

        [Fact]
        public void A_case_variant_of_a_real_folder_is_found_on_disk()
        {
            PublishedTreeView tree = PublishedTreeProbe.For(this.thisTree)!;

            // On a case-insensitive filesystem the probe still reports the name AS IT EXISTS, so the
            // check works the same on the developer's Windows machine and on the Linux server.
            Assert.Equal("Commodore", tree.FindCaseVariant("commodore/C64/250407"));
            Assert.Null(tree.FindCaseVariant("Commodore/C64/250407/Sheet1.png"));
        }

        // A hostile path must not make the probe read - and so fingerprint - anything outside the
        // tree. It goes through SubmissionPathRules like every other path.
        [Fact]
        public void Nothing_outside_the_tree_is_hashed_or_listed()
        {
            PublishedTreeView tree = PublishedTreeProbe.For(this.thisTree)!;

            Assert.Null(tree.HashOf("../secret.txt"));
            Assert.Null(tree.FindCaseVariant("../secret.txt"));
        }

        // No usable root is "the tree could not be consulted" - which the rules answer by refusing
        // every foreign file, the safe direction for a misconfigured service.
        [Fact]
        public void No_usable_root_gives_no_view_at_all()
        {
            Assert.Null(PublishedTreeProbe.For(null));
            Assert.Null(PublishedTreeProbe.For(Path.Combine(this.thisRoot, "does-not-exist")));
        }
    }
}
