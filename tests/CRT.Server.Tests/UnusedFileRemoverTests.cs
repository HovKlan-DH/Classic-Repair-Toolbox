using CRT.Server.Handlers.Submissions;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // UnusedFileRemover - the one place anything in a data tree is deleted because nothing uses
    // it (2026-09-25). It only ever removes FEWER files than it is given: the tree is read again,
    // a file used after all is kept, an unreadable tree removes nothing, and nothing goes through
    // a link. Each of those is a test here.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class UnusedFileRemoverTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-unused", Guid.NewGuid().ToString("N"));

        private const string Cited = "Commodore/C64/250407/sheet1.png";
        private const string Orphan = "Commodore/C64/250407/Scope baseline/old.png";
        private const string OtherOrphan = "Generic shared files/Component images/7408.jpg";

        public UnusedFileRemoverTests()
        {
            DataTreeBuilder.Master(this.thisRoot, DataTreeBuilder.Workbook);
            DataTreeBuilder.Board(this.thisRoot, DataTreeBuilder.Workbook, UnusedFileRemoverTests.Cited);
            DataTreeBuilder.Files(this.thisRoot, UnusedFileRemoverTests.Cited, UnusedFileRemoverTests.Orphan, UnusedFileRemoverTests.OtherOrphan);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        private bool Exists(string relative) => File.Exists(DataTreeBuilder.Full(this.thisRoot, relative));

        [Fact]
        public void An_unused_file_is_removed_with_the_folder_it_leaves_empty()
        {
            UnusedFileRemoval removal = UnusedFileRemover.Remove(this.thisRoot, [UnusedFileRemoverTests.Orphan], NullLogger.Instance);

            Assert.Equal([UnusedFileRemoverTests.Orphan], removal.Removed);
            Assert.False(File.Exists(Path.Combine(this.thisRoot, "Commodore", "C64", "250407", "Scope baseline", "old.png")));
            Assert.False(Directory.Exists(Path.Combine(this.thisRoot, "Commodore", "C64", "250407", "Scope baseline")));

            // The board folder is not empty, so it - and everything above it - stays.
            Assert.True(this.Exists(UnusedFileRemoverTests.Cited));
        }

        // Only what it is given: another orphan in the tree is left for the administrator's list.
        [Fact]
        public void Only_the_files_asked_for_are_removed()
        {
            UnusedFileRemover.Remove(this.thisRoot, [UnusedFileRemoverTests.Orphan], NullLogger.Instance);

            Assert.True(this.Exists(UnusedFileRemoverTests.OtherOrphan));
        }

        // Between a list being shown and the removal, something may start citing a file again.
        [Fact]
        public void A_file_that_is_used_after_all_is_kept()
        {
            UnusedFileRemoval removal = UnusedFileRemover.Remove(
                this.thisRoot, [UnusedFileRemoverTests.Cited, UnusedFileRemoverTests.Orphan], NullLogger.Instance);

            Assert.Equal([UnusedFileRemoverTests.Orphan], removal.Removed);
            Assert.Equal([UnusedFileRemoverTests.Cited], removal.Kept);
            Assert.True(this.Exists(UnusedFileRemoverTests.Cited));
        }

        [Fact]
        public void An_unreadable_tree_removes_nothing_and_says_why()
        {
            DataTreeBuilder.Master(this.thisRoot, DataTreeBuilder.Workbook, "Amstrad/CPC/464/Data CPC 464 v2.0.0.xlsx");

            UnusedFileRemoval removal = UnusedFileRemover.Remove(this.thisRoot, [UnusedFileRemoverTests.Orphan], NullLogger.Instance);

            Assert.Empty(removal.Removed);
            Assert.NotNull(removal.NotDoneBecause);
            Assert.True(this.Exists(UnusedFileRemoverTests.Orphan));
        }

        // A folder in the tree that is really a link elsewhere must not let a removal reach through
        // it. Asked through the delegate, so no link has to be created (Windows needs elevation).
        [Fact]
        public void Nothing_is_removed_through_a_link()
        {
            UnusedFileRemoval removal = UnusedFileRemover.Remove(
                this.thisRoot,
                [UnusedFileRemoverTests.Orphan],
                NullLogger.Instance,
                isLink: path => path.EndsWith("Scope baseline", StringComparison.Ordinal));

            Assert.Empty(removal.Removed);
            Assert.True(this.Exists(UnusedFileRemoverTests.Orphan));
        }

        [Fact]
        public void A_path_outside_the_tree_or_a_hidden_file_is_never_removed()
        {
            string outside = Path.Combine(Path.GetDirectoryName(this.thisRoot)!, "outside-" + Guid.NewGuid().ToString("N") + ".txt");
            File.WriteAllText(outside, "x");
            DataTreeBuilder.Files(this.thisRoot, "Commodore/C64/250407/.hidden");

            try
            {
                UnusedFileRemoval removal = UnusedFileRemover.Remove(
                    this.thisRoot, ["../" + Path.GetFileName(outside), "Commodore/C64/250407/.hidden"], NullLogger.Instance);

                Assert.Empty(removal.Removed);
                Assert.True(File.Exists(outside));
                Assert.True(this.Exists("Commodore/C64/250407/.hidden"));
            }
            finally
            {
                File.Delete(outside);
            }
        }

        [Fact]
        public void An_empty_list_does_not_read_the_tree()
        {
            UnusedFileRemoval removal = UnusedFileRemover.Remove(
                Path.Combine(this.thisRoot, "no-such-folder"), [], NullLogger.Instance);

            Assert.Same(UnusedFileRemoval.None, removal);
        }
    }
}
