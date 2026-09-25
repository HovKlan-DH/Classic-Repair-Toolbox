using CRT.Server.Handlers.Submissions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedSystemLister - the boards in the data tree a reviewer can be assigned to.
    //
    // A real temp folder, because the thing under test IS the folder walk: which folders count as
    // a board, and which are the shared-file folders that must never get a reviewer.
    // ###########################################################################################
    public sealed class PublishedSystemListerTests : IDisposable
    {
        private readonly string thisRoot;

        public PublishedSystemListerTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-lister-tests", Guid.NewGuid().ToString("N"));
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

        private void Board(string manufacturer, string hardware, string board, bool withWorkbook = true)
        {
            string folder = Path.Combine(this.thisRoot, manufacturer, hardware, board);
            Directory.CreateDirectory(folder);

            if (withWorkbook)
                File.WriteAllText(Path.Combine(folder, $"Data {hardware} {board} v2.0.0.xlsx"), "x");
        }

        [Fact]
        public void Every_board_folder_holding_a_workbook_is_listed_with_its_three_parts()
        {
            this.Board("Commodore", "C64", "250407");
            this.Board("Commodore", "C128", "310378");
            this.Board("Amstrad", "CPC", "464");

            IReadOnlyList<PublishedSystemLister.KnownSystem> systems = PublishedSystemLister.List(this.thisRoot);

            Assert.Equal(
                ["Amstrad/CPC/464", "Commodore/C128/310378", "Commodore/C64/250407"],
                systems.Select(system => system.SystemId));

            PublishedSystemLister.KnownSystem c64 = systems.Single(system => system.SystemId == "Commodore/C64/250407");
            Assert.Equal(("Commodore", "C64", "250407"), (c64.Manufacturer, c64.Hardware, c64.Board));
        }

        [Fact]
        public void The_SHARED_FILE_folders_are_never_systems()
        {
            // "Generic shared files" at the top and "Shared files" beside a manufacturer's boards
            // are not boards, and a reviewer assigned to one would review nothing. Both carry
            // sub-folders deep enough to look like a board to a naive walk.
            this.Board("Commodore", "C64", "250407");
            this.Board("Commodore", "Shared files", "Images");
            this.Board("Generic shared files", "Chips", "TTL");

            IReadOnlyList<PublishedSystemLister.KnownSystem> systems = PublishedSystemLister.List(this.thisRoot);

            Assert.Equal("Commodore/C64/250407", Assert.Single(systems).SystemId);
        }

        [Fact]
        public void A_folder_with_no_workbook_is_not_a_board()
        {
            this.Board("Commodore", "C64", "250407", withWorkbook: false);
            this.Board("Commodore", "C64", "250425");

            Assert.Equal("Commodore/C64/250425", Assert.Single(PublishedSystemLister.List(this.thisRoot)).SystemId);
        }

        [Fact]
        public void A_missing_or_blank_root_lists_nothing()
        {
            Assert.Empty(PublishedSystemLister.List(null));
            Assert.Empty(PublishedSystemLister.List(string.Empty));
            Assert.Empty(PublishedSystemLister.List(Path.Combine(this.thisRoot, "does-not-exist")));
        }
    }
}
