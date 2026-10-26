using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // BoardFilesFlow - every file a board uses as BETA holds it, for the Boards screen's Files
    // view (owner request, 2026-10-03: "just list all files"). Against a real temp tree - a CRT 4.0
    // one (2026-10-25): the board file and every name it lists whose bytes are stored.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class BoardFilesFlowTests : IDisposable
    {
        private const string BoardId = "Commodore/C64/250407";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-board-files", Guid.NewGuid().ToString("N"));

        public BoardFilesFlowTests()
        {
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

        private Task<IReadOnlyList<BoardFileEntry>?> BuildAsync(string boardId = BoardFilesFlowTests.BoardId, BoardFileSource source = BoardFileSource.Beta) =>
            BoardFilesFlow.BuildAsync(this.thisRoot, boardId, new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance), CancellationToken.None, source);

        // ###########################################################################################
        // The board file and every name it lists, by the board's names, each fetched from where its
        // bytes lie and sized from there. A listed name whose bytes the tree does not hold is not
        // listed as if it were; another board's files are not this board's.
        // ###########################################################################################
        [Fact]
        public async Task It_lists_the_board_file_and_every_stored_file_it_lists_by_name()
        {
            DataTreeBuilder.Board(this.thisRoot, BoardFilesFlowTests.BoardId,
                ("manual.pdf", "manual"),
                ("KiCad data/board.kicad_pcb", "(kicad_pcb)"),
                ("missing.pdf", "gone"));

            DataTreeBuilder.Board(this.thisRoot, "Commodore/C128/310378", ("other.pdf", "other"));

            File.Delete(DataTreeBuilder.Full(this.thisRoot, DataTreeBuilder.Stored("gone", "missing.pdf")));

            IReadOnlyList<BoardFileEntry>? files = await this.BuildAsync();

            Assert.NotNull(files);
            Assert.Equal(["KiCad data/board.kicad_pcb", "board.json", "manual.pdf"], files!.Select(file => file.Path));

            Assert.Equal("Commodore/C64/250407/board.json", files[1].PlaceInTree);
            Assert.Equal(DataTreeBuilder.Stored("manual", "manual.pdf"), files[2].PlaceInTree);

            Assert.All(files, file =>
            {
                Assert.Equal(BoardFileChange.Unchanged, file.Change);
                Assert.Equal(BoardFileSource.Beta, file.OpenFrom);

                // Each with its size as BETA holds it (owner request, 2026-10-04).
                Assert.Equal(new FileInfo(DataTreeBuilder.Full(this.thisRoot, file.PlaceInTree)).Length, file.SizeBytes);
            });
        }

        // The stable source's files (owner request, 2026-10-04): the same listing, each opened from
        // the stable source (the tree it was read from) and sized from it.
        [Fact]
        public async Task The_stable_sources_files_open_from_the_stable_source()
        {
            DataTreeBuilder.Board(this.thisRoot, BoardFilesFlowTests.BoardId, ("manual.pdf", "manual"));

            IReadOnlyList<BoardFileEntry>? files = await this.BuildAsync(source: BoardFileSource.Production);

            Assert.NotNull(files);
            Assert.NotEmpty(files!);
            Assert.All(files!, file =>
            {
                Assert.Equal(BoardFileSource.Production, file.OpenFrom);
                Assert.Equal(new FileInfo(DataTreeBuilder.Full(this.thisRoot, file.PlaceInTree)).Length, file.SizeBytes);
            });
        }

        [Fact]
        public async Task A_board_BETA_holds_nothing_of_has_no_files_to_list()
        {
            Assert.Null(await this.BuildAsync("Commodore/VIC-20/250403"));
        }
    }
}
