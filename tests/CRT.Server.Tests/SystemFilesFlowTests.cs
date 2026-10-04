using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // SystemFilesFlow - every file a system uses as BETA holds it, for the Systems screen's Files
    // view (owner request, 2026-10-03: "just list all files"). Against a real temp tree.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class SystemFilesFlowTests : IDisposable
    {
        private const string SystemId = "Commodore/C64/250407";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-system-files", Guid.NewGuid().ToString("N"));

        public SystemFilesFlowTests()
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

        private Task<IReadOnlyList<SystemFileEntry>?> BuildAsync(string systemId = SystemFilesFlowTests.SystemId) =>
            SystemFilesFlow.BuildAsync(this.thisRoot, systemId, new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance));

        // ###########################################################################################
        // The system's own folder - workbook, KiCad data, a file nothing cites - and the shared file
        // its board cites, all listed as they are. Another board's file and the shared file nobody
        // here cites are not the system's; a cited file BETA does not hold is not listed as if it were.
        // ###########################################################################################
        [Fact]
        public async Task It_lists_the_systems_own_folder_and_the_shared_files_it_cites_and_nothing_else()
        {
            DataTreeBuilder.Board(
                this.thisRoot,
                DataTreeBuilder.Workbook,
                "Commodore/C64/250407/manual.pdf",
                "Commodore/Shared files/74LS08.pdf",
                "Commodore/Shared files/missing.pdf");

            DataTreeBuilder.Files(
                this.thisRoot,
                "Commodore/C64/250407/manual.pdf",
                "Commodore/C64/250407/KiCad data/board.kicad_pcb",
                "Commodore/C64/250407/notes.txt",
                "Commodore/Shared files/74LS08.pdf",
                "Commodore/Shared files/7406.pdf",
                "Commodore/C128/310378/other.pdf");

            IReadOnlyList<SystemFileEntry>? files = await this.BuildAsync();

            Assert.NotNull(files);
            Assert.Equal(
                new[]
                {
                    "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx",
                    "Commodore/C64/250407/KiCad data/board.kicad_pcb",
                    "Commodore/C64/250407/manual.pdf",
                    "Commodore/C64/250407/notes.txt",
                    "Commodore/Shared files/74LS08.pdf"
                },
                files!.Select(file => file.Path));

            Assert.All(files!, file =>
            {
                Assert.Equal(SystemFileChange.Unchanged, file.Change);
                Assert.Equal(SystemFileSource.Beta, file.OpenFrom);

                // Each with its size as BETA holds it (owner request, 2026-10-04).
                Assert.Equal(new FileInfo(DataTreeBuilder.Full(this.thisRoot, file.Path)).Length, file.SizeBytes);
            });
        }

        // The stable source's files (owner request, 2026-10-04): the same listing, each opened from
        // the stable source (the tree it was read from) and sized from it.
        [Fact]
        public async Task The_stable_sources_files_open_from_the_stable_source()
        {
            DataTreeBuilder.Board(this.thisRoot, DataTreeBuilder.Workbook, "Commodore/C64/250407/manual.pdf");
            DataTreeBuilder.Files(this.thisRoot, "Commodore/C64/250407/manual.pdf");

            IReadOnlyList<SystemFileEntry>? files = await SystemFilesFlow.BuildAsync(
                this.thisRoot, "Commodore/C64/250407", new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                CancellationToken.None, SystemFileSource.Production);

            Assert.NotNull(files);
            Assert.NotEmpty(files!);
            Assert.All(files!, file =>
            {
                Assert.Equal(SystemFileSource.Production, file.OpenFrom);
                Assert.Equal(new FileInfo(DataTreeBuilder.Full(this.thisRoot, file.Path)).Length, file.SizeBytes);
            });
        }

        [Fact]
        public async Task A_system_BETA_holds_nothing_of_has_no_files_to_list()
        {
            Assert.Null(await this.BuildAsync("Commodore/VIC-20/250403"));
        }
    }
}
