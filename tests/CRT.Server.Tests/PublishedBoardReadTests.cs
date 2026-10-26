using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishedBoardRead - a published board's file read ONCE for an approval (code review,
    // 2026-10-26: the same multi-megabyte board file was parsed for its list three times, its renames
    // once and its calibrations once per schematic, all under the lock every other publish waits on).
    // Each answer must be the one the separate reads gave - checked against them, file for file.
    // ###########################################################################################
    public sealed class PublishedBoardReadTests : IDisposable
    {
        private const string BoardId = "Commodore/C64/250407";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-published-read", Guid.NewGuid().ToString("N"));

        public PublishedBoardReadTests()
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

        private string BoardFile => DataTreeBuilder.Full(this.thisRoot, DataTreeLayout.BoardFilePath(PublishedBoardReadTests.BoardId));

        [Fact]
        public void A_board_read_once_answers_as_the_separate_reads_did()
        {
            var board = new BoardData
            {
                Schematics =
                [
                    new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" },
                    new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Images/bottom.png" },
                ],
            };

            KiCadCalibrationEntry calibration = new() { SchematicName = "TOP", CadName = "board.kicad_pcb", OffsetX = 1.5, OffsetY = 2, ScaleX = 1, ScaleY = 1 };
            SchematicRenameRecord[] renames =
            [
                new("Sheet 1", "Top", "Images/top.png", "2026-October-1"),
                new(" ", "Nameless", "Images/x.png", "2026-October-2"),
            ];

            Directory.CreateDirectory(Path.GetDirectoryName(this.BoardFile)!);
            BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
                new BoardContent(board, [calibration], renames),
                name => name,
                [new BoardDocument.StoredFileRow { Name = "Images/top.png", Sha256 = DataTreeBuilder.Sha("top"), Size = 3 }]));

            PublishedBoardRead read = PublishedBoardRead.Of(this.thisRoot, PublishedBoardReadTests.BoardId);

            Assert.True(read.Exists);
            Assert.NotNull(read.Document);
            Assert.Equal(DataTreeBuilder.Sha("top"), read.Files!["images/TOP.png"].Sha256);
            Assert.Equal(PublishedBoardFiles.ByName(this.thisRoot, PublishedBoardReadTests.BoardId)!.Keys, read.Files.Keys);

            Assert.Equal(BoardComponentHighlightStorage.LoadSchematicRenames(this.BoardFile), read.Renames);
            Assert.Equal("Sheet 1", Assert.Single(read.Renames).From);

            string[] names = ["Top", "Bottom", " "];
            KiCadCalibrationEntry found = Assert.Single(read.CalibrationsFor(names));
            Assert.Equal(
                BoardComponentHighlightStorage.LoadKiCadCalibrations(this.BoardFile, names).Select(PublishedBoardReadTests.Values),
                read.CalibrationsFor(names).Select(PublishedBoardReadTests.Values));
            Assert.Equal(("Top", 1.5), (found.SchematicName, found.OffsetX));
        }

        private static (string, string, double, double, double, double, bool, bool) Values(KiCadCalibrationEntry entry) =>
            (entry.SchematicName, entry.CadName, entry.OffsetX, entry.OffsetY, entry.ScaleX, entry.ScaleY, entry.MirrorX, entry.MirrorY);

        // A board the tree does not hold has no files (a new board: every file an addition); one whose
        // file cannot be read is not known - never an empty list - and has no renames or calibrations.
        [Fact]
        public void A_board_the_tree_does_not_hold_has_no_files_and_an_unreadable_one_is_not_known()
        {
            PublishedBoardRead absent = PublishedBoardRead.Of(this.thisRoot, PublishedBoardReadTests.BoardId);

            Assert.False(absent.Exists);
            Assert.NotNull(absent.Files);
            Assert.Empty(absent.Files!);

            Directory.CreateDirectory(Path.GetDirectoryName(this.BoardFile)!);
            File.WriteAllText(this.BoardFile, "{ not a board file");

            PublishedBoardRead unreadable = PublishedBoardRead.Of(this.thisRoot, PublishedBoardReadTests.BoardId);

            Assert.True(unreadable.Exists);
            Assert.Null(unreadable.Document);
            Assert.Null(unreadable.Files);
            Assert.Empty(unreadable.Renames);
            Assert.Empty(unreadable.CalibrationsFor(["Top"]));
        }
    }
}
