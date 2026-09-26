using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers BoardTableFileCells - which table cells name a file, and which file each side of the
    // hover card shows (owner request, 2026-09-26: an image cell shows the picture, a changed one
    // the published and the new picture side by side, a PDF a link to open it). The Drafts tab and
    // the maintainer application both paint these answers, so they are decided - and tested - once.
    // ###########################################################################################
    public sealed class BoardTableFileCellsTests
    {
        // Exactly four columns hold file paths, checked against EVERY column of every sheet - so a
        // column added to the schema later cannot quietly start or stop showing a preview.
        [Fact]
        public void Exactly_the_four_file_columns_are_file_columns()
        {
            List<string> fileColumns = BoardWorkbookSchema.AllSheets
                .SelectMany(sheet => sheet.ColumnOrder.Select(column => (sheet.SheetName, column)))
                .Where(pair => BoardTableFileCells.IsFileColumn(pair.SheetName, pair.column))
                .Select(pair => $"{pair.SheetName} / {pair.column}")
                .ToList();

            Assert.Equal(
                new[]
                {
                    "Board schematics / Schematic image file",
                    "Component images / File",
                    "Component local files / File",
                    "Board local files / File",
                },
                fileColumns);
        }

        [Theory]
        [InlineData(null, "File")]
        [InlineData("Component links", "File")]
        [InlineData("Component images", "Name")]
        [InlineData("Component images", "file")]
        public void Other_columns_are_not_file_columns(string? sheet, string column)
        {
            Assert.False(BoardTableFileCells.IsFileColumn(sheet, column));
        }

        private static ComponentImageEntry Image(string pin, string file) =>
            new() { BoardLabel = "U8", Region = "PAL", Pin = pin, Name = "Pinout", File = file };

        private static BoardTableSheet ImagesOf(BoardData? published, BoardData draft) =>
            BoardTableDocument.Create(published, draft).FindSheet(BoardWorkbookSchema.SheetComponentImages)!;

        private static BoardTableCell FileCellOf(BoardTableRow row) =>
            row.Cells[row.Sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFile)];

        // A row the published board does not have: only the new file.
        [Fact]
        public void An_added_row_names_only_its_new_file()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                new BoardData(),
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/new.png")] });

            BoardTableFileCell file = BoardTableFileCells.Of(BoardTableFileCellsTests.FileCellOf(sheet.Rows.Single()))!;

            Assert.Null(file.PublishedPath);
            Assert.Equal("Commodore/C64/250407/new.png", file.CurrentPath);
            Assert.Equal("New file", BoardTableFileCells.Headline(file, sameContent: null));
        }

        // The case asked for by name: a changed picture shows both - the published one and the new.
        [Fact]
        public void A_row_changed_to_another_file_names_both()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/old.png")] },
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/new.png")] });

            BoardTableCell cell = BoardTableFileCellsTests.FileCellOf(sheet.Rows.Single());
            BoardTableFileCell file = BoardTableFileCells.Of(cell)!;

            Assert.Equal(BoardTableCellState.Modified, cell.State);
            Assert.Equal("Commodore/C64/250407/old.png", file.PublishedPath);
            Assert.Equal("Commodore/C64/250407/new.png", file.CurrentPath);
            Assert.False(file.IsSamePath);
            Assert.Equal("Changed to another file", BoardTableFileCells.Headline(file, sameContent: null));
        }

        // ###########################################################################################
        // *** THE SAME PATH IS STILL A COMPARISON. *** A picture replaced under its own name leaves
        // the cell unchanged - not orange - so only the bytes can tell. Both sides are named, and the
        // headline waits for the bytes.
        // ###########################################################################################
        [Fact]
        public void An_unchanged_file_cell_names_the_same_path_on_both_sides()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/u8.png")] },
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/u8.png")] });

            BoardTableFileCell file = BoardTableFileCells.Of(BoardTableFileCellsTests.FileCellOf(sheet.Rows.Single()))!;

            Assert.True(file.IsSamePath);
            Assert.Null(BoardTableFileCells.Headline(file, sameContent: null));
            Assert.Equal("Unchanged", BoardTableFileCells.Headline(file, sameContent: true));
            Assert.Equal("Replaced - same file name, new content", BoardTableFileCells.Headline(file, sameContent: false));
        }

        // A deleted row's ghost shows what was published, and nothing new.
        [Fact]
        public void A_deleted_row_names_only_the_published_file()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/gone.png")] },
                new BoardData());

            BoardTableRow ghost = Assert.Single(sheet.Rows, row => row.IsDeleted);
            BoardTableFileCell file = BoardTableFileCells.Of(BoardTableFileCellsTests.FileCellOf(ghost))!;

            Assert.Equal("Commodore/C64/250407/gone.png", file.PublishedPath);
            Assert.Null(file.CurrentPath);
            Assert.Equal("Removed", BoardTableFileCells.Headline(file, sameContent: null));
        }

        // Emptied in this version: the published file, removed.
        [Fact]
        public void A_file_cell_emptied_names_only_the_published_file()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Commodore/C64/250407/u8.png")] },
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", string.Empty)] });

            BoardTableFileCell file = BoardTableFileCells.Of(BoardTableFileCellsTests.FileCellOf(sheet.Rows.Single()))!;

            Assert.Equal("Commodore/C64/250407/u8.png", file.PublishedPath);
            Assert.Null(file.CurrentPath);
        }

        // Nothing to show: a blank file cell with nothing published, and any cell that is not a file.
        [Fact]
        public void A_blank_file_cell_and_other_columns_name_no_file()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                null,
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", string.Empty)] });

            BoardTableRow row = sheet.Rows.Single();

            Assert.Null(BoardTableFileCells.Of(BoardTableFileCellsTests.FileCellOf(row)));
            Assert.Null(BoardTableFileCells.Of(row.Cells[row.Sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColName)]));
            Assert.Null(BoardTableFileCells.Of(null));
        }

        // With nothing published at all (a new system), every file is new - no published side.
        [Fact]
        public void With_nothing_published_a_file_has_no_published_side()
        {
            BoardTableSheet sheet = BoardTableFileCellsTests.ImagesOf(
                null,
                new BoardData { ComponentImages = [BoardTableFileCellsTests.Image("1", "Test/HW/Board/a.png")] });

            BoardTableFileCell file = BoardTableFileCells.Of(BoardTableFileCellsTests.FileCellOf(sheet.Rows.Single()))!;

            Assert.Null(file.PublishedPath);
            Assert.Equal("Test/HW/Board/a.png", file.CurrentPath);
        }

        // ---- which files are drawn ------------------------------------------------------------

        [Theory]
        [InlineData("a.png")]
        [InlineData("Commodore/C64/250407/Board.JPG")]
        [InlineData("x.jpeg")]
        [InlineData("x.gif")]
        [InlineData("x.bmp")]
        [InlineData("x.webp")]
        public void Pictures_are_drawn(string path)
        {
            Assert.True(ImageFileTypes.IsDisplayable(path));
        }

        // A PDF or a web page is opened, never drawn; SVG is script-capable XML, not a picture.
        [Theory]
        [InlineData("manual.pdf")]
        [InlineData("notes.txt")]
        [InlineData("page.html")]
        [InlineData("drawing.svg")]
        [InlineData("noextension")]
        [InlineData("")]
        [InlineData(null)]
        public void Everything_else_is_not_drawn(string? path)
        {
            Assert.False(ImageFileTypes.IsDisplayable(path));
        }
    }
}
