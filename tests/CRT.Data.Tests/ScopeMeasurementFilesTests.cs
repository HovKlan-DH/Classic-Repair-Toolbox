using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ScopeMeasurementFiles - a board's notes on how its scope baselines were measured: its board files of
// the category "Scope measurements", as the live trees keep them (owner decision, 2026-10-26).
public sealed class ScopeMeasurementFilesTests
{
    [Fact]
    public void The_notes_are_the_boards_files_of_the_scope_measurements_category_in_order()
    {
        var board = new BoardData
        {
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Best-practises", Name = "Best-Practises", File = "Board local files/bp.txt" },
                new BoardLocalFileEntry { Category = "Scope measurements", Name = "Introduction", File = "Scope baseline/intro.txt" },
                new BoardLocalFileEntry { Category = " scope MEASUREMENTS ", Name = "References", File = "Scope baseline/refs.txt" },
                new BoardLocalFileEntry { Category = "Scope measurements", Name = "Half a row", File = " " },
            ],
        };

        Assert.Equal(["Introduction", "References"], ScopeMeasurementFiles.Of(board).Select(row => row.Name));
    }

    // A note can be of one region since 2026-11-02, as every board file: shown only in that region
    // (code review, 2026-11-02: linked under the other region's baselines).
    [Fact]
    public void A_note_of_one_region_is_listed_only_in_that_region()
    {
        var board = new BoardData
        {
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Scope measurements", Name = "Every", File = "a.txt" },
                new BoardLocalFileEntry { Category = "Scope measurements", Name = "NTSC's", File = "b.txt", Region = "NTSC" },
                new BoardLocalFileEntry { Category = "Scope measurements", Name = "PAL's", File = "c.txt", Region = "PAL" },
            ],
        };

        Assert.Equal(["Every", "PAL's"], ScopeMeasurementFiles.Of(board, "PAL").Select(row => row.Name));
        Assert.Equal(["Every", "NTSC's"], ScopeMeasurementFiles.Of(board, "ntsc").Select(row => row.Name));
    }

    [Fact]
    public void A_board_without_them_or_no_board_has_none()
    {
        Assert.Empty(ScopeMeasurementFiles.Of(new BoardData()));
        Assert.Empty(ScopeMeasurementFiles.Of(null));
    }
}
