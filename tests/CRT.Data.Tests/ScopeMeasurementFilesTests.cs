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

    [Fact]
    public void A_board_without_them_or_no_board_has_none()
    {
        Assert.Empty(ScopeMeasurementFiles.Of(new BoardData()));
        Assert.Empty(ScopeMeasurementFiles.Of(null));
    }
}
