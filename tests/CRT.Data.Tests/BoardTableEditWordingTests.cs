using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What the table says after a paste, a fill down or a replace (CRT 4.0, 2026-10-25) - the one
// place the Drafts tab and the Maintainer tab take these sentences from.
//
// The rule worth holding: whatever of the clipboard did NOT go in is always said, with why - a
// paste that quietly drops two columns is believed to have worked - and every change ends with
// how to take it back.
// ###########################################################################################
public sealed class BoardTableEditWordingTests
{
    private static IReadOnlyList<BoardTableRow> Rows(int count)
    {
        var board = new BoardData();

        for (int i = 1; i <= count; i++)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = $"U{i}", FriendlyName = "x", TechnicalNameOrValue = "y" });
        }

        return BoardTableDocument.Create(null, board).FindSheet(BoardWorkbookSchema.SheetComponents)!.Rows.ToList();
    }

    private static BoardTablePasteResult Result(
        int rows,
        int columns = 2,
        int changed = 4,
        int added = 0,
        int rowsLeftOut = 0,
        int columnsLeftOut = 0,
        bool headings = false) =>
        new(Rows(rows), 0, columns, changed, added, rowsLeftOut, columnsLeftOut, headings);

    [Fact]
    public void One_cell_pasted_into_one_cell_says_nothing()
    {
        Assert.Equal(string.Empty, BoardTableEditWording.Pasted(Result(rows: 1, columns: 1, changed: 1)));
        Assert.Equal(string.Empty, BoardTableEditWording.Pasted(Result(rows: 1, columns: 1, changed: 0)));
        Assert.Equal(string.Empty, BoardTableEditWording.Pasted(BoardTablePasteResult.Nothing));
    }

    [Fact]
    public void A_block_says_how_many_rows_it_went_into_and_how_to_take_it_back()
    {
        Assert.Equal("Pasted into 2 rows. Ctrl+Z undoes it.", BoardTableEditWording.Pasted(Result(rows: 2)));
    }

    [Fact]
    public void Everything_that_did_not_go_in_is_said_with_why()
    {
        Assert.Equal(
            "Pasted into 3 rows. 2 of them were added at the end of the sheet. " +
            "The first row copied was this sheet's column headings, so it was left out. " +
            "1 column did not fit and was left out. " +
            "4 rows were left out: while the table shows only some rows, a paste goes into those and adds none. " +
            "Ctrl+Z undoes it.",
            BoardTableEditWording.Pasted(Result(rows: 3, added: 2, rowsLeftOut: 4, columnsLeftOut: 1, headings: true)));
    }

    // A single cell with something left out is not silent either.
    [Fact]
    public void One_cell_with_columns_left_out_says_so()
    {
        Assert.Equal(
            "Pasted into 1 row. 2 columns did not fit and were left out. Ctrl+Z undoes it.",
            BoardTableEditWording.Pasted(Result(rows: 1, columns: 1, changed: 1, columnsLeftOut: 2)));
    }

    [Fact]
    public void A_paste_that_changed_nothing_says_so_and_offers_no_undo()
    {
        string said = BoardTableEditWording.Pasted(Result(rows: 2, changed: 0));

        Assert.Equal("The cells already held what was pasted - nothing changed.", said);
        Assert.DoesNotContain("Ctrl+Z", said);
    }

    [Fact]
    public void Fill_and_paste_into_cells_count_cells_and_say_how_to_take_it_back()
    {
        Assert.Equal("Pasted into 5 cells. Ctrl+Z undoes it.", BoardTableEditWording.PastedIntoCells(5));
        Assert.Equal("Filled 1 cell with the value at the top. Ctrl+Z undoes it.", BoardTableEditWording.Filled(1));
        Assert.Equal("Filled 7 cells with the value at the top. Ctrl+Z undoes it.", BoardTableEditWording.Filled(7));
        Assert.DoesNotContain("Ctrl+Z", BoardTableEditWording.Filled(0));
        Assert.DoesNotContain("Ctrl+Z", BoardTableEditWording.PastedIntoCells(0));
    }

    [Fact]
    public void A_replace_says_what_how_often_where_and_how_to_take_it_back()
    {
        Assert.Equal(
            "Replaced [6510] 3 times in 2 cells. Ctrl+Z undoes it.",
            BoardTableEditWording.Replaced(new BoardTableReplaceResult(3, 2, 1), "6510", everySheet: false));

        Assert.Equal(
            "Replaced [6510] 1 time in 1 cell. Ctrl+Z undoes it.",
            BoardTableEditWording.Replaced(new BoardTableReplaceResult(1, 1, 1), "6510", everySheet: true));

        Assert.Equal(
            "Replaced [6510] 4 times in 4 cells on 3 sheets. Ctrl+Z undoes it.",
            BoardTableEditWording.Replaced(new BoardTableReplaceResult(4, 4, 3), "6510", everySheet: true));
    }

    // The Maintainer tab shows the same table - a submission's, BETA's - so a red row's sentence
    // names nobody's draft (code review, 2026-10-25).
    [Fact]
    public void What_is_said_of_red_rows_fits_every_table_and_says_how_to_bring_one_back()
    {
        string[] sentences =
        [
            BoardTableEditWording.PasteOnDeletedRow,
            BoardTableEditWording.PasteOnDeletedRows,
            BoardTableEditWording.FillOnDeletedRows,
        ];

        Assert.All(sentences, sentence =>
        {
            Assert.DoesNotContain("draft", sentence, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("red row", sentence, StringComparison.Ordinal);
            Assert.Contains("Ctrl+Z", sentence, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_replace_finding_nothing_says_where_it_looked()
    {
        Assert.Equal("No row shown on this sheet contains [VIC].", BoardTableEditWording.Replaced(new BoardTableReplaceResult(0, 0, 0), "VIC", everySheet: false));
        Assert.Equal("No row shown on any sheet contains [VIC].", BoardTableEditWording.Replaced(new BoardTableReplaceResult(0, 0, 0), "VIC", everySheet: true));
    }
}
