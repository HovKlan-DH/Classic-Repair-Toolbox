using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The board table's search box (owner request, 2026-10-02: "how about the same search/filter as we
// have in the "Workbooks" tab, where it highlights what I search for") - BoardTableSearch, which
// decides the rows; the editor only paints them. Its grammar is the Workbooks tab's own
// WorklogSearchQuery: words ANDed, "quotes" for a phrase, -minus to leave out, any case.
//
// Each test is one case the project owner agreed before the code was written; the case number
// is in its comment.
// ###########################################################################################
public sealed class BoardTableSearchTests
{
    private static ComponentImageEntry Image(string label, string name, string file, string note = "") =>
        new() { BoardLabel = label, Name = name, File = file, Note = note };

    // The Component images sheet from the owner's screenshot, cut down.
    private static BoardData Board() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "RAM", TechnicalNameOrValue = "4164" },
            new ComponentEntry { BoardLabel = "U9", FriendlyName = "ROM", TechnicalNameOrValue = "x" }
        ],
        ComponentImages =
        [
            Image("CR13", "Pinout", "a/cr13.png"),
            Image("CR17", "Pinout", "a/cr17.png"),
            Image("R307", "Pinout (secondary)", "Generic shared files/Component images/resistor_3k3_5.png"),
            Image("R307", "Pinout (secondary)", "Generic shared files/Component images/resistor.png", "Resistors are non-polarized"),
            Image("U8", "Clock", "a/u8.png")
        ]
    };

    private static BoardTableDocument Document(BoardData? published = null) => BoardTableDocument.Create(published ?? Board(), Board());

    private static BoardTableSheet Images(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!;

    private static BoardTableSheet Components(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int Column(BoardTableSheet sheet, string column) => sheet.Columns.ToList().IndexOf(column);

    private static string Label(BoardTableRow row) => row.Cells[Column(row.Sheet, BoardWorkbookSchema.ColBoardLabel)].Text;

    // The labels of the rows a search shows on one sheet, in order (R307 twice when both show).
    private static List<string> Shown(BoardTableSearch search, BoardTableSheet sheet) =>
        sheet.Rows.Where(search.Shows).Select(Label).ToList();

    // Case 10: a word finds the rows with it in any cell, whatever its case.
    [Fact]
    public void A_word_finds_the_rows_with_it_in_any_cell_in_any_case()
    {
        BoardTableDocument document = Document();

        BoardTableSearch search = BoardTableSearch.For(document, "PiNoUt");

        Assert.True(search.IsActive);
        Assert.Equal(["CR13", "CR17", "R307", "R307"], Shown(search, Images(document)));
    }

    // Case 11: two words show only rows having both - and they may be in different cells.
    [Fact]
    public void Two_words_find_only_rows_having_both_even_in_different_cells()
    {
        BoardTableDocument document = Document();

        // "cr" is in the label, "pinout" in the name - never the same cell.
        Assert.Equal(["CR13", "CR17"], Shown(BoardTableSearch.For(document, "cr pinout"), Images(document)));
    }

    // Case 12: a quoted phrase is found only as that exact run - and within one cell.
    [Fact]
    public void A_quoted_phrase_finds_only_that_exact_phrase()
    {
        BoardTableDocument document = Document();

        Assert.Equal(["R307", "R307"], Shown(BoardTableSearch.For(document, "\"pinout (secondary)\""), Images(document)));

        // The label and the name side by side are two cells, not one phrase.
        Assert.Empty(Shown(BoardTableSearch.For(document, "\"cr13 pinout\""), Images(document)));
    }

    // Case 13: a word with a minus in front leaves out the rows containing it.
    [Fact]
    public void A_minus_word_leaves_out_rows_containing_it()
    {
        BoardTableDocument document = Document();

        Assert.Equal(["CR13", "CR17"], Shown(BoardTableSearch.For(document, "pinout -secondary"), Images(document)));
    }

    // Case 14: every cell is searched, numbers included - unlike the Workbooks tab, which skips
    // numbers - but not the +/~/-/! sign in front of a row, nor the column names.
    [Fact]
    public void Every_cell_is_searched_numbers_included_but_not_the_row_sign_or_the_column_names()
    {
        // U8 changed, so it carries the "~" sign.
        BoardData published = Board();
        published.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "DRAM", TechnicalNameOrValue = "4164" };
        BoardTableDocument document = Document(published);
        Assert.Equal("~", Components(document).Rows[0].Marker);

        Assert.Equal(["U8"], Shown(BoardTableSearch.For(document, "4164"), Components(document)));
        Assert.Equal(["CR13"], Shown(BoardTableSearch.For(document, "13"), Images(document)));

        Assert.Empty(Shown(BoardTableSearch.For(document, "~"), Components(document)));
        Assert.Empty(Shown(BoardTableSearch.For(document, "Friendly"), Components(document)));
    }

    // Case 18: a red (deleted) row is searched by the values it shows - the published ones.
    [Fact]
    public void A_red_row_is_searched_by_the_values_it_shows()
    {
        BoardTableDocument document = Document();
        BoardTableSheet images = Images(document);
        images.DeleteRow(images.Rows.Single(row => Label(row) == "U8"));
        Assert.True(images.Rows.Single(row => Label(row) == "U8").IsDeleted);

        Assert.Equal(["U8"], Shown(BoardTableSearch.For(document, "clock"), images));
    }

    // Case 19: a blank row just inserted is always shown, whatever the search.
    [Fact]
    public void A_blank_row_just_inserted_is_always_shown()
    {
        BoardTableDocument document = Document();
        BoardTableSheet images = Images(document);
        BoardTableRow blank = images.InsertRow(images.Rows[0]);

        BoardTableSearch search = BoardTableSearch.For(document, "clock");

        Assert.True(search.Shows(blank));
        Assert.Equal(["", "U8"], Shown(search, images));
    }

    // Case C (agreed: "your assumption is correct"): a row edited so it no longer matches stays
    // until the search is changed - it does not vanish under the contributor's typing. A row
    // inserted after the search was typed stays too.
    [Fact]
    public void A_row_edited_so_it_no_longer_matches_stays_until_the_search_changes()
    {
        BoardTableDocument document = Document();
        BoardTableSheet images = Images(document);
        BoardTableSearch search = BoardTableSearch.For(document, "pinout");

        BoardTableRow cr13 = images.Rows[0];
        cr13.Cells[Column(images, BoardWorkbookSchema.ColName)].Text = "Diode test";
        images.Refresh();
        BoardTableRow inserted = images.InsertRow(cr13);
        inserted.Cells[Column(images, BoardWorkbookSchema.ColBoardLabel)].Text = "D1";
        images.Refresh();

        Assert.True(search.Shows(cr13));
        Assert.True(search.Shows(inserted));

        // The search typed again: now it is judged on what the rows say.
        BoardTableSearch again = BoardTableSearch.For(document, "pinout");
        Assert.False(again.Shows(cr13));
        Assert.False(again.Shows(inserted));
    }

    // Case 15 (the rule half): an empty search is no search - every row shows and nothing is found
    // to mark; a search marks exactly the runs its words were found in.
    [Fact]
    public void An_empty_search_shows_every_row_and_marks_nothing()
    {
        BoardTableDocument document = Document();

        BoardTableSearch empty = BoardTableSearch.For(document, "   ");

        Assert.False(empty.IsActive);
        Assert.Equal(5, Shown(empty, Images(document)).Count);
        Assert.Empty(empty.HitsIn("Pinout (secondary)"));
    }

    [Fact]
    public void A_search_marks_exactly_the_runs_its_words_were_found_in()
    {
        BoardTableSearch search = BoardTableSearch.For(Document(), "pin second -clock");

        Assert.Equal(
            [(0, 3), (8, 6)],
            search.HitsIn("Pinout (secondary)").Select(hit => (hit.Start, hit.Length)));
        Assert.Empty(search.HitsIn("Clock"));
    }

    // Case 17 (the rule half): like the pills, a search covers every sheet - the tabs of sheets with
    // no match are hidden, and the table opens on the first sheet with one.
    //
    // *** A SEARCH MATCHING NOTHING ANYWHERE keeps only the sheet on screen (agreed case 3,
    // 2026-10-03). *** It hid no tab until then, each sheet empty, which read as a search of the one
    // sheet on screen; this test pinned that, and was changed with it.
    [Fact]
    public void Sheets_with_no_match_lose_their_tab_and_a_search_matching_nothing_keeps_only_the_sheet_on_screen()
    {
        BoardTableDocument document = Document();

        BoardTableSearch pinout = BoardTableSearch.For(document, "pinout");

        Assert.True(document.HasRowsShownBy(BoardTableRowKinds.None, pinout));
        Assert.Equal([BoardWorkbookSchema.SheetComponentImages], document.SheetsShown(BoardTableRowKinds.None, pinout, current: null).Select(sheet => sheet.Name));
        Assert.Equal(BoardWorkbookSchema.SheetComponentImages, document.SheetToShow(BoardWorkbookSchema.SheetComponents, BoardTableRowKinds.None, pinout).Name);

        BoardTableSearch nothing = BoardTableSearch.For(document, "zzz");
        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        Assert.Equal([components], document.SheetsShown(BoardTableRowKinds.None, nothing, current: components));
        Assert.Empty(document.SheetsShown(BoardTableRowKinds.None, nothing, current: null));

        // ...and the table stays on the sheet it was on, rather than jumping to the first.
        Assert.Equal(BoardWorkbookSchema.SheetComponents, document.SheetToShow(BoardWorkbookSchema.SheetComponents, BoardTableRowKinds.None, nothing).Name);

        // And no search keeps every tab, as before.
        Assert.Equal(document.Sheets.Count, document.SheetsShown(BoardTableRowKinds.None, BoardTableSearch.None, current: null).Count);
    }

    // Cases 3 and 5 (2026-10-03), the line above the table: said only when a search finds nothing in
    // ANY sheet - alone, or together with the picked pills. Picked pills alone showing nothing keep
    // the old rule (every tab, no line) - that was not asked for.
    [Fact]
    public void The_nothing_found_line_is_said_only_when_a_search_finds_nothing_in_any_sheet()
    {
        BoardData published = Board();
        published.ComponentImages[0] = Image("CR13", "Pinout", "a/old.png");
        BoardTableDocument document = Document(published);

        Assert.Equal("Nothing in any sheet matches \"zzz\"", BoardTableSearch.For(document, "zzz").NothingFoundLine(document, BoardTableRowKinds.None));
        Assert.Equal("Nothing in any sheet matches \"zzz\"", BoardTableSearch.For(document, "  zzz  ").NothingFoundLine(document, BoardTableRowKinds.None));

        // RAM is only on Components, where nothing is modified.
        Assert.Equal(
            "Nothing in any sheet matches both \"ram\" and the counts picked",
            BoardTableSearch.For(document, "ram").NothingFoundLine(document, BoardTableRowKinds.Modified));

        // Found somewhere: no line - with or without pills.
        Assert.Null(BoardTableSearch.For(document, "pinout").NothingFoundLine(document, BoardTableRowKinds.None));
        Assert.Null(BoardTableSearch.For(document, "pinout").NothingFoundLine(document, BoardTableRowKinds.Modified));

        // No search: no line, even with pills picked that show nothing.
        Assert.Null(BoardTableSearch.None.NothingFoundLine(document, BoardTableRowKinds.None));
        Assert.Null(BoardTableSearch.None.NothingFoundLine(document, BoardTableRowKinds.Added));
        Assert.Equal(document.Sheets, document.SheetsShown(BoardTableRowKinds.Added, BoardTableSearch.None, current: null));
    }

    // Case 16 (the rule half): a search and the picked pills together - a sheet keeps its tab only
    // with a row BOTH show.
    [Fact]
    public void A_sheet_keeps_its_tab_only_with_a_row_both_the_search_and_the_pills_show()
    {
        BoardData published = Board();
        published.ComponentImages[0] = Image("CR13", "Pinout", "a/old.png");
        BoardTableDocument document = Document(published);

        BoardTableSearch pinout = BoardTableSearch.For(document, "pinout");
        BoardTableSearch ram = BoardTableSearch.For(document, "ram");

        Assert.Equal([BoardWorkbookSchema.SheetComponentImages], document.SheetsShown(BoardTableRowKinds.Modified, pinout, current: null).Select(sheet => sheet.Name));

        // RAM is only on the Components sheet, where nothing is modified: nothing shows anywhere.
        Assert.False(document.HasRowsShownBy(BoardTableRowKinds.Modified, ram));
    }
}
