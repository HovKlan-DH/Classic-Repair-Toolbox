using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where "Replace all" replaces in a cell's text - the one rule its writing, its count and the table's preview go by
// (BoardTableReplaceText; owner request, 2026-11-07: the replacement drawn "so I can see directly what the new text
// will look like", and "a counter badge for how many places it will replace").
// ###########################################################################################
public sealed class BoardTableReplaceTextTests
{
    // Left to right, never overlapping, capitals ignored - exactly where String.Replace replaces.
    [Theory]
    [InlineData("part U1", "part", 1)]
    [InlineData("PART part Part", "part", 3)]
    [InlineData("aaa", "aa", 1)]
    [InlineData("aaaa", "aa", 2)]
    [InlineData("nothing", "part", 0)]
    [InlineData("", "part", 0)]
    [InlineData("part", "", 0)]
    public void The_occurrences_are_the_ones_a_replace_replaces(string text, string find, int expected)
    {
        Assert.Equal(expected, BoardTableReplaceText.Occurrences(text, find));

        if (find.Length > 0)
        {
            string replaced = text.Replace(find, "#", StringComparison.OrdinalIgnoreCase);
            Assert.Equal(expected, replaced.Count(ch => ch == '#'));
        }
    }

    [Fact]
    public void The_text_is_cut_where_a_replace_cuts_it_covering_all_of_it()
    {
        Assert.Equal(
            [("PART", true), (" U1 and ", false), ("part", true)],
            BoardTableReplaceText.Split("PART U1 and part", "part"));

        Assert.Equal([("aa", true), ("a", false)], BoardTableReplaceText.Split("aaa", "aa"));
        Assert.Equal([("no find", false)], BoardTableReplaceText.Split("no find", "xyz"));
        Assert.Equal([("whole", false)], BoardTableReplaceText.Split("whole", ""));
        Assert.Empty(BoardTableReplaceText.Split("", "x"));

        // Joined again, nothing is lost or doubled.
        IReadOnlyList<(string Text, bool IsFound)> parts = BoardTableReplaceText.Split("x part y PART z", "part");
        Assert.Equal("x part y PART z", string.Concat(parts.Select(part => part.Text)));
    }

    // ###########################################################################################
    // The count on "Replace all" is what ReplaceAll then replaces - the rows shown, the columns shown, red rows never,
    // and a place the replacement would leave as it was (the same text in other capitals) none.
    // ###########################################################################################
    [Fact]
    public void The_count_is_what_Replace_all_then_replaces()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", FriendlyName = "part", TechnicalNameOrValue = "part part" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U2", FriendlyName = "VIC", TechnicalNameOrValue = "PART" });

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

        Assert.Equal(4, document.CountReplacements("part", "chip", [sheet], _ => true));

        // "part" for "part": U1's cells stay as they are, U2's "PART" becomes "part".
        Assert.Equal(1, document.CountReplacements("part", "part", [sheet], _ => true));
        Assert.Equal(1, document.CountReplacements("part", "Part", [sheet], row => row.Cells[0].Text == "U2"));
        Assert.Equal(0, document.CountReplacements("", "chip", [sheet], _ => true));

        int counted = document.CountReplacements("part", "chip", [sheet], _ => true);
        BoardTableReplaceResult result = document.ReplaceAll("part", "chip", [sheet], _ => true);

        Assert.Equal(counted, result.Occurrences);
        Assert.Equal(0, document.CountReplacements("part", "chip", [sheet], _ => true));
    }

    // ###########################################################################################
    // "Find" is the text AS TYPED (owner request, 2026-11-07: "only search for the string as-is and do not 'AND'
    // multiple words or exclude with minus") - one phrase, a minus and quotes part of it.
    // ###########################################################################################
    [Fact]
    public void Find_searches_for_the_text_as_typed()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU VIC" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U2", FriendlyName = "VIC CPU" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U3", FriendlyName = "-VIC \"x\"" });

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        List<BoardTableRow> rows = document.FindSheet(BoardWorkbookSchema.SheetComponents)!.Rows.ToList();

        string[] Shown(BoardTableSearch search) => rows.Where(search.Shows).Select(row => row.Cells[0].Text).ToArray();

        Assert.Equal(["U1"], Shown(BoardTableSearch.Literal(document, "cpu vic")));
        Assert.Equal(["U3"], Shown(BoardTableSearch.Literal(document, "-vic")));
        Assert.Equal(["U3"], Shown(BoardTableSearch.Literal(document, "\"x\"")));
        Assert.False(BoardTableSearch.Literal(document, "").IsActive);

        // The search box's grammar finds both rows for the two words, and leaves out with a minus.
        Assert.Equal(["U1", "U2"], Shown(BoardTableSearch.For(document, "cpu vic")));
        Assert.Equal(["U1", "U2"], Shown(BoardTableSearch.For(document, "cpu -\"x\"")));
    }

    [Fact]
    public void A_literal_query_is_one_term_as_typed()
    {
        WorklogSearchQuery query = WorklogSearchQuery.Literal("  two words ");

        WorklogSearchTerm term = Assert.Single(query.Terms);
        Assert.Equal("  two words ", term.Text);
        Assert.False(term.IsExcluded);
        Assert.True(WorklogSearchQuery.Literal(null).IsEmpty);
        Assert.True(WorklogSearchQuery.Literal(string.Empty).IsEmpty);
    }
}
