using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// What the table says about the checks' problems, apart from the cells (2026-10-02) -
// BoardTableProblemWording.
public sealed class BoardTableProblemWordingTests
{
    private static BoardDataProblem Problem(BoardProblemLevel level, string message) =>
        new(level, "code", string.Empty, message, Sheet: null, Index: -1, Column: null);

    [Fact]
    public void Nothing_outside_the_sheets_says_nothing()
    {
        Assert.Null(BoardTableProblemWording.OutsideSheets([]));
    }

    // Errors first, at most three, then how many more - each line saying error or warning.
    [Fact]
    public void Problems_outside_the_sheets_are_listed_worst_first_and_cut_short()
    {
        string text = BoardTableProblemWording.OutsideSheets(
        [
            Problem(BoardProblemLevel.Warning, "w1"),
            Problem(BoardProblemLevel.Error, "e1"),
            Problem(BoardProblemLevel.Warning, "w2"),
            Problem(BoardProblemLevel.Error, "e2"),
            Problem(BoardProblemLevel.Warning, "w3")
        ])!;

        string[] lines = text.Split(Environment.NewLine);

        Assert.Contains("\"Mark components manually\"", lines[0], StringComparison.Ordinal);
        Assert.Equal(["Error: e1", "Error: e2", "Warning: w1"], lines[1..4]);
        Assert.Equal("... and 2 more.", lines[4]);
    }

    // The table is not narrowed any more (owner request, 2026-11-13): the cursor is put on the first error, and the words
    // say so only when it is - the first error can be on a schematic, in no row.
    [Fact]
    public void The_blocked_submit_counts_its_errors_and_says_whether_the_cursor_is_on_the_first()
    {
        Assert.StartsWith("This draft has 1 error to fix", BoardTableProblemWording.SubmitBlocked(1, onFirst: true), StringComparison.Ordinal);
        Assert.StartsWith("This draft has 3 errors to fix", BoardTableProblemWording.SubmitBlocked(3, onFirst: false), StringComparison.Ordinal);

        Assert.Contains("The cursor is on it.", BoardTableProblemWording.SubmitBlocked(1, onFirst: true), StringComparison.Ordinal);
        Assert.Contains("The cursor is on the first.", BoardTableProblemWording.SubmitBlocked(2, onFirst: true), StringComparison.Ordinal);
        Assert.DoesNotContain("cursor", BoardTableProblemWording.SubmitBlocked(2, onFirst: false), StringComparison.Ordinal);
        Assert.DoesNotContain("Only the rows", BoardTableProblemWording.SubmitBlocked(2, onFirst: true), StringComparison.Ordinal);

        // A pill or the search hiding every error (code review, 2026-11-13): no "red corner" nobody can see - where they are.
        string hidden = BoardTableProblemWording.SubmitBlocked(2, onFirst: false, hiddenByFilter: true);
        Assert.Contains("pick \"Errors\" to see them", hidden, StringComparison.Ordinal);
        Assert.DoesNotContain("red corner", hidden, StringComparison.Ordinal);
        Assert.Contains("red corner", BoardTableProblemWording.SubmitBlocked(2, onFirst: true, hiddenByFilter: true), StringComparison.Ordinal);
    }

    // Hobbyist users: no pronoun for a contributor, and plain ASCII punctuation in what they read.
    [Fact]
    public void The_words_use_plain_punctuation()
    {
        string all = BoardTableProblemWording.SubmitBlocked(2, true) + BoardTableProblemWording.OutsideSheets([Problem(BoardProblemLevel.Error, "x")]);

        Assert.All(all, character => Assert.True(character < 128, $"[{character}] is not plain ASCII"));
    }

    // ###########################################################################################
    // The problems in the sheets, as text, for the picked pill (code review, 2026-10-13): errors
    // first, in sheet order, each naming its sheet, row and column; capped, with "and N more".
    // ###########################################################################################
    [Fact]
    public void The_problems_in_the_sheets_are_listed_errors_first_with_where_they_are()
    {
        BoardDataProblem warning = new(BoardProblemLevel.Warning, "image.orphan", "U9", "Not shown anywhere.", BoardWorkbookSchema.SheetComponentImages, 3, BoardWorkbookSchema.ColBoardLabel);
        BoardDataProblem error = new(BoardProblemLevel.Error, "link.url", "U1", "Not a web address.", BoardWorkbookSchema.SheetComponentLinks, 1, BoardWorkbookSchema.ColUrl);
        BoardDataProblem outside = new(BoardProblemLevel.Error, "highlight.orphan", "U7", "On no schematic.", null, -1, null);

        string? both = BoardTableProblemWording.InSheets([warning, error, outside], errors: true, warnings: true);
        string[] lines = both!.Split(Environment.NewLine);

        Assert.Equal(2, lines.Length);
        Assert.Equal($"Error - \"Components: Links\", U1, {BoardWorkbookSchema.ColUrl}: Not a web address.", lines[0]);
        Assert.StartsWith("Warning - \"Components: Pictures or Oscilloscope baseline\", U9", lines[1]);

        Assert.DoesNotContain("Warning", BoardTableProblemWording.InSheets([warning, error], errors: true, warnings: false));
        Assert.Null(BoardTableProblemWording.InSheets([warning, error], errors: false, warnings: false));
        Assert.Null(BoardTableProblemWording.InSheets([outside], errors: true, warnings: true));
    }

    [Fact]
    public void A_long_list_of_problems_is_capped()
    {
        List<BoardDataProblem> many = Enumerable.Range(0, BoardTableProblemWording.InSheetsShown + 5)
            .Select(i => new BoardDataProblem(BoardProblemLevel.Warning, "x", $"U{i}", "msg", BoardWorkbookSchema.SheetComponents, i, null))
            .ToList();

        string[] lines = BoardTableProblemWording.InSheets(many, errors: false, warnings: true)!.Split(Environment.NewLine);

        Assert.Equal(BoardTableProblemWording.InSheetsShown + 1, lines.Length);
        Assert.Equal("... and 5 more.", lines[^1]);
    }

    // In the order the table's tabs stand in, which is the table's own - not the order a workbook is written in
    // (code review, 2026-11-05).
    [Fact]
    public void The_problems_in_the_sheets_are_listed_in_the_order_of_the_tabs()
    {
        BoardDataProblem image = new(BoardProblemLevel.Error, "image.a", "U9", "An image.", BoardWorkbookSchema.SheetComponentImages, 0, BoardWorkbookSchema.ColFile);
        BoardDataProblem credit = new(BoardProblemLevel.Error, "credit.a", "Me", "A credit.", BoardWorkbookSchema.SheetCredits, 0, null);

        string[] lines = BoardTableProblemWording.InSheets([image, credit], errors: true, warnings: false)!.Split(Environment.NewLine);

        Assert.StartsWith("Error - \"Board: Credits\"", lines[0]);
        Assert.StartsWith("Error - \"Components: Pictures or Oscilloscope baseline\"", lines[1]);
    }
}
