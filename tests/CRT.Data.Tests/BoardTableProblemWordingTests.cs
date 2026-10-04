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

        Assert.Contains("label editor", lines[0], StringComparison.Ordinal);
        Assert.Equal(["Error: e1", "Error: e2", "Warning: w1"], lines[1..4]);
        Assert.Equal("... and 2 more.", lines[4]);
    }

    [Fact]
    public void The_blocked_submit_counts_its_errors_and_says_whether_the_table_is_filtered()
    {
        Assert.StartsWith("This draft has 1 error to fix", BoardTableProblemWording.SubmitBlocked(1, rowsFiltered: true), StringComparison.Ordinal);
        Assert.StartsWith("This draft has 3 errors to fix", BoardTableProblemWording.SubmitBlocked(3, rowsFiltered: false), StringComparison.Ordinal);

        Assert.Contains("Only the rows with errors are shown", BoardTableProblemWording.SubmitBlocked(2, rowsFiltered: true), StringComparison.Ordinal);
        Assert.DoesNotContain("Only the rows", BoardTableProblemWording.SubmitBlocked(2, rowsFiltered: false), StringComparison.Ordinal);
    }

    // Hobbyist users: no pronoun for a contributor, and plain ASCII punctuation in what they read.
    [Fact]
    public void The_words_use_plain_punctuation()
    {
        string all = BoardTableProblemWording.SubmitBlocked(2, true) + BoardTableProblemWording.OutsideSheets([Problem(BoardProblemLevel.Error, "x")]);

        Assert.All(all, character => Assert.True(character < 128, $"[{character}] is not plain ASCII"));
    }
}
