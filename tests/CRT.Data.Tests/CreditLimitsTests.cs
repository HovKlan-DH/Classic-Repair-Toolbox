using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// CreditLimits - a credit's short fields are one short line (owner request, 2026-11-15: "it was not the intention to state
// everything in there, as this should be a one-liner only"; asked: "What" 60, "Which part of it" 60, "Name or handle" 50,
// "Contact" 100). An error in both scopes on a credit not published as it stands - the table marks the cell, the server refuses
// it - and nothing said of one published as it stands, so what is published stays as it is.
// ###########################################################################################
public sealed class CreditLimitsTests
{
    private static CreditEntry Credit(int category = 10, int subCategory = 10, int name = 10, int contact = 10) => new()
    {
        Category = new string('c', category),
        SubCategory = new string('s', subCategory),
        NameOrHandle = new string('n', name),
        Contact = new string('w', contact),
        MainContribution = "Drew the bottom side",
    };

    [Fact]
    public void Each_field_may_be_as_long_as_its_limit_and_no_longer()
    {
        CreditEntry atLimit = Credit(CreditLimits.Category, CreditLimits.SubCategory, CreditLimits.NameOrHandle, CreditLimits.Contact);
        Assert.Empty(CreditLimits.TooLong(atLimit, judged: true));

        CreditEntry over = Credit(CreditLimits.Category + 1, CreditLimits.SubCategory + 2, CreditLimits.NameOrHandle + 3, CreditLimits.Contact + 4);

        Assert.Equal(
            [
                (BoardWorkbookSchema.ColCategory, 60, 61),
                (BoardWorkbookSchema.ColSubCategory, 60, 62),
                (BoardWorkbookSchema.ColNameOrHandle, 50, 53),
                (BoardWorkbookSchema.ColContact, 100, 104),
            ],
            CreditLimits.TooLong(over, judged: true));
    }

    // ###########################################################################################
    // *** The main contribution at most 500 characters (owner decision, 2026-11-16: "500 characters? Shouldn't this be enough to
    // state the most important things? ... let us start with this"), counted as its words - formatting marks no characters. ***
    // An error on its cell in both scopes, as the short fields, saying what to keep.
    // ###########################################################################################
    [Fact]
    public void The_main_contribution_may_be_500_characters_of_its_words()
    {
        CreditEntry atLimit = new() { Category = "Board data", NameOrHandle = "Anna", MainContribution = "**" + new string('w', CreditLimits.MainContribution) + "**" };
        Assert.Equal(500, CreditLimits.MainContribution);
        Assert.Empty(CreditLimits.TooLong(atLimit, judged: true));

        CreditEntry over = new() { Category = "Board data", NameOrHandle = "Anna", MainContribution = new string('w', CreditLimits.MainContribution + 1) };
        Assert.Equal([(BoardWorkbookSchema.ColMainContribution, 500, 501)], CreditLimits.TooLong(over, judged: true));

        BoardDataProblem refused = Assert.Single(
            BoardDataChecks.Check(BoardCheckRows.From(new BoardData { Credits = [over] }, PublishedCredits.None), null, BoardCheckScope.Everything),
            problem => problem.Code == "credit.too_long");
        Assert.Equal(BoardWorkbookSchema.ColMainContribution, refused.Column);
        Assert.Contains("most important updates", refused.Message, StringComparison.Ordinal);
    }

    // Counted as kept: the spaces round a value are no part of it.
    [Fact]
    public void Spaces_round_a_value_are_not_counted()
    {
        var credit = new CreditEntry { Category = "  " + new string('c', CreditLimits.Category) + "  ", NameOrHandle = "Anna" };

        Assert.Empty(CreditLimits.TooLong(credit, judged: true));
    }

    // What is published stays as it is - one row for each published credit - and the finalise (null) judges nothing, its rows
    // judged as the submission was created.
    [Fact]
    public void A_credit_published_as_it_stands_is_not_held_to_the_limits()
    {
        static IEnumerable<BoardDataProblem> TooLong(PublishedCredits? published, params CreditEntry[] credits) =>
            BoardDataChecks.Check(BoardCheckRows.From(new BoardData { Credits = [.. credits] }, published), null, BoardCheckScope.Everything)
                .Where(problem => problem.Code == "credit.too_long");

        Assert.Empty(TooLong(PublishedCredits.Of([Credit(category: 80)]), Credit(category: 80)));
        Assert.Empty(TooLong(null, Credit(category: 80)));
        Assert.Single(TooLong(PublishedCredits.Of([Credit(category: 79)]), Credit(category: 80)));

        // A second like the published one is new, and held to them.
        Assert.Equal(1, Assert.Single(TooLong(PublishedCredits.Of([Credit(category: 80)]), Credit(category: 80), Credit(category: 80))).Index);
    }

    // An error on the very cell, in both scopes - so the server refuses what the table marks.
    [Fact]
    public void A_field_too_long_is_an_error_on_its_cell_on_both_sides()
    {
        var board = new BoardData { Credits = [Credit(), Credit(contact: 101)] };

        foreach (BoardCheckScope scope in new[] { BoardCheckScope.Submission, BoardCheckScope.Everything })
        {
            BoardDataProblem problem = Assert.Single(
                BoardDataChecks.Check(BoardCheckRows.From(board, PublishedCredits.None), null, scope),
                each => each.Code == "credit.too_long");

            Assert.Equal(BoardProblemLevel.Error, problem.Level);
            Assert.Equal((BoardWorkbookSchema.SheetCredits, 1, BoardWorkbookSchema.ColContact), (problem.Sheet, problem.Index, problem.Column));
            Assert.Contains("101 characters", problem.Message, StringComparison.Ordinal);
            Assert.Contains("at most 100", problem.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void The_server_refuses_a_submission_with_a_field_too_long()
    {
        var manifest = new SubmissionManifest
        {
            BoardId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = "Credited the bottom side",
            Rows = new SubmissionRows { Credits = [Credit(category: 61)] },
        };

        ValidationFinding refused = Assert.Single(SubmissionValidator.Validate(manifest, [], PublishedCredits.None), finding => finding.Code == "credit.too_long");

        Assert.Equal(ValidationSeverity.Error, refused.Severity);
        Assert.Equal(BoardWorkbookSchema.ColCategory, refused.Subject);
    }

    // Every credit published today fits - the limits were chosen against them (measured 2026-11-15: the longest "What" and
    // "Which part of it" 33, "Name or handle" 16, "Contact" 60).
    [Fact]
    public void The_limits_leave_room_for_what_is_published()
    {
        Assert.Empty(CreditLimits.TooLong(
            new CreditEntry
            {
                Category = "Oscilloscope baseline measurement",
                SubCategory = "VIC; RF Modulator & AV connectors",
                NameOrHandle = "CMack/Deadweasel",
                Contact = "https://github.com/Deadweasel/Commodore-64-mainboard-layouts",
            },
            judged: true));
    }
}
