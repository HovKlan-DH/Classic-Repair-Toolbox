using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The automatic comments the app writes into an entry's own Comments list when it is created, its
// state is flipped, or its category is changed - an audit trail sitting beside the user's own
// comments, so a worklog explains its own history rather than only showing its current state.
//
// The exact wording is pinned here because it is user-visible text that two callers share (the
// quick create card and the full editor). Building the strings in WorklogManager rather than at
// those call sites is what stops the two drifting apart, and these tests are what stop the
// wording changing by accident.
public class WorklogAutomaticCommentTests
{
    // ------------------------------------------------------------- wording

    [Fact]
    public void The_created_comment_reads_worklog_created()
    {
        Assert.Equal("Worklog created", WorklogManager.CreatedCommentText);
    }

    [Theory]
    [InlineData("Open", "Worklog opened")]
    [InlineData("Closed", "Worklog closed")]
    public void A_state_change_is_described_in_the_past_tense(string state, string expected)
    {
        Assert.Equal(expected, WorklogManager.BuildStateChangedCommentText(state));
    }

    // An unrecognised state records nothing rather than claiming the worklog was opened. A value
    // from a future build would otherwise be logged as the wrong event, which is worse than an
    // absent line in an audit trail.
    [Theory]
    [InlineData("Pending")]
    [InlineData("RuledOut")]
    [InlineData("")]
    [InlineData("open")]
    public void An_unrecognised_state_produces_no_comment(string state)
    {
        Assert.Null(WorklogManager.BuildStateChangedCommentText(state));
    }

    // The category name is quoted. A bare "Worklog changed to Note" reads as a sentence missing a
    // word; the quotes mark the value as the literal category name.
    [Theory]
    [InlineData("Note")]
    [InlineData("Cosmetic")]
    [InlineData("Issue")]
    public void A_category_change_quotes_the_category_name(string category)
    {
        Assert.Equal($"Worklog changed to \"{category}\"", WorklogManager.BuildCategoryChangedCommentText(category));
    }

    // ------------------------------------------------------------- appending

    [Fact]
    public void An_automatic_comment_is_appended_with_its_text_and_a_date()
    {
        var comments = new List<WorklogCommentRecord>();

        var added = WorklogManager.AppendAutomaticComment(comments, "Worklog created");

        Assert.NotNull(added);
        Assert.Single(comments);
        Assert.Equal("Worklog created", comments[0].Text);
        Assert.NotEqual(default, comments[0].Date);
    }

    // The first comment takes id 1, not 0 - matching the editor's own Add-comment, whose delete
    // and edit handlers match rows by id.
    [Fact]
    public void The_first_comment_gets_id_one()
    {
        var comments = new List<WorklogCommentRecord>();

        WorklogManager.AppendAutomaticComment(comments, "Worklog created");

        Assert.Equal(1, comments[0].Id);
    }

    // Ids come from the highest EXISTING id, not from the count. A list whose earlier comments
    // have been deleted would otherwise reuse an id that a row still on screen might hold.
    [Fact]
    public void A_new_comment_takes_the_next_id_after_the_highest_in_use()
    {
        var comments = new List<WorklogCommentRecord>
        {
            new() { Id = 4, Text = "User comment", Date = DateTime.Now },
        };

        var added = WorklogManager.AppendAutomaticComment(comments, "Worklog closed");

        Assert.Equal(5, added!.Id);
    }

    // Blank text adds nothing: BuildStateChangedCommentText can decline to describe a state, and
    // an empty row in the comment list would be worse than no row at all.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Blank_text_adds_no_comment(string? text)
    {
        var comments = new List<WorklogCommentRecord>();

        var added = WorklogManager.AppendAutomaticComment(comments, text);

        Assert.Null(added);
        Assert.Empty(comments);
    }

    [Fact]
    public void The_comment_text_is_trimmed()
    {
        var comments = new List<WorklogCommentRecord>();

        WorklogManager.AppendAutomaticComment(comments, "  Worklog opened  ");

        Assert.Equal("Worklog opened", comments[0].Text);
    }

    // Existing comments are kept - the automatic ones sit alongside the user's own, they do not
    // replace them.
    [Fact]
    public void Appending_keeps_the_comments_already_there()
    {
        var comments = new List<WorklogCommentRecord>
        {
            new() { Id = 1, Text = "Replaced C12", Date = DateTime.Now },
        };

        WorklogManager.AppendAutomaticComment(comments, "Worklog closed");

        Assert.Equal(2, comments.Count);
        Assert.Equal("Replaced C12", comments[0].Text);
        Assert.Equal("Worklog closed", comments[1].Text);
    }
}
