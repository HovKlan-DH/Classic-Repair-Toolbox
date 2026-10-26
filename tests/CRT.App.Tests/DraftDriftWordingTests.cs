using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Every sentence the drift warning says (code review, 2026-10-13). The Drafts row, the main
// window's notice and the window all take their words from DraftDriftWording, so these pin the
// promises the words make - none of which may be "your edits are still applied".
// ###########################################################################################
public sealed class DraftDriftWordingTests
{
    [Theory]
    [InlineData(DraftDriftState.OfficialIsNewer, true)]
    [InlineData(DraftDriftState.Changed, true)]
    [InlineData(DraftDriftState.OfficialIsNewer, false)]
    [InlineData(DraftDriftState.Changed, false)]
    public void No_sentence_promises_the_edits_are_applied_on_top(DraftDriftState state, bool canBringUpToDate)
    {
        Assert.DoesNotContain("applied", DraftDriftWording.Summary(state, canBringUpToDate));
        Assert.DoesNotContain("applied", DraftDriftWording.Line(state, canBringUpToDate));
    }

    [Fact]
    public void Updated_is_said_only_when_the_published_revision_is_known_to_be_later()
    {
        Assert.Contains("has been updated", DraftDriftWording.Line(DraftDriftState.OfficialIsNewer, true));
        Assert.Contains("has changed", DraftDriftWording.Line(DraftDriftState.Changed, true));
        Assert.DoesNotContain("updated", DraftDriftWording.Line(DraftDriftState.Changed, false));
    }

    [Fact]
    public void Only_a_real_difference_is_a_warning()
    {
        Assert.True(DraftDriftWording.IsWarning(DraftDriftState.OfficialIsNewer));
        Assert.True(DraftDriftWording.IsWarning(DraftDriftState.Changed));
        Assert.False(DraftDriftWording.IsWarning(DraftDriftState.InSync));
        Assert.False(DraftDriftWording.IsWarning(DraftDriftState.Unknown));
    }

    [Fact]
    public void A_draft_without_a_base_is_told_what_sending_it_would_do()
    {
        string summary = DraftDriftWording.Summary(DraftDriftState.OfficialIsNewer, canBringUpToDate: false);

        Assert.Contains("does not have that update", summary);
        Assert.Contains("the maintainer sees the update's changes as changes you made", summary);
    }

    [Fact]
    public void Rows_are_said_as_the_contributors_with_a_base_and_as_differences_without()
    {
        var removed = new BoardRowChange { Section = "Components", DisplayLabel = "U9", Kind = BoardRowChangeKind.Deleted };
        var changed = new BoardRowChange { Section = "Components", DisplayLabel = "U8", Kind = BoardRowChangeKind.Modified, ChangedFields = ["Description"] };

        Assert.StartsWith("You removed this", DraftDriftWording.OwnRow(removed, comparedWithBase: true));
        Assert.StartsWith("In the published data, not in your draft", DraftDriftWording.OwnRow(removed, comparedWithBase: false));
        Assert.Equal("You changed: Description", DraftDriftWording.OwnRow(changed, comparedWithBase: true));
        Assert.Equal("Changed in the update: Description", DraftDriftWording.UpdateRow(changed));
    }

    [Fact]
    public void The_outcome_counts_what_was_kept_and_points_at_the_rows_both_changed()
    {
        string clean = DraftDriftWording.Outcome(new DraftRebaseReport { Outcome = DraftRebaseOutcome.BroughtUpToDate, ChangesKept = 3 });
        string met = DraftDriftWording.Outcome(new DraftRebaseReport
        {
            Outcome = DraftRebaseOutcome.BroughtUpToDate,
            ChangesKept = 1,
            Conflicts = [new DraftRebaseConflict("Components", "U8", "both")],
        });

        Assert.Equal("Your draft now has the published update. All 3 changes of yours were kept.", clean);
        Assert.Contains("1 of them was also changed by the update", met);
        Assert.True(DraftDriftWording.IsSuccess(new DraftRebaseReport { Outcome = DraftRebaseOutcome.AlreadyUpToDate }));
    }

    // The draft has the update but could not be noted as up to date (code review, 2026-10-13): not
    // a success - the button stays - and it says pressing it again finishes the job.
    [Fact]
    public void An_update_that_could_not_be_noted_asks_to_press_the_button_again()
    {
        var report = new DraftRebaseReport { Outcome = DraftRebaseOutcome.NotNoted, ChangesKept = 2 };

        Assert.False(DraftDriftWording.IsSuccess(report));
        Assert.Equal(
            "Your draft now has the published update, but CRT could not note that it does, so the warning stays. " +
            "Press \"Bring my draft up to date\" again to finish - nothing of yours is lost.",
            DraftDriftWording.Outcome(report));
    }

    [Theory]
    [InlineData(DraftRebaseOutcome.ChangedOnDisk)]
    [InlineData(DraftRebaseOutcome.WriteFailed)]
    [InlineData(DraftRebaseOutcome.PublishedUnreadable)]
    [InlineData(DraftRebaseOutcome.DraftUnreadable)]
    [InlineData(DraftRebaseOutcome.NoBase)]
    public void Every_failure_says_nothing_was_changed(DraftRebaseOutcome outcome)
    {
        var report = new DraftRebaseReport { Outcome = outcome };

        Assert.False(DraftDriftWording.IsSuccess(report));
        Assert.Contains("Nothing was changed", DraftDriftWording.Outcome(report));

        // A CRT 4.0 draft is a board file no spreadsheet program edits (2026-10-25), so no failure
        // sends the contributor to close it in Excel.
        Assert.DoesNotContain("Excel", DraftDriftWording.Outcome(report));
    }
}
