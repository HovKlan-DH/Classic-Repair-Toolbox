using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What a draft held in memory says about itself (DraftSaveWording) - CRT 4.0 step 1, 2026-10-25:
// "The header says Saved, Saving, or "Could not save - <reason>" with Try again".
// ###########################################################################################
public sealed class DraftSaveWordingTests
{
    [Theory]
    [InlineData(DraftSaveState.Saved, "Saved")]
    [InlineData(DraftSaveState.Pending, "Saving...")]
    [InlineData(DraftSaveState.Saving, "Saving...")]
    public void A_drafts_state_reads_Saved_or_Saving(DraftSaveState state, string expected)
    {
        Assert.Equal(expected, DraftSaveWording.StateLine(state, failure: null));
    }

    // A failed save names its reason, in words a hobbyist acts on - and no program.
    [Theory]
    [InlineData(DraftWorkbookEditOutcome.ChangedOnDisk, "Could not save - the draft file was changed somewhere else.")]
    [InlineData(DraftWorkbookEditOutcome.NoDraft, "Could not save - this draft no longer exists.")]
    [InlineData(DraftWorkbookEditOutcome.WriteFailed, "Could not save - the draft file could not be written. If another program has it open, close it there, then try again.")]
    public void A_failed_save_says_Could_not_save_and_why(DraftWorkbookEditOutcome failure, string expected)
    {
        Assert.Equal(expected, DraftSaveWording.StateLine(DraftSaveState.Failed, failure));
        Assert.DoesNotContain("Excel", expected, StringComparison.Ordinal);
    }

    // The bar after a refused save offers both versions, and says the way back from the one taken.
    [Fact]
    public void The_outside_change_bar_offers_both_versions_and_the_way_back()
    {
        Assert.Contains("Use the draft as it was saved there", DraftSaveWording.ChangedOutside, StringComparison.Ordinal);
        Assert.Contains("keep the table's version", DraftSaveWording.ChangedOutside, StringComparison.Ordinal);
        Assert.Contains("Ctrl+Z", DraftSaveWording.ChangedOutside, StringComparison.Ordinal);
        Assert.Contains("Ctrl+Z", DraftSaveWording.UsedSavedVersion, StringComparison.Ordinal);
        Assert.Contains("Ctrl+Z", DraftSaveWording.UpdatedFromOutside, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // Another editor's save that could not be made (DraftEdits took it back out) says why, where it
    // used to say "check the logfile": the one way the draft is "changed outside" then is the table
    // waiting with a change while another program wrote the file - so it says where to settle that.
    // ###########################################################################################
    [Theory]
    [InlineData(DraftWorkbookEditOutcome.ChangedOnDisk, "The draft was changed outside CRT while changes in the table view waited to be saved. Choose there which version to keep, then save again.")]
    [InlineData(DraftWorkbookEditOutcome.NoDraft, "This draft no longer exists.")]
    [InlineData(DraftWorkbookEditOutcome.WriteFailed, "The draft file could not be written. If another program has it open, close it there, then save again.")]
    public void Another_editors_save_that_could_not_be_made_says_why(DraftWorkbookEditOutcome failure, string expected)
    {
        Assert.Equal(expected, DraftSaveWording.EditNotSaved(failure));
        Assert.DoesNotContain("Excel", expected, StringComparison.Ordinal);
    }

    // The row's Undo and Redo name the step they take, as the history describes it.
    [Theory]
    [InlineData(null, "Nothing to undo", "Nothing to redo")]
    [InlineData(" ", "Nothing to undo", "Nothing to redo")]
    [InlineData("Changes to [U8]", "Undo: Changes to [U8]", "Redo: Changes to [U8]")]
    public void Undo_and_redo_name_their_step(string? description, string undo, string redo)
    {
        Assert.Equal(undo, DraftSaveWording.UndoTip(description));
        Assert.Equal(redo, DraftSaveWording.RedoTip(description));
    }
}
