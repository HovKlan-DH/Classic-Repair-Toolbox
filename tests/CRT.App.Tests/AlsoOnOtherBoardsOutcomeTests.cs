using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// AlsoOnOtherBoardsOutcome - what "Also on other boards" did, board by board, in words: a board where
// anything changed is "Changed"; what was not changed is said - the board as a whole when nothing on it
// changed for one reason, else each component by its label - and each file not put is said by the file.
// ###########################################################################################
public sealed class AlsoOnOtherBoardsOutcomeTests
{
    [Fact]
    public void A_board_where_nothing_changed_for_one_reason_is_said_as_a_whole()
    {
        var outcome = new AlsoOnOtherBoardsOutcome();
        outcome.Board("B", changed: false, [("U9", ComponentPaneWording.AlsoNotSaved), ("U10", ComponentPaneWording.AlsoNotSaved)]);

        Assert.Equal("Not changed on B - its draft could not be saved.", outcome.Said);
    }

    // A board where something changed - or where components failed for different reasons - says each
    // component not changed, so none hides behind another or behind a file that went in.
    [Fact]
    public void A_component_not_changed_beside_a_change_is_said_by_its_label()
    {
        var outcome = new AlsoOnOtherBoardsOutcome();
        outcome.Board("B", changed: true, [("U10", ComponentPaneWording.AlsoNotSaved)]);
        outcome.Board("C", changed: false, [("U1", ComponentPaneWording.AlsoNotSaved), ("U2", ComponentPaneWording.AlsoAlreadySo)]);

        Assert.Equal(
            "Changed on 1 board, in its own draft: B. [U10] not changed on B - its draft could not be saved. [U1] not changed on C - its draft could not be saved. [U2] not changed on C - it is like that there already.",
            outcome.Said);
    }

    [Fact]
    public void A_board_without_a_draft_and_a_file_not_put_are_said()
    {
        var outcome = new AlsoOnOtherBoardsOutcome();
        outcome.NoDraft("D");
        outcome.FileNotPut("Datasheet (6510.pdf)", "E", ComponentPaneWording.AlsoFileOwnVersion);
        outcome.Board("E", changed: false, []);

        Assert.Equal(
            "Not changed on D - no draft could be made of it. [Datasheet (6510.pdf)] not put on E - that board's draft has a version of its own.",
            outcome.Said);
    }
}
