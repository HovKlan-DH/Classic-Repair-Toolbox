using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// BoardOrderDisplay - Account > "Order of boards" (owner request, 2026-10-04: "sort the list of
// systems, which then gets saved to both sources (BETA + stable) after my save"). What a save says
// must match what the server did to each list: a claim of "saved in the stable source" that did
// not happen would send the project owner looking for an order CRT never got.
// ###########################################################################################
public sealed class BoardOrderDisplayTests
{
    [Fact]
    public void A_list_is_reordered_only_when_a_board_moved()
    {
        Assert.False(BoardOrderDisplay.IsReordered(["A/B/1", "A/B/2"], ["a/b/1", "A/B/2"]));
        Assert.True(BoardOrderDisplay.IsReordered(["A/B/1", "A/B/2"], ["A/B/2", "A/B/1"]));
    }

    // ###########################################################################################
    // A main Excel data file listing one board on two rows (two workbooks in one board folder)
    // sent that id twice, which the server refuses as "in the list more than once" - so such a list
    // could never be saved (code review, 2026-10-04). Each board is sent once, where it first
    // appears, in any case.
    // ###########################################################################################
    [Fact]
    public void A_board_listed_on_two_rows_is_sent_once_where_it_first_appears()
    {
        Assert.Equal(
            ["A/B/2", "A/B/1", "A/B/3"],
            BoardOrderDisplay.OrderToSend(["A/B/2", "A/B/1", "a/b/2", "A/B/3", "A/B/1"]));
    }

    [Theory]
    [InlineData(true, true, null, "Saved. CRT's drop-down lists in BETA and the stable source are in the new order.", false)]
    [InlineData(true, null, null, "Saved. CRT's drop-down list in BETA is in the new order (this server has no stable source).", false)]
    [InlineData(true, false, null, "Saved. CRT's drop-down list in BETA is in the new order; the stable source's was already in it.", false)]
    [InlineData(false, true, null, "Saved. BETA's list was already in this order; the stable source's list is now in it too.", false)]
    [InlineData(false, false, null, "Nothing to save - the lists in BETA and the stable source were already in this order.", false)]
    [InlineData(false, null, null, "Nothing to save - BETA's list was already in this order.", false)]
    public void What_a_save_did_is_said_for_each_list(bool beta, bool? stable, string? problem, string text, bool isError)
    {
        Assert.Equal((text, isError), BoardOrderDisplay.Saved(new BoardOrderAnswer(beta, stable, problem)));
    }

    // ###########################################################################################
    // *** A LIST THAT COULD NOT BE WRITTEN IS NEVER CALLED "ALREADY IN THIS ORDER". *** The server
    // answers StableChanged false both when the stable list already had the order and when it could
    // not be written - the problem tells them apart, and turns the line red.
    // ###########################################################################################
    [Fact]
    public void A_stable_list_that_could_not_be_written_is_said_as_such_and_in_red()
    {
        (string text, bool isError) = BoardOrderDisplay.Saved(new BoardOrderAnswer(true, false, "The stable source's list could not be changed: locked."));

        Assert.True(isError);
        Assert.Equal("Saved. CRT's drop-down list in BETA is in the new order. But: The stable source's list could not be changed: locked.", text);
        Assert.DoesNotContain("already", text, StringComparison.Ordinal);
    }
}
