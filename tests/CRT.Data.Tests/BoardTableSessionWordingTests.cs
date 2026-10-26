using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What the Drafts tab's table says about a change it handed the draft held in memory
// (BoardTableSessionWording, CRT 4.0 step 1, 2026-10-25).
// ###########################################################################################
public sealed class BoardTableSessionWordingTests
{
    [Fact]
    public void A_tables_step_names_its_sheet_for_the_drafts_history()
    {
        Assert.Equal("Edit in the table, Components", BoardTableSessionWording.Step("Components"));
        Assert.Equal("Edit in the table", BoardTableSessionWording.Step(null));
        Assert.Equal("Undone: Edit in the table, Components.", BoardTableSessionWording.Undone(BoardTableSessionWording.Step("Components")));
        Assert.Equal("Redone: Edit in the table.", BoardTableSessionWording.Redone(BoardTableSessionWording.Step(" ")));
    }

    [Fact]
    public void What_went_with_a_removed_schematic_image_is_said_with_the_way_back()
    {
        Assert.Equal(
            "What was drawn on [Top] went with it: 12 highlights and its KiCad calibration. Ctrl+Z brings it back.",
            BoardTableSessionWording.WentWith(new DraftEditLoss(["Top"], 12, 1)));

        Assert.Equal(
            "What was drawn on [Top] went with it: 1 highlight. Ctrl+Z brings it back.",
            BoardTableSessionWording.WentWith(new DraftEditLoss(["Top"], 1, 0)));

        Assert.Equal(
            "What was drawn on [Top], [Bottom] and [Clock] went with them: 3 highlights and 2 KiCad calibrations. Ctrl+Z brings it back.",
            BoardTableSessionWording.WentWith(new DraftEditLoss(["Top", "Bottom", "Clock"], 3, 2)));

        // Nothing drawn on it, or nothing removed: nothing to say.
        Assert.Null(BoardTableSessionWording.WentWith(new DraftEditLoss(["Top"], 0, 0)));
        Assert.Null(BoardTableSessionWording.WentWith(DraftEditLoss.None));
    }
}
