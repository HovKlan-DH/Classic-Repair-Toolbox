using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The Workbooks tab's board pane lets the user drag its schematic previews into their own order,
// persisted per workbook. This class is the whole of that ordering rule - the tab only joins the
// result back onto its grouping objects and draws it.
//
// The rules worth knowing before reading these, all of which exist because the SET of schematics
// shown changes underneath a stored order (worklogs get added and deleted, and a search filters the
// pane further): the order is stored as NAMES rather than indices, an unknown name sorts to the
// bottom, a stored name that is not currently shown is skipped, and the whole thing degrades to
// "leave it alone" when nothing has ever been dragged.
public sealed class WorkbookSchematicOrderTests
{
    // The no-stored-order case, which is every workbook until someone drags something. The caller's
    // own order (the pane's alphabetical grouping) has to survive untouched, or this feature would
    // silently rearrange every workbook that has never used it.
    [Fact]
    public void No_stored_order_leaves_the_schematics_exactly_as_they_arrived()
    {
        var shown = new[] { "Motherboard", "Power supply", "Video" };

        Assert.Equal(shown, WorkbookSchematicOrder.Apply(shown, new List<string>()));
        Assert.Equal(shown, WorkbookSchematicOrder.Apply(shown, null));
    }

    [Fact]
    public void A_stored_order_overrides_the_alphabetical_one()
    {
        var shown = new[] { "Motherboard", "Power supply", "Video" };
        var stored = new List<string> { "Video", "Motherboard", "Power supply" };

        Assert.Equal(
            new[] { "Video", "Motherboard", "Power supply" },
            WorkbookSchematicOrder.Apply(shown, stored));
    }

    // A schematic that has just received its first worklog is not in the stored order yet. It must
    // arrive at the BOTTOM rather than in its alphabetical place among the ordered ones: slotting it
    // into the middle would shift everything below it down by one, for a schematic the user has
    // never positioned.
    [Fact]
    public void A_schematic_missing_from_the_stored_order_goes_to_the_bottom()
    {
        var shown = new[] { "Audio", "Motherboard", "Video" };
        var stored = new List<string> { "Video", "Motherboard" };

        Assert.Equal(
            new[] { "Video", "Motherboard", "Audio" },
            WorkbookSchematicOrder.Apply(shown, stored));
    }

    // Deleting a schematic's last worklog stops it being shown, but its name stays in the stored
    // order. That must not leave a hole or throw - and, per the next test, its position must still
    // be honoured if it ever comes back.
    [Fact]
    public void A_stored_name_that_is_no_longer_shown_is_skipped()
    {
        var shown = new[] { "Motherboard", "Video" };
        var stored = new List<string> { "Video", "Power supply", "Motherboard" };

        Assert.Equal(
            new[] { "Video", "Motherboard" },
            WorkbookSchematicOrder.Apply(shown, stored));
    }

    // The pay-off for storing names rather than indices: a schematic whose worklogs were all deleted
    // and later added again returns to the position the user put it in, rather than to the bottom.
    [Fact]
    public void A_schematic_that_comes_back_returns_to_its_stored_position()
    {
        var stored = new List<string> { "Video", "Power supply", "Motherboard" };

        Assert.Equal(
            new[] { "Video", "Power supply", "Motherboard" },
            WorkbookSchematicOrder.Apply(new[] { "Motherboard", "Power supply", "Video" }, stored));
    }

    // Board Excel files arrive from the server independently of app releases and nothing normalises
    // their casing, so every schematic-name lookup in this app is case-insensitive. A stored order
    // written before a board file changed the casing of a name must still apply.
    [Fact]
    public void Names_are_matched_without_regard_to_case()
    {
        var shown = new[] { "Motherboard", "Video" };
        var stored = new List<string> { "VIDEO", "motherboard" };

        Assert.Equal(new[] { "Video", "Motherboard" }, WorkbookSchematicOrder.Apply(shown, stored));
    }

    // index.json is a plain file a user can hand-edit, so neither a duplicate nor a blank may
    // corrupt the pane. A duplicate must not place the same preview twice (which, in the tab, would
    // throw on the dictionary join) and a blank must not consume a slot.
    [Fact]
    public void A_hand_edited_order_with_duplicates_or_blanks_still_produces_each_schematic_once()
    {
        var shown = new[] { "Motherboard", "Video" };
        var stored = new List<string> { "Video", "", "  ", "Video", "Motherboard" };

        Assert.Equal(new[] { "Video", "Motherboard" }, WorkbookSchematicOrder.Apply(shown, stored));
    }

    [Fact]
    public void Moving_a_schematic_down_puts_it_at_the_target_position()
    {
        var displayed = new[] { "A", "B", "C", "D" };

        Assert.Equal(
            new[] { "B", "C", "A", "D" },
            WorkbookSchematicOrder.ApplyMove(displayed, "A", 2));
    }

    [Fact]
    public void Moving_a_schematic_up_puts_it_at_the_target_position()
    {
        var displayed = new[] { "A", "B", "C", "D" };

        Assert.Equal(
            new[] { "D", "A", "B", "C" },
            WorkbookSchematicOrder.ApplyMove(displayed, "D", 0));
    }

    // The caller persists unconditionally rather than testing for a no-op first, so a drop back onto
    // the schematic's own slot has to come back unchanged rather than shifting anything.
    [Fact]
    public void Dropping_a_schematic_where_it_already_is_changes_nothing()
    {
        var displayed = new[] { "A", "B", "C" };

        Assert.Equal(displayed, WorkbookSchematicOrder.ApplyMove(displayed, "B", 1));
    }

    // A drop past either end of the pane (the pointer released above the first preview or below the
    // last) clamps rather than throwing or dropping the schematic out of the list.
    [Fact]
    public void A_drop_beyond_either_end_clamps_into_the_list()
    {
        var displayed = new[] { "A", "B", "C" };

        Assert.Equal(new[] { "B", "C", "A" }, WorkbookSchematicOrder.ApplyMove(displayed, "A", 99));
        Assert.Equal(new[] { "C", "A", "B" }, WorkbookSchematicOrder.ApplyMove(displayed, "C", -5));
    }

    // A refresh can rebuild the pane mid-drag, so by the time the drop is committed the dragged
    // schematic may no longer be shown. That must leave the order alone rather than inventing one.
    [Fact]
    public void Moving_a_schematic_that_is_not_shown_leaves_the_order_alone()
    {
        var displayed = new[] { "A", "B" };

        Assert.Equal(displayed, WorkbookSchematicOrder.ApplyMove(displayed, "Gone", 0));
    }

    [Fact]
    public void A_single_schematic_cannot_be_reordered()
    {
        Assert.Equal(new[] { "A" }, WorkbookSchematicOrder.ApplyMove(new[] { "A" }, "A", 0));
        Assert.Empty(WorkbookSchematicOrder.Apply(Array.Empty<string>(), new List<string> { "A" }));
    }

    // ---------------------------------------------------------------------------------------------
    // ResolveDropIndex - which slot the pointer is over.
    //
    // Schematic previews are NOT uniform-height rows (each is as tall as its image needs), which is
    // why this cannot be the "pointerY / rowHeight" division the editor's photo list can use. The
    // rule is the standard one: the drop lands before the first preview whose vertical midpoint the
    // pointer has not yet passed.
    // ---------------------------------------------------------------------------------------------

    // Heights 100/200/50 with a 10px gap put the previews at y=0-100, 110-310 and 320-370, so the
    // midpoints are at 50, 210 and 345.
    [Theory]
    [InlineData(0, 0)]      // above everything
    [InlineData(49, 0)]     // just above the first midpoint
    [InlineData(51, 1)]     // just past it
    [InlineData(209, 1)]    // just above the second midpoint
    [InlineData(211, 2)]    // just past it
    [InlineData(1000, 2)]   // below everything, clamped to the last slot
    public void The_drop_lands_at_the_preview_whose_midpoint_the_pointer_has_passed(double pointerY, int expectedIndex)
    {
        var heights = new List<double> { 100, 200, 50 };

        Assert.Equal(expectedIndex, WorkbookSchematicOrder.ResolveDropIndex(heights, 10, pointerY));
    }

    // The gap between previews is part of the distance the pointer travels. Ignoring it drifts every
    // midpoint below the first upward by one gap per preview, which is what makes a drop near the
    // bottom of a long pane land a slot early - so this pins the spacing being counted.
    //
    // Read against the SECOND midpoint, the first one spacing can move, over four previews so that
    // neither answer is the saturated last slot (with three, everything past the second midpoint
    // clamps to index 2 and the two spacings agree no matter what).
    //
    // Four 100px previews put the midpoints at 50/150/250/350 with no gaps, and at 50/190/330/470
    // with a 40px gap between each. A pointer at 160 has passed the second midpoint in the first
    // case but not in the second, so it lands in a different slot depending only on whether the
    // spacing was counted - which is exactly the drift that makes a drop near the bottom of a long
    // pane land a slot early.
    [Fact]
    public void The_spacing_between_previews_counts_toward_where_a_drop_lands()
    {
        var heights = new List<double> { 100, 100, 100, 100 };

        Assert.Equal(2, WorkbookSchematicOrder.ResolveDropIndex(heights, 0, 160));
        Assert.Equal(1, WorkbookSchematicOrder.ResolveDropIndex(heights, 40, 160));
    }

    [Fact]
    public void An_empty_pane_drops_at_the_first_slot_rather_than_throwing()
    {
        Assert.Equal(0, WorkbookSchematicOrder.ResolveDropIndex(new List<double>(), 10, 500));
    }
}
