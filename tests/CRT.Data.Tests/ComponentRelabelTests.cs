using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A component relabelled keeps its highlights (code review, 2026-10-13). A typo fixed in the
// Components sheet ("U8" -> "U08") left the rectangles on "U8": a stray rectangle and an unmarked
// chip on every machine after publishing. The rule is the count's own: a KEY edited alone is the
// row renamed; a key and another cell is a new row and a deleted one.
// ###########################################################################################
public sealed class ComponentRelabelTests
{
    private static ComponentEntry Component(string label, string description = "CPU", string region = "") =>
        new() { BoardLabel = label, Description = description, Region = region };

    private static ComponentHighlightEntry Highlight(string label) =>
        new() { SchematicName = "Sheet 1", BoardLabel = label, X = "1", Y = "1", Width = "5", Height = "5" };

    [Fact]
    public void A_label_edited_alone_is_followed()
    {
        IReadOnlyDictionary<string, string> map = ComponentRelabel.Between([Component("U8"), Component("U9", "VIC")], [Component("U08"), Component("U9", "VIC")]);

        Assert.Equal("U08", Assert.Single(map).Value);
        Assert.Equal(
            ["U08", "U9"],
            ComponentRelabel.FollowHighlights([Highlight("U8"), Highlight("U9")], map).Select(h => h.BoardLabel).ToArray());
    }

    [Fact]
    public void A_label_and_another_cell_edited_together_is_not_a_rename()
    {
        Assert.Empty(ComponentRelabel.Between([Component("U8", "CPU")], [Component("U08", "6510 CPU")]));
    }

    [Fact]
    public void A_label_another_component_still_carries_is_not_moved()
    {
        // Regional twins share a label: renaming one leaves the other's - and the highlights - put.
        Assert.Empty(ComponentRelabel.Between(
            [Component("U8", region: "PAL"), Component("U8", region: "NTSC")],
            [Component("U08", region: "PAL"), Component("U8", region: "NTSC")]));
    }

    [Fact]
    public void A_capitalisation_or_region_only_edit_moves_nothing()
    {
        Assert.Empty(ComponentRelabel.Between([Component("u8")], [Component("U8")]));
        Assert.Empty(ComponentRelabel.Between([Component("U8", region: "PAL")], [Component("U8", region: "NTSC")]));
    }
}
