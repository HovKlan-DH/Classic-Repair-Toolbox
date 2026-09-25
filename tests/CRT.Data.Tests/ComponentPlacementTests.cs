using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where a component's rows go in the Components sheet (owner request, 2026-09-24).
//
// Row order is what the user SEES: the main window's component list, its category list and the
// Overview all show components in the sheet's own order. So: a NEW component joins its category
// in natural label order (C1, C2, C10), a new category goes last, and an EDITED component stays
// exactly where it is - it used to jump to the bottom on every save.
// ###########################################################################################
public sealed class ComponentPlacementTests
{
    private static ComponentEntry C(string label, string category = "Capacitor") =>
        new() { BoardLabel = label, Category = category };

    private static List<string> Labels(IEnumerable<ComponentEntry> components) =>
        components.Select(component => component.BoardLabel).ToList();

    // ------------------------------------------------------------------ InsertionIndex

    [Fact]
    public void A_new_component_goes_into_its_category_in_natural_label_order()
    {
        var components = new List<ComponentEntry> { C("U1", "IC"), C("C1"), C("C3"), C("C10"), C("U2", "IC") };

        Assert.Equal(2, ComponentPlacement.InsertionIndex(components, "Capacitor", "C2"));

        // Natural, not ordinal: C4 comes before C10, which string ordering would get wrong.
        Assert.Equal(3, ComponentPlacement.InsertionIndex(components, "Capacitor", "C4"));
    }

    [Fact]
    public void A_label_sorting_after_the_whole_category_goes_just_after_that_categorys_last_row()
    {
        // Not at the very end - the IC rows after the capacitors stay after them.
        var components = new List<ComponentEntry> { C("C1"), C("C2"), C("U1", "IC") };

        Assert.Equal(2, ComponentPlacement.InsertionIndex(components, "Capacitor", "C99"));
    }

    [Fact]
    public void A_label_sorting_before_the_whole_category_goes_to_its_top()
    {
        var components = new List<ComponentEntry> { C("U1", "IC"), C("C5"), C("C6") };

        Assert.Equal(1, ComponentPlacement.InsertionIndex(components, "Capacitor", "C2"));
    }

    [Fact]
    public void A_category_nobody_uses_yet_goes_at_the_end_becoming_the_last_category()
    {
        var components = new List<ComponentEntry> { C("C1"), C("U1", "IC") };

        Assert.Equal(2, ComponentPlacement.InsertionIndex(components, "Resistor", "R1"));
    }

    [Fact]
    public void Categories_match_trimmed_and_case_insensitively()
    {
        var components = new List<ComponentEntry> { C("C1", "capacitor "), C("U1", "IC") };

        Assert.Equal(1, ComponentPlacement.InsertionIndex(components, " Capacitor", "C2"));
    }

    [Fact]
    public void An_empty_sheet_takes_the_first_component_at_the_top()
    {
        Assert.Equal(0, ComponentPlacement.InsertionIndex([], "IC", "U1"));
    }

    // ------------------------------------------------------------------ ReplaceComponent

    [Fact]
    public void An_EDITED_component_stays_exactly_where_it_was()
    {
        // The reported drift: C1 edited after C2 was added moved below C2.
        var components = new List<ComponentEntry> { C("C1"), C("C2"), C("U1", "IC") };

        List<ComponentEntry> result = ComponentPlacement.ReplaceComponent(
            components, "C1", [new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Edited" }]);

        Assert.Equal(["C1", "C2", "U1"], Labels(result));
        Assert.Equal("Edited", result[0].FriendlyName);
    }

    [Fact]
    public void An_edited_component_stays_put_even_if_its_category_changed()
    {
        // Existing rows are never re-sorted behind the user's back - only NEW rows are placed.
        var components = new List<ComponentEntry> { C("C1"), C("C2"), C("U1", "IC") };

        List<ComponentEntry> result = ComponentPlacement.ReplaceComponent(components, "C1", [C("C1", "IC")]);

        Assert.Equal(["C1", "C2", "U1"], Labels(result));
    }

    [Fact]
    public void A_NEW_component_is_placed_by_category_and_label()
    {
        var components = new List<ComponentEntry> { C("C1"), C("C3"), C("U1", "IC") };

        List<ComponentEntry> result = ComponentPlacement.ReplaceComponent(components, "C2", [C("C2")]);

        Assert.Equal(["C1", "C2", "C3", "U1"], Labels(result));
    }

    [Fact]
    public void A_component_with_one_row_per_region_is_kept_together_in_its_own_order()
    {
        var components = new List<ComponentEntry>
        {
            C("U1", "IC"),
            new() { BoardLabel = "U2", Category = "IC", Region = "PAL" },
            new() { BoardLabel = "U2", Category = "IC", Region = "NTSC" },
            C("U3", "IC"),
        };

        List<ComponentEntry> result = ComponentPlacement.ReplaceComponent(
            components,
            "u2",
            [
                new() { BoardLabel = "U2", Category = "IC", Region = "NTSC" },
                new() { BoardLabel = "U2", Category = "IC", Region = "PAL" },
            ]);

        Assert.Equal(["U1", "U2", "U2", "U3"], Labels(result));
        Assert.Equal(["", "NTSC", "PAL", ""], result.Select(component => component.Region));
    }

    [Fact]
    public void A_blank_label_removes_nothing()
    {
        // A blank label identifies no component; treating it as a wildcard would clear every row
        // whose own label happened to be blank.
        var components = new List<ComponentEntry> { C(""), C("C1") };

        List<ComponentEntry> result = ComponentPlacement.ReplaceComponent(components, "", [C("C2")]);

        Assert.Equal(3, result.Count);
    }

    // ------------------------------------------------------------------ Sheets without a category

    [Fact]
    public void A_new_components_files_sit_among_their_neighbours_in_label_order()
    {
        var files = new List<ComponentLocalFileEntry>
        {
            new() { BoardLabel = "U1", Name = "a" },
            new() { BoardLabel = "U10", Name = "b" },
        };

        List<ComponentLocalFileEntry> result = ComponentPlacement.ReplaceRowsForLabel(
            files, "U2", file => file.BoardLabel, [new ComponentLocalFileEntry { BoardLabel = "U2", Name = "new" }]);

        Assert.Equal(["U1", "U2", "U10"], result.Select(file => file.BoardLabel));
    }

    [Fact]
    public void An_edited_components_files_stay_where_they_were()
    {
        var files = new List<ComponentLocalFileEntry>
        {
            new() { BoardLabel = "U9", Name = "a" },
            new() { BoardLabel = "U1", Name = "b" },
        };

        List<ComponentLocalFileEntry> result = ComponentPlacement.ReplaceRowsForLabel(
            files, "U9", file => file.BoardLabel, [new ComponentLocalFileEntry { BoardLabel = "U9", Name = "edited" }]);

        Assert.Equal(["U9", "U1"], result.Select(file => file.BoardLabel));
        Assert.Equal("edited", result[0].Name);
    }

    [Fact]
    public void A_new_REGIONAL_variant_goes_straight_after_its_twin_whatever_its_category()
    {
        // Reported: an NTSC variant of U1 typed in with no category went to the END of the sheet
        // (a blank category is one nobody uses) and looked as though it had disappeared.
        var components = new List<ComponentEntry>
        {
            C("U1", "IC"), C("U2", "IC"), C("C1"), C("C2"),
        };

        Assert.Equal(1, ComponentPlacement.InsertionIndex(components, "", "u1"));
        Assert.Equal(1, ComponentPlacement.InsertionIndex(components, "Capacitor", "U1"));
    }

    [Fact]
    public void A_variant_goes_after_the_LAST_of_its_twins()
    {
        var components = new List<ComponentEntry>
        {
            new() { BoardLabel = "U1", Category = "IC", Region = "PAL" },
            new() { BoardLabel = "U1", Category = "IC", Region = "NTSC" },
            C("U2", "IC"),
        };

        Assert.Equal(2, ComponentPlacement.InsertionIndex(components, "IC", "U1"));
    }
}
