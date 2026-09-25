using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for BoardDraftNaturalKeys - the row-identity rule BoardDraftApplier
// merges by, and (per NewContributeStrategy.md's "Retiring the UUIDs") what a future server-side
// diff will pair rows on too. The one thing worth pinning here is that keys are built from
// TRIMMED, case-preserved values with a separator that cannot appear in ordinary board data, so
// two rows that a human would call "the same row" always produce the same key even with stray
// whitespace, while genuinely different values never collide by accident.
public sealed class BoardDraftNaturalKeysTests
{
    [Fact]
    public void ForComponent_trims_surrounding_whitespace()
    {
        Assert.Equal(BoardDraftNaturalKeys.ForComponent("U8"), BoardDraftNaturalKeys.ForComponent("  U8  "));
    }

    [Fact]
    public void ForComponent_treats_null_as_empty()
    {
        Assert.Equal(BoardDraftNaturalKeys.ForComponent(string.Empty), BoardDraftNaturalKeys.ForComponent(null!));
    }

    [Fact]
    public void ForComponentImage_is_sensitive_to_every_one_of_its_four_parts()
    {
        string baseline = BoardDraftNaturalKeys.ForComponentImage("U8", "CPU", "1", "Clock");

        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U9", "CPU", "1", "Clock"));
        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U8", "PLA", "1", "Clock"));
        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U8", "CPU", "2", "Clock"));
        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U8", "CPU", "1", "Reset"));
    }

    [Fact]
    public void ForComponentImage_does_not_let_field_boundaries_shift_and_collide()
    {
        // Without a separator that cannot appear in the fields themselves, ("U8", "CPU1") and
        // ("U8C", "PU1") could concatenate to the same string. The separator is what stops that.
        string a = BoardDraftNaturalKeys.ForComponentImage("U8", "CPU1", string.Empty, string.Empty);
        string b = BoardDraftNaturalKeys.ForComponentImage("U8C", "PU1", string.Empty, string.Empty);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void ForComponentHighlight_distinguishes_the_same_label_on_different_schematics()
    {
        string mainBoard = BoardDraftNaturalKeys.ForComponentHighlight("Main board", "J1");
        string daughterBoard = BoardDraftNaturalKeys.ForComponentHighlight("Daughterboard", "J1");

        Assert.NotEqual(mainBoard, daughterBoard);
    }

    [Fact]
    public void ForCredit_distinguishes_two_different_people_in_the_same_category_and_subcategory()
    {
        string alex = BoardDraftNaturalKeys.ForCredit("Schematics", "Tracing", "Alex");
        string sam = BoardDraftNaturalKeys.ForCredit("Schematics", "Tracing", "Sam");

        Assert.NotEqual(alex, sam);
    }

    // ------------------------------------------------------------------ Component region (2026-09-24)

    [Fact]
    public void A_component_WITHOUT_a_region_keys_exactly_as_it_always_did()
    {
        // Nearly every component has no region; their keys must not move.
        Assert.Equal("U8", BoardDraftNaturalKeys.ForComponent("U8"));
        Assert.Equal("U8", BoardDraftNaturalKeys.ForComponent(" U8 ", "  "));
        Assert.Equal("U8", BoardDraftNaturalKeys.ForRow(new ComponentEntry { BoardLabel = "U8" }));
    }

    [Fact]
    public void The_same_label_in_two_REGIONS_is_two_rows()
    {
        // A regionalised component is one row per region. Keyed on the label alone they were one
        // row to every comparison, and an added NTSC variant was invisible.
        string pal = BoardDraftNaturalKeys.ForRow(new ComponentEntry { BoardLabel = "U8", Region = "PAL" });
        string ntsc = BoardDraftNaturalKeys.ForRow(new ComponentEntry { BoardLabel = "U8", Region = "NTSC" });

        Assert.NotEqual(pal, ntsc);
        Assert.NotEqual("U8", pal);
        Assert.Equal("U8", pal.Split(BoardDraftNaturalKeys.Separator)[0]);
    }
}
