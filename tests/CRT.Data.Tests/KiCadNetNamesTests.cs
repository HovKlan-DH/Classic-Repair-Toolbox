using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// KiCadNetNames - the KiCad net names the board's "Important signals" sheet can use, and which of
// that sheet's rows find no net (owner request, 2026-10-10: the KiCad data window lists what the
// logfile listed as "normalized KiCad net name [...]").
//
// The resolver is the Schematics tab's own rule for its "Important signals" panel (the panel calls
// it), so these pin the panel's behaviour too: a row's net name matches a net's own name or its
// normalised name, trimmed, case-insensitively, across every PCB file.
public sealed class KiCadNetNamesTests
{
    private static KiCadImportantSignalEntry Signal(string displayName, string netName) =>
        new() { DisplayName = displayName, KiCadNetName = netName };

    // ------------------------------------------------------------------ the resolver

    [Fact]
    public void A_hierarchical_net_name_and_its_normalised_name_both_find_the_normalised_one()
    {
        Dictionary<string, string> resolver = KiCadNetNames.Resolver([("/Sheet1/CLK", "CLK")]);

        Assert.Equal("CLK", resolver["/Sheet1/CLK"]);
        Assert.Equal("CLK", resolver["CLK"]);
    }

    [Fact]
    public void The_resolver_trims_and_ignores_case()
    {
        Dictionary<string, string> resolver = KiCadNetNames.Resolver([("  /Sheet1/Clk ", " CLK ")]);

        Assert.Equal("CLK", resolver["/sheet1/clk"]);
        Assert.Equal("CLK", resolver["clk"]);
    }

    // The panel never found a net with no normalised name, and neither may the window.
    [Fact]
    public void A_net_without_a_normalised_name_is_no_net()
    {
        Dictionary<string, string> resolver = KiCadNetNames.Resolver([("GND", null), ("VCC", "  "), (null, "A0")]);

        Assert.False(resolver.ContainsKey("GND"));
        Assert.False(resolver.ContainsKey("VCC"));
        Assert.Equal("A0", resolver["A0"]);
    }

    // ------------------------------------------------------------------ the names listed

    [Fact]
    public void The_names_are_the_normalised_ones_once_each_in_natural_order()
    {
        List<string> names = KiCadNetNames.NormalizedNames(
        [
            ("/A10", "A10", false), ("/A2", "A2", false), ("/a2", "a2", false), ("/A1", "A1", false), ("", "", false), ("X", null, false), ("+5V", "+5V", false),
        ]);

        // A person counts A1, A2, A10 - and a name differing only in case is the same name.
        Assert.Equal(["+5V", "A1", "A2", "A10"], names);
    }

    // ###########################################################################################
    // *** KiCad's "NO NET" IS NOT LISTED (owner decision, 2026-10-10: "yes, hide it"). *** The
    // loader marks it (KiCadNetRef.IsNoNet, KiCadRawProjectLoaderTests): an older .kicad_pcb
    // declares it as (net 0 ""), and a net with no name is named by its number, "0".
    // ###########################################################################################
    [Fact]
    public void KiCads_no_net_is_not_listed()
    {
        Assert.Equal(["GND"], KiCadNetNames.NormalizedNames([("0", "0", true), ("GND", "GND", false)]));
    }

    // ###########################################################################################
    // *** A REAL NET NAMED "0" IS LISTED (code review, 2026-10-10). *** KiCad 9 refers to every net
    // by its name alone - (net "0") for SPICE's ground - so it arrives as id "0", name "0", exactly
    // like "no net" did. Told apart by id and name, it was hidden too; the loader's mark is what
    // tells them apart now.
    // ###########################################################################################
    [Fact]
    public void A_KiCad_9_net_named_0_is_listed()
    {
        Assert.Equal(
            ["0", "+5V", "GND"],
            KiCadNetNames.NormalizedNames([("GND", "GND", false), ("+5V", "+5V", false), ("0", "0", false)]));
    }

    // Hidden from the list, but still a net: a row naming "0" finds it, as the panel always let it.
    [Fact]
    public void Net_0_is_not_listed_but_a_row_naming_it_still_finds_it()
    {
        KiCadNetNameReport report = KiCadNetNames.Build([("0", "0", true), ("GND", "GND", false)], [Signal("Unconnected", "0")]);

        Assert.Equal(["GND"], report.NetNames);
        Assert.Empty(report.UnresolvedSignals);
    }

    // ------------------------------------------------------------------ the report

    [Fact]
    public void A_row_naming_a_known_net_marks_it_used()
    {
        KiCadNetNameReport report = KiCadNetNames.Build(
            [("/Sheet1/CLK", "CLK", false), ("GND", "GND", false)],
            [Signal("Clock", "/Sheet1/CLK")]);

        Assert.True(report.IsUsed("CLK"));
        Assert.True(report.IsUsed("clk"));
        Assert.False(report.IsUsed("GND"));
        Assert.Empty(report.UnresolvedSignals);
    }

    // The silent failure the window now says out loud: the panel drops such a row with no word.
    [Fact]
    public void A_row_naming_a_net_the_KiCad_data_does_not_have_is_unresolved_in_sheet_order()
    {
        KiCadNetNameReport report = KiCadNetNames.Build(
            [("PHI2", "PHI2", false)],
            [Signal("Phase 2", "PH12"), Signal("Phi", " PHI2 "), Signal("Reset", "RES")]);

        Assert.Equal(
            ["Phase 2: PH12", "Reset: RES"],
            report.UnresolvedSignals.Select(signal => $"{signal.DisplayName}: {signal.KiCadNetName}"));
        Assert.True(report.IsUsed("PHI2"));
    }

    // Half a row lights up nothing either way, and the panel skips it; the table's checks point at it.
    [Fact]
    public void A_row_missing_its_display_name_or_net_name_is_skipped_not_reported()
    {
        KiCadNetNameReport report = KiCadNetNames.Build(
            [("CLK", "CLK", false)],
            [Signal("", "NOPE"), Signal("Clock", "  "), Signal("  ", "CLK")]);

        Assert.Empty(report.UnresolvedSignals);
        Assert.Empty(report.UsedNetNames);
    }

    [Fact]
    public void No_nets_and_no_rows_is_an_empty_report()
    {
        KiCadNetNameReport report = KiCadNetNames.Build([], []);

        Assert.Empty(report.NetNames);
        Assert.Empty(report.UsedNetNames);
        Assert.Empty(report.UnresolvedSignals);
    }

    // ---- the choices an important signal picks from (slice 4a) ---------------------------------

    // The names to pick are the report's; a net is found by the panel's own rule - its own name or its
    // normalised one, trimmed, capitals ignored - and net 0 is found though never offered.
    [Fact]
    public void The_choices_offer_the_normalised_names_and_find_a_net_as_the_panel_does()
    {
        KiCadNetChoices choices = KiCadNetNames.Choices([("/Sheet1/CLK", "CLK", false), ("A10", "A10", false), ("A2", "A2", false), ("", "0", true)]);

        Assert.Equal(["A2", "A10", "CLK"], choices.Names);
        Assert.True(choices.HasNets);
        Assert.True(choices.Finds(" clk "));
        Assert.True(choices.Finds("/Sheet1/CLK"));
        Assert.True(choices.Finds("0"));
        Assert.False(choices.Finds("NOPE"));

        // Half a row is the table's to point at, not the picker's.
        Assert.True(choices.Finds("  "));
    }

    // A board without KiCad data offers nothing and calls no typed net missing.
    [Fact]
    public void Without_KiCad_data_every_net_may_be_typed()
    {
        Assert.False(KiCadNetChoices.None.HasNets);
        Assert.Empty(KiCadNetChoices.None.Names);
        Assert.True(KiCadNetChoices.None.Finds("ANYTHING"));
        Assert.True(KiCadNetNames.Choices([]).Finds("ANYTHING"));
    }
}
