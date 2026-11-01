using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ImportantSignalSuggestions - important signals suggested from the board's KiCad net names (CRT 4.0
// step 2: "pick nets from a searchable KiCad list, with suggestions"): a net whose WHOLE name says
// what it is - supply, ground, clock, control line, bus line - named as the live boards name theirs,
// and a net a row already has never offered again.
// ###########################################################################################
public sealed class ImportantSignalSuggestionsTests
{
    // ###########################################################################################
    // EVERY NET IS TO BE SEEN (owner request, 2026-11-01: "as the last data in here I would like to see
    // all other nets not already used ... so all nets will be visible in here"): the nets neither a row
    // nor a suggestion has. A row, a suggestion and an other net together are the board's nets.
    // ###########################################################################################
    [Fact]
    public void The_other_nets_are_every_net_neither_a_row_nor_a_suggestion_has()
    {
        KiCadNetChoices choices = KiCadNetNames.Choices([("/RESET", "RESET", false), ("PHI2", "PHI2", false), ("N$12", "N$12", false), ("U7_3", "U7_3", false)]);
        KiCadImportantSignalEntry[] rows = [new() { DisplayName = "My clock", KiCadNetName = "PHI2" }];

        List<string> others = ImportantSignalSuggestions.OtherNets(choices, rows);

        Assert.Equal(["N$12", "U7_3"], others);
        Assert.Equal(
            choices.Names.Order(StringComparer.Ordinal),
            rows.Select(row => row.KiCadNetName)
                .Concat(ImportantSignalSuggestions.For(choices, rows).Select(row => row.KiCadNetName))
                .Concat(others)
                .Order(StringComparer.Ordinal));
    }

    [Theory]
    [InlineData("+5V", "+5VDC")]
    [InlineData("5V", "+5VDC")]
    [InlineData("+12V", "+12VDC")]
    [InlineData("+9V", "+9VDC")]
    [InlineData("+9V_UNREG", "+9VDC (unregulated)")]
    [InlineData("CAN+5V", "+5VDC CAN")]
    [InlineData("9VAC~", "9VAC")]
    [InlineData("'9VAC~", "9VAC")]
    [InlineData("9V_AC1", "9VAC")]
    [InlineData("9VAC2'", "9VAC")]
    [InlineData("GND", "Ground")]
    [InlineData("GND'", "Ground")]
    [InlineData("Earth", "Ground")]
    [InlineData("0V", "Ground")]
    [InlineData("PHI2", "Clock; PHI2")]
    [InlineData("ΦDOT", "Clock; DOT Clock")]
    [InlineData("~{ΦDOT}", "Clock; DOT Clock")]
    [InlineData("DOT_CLOCK", "Clock; DOT Clock")]
    [InlineData("ΦCOLOR", "Clock; PHI Color")]
    [InlineData("COLOR_CLOCK", "Clock; PHI Color")]
    [InlineData("1MHZ", "Clock; 1 MHz")]
    [InlineData("~{RESET}", "RESET")]
    [InlineData("~{RESET'}", "RESET")]
    [InlineData("~{IRQ}", "IRQ")]
    [InlineData("R/~{W}", "R/W")]
    [InlineData("R{slash}~{W}", "R/W")]
    [InlineData("R~{W}", "R/W")]
    [InlineData("~{AEC}", "AEC")]
    [InlineData("~{NMI}", "NMI")]
    [InlineData("A0", "A0")]
    [InlineData("A15", "A15")]
    [InlineData("D7", "D7")]
    public void A_net_whose_name_says_what_it_is_is_suggested_under_the_live_boards_name(string net, string displayName) =>
        Assert.Equal(displayName, ImportantSignalSuggestions.NameOf(net)?.DisplayName);

    // Only the WHOLE name counts - a net that merely contains a known name is no suggestion.
    [Theory]
    [InlineData("IN_GND")]
    [InlineData("ATN_OUT")]
    [InlineData("A16")]
    [InlineData("D8")]
    [InlineData("MA0")]
    [InlineData("Net-(U8-Pad3)")]
    [InlineData("")]
    [InlineData(null)]
    public void A_net_whose_name_says_nothing_known_is_not_suggested(string? net) =>
        Assert.Null(ImportantSignalSuggestions.NameOf(net));

    private static KiCadNetChoices Choices(params string[] names) =>
        KiCadNetNames.Choices(names.Select(name => ((string?)name, (string?)name, false)));

    // ###########################################################################################
    // The rows suggested: in the catalogue's order (supplies, ground, clocks, control lines, buses),
    // a name's nets in natural order - and never a net a row already has, by the panel's own rule.
    // ###########################################################################################
    [Fact]
    public void The_suggestions_come_in_the_catalogues_order_and_leave_out_the_nets_rows_have()
    {
        KiCadNetChoices choices = Choices("D1", "A10", "A2", "~{RESET}", "PHI2", "GND", "9VAC2", "9VAC1", "+5V", "MA0", "D0");

        List<KiCadImportantSignalEntry> existing =
        [
            new() { DisplayName = "Data", KiCadNetName = " d0 " },
            new() { DisplayName = "Typo", KiCadNetName = "NOT A NET" },
        ];

        Assert.Equal(
            [
                "+5VDC:+5V",
                "9VAC:9VAC1",
                "9VAC:9VAC2",
                "Ground:GND",
                "Clock; PHI2:PHI2",
                "RESET:~{RESET}",
                "A2:A2",
                "A10:A10",
                "D1:D1",
            ],
            ImportantSignalSuggestions.For(choices, existing).Select(row => $"{row.DisplayName}:{row.KiCadNetName}"));
    }

    // ###########################################################################################
    // A NEW SIGNAL PICKS ITS NET (owner request, 2026-10-29 - the row's net is shown as text, never
    // typed): every net of the board in the picker's order but the ones a row has, by the panel's
    // rule (" d0 " is D0); the name it starts with is its suggestion's, else the net as CRT spells it.
    // ###########################################################################################
    [Fact]
    public void A_new_signal_picks_from_every_net_no_row_has_and_starts_with_its_suggested_name()
    {
        KiCadNetChoices choices = Choices("D0", "D1", "Net-(U8-Pad3)", "~{RESET}");

        Assert.Equal(
            ["D1", "Net-(U8-Pad3)", "~{RESET}"],
            ImportantSignalSuggestions.NetsToAdd(choices, [new() { DisplayName = "Data", KiCadNetName = " d0 " }]));
        Assert.Empty(ImportantSignalSuggestions.NetsToAdd(KiCadNetChoices.None, []));

        Assert.Equal("RESET", ImportantSignalSuggestions.StartingNameOf("~{RESET}"));
        Assert.Equal("Net-(U8-Pad3)", ImportantSignalSuggestions.StartingNameOf("Net-(U8-Pad3)"));
        Assert.Equal("~SYNC", ImportantSignalSuggestions.StartingNameOf("~{SYNC}"));
    }

    [Fact]
    public void A_board_without_KiCad_nets_gets_no_suggestions() =>
        Assert.Empty(ImportantSignalSuggestions.For(KiCadNetChoices.None, []));

    [Fact]
    public void A_net_is_found_by_its_own_or_its_normalised_name()
    {
        KiCadNetChoices choices = KiCadNetNames.Choices([("/CPU/PHI2", "PHI2", false)]);

        Assert.Equal("PHI2", choices.NormalizedOf(" /CPU/PHI2 "));
        Assert.Equal("PHI2", choices.NormalizedOf("phi2"));
        Assert.Null(choices.NormalizedOf("PHI0"));
        Assert.Null(choices.NormalizedOf(" "));
    }
}
