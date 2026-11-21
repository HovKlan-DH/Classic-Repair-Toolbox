using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentVocabulary - the words CRT 4.0's new component data is written in (step 2,
// 2026-10-25). The words ARE the data (a board file stores "Datasheet", not a code), so what a
// check calls known is pinned here, with the rules every value of a board follows: empty is fine,
// capitals and surrounding spaces do not matter.
// ###########################################################################################
public sealed class ComponentVocabularyTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Drop-in replacement")]
    [InlineData("  needs MODIFICATION ")]
    public void Empty_or_a_known_word_in_any_capitals_is_known(string? value)
    {
        Assert.True(ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.ReplacementKinds, value));
    }

    [Theory]
    [InlineData("My cousin")]
    [InlineData("Drop-in")]
    public void Anything_else_is_not(string value)
    {
        Assert.False(ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.ReplacementKinds, value));
    }

    // The plan's words (2026-10-24), each list in the order the component pane offers it - with
    // "Other" last wherever it is offered.
    [Fact]
    public void The_lists_hold_the_plans_words()
    {
        Assert.Equal(["Drop-in replacement", "Needs modification", "Improvement"], ComponentVocabulary.ReplacementKinds);
        // "Unknown" last (owner request, 2026-11-01).
        Assert.Equal("Unknown", ComponentVocabulary.PinDirections[^1]);
        Assert.Equal("Other", ComponentVocabulary.Packages[^1]);
        Assert.Contains("DIP-40", ComponentVocabulary.Packages);
        Assert.Contains("TO-220", ComponentVocabulary.Packages);
    }

    // Plain ASCII and no word twice - they are shown in drop-down lists and written into files.
    [Fact]
    public void Every_word_is_plain_ASCII_and_listed_once()
    {
        foreach (IReadOnlyList<string> words in new[]
                 {
                     ComponentVocabulary.PinDirections,
                     ComponentVocabulary.ReplacementKinds,
                     ComponentVocabulary.Packages,
                 })
        {
            Assert.All(words, word => Assert.True(word.All(character => character is >= ' ' and <= '~'), word));
            Assert.Equal(words.Count, words.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }
    }

    // ###########################################################################################
    // A case's pin count, by its name - what the pane numbers the pins and draws the case from. A
    // chip's case says it, however it is typed (ChipCaseOf, below); TO-92 and TO-220 have three; a
    // SIP's, "Other"'s, a SOT-23's and anything else's cannot be told.
    // ###########################################################################################
    [Theory]
    [InlineData("DIP-40", 40)]
    [InlineData("DIP-24 wide", 24)]
    [InlineData(" dip 28 ", 28)]
    [InlineData("DIP16", 16)]
    [InlineData("SOIC-14", 14)]
    [InlineData("PLCC-44", 44)]
    [InlineData("QFP-44", 44)]
    [InlineData("TO-92", 3)]
    [InlineData("to-220", 3)]
    [InlineData("SIP", null)]
    [InlineData("Other", null)]
    [InlineData("DIP", null)]
    [InlineData("DIP-7", null)]
    [InlineData("DIP-0", null)]
    [InlineData("SOT-23", null)]
    [InlineData("", null)]
    [InlineData(null, null)]
    public void A_cases_pin_count_is_read_off_its_name(string? package, int? count)
    {
        Assert.Equal(count, ComponentVocabulary.PinCountOf(package));
    }

    // ###########################################################################################
    // THE CASES A CHIP IS DRAWN IN (owner request, 2026-11-18: the Amiga "has a lot of SMD components"): two rows - a DIP
    // and the surface-mounted cases numbered as one, whatever follows the count - and four sides, square: a PLCC with an
    // odd number of pins a side (the rectangular PLCC-32 has eight and nine), whatever follows its count (code review,
    // 2026-11-18: "PLCC-84 socket", the A500's Fat Agnus, was none), a QFP a quarter of its pins a side, only as the plain
    // name ("QFP-100 14x20mm" is rectangular). A SOT-23 is three pins of no row.
    // ###########################################################################################
    [Theory]
    [InlineData("PLCC-84 socket", ChipCaseKind.Plcc, 84)]
    [InlineData("DIP-40", ChipCaseKind.TwoRows, 40)]
    [InlineData("DIP-24 wide", ChipCaseKind.TwoRows, 24)]
    [InlineData("SOIC-14", ChipCaseKind.TwoRows, 14)]
    [InlineData("soic16", ChipCaseKind.TwoRows, 16)]
    [InlineData("SO-8", ChipCaseKind.TwoRows, 8)]
    [InlineData("SOP-28", ChipCaseKind.TwoRows, 28)]
    [InlineData("SSOP-20", ChipCaseKind.TwoRows, 20)]
    [InlineData("TSSOP-16", ChipCaseKind.TwoRows, 16)]
    [InlineData("SOIC-16 wide", ChipCaseKind.TwoRows, 16)]
    [InlineData("PLCC-20", ChipCaseKind.Plcc, 20)]
    [InlineData("PLCC-84", ChipCaseKind.Plcc, 84)]
    [InlineData("plcc 68", ChipCaseKind.Plcc, 68)]
    [InlineData("QFP-44", ChipCaseKind.Qfp, 44)]
    [InlineData("LQFP-64", ChipCaseKind.Qfp, 64)]
    [InlineData("TQFP-100", ChipCaseKind.Qfp, 100)]
    [InlineData("PQFP-208", ChipCaseKind.Qfp, 208)]
    public void A_chips_case_is_read_off_its_name(string package, ChipCaseKind kind, int pins)
    {
        Assert.Equal(new ChipCase(kind, pins), ComponentVocabulary.ChipCaseOf(package));
    }

    [Theory]
    [InlineData("PLCC-32")]
    [InlineData("PLCC-32 socket")]
    [InlineData("PLCC-18")]
    [InlineData("QFP-42")]
    [InlineData("QFP-100 14x20mm")]
    [InlineData("SOIC-7")]
    [InlineData("SOT-23")]
    [InlineData("SOT-223")]
    [InlineData("TO-92")]
    [InlineData("SIP")]
    [InlineData("Other")]
    [InlineData("QFN-32")]
    [InlineData("")]
    [InlineData(null)]
    public void What_is_no_chips_case_CRT_can_draw_is_none(string? package)
    {
        Assert.Null(ComponentVocabulary.ChipCaseOf(package));
    }

    // Every chip's case the pane offers can be counted and drawn, and two names of one case are equal.
    [Fact]
    public void Every_offered_chip_case_can_be_counted_and_two_names_of_one_case_are_equal()
    {
        foreach (string package in ComponentVocabulary.Packages.Where(word => word.StartsWith("DIP", StringComparison.Ordinal) ||
                                                                              word.StartsWith("SOIC", StringComparison.Ordinal) ||
                                                                              word.StartsWith("PLCC", StringComparison.Ordinal) ||
                                                                              word.StartsWith("QFP", StringComparison.Ordinal)))
        {
            Assert.NotNull(ComponentVocabulary.ChipCaseOf(package));
            Assert.Equal(ComponentVocabulary.ChipCaseOf(package)!.Pins, ComponentVocabulary.PinCountOf(package));
        }

        Assert.Equal(ComponentVocabulary.ChipCaseOf("DIP-14"), ComponentVocabulary.ChipCaseOf("SOIC-14"));
        Assert.Equal(ComponentVocabulary.ChipCaseOf("QFP-44"), ComponentVocabulary.ChipCaseOf("LQFP-44"));
        Assert.NotEqual(ComponentVocabulary.ChipCaseOf("QFP-44"), ComponentVocabulary.ChipCaseOf("PLCC-44"));
        Assert.Equal(21, ComponentVocabulary.ChipCaseOf("PLCC-84")!.PerSide);
        Assert.Equal(20, ComponentVocabulary.ChipCaseOf("DIP-40")!.PerSide);
    }

    // ###########################################################################################
    // T/DIV and V/DIV as the scope baseline editor offers them (CRT 4.0 step 2, slice 3b): the
    // 1-2-5 steps every scope uses, spelt as the converted data spells them - "50us", "1ms",
    // "500ns", "1V", "500mV" - so a value picked reads like the 13,862 baselines already published.
    // ###########################################################################################
    [Fact]
    public void The_time_and_volt_divisions_are_the_scope_steps_spelt_as_the_data_spells_them()
    {
        Assert.Contains("50us", ComponentVocabulary.TimeDivisions);
        Assert.Contains("1ms", ComponentVocabulary.TimeDivisions);
        Assert.Contains("500ns", ComponentVocabulary.TimeDivisions);
        Assert.Equal("1ns", ComponentVocabulary.TimeDivisions[0]);
        Assert.Equal("10s", ComponentVocabulary.TimeDivisions[^1]);
        Assert.Equal(31, ComponentVocabulary.TimeDivisions.Count);

        Assert.Equal(["1mV", "2mV", "5mV", "10mV", "20mV", "50mV", "100mV", "200mV", "500mV", "1V", "2V", "5V", "10V"], ComponentVocabulary.VoltDivisions);
        Assert.Equal(ComponentVocabulary.TimeDivisions.Count, ComponentVocabulary.TimeDivisions.Distinct().Count());
    }

    // ---- a pin against its case (owner request, 2026-11-01) --------------------------------------

    [Theory]
    [InlineData("48", 48)]
    [InlineData(" 2 ", 2)]
    [InlineData("2 [64K]", 2)]
    [InlineData("2[64K]", 2)]
    public void A_pins_number_is_what_its_text_starts_with(string pin, int number)
    {
        Assert.Equal(number, ComponentVocabulary.PinNumberOf(pin));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("123456789012")]
    // Digits running straight into letters are a name: a regulator's "12V" is no pin 12, and "1A" is one
    // of a 7400's named pins (code review, 2026-11-01).
    [InlineData("12V")]
    [InlineData("5VDC")]
    [InlineData("1A")]
    public void A_pin_starting_with_no_number_has_none(string? pin)
    {
        Assert.Null(ComponentVocabulary.PinNumberOf(pin));
    }

    [Theory]
    [InlineData("DIP-24", "24", true)]
    [InlineData("DIP-24", "48", false)]
    [InlineData("DIP-24", "0", false)]
    [InlineData("DIP-24 wide", "25", false)]
    [InlineData("TO-92", "3", true)]
    [InlineData("TO-92", "4", false)]
    [InlineData("", "48", true)]
    [InlineData("SIP", "48", true)]
    [InlineData("DIP-24", "A1", true)]
    [InlineData("DIP-24", "", true)]
    [InlineData("TO-220", "12V", true)]
    public void A_pin_is_off_its_case_only_when_both_are_known(string package, string pin, bool on)
    {
        Assert.Equal(on, ComponentVocabulary.IsOnPackage(package, pin));
    }

    // ###########################################################################################
    // A CRYSTAL OSCILLATOR'S CAN (owner request, 2026-11-18 - the Amiga 500's X1 and the Amiga 1000's OSC1, and the Open128's
    // Y1): on a DIP-14's places with a pin at each corner, 1, 7, 8 and 14, or on a DIP-8's with 1, 4, 5 and 8 - as typed, or
    // as KiCad names its footprint. No chip's case, so nothing is counted: the pane offers no "Add the 14 pins".
    // ###########################################################################################
    [Theory]
    [InlineData("Oscillator DIP-14", 14, new[] { 1, 7, 8, 14 })]
    [InlineData(" oscillator_dip-14 ", 14, new[] { 1, 7, 8, 14 })]
    [InlineData("Oscillator DIP14", 14, new[] { 1, 7, 8, 14 })]
    [InlineData("Oscillator DIP-8", 8, new[] { 1, 4, 5, 8 })]
    [InlineData("OSCILLATOR-DIP-8", 8, new[] { 1, 4, 5, 8 })]
    public void An_oscillators_can_has_a_pin_at_each_corner_of_its_DIPs_places(string package, int positions, int[] pins)
    {
        OscillatorCan can = Assert.IsType<OscillatorCan>(ComponentVocabulary.OscillatorOf(package));

        Assert.Equal(positions, can.Positions);
        Assert.Equal(pins, can.Pins);
        Assert.Null(ComponentVocabulary.ChipCaseOf(package));
        Assert.Null(ComponentVocabulary.PinCountOf(package));
    }

    [Theory]
    [InlineData("Oscillator DIP-16")]
    [InlineData("Oscillator")]
    [InlineData("Oscillator SMD-4")]
    [InlineData("DIP-14")]
    [InlineData("HC-49")]
    [InlineData("")]
    [InlineData(null)]
    public void What_is_no_oscillators_can_is_none(string? package)
    {
        Assert.Null(ComponentVocabulary.OscillatorOf(package));
    }

    // The pane offers both, and each is one CRT reads as a can.
    [Fact]
    public void The_pane_offers_both_oscillator_cans()
    {
        Assert.NotNull(ComponentVocabulary.OscillatorOf(Assert.Single(ComponentVocabulary.Packages, word => word == "Oscillator DIP-14")));
        Assert.NotNull(ComponentVocabulary.OscillatorOf(Assert.Single(ComponentVocabulary.Packages, word => word == "Oscillator DIP-8")));
    }

    // A pin of a can is on it only at a corner - pin 2 of a full-size can stands on a place with no pin - and a named pin says
    // nothing, as on any case.
    [Theory]
    [InlineData("Oscillator DIP-14", "1", true)]
    [InlineData("Oscillator DIP-14", "8", true)]
    [InlineData("Oscillator DIP-14", "14", true)]
    [InlineData("Oscillator DIP-14", "2", false)]
    [InlineData("Oscillator DIP-14", "15", false)]
    [InlineData("Oscillator DIP-8", "5", true)]
    [InlineData("Oscillator DIP-8", "7", false)]
    [InlineData("Oscillator DIP-14", "OUT", true)]
    public void A_pin_of_an_oscillators_can_is_on_it_only_at_a_corner(string package, string pin, bool on)
    {
        Assert.Equal(on, ComponentVocabulary.IsOnPackage(package, pin));
    }

    // "Analog input" and "Analog output" went again (owner request, 2026-11-02: "it can be covered by
    // the existing two choices") - a board already holding one keeps it, shown as it is. In the owner's order
    // (2026-11-09).
    [Fact]
    public void A_pin_carrying_a_voltage_level_is_an_input_or_an_output()
    {
        Assert.Equal(["Not connected", "Input", "Output", "Input/output", "Ground", "Power", "Other", "Unknown"], ComponentVocabulary.PinDirections);
    }

    // One rule for "the same pin" - the popup lighting the pin of its image, a new baseline taking its
    // signal name. "2 [64K]" is pin 2, "1A" a name of its own.
    [Theory]
    [InlineData("2", "2", true)]
    [InlineData(" a1 ", "A1", true)]
    [InlineData("2 [64K]", "2", true)]
    [InlineData("1A", "1", false)]
    [InlineData("2", "3", false)]
    [InlineData("", "", false)]
    [InlineData(null, "2", false)]
    public void Two_texts_name_one_pin_by_their_words_or_their_number(string? first, string? second, bool same)
    {
        Assert.Equal(same, ComponentVocabulary.IsSamePin(first, second));
    }

    // Owner request, 2026-11-11: a package typed in the table ("Diskette") is offered in the pane's drop-down too, and the
    // same for a pin's direction - after the usual words, alphabetically, each once whatever its capitals, in the board's
    // first spelling. "Unknown" stays the last direction.
    [Fact]
    public void A_board_offers_its_own_packages_and_directions_after_the_usual_ones()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "D1", Package = "Diskette" });
        board.Components.Add(new ComponentEntry { BoardLabel = "D2", Package = " diskette " });
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Package = "dip-40" });
        board.Components.Add(new ComponentEntry { BoardLabel = "X1", Package = "Can" });
        board.Components.Add(new ComponentEntry { BoardLabel = "X2", Package = "" });
        board.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Direction = "Open collector" });
        board.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U1", Pin = "2", Direction = "input" });
        board.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U1", Pin = "3", Direction = "Bidirectional" });

        Assert.Equal([.. ComponentVocabulary.Packages, "Can", "Diskette"], ComponentVocabulary.PackagesOf(board));
        Assert.Equal(
            ["Not connected", "Input", "Output", "Input/output", "Ground", "Power", "Other", "Bidirectional", "Open collector", "Unknown"],
            ComponentVocabulary.PinDirectionsOf(board));
    }

    [Fact]
    public void A_board_using_only_the_usual_words_offers_just_them()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Package = "DIP-40" });

        Assert.Same(ComponentVocabulary.Packages, ComponentVocabulary.PackagesOf(board));
        Assert.Same(ComponentVocabulary.PinDirections, ComponentVocabulary.PinDirectionsOf(board));
        Assert.Same(ComponentVocabulary.Packages, ComponentVocabulary.PackagesOf(null));
    }

    // ###########################################################################################
    // A baseline's "Expected reading" offers the readings the board's baselines already give (owner request, 2026-11-18: "if I
    // already have typed "High" previously for this board, then when typing "h" it will show the "High"") - every component's,
    // every region's, alphabetically, each once whatever its capitals, in the board's first spelling; a picture's reading too,
    // as nothing tells a reading by what its row is.
    // ###########################################################################################
    [Fact]
    public void A_baselines_expected_reading_offers_the_readings_the_board_already_gives()
    {
        var board = new BoardData();
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U8", Pin = "1", ExpectedOscilloscopeReading = "Pulsing" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U9", Region = "NTSC", Pin = "2", ExpectedOscilloscopeReading = " High " });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U9", Pin = "3", ExpectedOscilloscopeReading = "high" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U9", Pin = "4", ExpectedOscilloscopeReading = "" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U1", Pin = "5", ExpectedOscilloscopeReading = "Low (or floating?)" });

        Assert.Equal(["High", "Low (or floating?)", "Pulsing"], ComponentVocabulary.ReadingsOf(board));
        Assert.Empty(ComponentVocabulary.ReadingsOf(new BoardData()));
        Assert.Empty(ComponentVocabulary.ReadingsOf(null));
    }
}
