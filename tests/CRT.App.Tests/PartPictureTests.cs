using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which picture of the part a component gets (PartPicture; owner request, 2026-11-09: "Actually all of the current
// components for all boards - can you draw them"). The shape comes from the package, electrolytic or not from the friendly
// name and the value, the polarity from the pins' names or KiCad's convention - and whatever the picture could show wrong
// draws nothing at all.
// ###########################################################################################
public sealed class PartPictureTests
{
    private static ComponentEntry Part(string category, string package, string friendly = "", string value = "", string label = "C1", string description = "") =>
        new() { BoardLabel = label, Category = category, Package = package, FriendlyName = friendly, TechnicalNameOrValue = value, Description = description };

    // Pins 1, 2, ... named as given ("" for a pin with no name).
    private static List<ComponentPinEntry> Pins(params string[] names) =>
        names.Select((name, index) => new ComponentPinEntry { BoardLabel = "C1", Pin = (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), Name = name }).ToList();

    private static List<(int Number, string Name)> LeadsOf(PartPicture.Picture? picture) =>
        picture!.Leads.Select(lead => (lead.Number, lead.Name)).ToList();

    // ---- capacitors ---------------------------------------------------------------------------

    // The data's radial electrolytics: pins with no names, so KiCad's convention says pin 1 is plus - the stripe is pin 2's.
    [Fact]
    public void A_radial_electrolytic_has_its_minus_stripe_on_pin_2_and_its_value_printed()
    {
        PartPicture.Picture? picture = PartPicture.Of(Part("Capacitor", "Radial", "Electrolytic Radial", "10uF 16V"), Pins("", ""));

        Assert.Equal(PartPicture.Shape.RadialElectrolytic, picture!.Shape);
        Assert.Equal([(1, "+"), (2, "-")], LeadsOf(picture));
        Assert.Equal(1, picture.MarkedLead);
        Assert.Equal("10µF\n16V", picture.Marking);
    }

    // A pin named for its polarity says which lead is which - "-" on pin 1 moves the stripe there, and either name alone
    // is enough.
    [Theory]
    [InlineData("-", "+", 0)]
    [InlineData("-", "", 0)]
    [InlineData("", "+", 0)]
    [InlineData("Positive", "Negative", 1)]
    [InlineData("+", "", 1)]
    [InlineData("neg", "pos", 0)]
    public void The_pins_names_say_which_lead_is_minus(string first, string second, int minusLead)
    {
        PartPicture.Picture? picture = PartPicture.Of(Part("Capacitor", "Axial", "Electrolytic Axial", "1000uF 25V"), Pins(first, second));

        string[] typed = [first, second];

        Assert.Equal(PartPicture.Shape.AxialElectrolytic, picture!.Shape);
        Assert.Equal(minusLead, picture.MarkedLead);
        Assert.Equal(typed[minusLead].Length > 0 ? typed[minusLead] : "-", picture.Leads[minusLead].Name);
        Assert.Equal(typed[1 - minusLead].Length > 0 ? typed[1 - minusLead] : "+", picture.Leads[1 - minusLead].Name);
    }

    // A lead with no name of its own is given its role; one typed keeps its words.
    [Fact]
    public void A_lead_with_no_name_is_given_its_role_and_a_typed_name_is_kept()
    {
        PartPicture.Picture? picture = PartPicture.Of(Part("Capacitor", "Radial", "Electrolytic Radial", "10uF"), Pins("Negative", ""));

        Assert.Equal([(1, "Negative"), (2, "+")], LeadsOf(picture));
    }

    // Names that contradict each other leave the minus side unknown - and a wrong minus side is worse than no picture.
    [Theory]
    [InlineData("+", "+")]
    [InlineData("-", "-")]
    [InlineData("+ ", "Positive")]
    public void Names_that_contradict_each_other_draw_nothing(string first, string second)
    {
        Assert.Null(PartPicture.Of(Part("Capacitor", "Radial", "Electrolytic Radial", "10uF"), Pins(first, second)));
    }

    // The data's C128DCR has a ceramic one whose DESCRIPTION says "not 1uF electrolytic listed on parts list": the kind is
    // read from the friendly name and the value, never the description.
    [Fact]
    public void A_description_mentioning_electrolytic_makes_no_ceramic_capacitor_polarised()
    {
        PartPicture.Picture? picture = PartPicture.Of(
            Part("Capacitor", "Radial", "Ceramic", ".1 µF", description: ".1uF ceramic 104Z verified -  not 1uF electrolytic listed on parts list."),
            Pins("", ""));

        Assert.Equal(PartPicture.Shape.RadialCapacitor, picture!.Shape);
        Assert.Null(picture.MarkedLead);
        Assert.Equal([(1, ""), (2, "")], LeadsOf(picture));
        Assert.Equal("104", picture.Marking);
    }

    // With no word for its kind, 10 uF or more is an electrolytic; anything less is drawn without a polarity - 1 uF and
    // 2.2 uF ceramic discs are in the data, and in doubt the plain look is never wrong about a minus side.
    [Theory]
    [InlineData("Radial", "10uF 16V", PartPicture.Shape.RadialElectrolytic)]
    [InlineData("Radial", "100u", PartPicture.Shape.RadialElectrolytic)]
    [InlineData("Radial", "4.7uF 25V", PartPicture.Shape.RadialCapacitor)]
    [InlineData("Axial", "1uF 50V", PartPicture.Shape.AxialCapacitor)]
    [InlineData("Axial", "2200uF 16V", PartPicture.Shape.AxialElectrolytic)]
    [InlineData("Disc", "1uF 50V 10%", PartPicture.Shape.DiscCapacitor)]
    [InlineData("Radial", "N/A", PartPicture.Shape.RadialCapacitor)]
    public void With_no_word_for_its_kind_only_10uF_or_more_is_electrolytic(string package, string value, PartPicture.Shape shape)
    {
        Assert.Equal(shape, PartPicture.Of(Part("Capacitor", package, string.Empty, value), Pins("", ""))!.Shape);
    }

    [Fact]
    public void A_disc_is_printed_with_its_code_and_voltage_and_has_no_polarity()
    {
        PartPicture.Picture? picture = PartPicture.Of(Part("Capacitor", "Disc", "Ceramic", "100nF 25V"), Pins("", ""));

        Assert.Equal(PartPicture.Shape.DiscCapacitor, picture!.Shape);
        Assert.Null(picture.MarkedLead);
        Assert.Equal("104\n25V", picture.Marking);
    }

    // A disc is never an electrolytic, nor anything with two kinds named.
    [Theory]
    [InlineData("Disc", "Electrolytic")]
    [InlineData("Radial", "Electrolytic ceramic")]
    public void A_kind_the_shape_cannot_have_draws_nothing(string package, string friendly)
    {
        Assert.Null(PartPicture.Of(Part("Capacitor", package, friendly, "10uF"), Pins("", "")));
    }

    // With no package, only a friendly name that says the shape draws one: the 540 "Ceramic" ones may be a disc, a dipped
    // or an axial part, and "Electrolytic" alone may stand or lie.
    [Theory]
    [InlineData("Electrolytic Radial", PartPicture.Shape.RadialElectrolytic)]
    [InlineData("Electrolytic Axial", PartPicture.Shape.AxialElectrolytic)]
    [InlineData("Ceramic Axial", PartPicture.Shape.AxialCapacitor)]
    [InlineData("Ceramic disc", PartPicture.Shape.DiscCapacitor)]
    [InlineData("Ceramic", null)]
    [InlineData("Electrolytic", null)]
    [InlineData("Film", null)]
    [InlineData("", null)]
    [InlineData("Radial axial", null)]
    public void With_no_package_only_a_shape_in_the_friendly_name_draws_a_capacitor(string friendly, PartPicture.Shape? shape)
    {
        Assert.Equal(shape, PartPicture.Of(Part("Capacitor", string.Empty, friendly, "10uF 16V"), Pins("", ""))?.Shape);
    }

    // Shapes that vary too much to draw - a trimmer, a tantalum bead, a bipolar electrolytic - draw nothing.
    [Theory]
    [InlineData("Trimmer")]
    [InlineData("Trimmer capacitor")]
    [InlineData("Tantalum")]
    [InlineData("Electrolytic bipolar")]
    [InlineData("Variable capacitor")]
    public void A_capacitor_of_a_kind_not_drawn_draws_nothing(string friendly)
    {
        Assert.Null(PartPicture.Of(Part("Capacitor", "Radial", friendly, "40pF"), Pins("", "")));
    }

    // The data says something the picture cannot show: a third pin on a two-legged capacitor.
    [Fact]
    public void A_pin_the_part_has_no_lead_for_draws_nothing()
    {
        Assert.Null(PartPicture.Of(Part("Capacitor", "Disc", "Ceramic", "100nF"), Pins("", "", "")));
        Assert.Null(PartPicture.Of(Part("Transistor", "TO-92", "NPN", "BC547"), Pins("E", "B", "C", "")));
    }

    // With no pins listed, the leads are the case's, numbered as it numbers them - and given their roles.
    [Fact]
    public void A_part_with_no_pins_listed_has_the_cases_leads()
    {
        Assert.Equal([(1, "+"), (2, "-")], LeadsOf(PartPicture.Of(Part("Capacitor", "Radial", "Electrolytic Radial", "47uF"), [])));
        Assert.Equal([(1, "Cathode"), (2, "Anode")], LeadsOf(PartPicture.Of(Part("Diode", "DO-35", value: "1N914"), [])));
        Assert.Equal([(1, ""), (2, ""), (3, "")], LeadsOf(PartPicture.Of(Part("Transistor", "TO-92", "NPN", "2SC1815"), [])));
    }

    // ###########################################################################################
    // WHAT A TRANSISTOR'S LETTERS STAND FOR, under its picture (owner request, 2026-11-09: "E = Emitter, C = Collector, B =
    // Base"), in the leads' order, each once - a field-effect transistor's G, D, S too. Only when every named lead is one
    // of a set's letters: a regulator's IN, GND, OUT, and a mix, explain nothing; nor do leads with no names.
    // ###########################################################################################
    [Fact]
    public void A_transistors_letters_are_explained_in_the_leads_order()
    {
        static string[] Legend(PartPicture.Picture? picture) =>
            picture!.LegendLines.Select(line => $"{line.Letter} = {line.Meaning}").ToArray();

        Assert.Equal(
            ["E = Emitter", "C = Collector", "B = Base"],
            Legend(PartPicture.Of(Part("Transistor", "TO-92", "NPN", "2SC1815"), Pins("E", "C", "B"))));
        Assert.Equal(
            ["B = Base", "C = Collector", "E = Emitter"],
            Legend(PartPicture.Of(Part("Transistor", "TO-220", "NPN Power", "TIP29A"), Pins("b", "C", "E"))));
        Assert.Equal(
            ["G = Gate", "D = Drain", "S = Source"],
            Legend(PartPicture.Of(Part("Transistor", "TO-220", "MOSFET", "IRF540"), Pins("G", "D", "S"))));
        Assert.Equal(["E = Emitter", "C = Collector"], Legend(PartPicture.Of(Part("Transistor", "TO-92", "NPN", "BC547"), Pins("E", "C", ""))));

        Assert.Empty(Legend(PartPicture.Of(Part("Misc", "TO-220", "Adjustable voltage regulator", "LM317"), Pins("ADJ", "OUT", "IN"))));
        Assert.Empty(Legend(PartPicture.Of(Part("Transistor", "TO-92", "NPN", "BC547"), Pins("E", "G", "B"))));
        Assert.Empty(Legend(PartPicture.Of(Part("Transistor", "TO-92", "NPN", "BC547"), Pins("", "", ""))));
        Assert.Empty(Legend(PartPicture.Of(Part("Capacitor", "Disc", "Ceramic", "100nF"), Pins("", ""))));
    }

    // ---- 78xx regulators ----------------------------------------------------------------------

    private static PartPicture.Picture? Regulator(string value, IReadOnlyList<ComponentPinEntry> pins, Dictionary<string, string>? nets = null) =>
        PartPicture.Of(Part("Misc", "TO-220", "Voltage regulator", value), pins, nets);

    private static string[] LegendOf(PartPicture.Picture? picture) =>
        picture!.LegendLines.Select(line => $"{line.Letter} = {line.Meaning}").ToArray();

    // ###########################################################################################
    // A 78xx IN TO-220 IS INPUT - GROUND - OUTPUT FROM LEFT TO RIGHT, its legs named and explained with the voltage it gives
    // (owner question, 2026-11-09: "is it always implied that 1=Input, 2=Ground, 3=Output? If so, then add this as info also,
    // and there is then also a 5V variant of the same"). With no pin names and no KiCad nets, the datasheet's numbers 1, 2,
    // 3 - as the C128 310378's U59 is.
    // ###########################################################################################
    [Theory]
    [InlineData("7812", "+12 V")]
    [InlineData("7805", "+5 V")]
    [InlineData("LM7805CT", "+5 V")]
    [InlineData("L7812CV", "+12 V")]
    [InlineData("78M05", "+5 V")]
    public void A_78xx_is_input_ground_output_from_the_left_and_says_its_voltage(string value, string voltage)
    {
        PartPicture.Picture? picture = Regulator(value, Pins("", "", ""));

        Assert.Equal(PartPicture.Shape.TO220, picture!.Shape);
        Assert.Equal([(1, "IN"), (2, "GND"), (3, "OUT")], LeadsOf(picture));
        Assert.Equal(["IN = Input", "GND = Ground", $"OUT = Output, {voltage}"], LegendOf(picture));
        Assert.Equal("Printed side facing you, legs down", picture.Caption);
    }

    // ###########################################################################################
    // NUMBERED AS THE BOARD NUMBERS IT, by its KiCad nets: the C64 250407's and KU-14194HB's footprint numbers VR1 and VR2 1 =
    // input, 2 = output, 3 = ground (measured 2026-11-09), so the legs left to right are its pins 1, 3, 2; the C128 310378's
    // U59 is 1, 2, 3. Ground by its net's name, the output by the net naming the voltage - "+15V" is no 5 V.
    // ###########################################################################################
    [Fact]
    public void A_78xx_is_numbered_as_the_boards_KiCad_nets_say()
    {
        PartPicture.Picture? vr2 = Regulator("7805", Pins("", "", ""), new() { ["1"] = "+9V_UNREG", ["2"] = "CAN+5V", ["3"] = "Earth" });
        Assert.Equal([(1, "IN"), (3, "GND"), (2, "OUT")], LeadsOf(vr2));

        PartPicture.Picture? vr1 = Regulator("7812", Pins("", "", ""), new() { ["1"] = "Net-(C88-Pad1)", ["2"] = "+12V", ["3"] = "Earth" });
        Assert.Equal([(1, "IN"), (3, "GND"), (2, "OUT")], LeadsOf(vr1));

        PartPicture.Picture? u59 = Regulator("7812", Pins("", "", ""), new() { ["1"] = "Net-(C101-Pad1)", ["2"] = "GND", ["3"] = "+12V" });
        Assert.Equal([(1, "IN"), (2, "GND"), (3, "OUT")], LeadsOf(u59));

        // Nets that do not settle it draw nothing - a leg named wrong is worse than none.
        Assert.Null(Regulator("7805", Pins("", "", ""), new() { ["1"] = "+15V", ["2"] = "GND", ["3"] = "VCC" }));
        Assert.Null(Regulator("7805", Pins("", "", ""), new() { ["1"] = "GND", ["2"] = "GND", ["3"] = "+5V" }));
    }

    // The pins' own names come first - all three, or nothing is drawn; a 78L05 (a TO-92 the other way round) is no 78xx.
    [Fact]
    public void A_78xx_takes_its_pins_names_and_draws_nothing_on_half_of_them()
    {
        Assert.Equal([(3, "IN"), (2, "GND"), (1, "OUT")], LeadsOf(Regulator("7805", Pins("Vout", "Ground", "Vin"))));
        Assert.Null(Regulator("7805", Pins("IN", "", "")));
        Assert.Null(Regulator("7805", Pins("IN", "IN", "OUT")));

        PartPicture.Picture? low = Regulator("78L05", Pins("", "", ""));
        Assert.Equal([(1, ""), (2, ""), (3, "")], LeadsOf(low));
        Assert.Empty(low!.LegendLines);
    }

    // ---- diodes -------------------------------------------------------------------------------

    [Theory]
    [InlineData("DO-35", "1N914", PartPicture.Shape.GlassDiode)]
    [InlineData("DO-41", "1N4001", PartPicture.Shape.EpoxyDiode)]
    [InlineData("Axial", "1N4148", PartPicture.Shape.GlassDiode)]
    [InlineData("Axial", "1N4371", PartPicture.Shape.GlassDiode)]
    [InlineData("Axial", "1N4001", PartPicture.Shape.EpoxyDiode)]
    [InlineData("Axial", "1N5408", PartPicture.Shape.EpoxyDiode)]
    [InlineData("Axial", "RD6.8EB", PartPicture.Shape.GlassDiode)]
    public void A_diode_is_glass_or_black_epoxy_by_its_package_or_family(string package, string value, PartPicture.Shape shape)
    {
        Assert.Equal(shape, PartPicture.Of(Part("Diode", package, value: value), Pins("", ""))!.Shape);
    }

    // ###########################################################################################
    // KiCad's convention, which every named pin in the data agrees with: pin 1 is the cathode, the band's side - its
    // DO-35 and DO-41 footprints make pad 1 the square one, and on the C128 310378 every diode to ground has pad 2 there
    // (checked 2026-11-09). Written in full, in English, however the data spells it (owner request, 2026-11-09: "Isn't
    // this "Cathode" instead then? ... do write the full name of it instead of a single letter").
    // ###########################################################################################
    [Fact]
    public void A_diodes_cathode_is_pin_1_unless_its_pins_say_otherwise_and_both_are_named_in_full()
    {
        PartPicture.Picture? convention = PartPicture.Of(Part("Diode", "DO-35", value: "1N914"), Pins("", ""));
        Assert.Equal(0, convention!.MarkedLead);
        Assert.Equal([(1, "Cathode"), (2, "Anode")], LeadsOf(convention));

        PartPicture.Picture? named = PartPicture.Of(Part("Diode", "Axial", value: "1N4148"), Pins("A", "KATODE"));
        Assert.Equal(1, named!.MarkedLead);
        Assert.Equal([(1, "Anode"), (2, "Cathode")], LeadsOf(named));
        Assert.Equal([(1, "Cathode"), (2, "Anode")], LeadsOf(PartPicture.Of(Part("Diode", "DO-35", value: "1N914"), Pins("k", "a"))));

        // A name that says something else is kept as typed, KiCad's convention marking pin 1.
        PartPicture.Picture? other = PartPicture.Of(Part("Diode", "DO-35", value: "1N914"), Pins("To +5V", "From pin 3"));
        Assert.Equal([(1, "To +5V"), (2, "From pin 3")], LeadsOf(other));
        Assert.Equal(0, other!.MarkedLead);

        Assert.Equal(1, PartPicture.Of(Part("Diode", "DO-41", value: "1N4001"), Pins("Anode", ""))!.MarkedLead);
        Assert.Null(PartPicture.Of(Part("Diode", "DO-41", value: "1N4001"), Pins("K", "Cathode")));
    }

    [Theory]
    [InlineData("Diode", "", "", "1N914")]
    [InlineData("Diode", "Axial", "Bridge rectifier", "DBA10B")]
    [InlineData("Diode", "Axial", "LED", "Red")]
    [InlineData("Rectifier", "", "Bridge rectifier", "S2VB10/100V")]
    [InlineData("Transistor", "DO-35", "", "1N914")]
    public void A_diode_with_no_package_a_bridge_or_an_LED_draws_nothing(string category, string package, string friendly, string value)
    {
        Assert.Null(PartPicture.Of(Part(category, package, friendly, value), Pins("", "")));
    }

    // ---- three legs ---------------------------------------------------------------------------

    [Fact]
    public void A_TO92_transistor_has_its_three_legs_named_and_its_type_printed()
    {
        PartPicture.Picture? picture = PartPicture.Of(Part("Transistor", "TO-92", "NPN", "2SC1815", "Q1"), Pins("E", "C", "B"));

        Assert.Equal(PartPicture.Shape.TO92, picture!.Shape);
        Assert.Equal([(1, "E"), (2, "C"), (3, "B")], LeadsOf(picture));
        Assert.Equal("2SC1815", picture.Marking);
        Assert.Equal("Flat side facing you, legs down", picture.Caption);
    }

    // A TO-220 whatever its category - a power transistor or a regulator - its type printed, else its part number when short.
    [Fact]
    public void A_TO220_is_drawn_for_a_regulator_too()
    {
        PartPicture.Picture? regulator = PartPicture.Of(Part("Misc", "TO-220", "12VDC voltage regulator", "7812", "VR1"), Pins("IN", "GND", "OUT"));

        Assert.Equal(PartPicture.Shape.TO220, regulator!.Shape);
        Assert.Equal([(1, "IN"), (2, "GND"), (3, "OUT")], LeadsOf(regulator));
        Assert.Equal("7812", regulator.Marking);

        ComponentEntry noType = new() { Category = "Transistor", Package = "to-220", TechnicalNameOrValue = "a power transistor", PartNumber = "902694-01" };
        Assert.Equal("902694-01", PartPicture.Of(noType, [])!.Marking);
    }

    // Pins named rather than numbered (the C64 250469's "B", "C", "E") cannot be placed on the legs: the legs carry the
    // case's numbers only.
    [Fact]
    public void Pins_named_rather_than_numbered_leave_the_legs_numbers_only()
    {
        List<ComponentPinEntry> lettered =
        [
            new() { Pin = "B", Name = "B" },
            new() { Pin = "C", Name = "C" },
            new() { Pin = "E", Name = "E" },
        ];

        Assert.Equal([(1, ""), (2, ""), (3, "")], LeadsOf(PartPicture.Of(Part("Transistor", "TO-220", "NPN", "TIP29A"), lettered)));
    }

    // A crystal's third pin is its case: its leg stands between the other two, as the board files' KiCad footprints have it.
    [Fact]
    public void A_crystal_has_its_case_leg_in_the_middle_and_its_frequency_printed()
    {
        PartPicture.Picture? hc49 = PartPicture.Of(Part("Misc", "HC-49", "Crystal oscillator", string.Empty, "Y2", "17.73447 MHz (PAL)"), Pins("", "", ""));

        Assert.Equal(PartPicture.Shape.Crystal, hc49!.Shape);
        Assert.Equal([1, 3, 2], hc49.Leads.Select(lead => lead.Number));
        Assert.Equal("17.73447\nMHz", hc49.Marking);

        PartPicture.Picture? hc18 = PartPicture.Of(Part("Misc", "HC-18", "Crystal oscillator", "XTAL 16 MHz", "X1"), Pins("", ""));
        Assert.Equal([1, 2], hc18!.Leads.Select(lead => lead.Number));
        Assert.Equal("16\nMHz", hc18.Marking);

        Assert.Null(PartPicture.Of(Part("Misc", "HC-49", "Crystal oscillator"), Pins("", "", "", "")));
    }

    // An EMI filter's ground is its middle leg - pin 2 on the Open128's EMI filters, pin 3 on its ferrite-bead ones.
    [Fact]
    public void An_EMI_filters_ground_leg_stands_in_the_middle()
    {
        PartPicture.Picture? filter = PartPicture.Of(Part("Misc", "EMI filter", "EMI filter", "100pF", "EMI1"), Pins("", "", ""));
        Assert.Equal([1, 2, 3], filter!.Leads.Select(lead => lead.Number));
        Assert.Equal("101", filter.Marking);

        PartPicture.Picture? bead = PartPicture.Of(Part("Misc", "EMI filter", "Ferrite bead", string.Empty, "FB10"), Pins("", "", "GND"));
        Assert.Equal([(1, ""), (3, "GND"), (2, "")], LeadsOf(bead));
    }

    // ---- many legs ----------------------------------------------------------------------------

    [Fact]
    public void A_resistor_pack_has_a_leg_per_pin_its_first_marked_and_its_value_printed()
    {
        PartPicture.Picture? picture = PartPicture.Of(Part("Resistor", "SIP", value: "3K3Ω, 6pin", label: "RP1"), Pins("common", "R1", "R2", "R3", "R4", "R5"));

        Assert.Equal(PartPicture.Shape.ResistorNetwork, picture!.Shape);
        Assert.Equal(6, picture.Leads.Count);
        Assert.Equal(0, picture.MarkedLead);
        Assert.Equal("3K3Ω", picture.Marking);
        Assert.Equal("common", picture.Leads[0].Name);
    }

    // A single resistor is ResistorColorCode's; a pack with no pins listed has no count to draw; a regulator typed as SIP
    // is no resistor pack.
    [Fact]
    public void A_single_resistor_a_pack_with_no_pins_and_a_SIP_of_another_kind_draw_nothing()
    {
        Assert.Null(PartPicture.Of(Part("Resistor", "Axial", value: "1KΩ 1/4W 5%", label: "R1"), Pins("", "")));
        Assert.Null(PartPicture.Of(Part("Resistor", "SIP", value: "3.3KΩ", label: "RP1"), []));
        Assert.Null(PartPicture.Of(Part("Misc", "SIP", "5VDC voltage regulator", "7805", "REG1"), Pins("", "", "")));
    }

    [Fact]
    public void A_pin_header_has_a_post_per_pin_with_its_name()
    {
        PartPicture.Picture? jumper = PartPicture.Of(Part("Misc", "Pin header", "Jumper", label: "J4"), Pins("A", "B"));

        Assert.Equal(PartPicture.Shape.PinHeader, jumper!.Shape);
        Assert.Equal([(1, "A"), (2, "B")], LeadsOf(jumper));
        Assert.Equal(7, PartPicture.Of(Part("Connector", "Pin header", "Internal serial bus", "7 pin", "CN7"), Pins("", "", "", "", "", "", ""))!.Leads.Count);
        Assert.Null(PartPicture.Of(Part("Misc", "Pin header", "Jumper", label: "J4"), []));
    }

    // ---- beads and coils ----------------------------------------------------------------------

    [Theory]
    [InlineData("Misc", "Ferrite bead", "", "FB1", PartPicture.Shape.FerriteBead)]
    [InlineData("Misc", "Coil inductor", "2.2uH", "L1", PartPicture.Shape.Coil)]
    [InlineData("Misc", "Choke coil", "", "L4", PartPicture.Shape.Coil)]
    [InlineData("Misc", "", "", "FB7", PartPicture.Shape.FerriteBead)]
    [InlineData("Misc", "", "", "L3", PartPicture.Shape.Coil)]
    [InlineData("Misc", "", "", "J1", null)]
    [InlineData("Transistor", "", "", "Q1", null)]
    [InlineData("Misc", "Ferrite bead", "", "FB2", null)]
    public void An_axial_part_of_another_kind_is_a_ferrite_bead_or_a_coil_by_its_names(string category, string friendly, string value, string label, PartPicture.Shape? shape)
    {
        // The last case has no package: a bead may also be an EMI filter's three legs, so none is guessed.
        string package = label == "FB2" ? string.Empty : "Axial";

        Assert.Equal(shape, PartPicture.Of(Part(category, package, friendly, value, label), Pins("", ""))?.Shape);
    }

    // The chips are PinoutDrawing's, drawn from above.
    [Theory]
    [InlineData("IC", "DIP-40")]
    [InlineData("IC", "")]
    [InlineData("Connector", "")]
    [InlineData("Misc", "Other")]
    public void A_chip_and_a_part_with_no_drawn_case_draw_nothing(string category, string package)
    {
        Assert.Null(PartPicture.Of(Part(category, package, "Something", "6510"), Pins("", "")));
    }

    // A name that only repeats its pin's number - "1" for pin 1, as 49 of the data's filters and beads have it - is none.
    [Fact]
    public void A_name_that_only_repeats_its_number_is_no_name()
    {
        Assert.Equal([(1, ""), (2, ""), (3, "")], LeadsOf(PartPicture.Of(Part("Misc", "EMI filter", "EMI filter", "100pF"), Pins("1", "2", "3"))));
    }

    // ---- reading the value --------------------------------------------------------------------

    [Theory]
    [InlineData("100nF 25V", 100000.0)]
    [InlineData("0.1 µF/25V", 100000.0)]
    [InlineData(".1 µF", 100000.0)]
    [InlineData("470pF 50V 10%", 470.0)]
    [InlineData("1.8nF 50V 10%", 1800.0)]
    [InlineData("4u7", 4700000.0)]
    [InlineData("10 μF/16V", 10000000.0)]
    [InlineData("1000uF 25V", 1000000000.0)]
    [InlineData("100n", 100000.0)]
    [InlineData("N/A", null)]
    [InlineData("", null)]
    [InlineData("100", null)]
    [InlineData("10 uH", null)]
    public void A_capacitance_is_read_as_the_data_writes_it(string value, double? picofarads)
    {
        Assert.Equal(picofarads is double expected ? (decimal)expected : null, PartPicture.Picofarads(value));
    }

    // The three-figure code a ceramic capacitor is printed with: two digits and the zeros after them, in picofarads.
    [Theory]
    [InlineData(100000, "104")]
    [InlineData(470, "471")]
    [InlineData(1800, "182")]
    [InlineData(2200, "222")]
    [InlineData(100, "101")]
    [InlineData(1000000, "105")]
    [InlineData(2200000, "225")]
    [InlineData(82, "82")]
    [InlineData(4.7, "4.7")]
    [InlineData(1234, "")]
    [InlineData(0, "")]
    public void A_ceramic_capacitor_is_printed_with_its_three_figure_code(double picofarads, string code)
    {
        Assert.Equal(code, PartPicture.CapacitorCode((decimal)picofarads));
    }

    [Theory]
    [InlineData("100nF 25V", "25V")]
    [InlineData("0.1 µF/50V", "50V")]
    [InlineData("10uF", "")]
    [InlineData("220nF 100V 20%", "100V")]
    public void The_voltage_is_read_after_the_value(string value, string volts)
    {
        Assert.Equal(volts, PartPicture.Voltage(value));
    }

    [Theory]
    [InlineData(10000000, "10µF")]
    [InlineData(4700000, "4.7µF")]
    [InlineData(100000, "0.1µF")]
    public void An_electrolytic_is_printed_in_microfarads(double picofarads, string text)
    {
        Assert.Equal(text, PartPicture.MicrofaradText((decimal)picofarads));
    }
}
