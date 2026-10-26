using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadBoardComponents - a new board's components from its KiCad data (CRT 4.0 step 3: "one press
// adds every KiCad component that isn't on the board yet: label and value from KiCad, category
// guessed from the reference prefix, package from the footprint name ... a datasheet link from the
// KiCad symbol"). The footprints' names and values are the live boards' (the shipped KiCad files,
// 2026-10-26).
// ###########################################################################################
public sealed class KiCadBoardComponentsTests
{
    // `count` pads in two rows `spacing` mm apart, numbered as a DIP is: down one side, up the other.
    private static List<KiCadPadPlace> DipPads(int count, double spacing)
    {
        var pads = new List<KiCadPadPlace>();
        int perSide = count / 2;

        for (int index = 0; index < perSide; index++)
            pads.Add(new KiCadPadPlace((index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), 0, index * 2.54));

        for (int index = 0; index < perSide; index++)
            pads.Add(new KiCadPadPlace((count - index).ToString(System.Globalization.CultureInfo.InvariantCulture), spacing, index * 2.54));

        return pads;
    }

    private static List<KiCadPadPlace> Pads(int count) =>
        [.. Enumerable.Range(1, count).Select(number => new KiCadPadPlace(number.ToString(System.Globalization.CultureInfo.InvariantCulture), number, 0))];

    private static KiCadFootprintInfo Footprint(string reference, string value, string footprint, int pads, string datasheet = "") =>
        new(reference, value, footprint, datasheet, Pads(pads));

    [Theory]
    [InlineData("C12", "Capacitor")]
    [InlineData("CT1", "Capacitor")]
    [InlineData("R7", "Resistor")]
    [InlineData("RP4", "Resistor")]
    [InlineData("U7", "IC")]
    [InlineData("UE12", "IC")]
    [InlineData("IC14", "IC")]
    [InlineData("CR5", "Diode")]
    [InlineData("D1", "Diode")]
    [InlineData("CN3", "Connector")]
    [InlineData("CNKYB1", "Connector")]
    [InlineData("P2", "Connector")]
    [InlineData("J1", "Connector")]
    [InlineData("Q3", "Transistor")]
    [InlineData("TR5", "Transistor")]
    [InlineData("FB15", "Misc")]
    [InlineData("EMI40", "Misc")]
    [InlineData("Y1", "Misc")]
    [InlineData("L4", "Misc")]
    [InlineData("SW1", "Misc")]
    [InlineData("VR2", "Misc")]
    [InlineData("JP1", "Misc")]
    [InlineData(" c56 ", "Capacitor")]
    [InlineData("", "Misc")]
    public void The_category_is_the_one_the_live_boards_use_for_the_prefix(string reference, string category)
    {
        Assert.Equal(category, KiCadBoardComponents.CategoryOf(reference));
    }

    // ###########################################################################################
    // The package from the footprint's name - when the pads agree with it. A DIP's rows 600 mil
    // apart make a 24- or 28-pin DIP "wide", whether the name says W15.24 or not.
    // ###########################################################################################
    [Theory]
    [InlineData("Package_DIP:DIP-14_W7.62mm", 14, 7.62, "DIP-14")]
    [InlineData("KU-14194HB-RevB-KiCad:DIP16", 16, 7.62, "DIP-16")]
    [InlineData("c64c-library:PDIP-40", 40, 15.24, "DIP-40")]
    [InlineData("Package_DIP:DIP-28_W15.24mm_Socket", 28, 15.24, "DIP-28 wide")]
    [InlineData("Package_DIP:DIP-28_W7.62mm", 28, 7.62, "DIP-28")]
    [InlineData("KU-14194HB-RevB-KiCad:DIP24", 24, 15.24, "DIP-24 wide")]
    [InlineData("KU-14194HB-RevB-KiCad:DIP24", 24, 7.62, "DIP-24")]
    [InlineData("Open128:Oscillator_DIP-14", 4, 7.62, "")]
    [InlineData("KU-14194HB-RevB-KiCad:DIP4_DIP14_but_4_pins", 4, 7.62, "DIP-4")]
    public void A_DIP_is_named_by_its_pins_and_rows(string footprint, int pads, double spacing, string package)
    {
        Assert.Equal(package, KiCadBoardComponents.PackageOf(footprint, DipPads(pads, spacing)));
    }

    // ###########################################################################################
    // A footprint's name is the package only where the pads agree and the name means it (code review,
    // 2026-10-26): a five-pin TO-220 is no TO-220 the pane draws with three; a DIP is a name STARTING
    // with it ("SOIC-DIP-8" is surface-mounted, "Oscillator_DIP-14" a can); a run of digits too long
    // for a number is no count - it threw, and the whole board's components with it.
    // ###########################################################################################
    [Theory]
    [InlineData("Package_TO_SOT_THT:TO-220-5_Vertical", 5, "")]
    [InlineData("Package_TO_SOT_THT:TO-92-2", 2, "")]
    [InlineData("Package_TO_SOT_THT:TO-220-3_Vertical", 3, "TO-220")]
    [InlineData("MyLib:SOIC-DIP-8", 8, "")]
    [InlineData("Open128:Oscillator_DIP-14", 14, "")]
    [InlineData("MyLib:DIP_20240108123", 20, "")]
    public void A_name_is_the_package_only_where_the_pads_agree_and_the_name_means_it(string footprint, int pads, string package)
    {
        Assert.Equal(package, KiCadBoardComponents.PackageOf(footprint, Pads(pads)));
    }

    [Theory]
    [InlineData("Resistor_THT:R_Array_SIP8", "SIP")]
    [InlineData("Package_TO_SOT_THT:TO-92_Inline_Wide", "TO-92")]
    [InlineData("Package_TO_SOT_THT:TO-92Flat", "TO-92")]
    [InlineData("Open128:TO-220-3_Horizontal_TabDown", "TO-220")]
    [InlineData("KU-14194HB-RevB-KiCad:VOTLAGE_REGULATOR_TO220", "TO-220")]
    [InlineData("Capacitor_THT:C_Disc_D10.0mm_W2.5mm_P5.00mm", "")]
    [InlineData("KU-14194HB-RevB-KiCad:RESISTOR", "")]
    [InlineData("", "")]
    public void Other_packages_are_named_only_when_the_footprint_says(string footprint, string package)
    {
        Assert.Equal(package, KiCadBoardComponents.PackageOf(footprint, Pads(3)));
    }

    // ###########################################################################################
    // A datasheet is a web address or nothing: KiCad's "~", a library path and a sentence are none;
    // "www." opens as "https://".
    // ###########################################################################################
    [Theory]
    [InlineData("http://www.ti.com/lit/gpn/sn74LS373", "http://www.ti.com/lit/gpn/sn74LS373")]
    [InlineData(" https://archive.org/download/tms4164datasheetnov85/TMS4164-datasheet-nov85.pdf ", "https://archive.org/download/tms4164datasheetnov85/TMS4164-datasheet-nov85.pdf")]
    [InlineData("www.ti.com/lit/ds/symlink/sn74ls07.pdf", "https://www.ti.com/lit/ds/symlink/sn74ls07.pdf")]
    [InlineData("~", "")]
    [InlineData("74xx/74hc_hct74.pdf", "")]
    [InlineData("see http://www.ti.com", "")]
    [InlineData("", "")]
    public void A_datasheet_is_a_web_address_or_nothing(string datasheet, string link)
    {
        Assert.Equal(link, KiCadBoardComponents.LinkOf(datasheet));
    }

    // ###########################################################################################
    // The components are the first PCB file's footprints, once per reference: no placeholder, logo,
    // mounting hole or net tie; the value the footprint's, else the symbol's; the datasheet the
    // symbol's, else the footprint's. Listed chips first, each category in natural label order.
    // ###########################################################################################
    [Fact]
    public void The_footprints_are_the_components_with_their_symbols_values_and_datasheets()
    {
        KiCadFootprintInfo[] footprints =
        [
            Footprint("C10", "0.1 µF", "Capacitor_THT:C_Disc_D5.0mm_W2.5mm_P5.00mm", 2),
            Footprint("C2", "~", "Capacitor_THT:C_Disc_D5.0mm_W2.5mm_P5.00mm", 2),
            new("U12", "74LS373", "Package_DIP:DIP-20_W7.62mm", "~", DipPads(20, 7.62)),
            Footprint("U12", "duplicate", "Package_DIP:DIP-20_W7.62mm", 20),
            Footprint("Q2", "2SC1815", "Package_TO_SOT_THT:TO-92_Inline_Wide", 3, "https://media.digikey.com/2SC1815.pdf"),
            Footprint("REF**", "MountingHole_4.3mm_M4", "MountingHole:MountingHole_4.3mm_M4", 1),
            Footprint("H1", "MountingHole", "MountingHole:MountingHole_3.2mm", 1),
            Footprint("LOGO1", "C128-logo", "Open128:C128-logo", 0),
            Footprint("NT1", "Net-Tie", "NetTie:NetTie-2_SMD_Pad0.5mm", 2),
            Footprint("", "", "c64c-library:CAP_DISC_P100mil", 2),
        ];

        KiCadSymbolInfo[] symbols =
        [
            new("U12", "74LS373", "http://www.ti.com/lit/gpn/sn74LS373"),
            new("C2", "1 µF", "~"),
            new("Q2", "2SC1815", "~"),
            new("#PWR01", "+5V", ""),
        ];

        List<KiCadBoardComponent> components = KiCadBoardComponents.Of(footprints, symbols);

        Assert.Equal(
            [
                new KiCadBoardComponent("U12", "74LS373", "IC", "DIP-20", "http://www.ti.com/lit/gpn/sn74LS373"),
                new KiCadBoardComponent("Q2", "2SC1815", "Transistor", "TO-92", "https://media.digikey.com/2SC1815.pdf"),
                new KiCadBoardComponent("C2", "1 µF", "Capacitor", "", ""),
                new KiCadBoardComponent("C10", "0.1 µF", "Capacitor", "", ""),
            ],
            components);
    }

    // A project with no PCB file: its schematic's symbols, without power symbols and unannotated ones.
    [Fact]
    public void Without_a_PCB_the_symbols_are_the_components()
    {
        List<KiCadBoardComponent> components = KiCadBoardComponents.Of(
            [],
            [new("U1", "6510", "~"), new("U1", "6510", "~"), new("#PWR03", "GND", ""), new("U?", "74LS08", ""), new("R1", "10k", "")]);

        Assert.Equal(
            [new KiCadBoardComponent("U1", "6510", "IC", "", ""), new KiCadBoardComponent("R1", "10k", "Resistor", "", "")],
            components);
    }

    // ###########################################################################################
    // Adding them: only the ones the board lacks (by label, any region, capitals ignored), each placed
    // in its category in label order, a new category after the board's own; pins for one with a
    // package - never its net - and a "Datasheet for <value>" link for one with a datasheet.
    // ###########################################################################################
    [Fact]
    public void Adding_puts_in_what_the_board_lacks_with_pins_and_datasheet_links()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "u1", Category = "IC", FriendlyName = "CPU", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U3", Category = "IC" },
            ],
        };

        KiCadBoardComponent[] fromKiCad =
        [
            new("U1", "6510", "IC", "DIP-40", ""),
            new("U2", "74LS373", "IC", "DIP-20", "http://www.ti.com/lit/gpn/sn74LS373"),
            new("C1", "", "Capacitor", "", "https://example.org/cap.pdf"),
        ];

        var asked = new List<string>();

        BoardData result = KiCadBoardComponents.WithComponents(
            board,
            fromKiCad,
            label =>
            {
                asked.Add(label);
                return [new KiCadPin("1", "~OE", "Input", "/Bus/~OE"), new KiCadPin("10", "GND", "Ground", "GND")];
            },
            out KiCadComponentsAdded added);

        Assert.Equal(new KiCadComponentsAdded(2, 1, 2), added);
        Assert.Equal(["U2"], asked);

        Assert.Equal(["u1", "U2", "U3", "C1"], result.Components.Select(row => row.BoardLabel));
        ComponentEntry u2 = result.Components[1];
        Assert.Equal(("74LS373", "IC", "DIP-20", ""), (u2.TechnicalNameOrValue, u2.Category, u2.Package, u2.Region));

        Assert.Equal(
            [("U2", "1", "~OE", "Input"), ("U2", "10", "GND", "Ground")],
            result.ComponentPins.Select(row => (row.BoardLabel, row.Pin, row.Name, row.Direction)));

        Assert.Equal(
            [("C1", "Datasheet", "https://example.org/cap.pdf"), ("U2", "Datasheet for 74LS373", "http://www.ti.com/lit/gpn/sn74LS373")],
            result.ComponentLinks.Select(row => (row.BoardLabel, row.Name, row.Url)));

        // The board as it was is left alone.
        Assert.Equal(2, board.Components.Count);
    }

    // Nothing to add: the very same board, so no step in the draft's history.
    [Fact]
    public void A_board_that_has_them_all_is_the_same_board()
    {
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }] };

        BoardData result = KiCadBoardComponents.WithComponents(board, [new("U1", "6510", "IC", "DIP-40", "")], _ => [], out KiCadComponentsAdded added);

        Assert.Same(board, result);
        Assert.Equal(new KiCadComponentsAdded(0, 0, 0), added);
    }

    // Pin and link rows a label already has - left over from a component deleted in the table - are
    // kept as they are, never doubled.
    [Fact]
    public void Rows_a_label_already_has_are_not_doubled()
    {
        var board = new BoardData
        {
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U2", Pin = "1", Name = "A0" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U2", Name = "Mine", Url = "https://example.org/mine" }],
        };

        BoardData result = KiCadBoardComponents.WithComponents(
            board,
            [new("U2", "74LS373", "IC", "DIP-20", "http://www.ti.com/lit/gpn/sn74LS373")],
            _ => [new KiCadPin("1", "~OE", "Input", "")],
            out KiCadComponentsAdded added);

        Assert.Equal(new KiCadComponentsAdded(1, 0, 0), added);
        Assert.Equal(["A0"], result.ComponentPins.Select(row => row.Name));
        Assert.Equal(["Mine"], result.ComponentLinks.Select(row => row.Name));
    }
}
