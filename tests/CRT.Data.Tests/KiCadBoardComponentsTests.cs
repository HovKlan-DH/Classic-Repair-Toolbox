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
    // What KiCad knows but has no footprint for needs marking by hand (owner request, 2026-11-01) - never
    // a placeholder, once each, in natural order; with no PCB file at all, every symbol.
    [Fact]
    public void The_parts_only_the_schematic_has_are_named()
    {
        KiCadFootprintInfo[] footprints = [new("U1", string.Empty, string.Empty, string.Empty, [new KiCadPadPlace("1", 0, 0)])];
        KiCadSymbolInfo[] symbols =
        [
            new("U1", string.Empty, string.Empty),
            new("Y10", string.Empty, string.Empty),
            new("Y2", string.Empty, string.Empty),
            new("Y2", string.Empty, string.Empty),
            new("#PWR01", string.Empty, string.Empty),
            new("U?", string.Empty, string.Empty),
        ];

        Assert.Equal(["Y2", "Y10"], KiCadBoardComponents.OnlyInSchematic(footprints, symbols));
        Assert.Equal(["U1", "Y2", "Y10"], KiCadBoardComponents.OnlyInSchematic([], symbols));
    }

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

    // ###########################################################################################
    // A mounting hole is what a footprint IS, never what its name mentions (owner report, 2026-11-01: the
    // Open128's "6 components are not labelled yet" held four D-sub connectors, whose footprints say where
    // their own holes are - CN3, CN4, CN5 and CN10 could not be added or marked from KiCad at all).
    // ###########################################################################################
    [Theory]
    [InlineData("Connector_Dsub:DSUB-9_Pins_Horizontal_P2.77x2.84mm_EdgePinOffset14.56mm_Housed_MountingHolesOffset8.20mm", true)]
    [InlineData("Connector_Dsub:DSUB-25_Socket_Horizontal_P2.77x2.84mm_EdgePinOffset7.70mm_Housed_MountingHolesOffset9.12mm", true)]
    [InlineData("MountingHole:MountingHole_3.2mm_M3", false)]
    [InlineData("Mounting_Holes:MountingHole_3mm", false)]
    [InlineData("MountingHole:Hole_3mm_Pad", false)]
    [InlineData("MountingHole_3mm", false)]
    public void Only_a_footprint_that_is_a_mounting_hole_is_taken_as_one(string footprint, bool isComponent)
    {
        Assert.Equal(isComponent, KiCadBoardComponents.IsComponent(Footprint("CN3", "DB9", footprint, 2)));
    }

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
    // A J is a jumper on the live boards and in every shipped KiCad project (owner decision,
    // 2026-10-31) - not KiCad's connector.
    [InlineData("J1", "Misc")]
    [InlineData("J17", "Misc")]
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
    // WHAT A RESISTOR, A CAPACITOR OR A FERRITE BEAD COMES IN (owner request, 2026-11-03: "what package is a
    // resistor? Or a ferrite bead or a capacitor?"), by the footprints the six published boards with a PCB
    // file call them. A guess only where the name says it, and for a resistor by its reference; "CAP_P5.08mm"
    // says only its pitch - the 125 capacitors of KU-14194HB stay empty rather than guessed.
    // ###########################################################################################
    [Theory]
    [InlineData("Capacitor_THT:C_Disc_D10.0mm_W2.5mm_P5.00mm", "C1", "Disc")]
    [InlineData("c64c-library:CAP_DISC_P100mil", "C1", "Disc")]
    [InlineData("Capacitor_THT:CP_Radial_Tantal_D5.0mm_P5.00mm", "C1", "Radial")]
    [InlineData("KU-14194HB-RevB-KiCad:CAP_ELYT_RADIAL_P2.54mm", "C1", "Radial")]
    [InlineData("KU-14194HB-RevB-KiCad:CAPACITOR_RADIAL_P13.97_L9.9_D3.6", "C1", "Radial")]
    [InlineData("Capacitor_THT:CP_Axial_L11.0mm_D6.0mm_P18.00mm_Horizontal", "C1", "Axial")]
    [InlineData("KU-14194HB-RevB-KiCad:LARGE_ELECTROLYTIC_33x18_P38.10mm_AXIAL_TH", "C1", "Axial")]
    [InlineData("Open128:R_Axial_DIN0309_L9.0mm_D2.5mm_P12.70mm_Horizontal", "R1", "Axial")]
    [InlineData("KU-14194HB-RevB-KiCad:RESISTOR", "R1", "Axial")]
    [InlineData("KU-14194HB-RevB-KiCad:RESISTOR", "L1", "Axial")]
    [InlineData("c64c-library:RESISTOR_P200mil_STANDING", "R1", "Axial")]
    [InlineData("c64c-library:POWER_RESISTOR", "R1", "Axial")]
    [InlineData("Issue 4B:Passive_small", "R1", "Axial")]
    [InlineData("KU-14194HB-RevB-KiCad:FERRITE_BEAD_TH_P12.7mm", "FB1", "Axial")]
    [InlineData("c64c-library:FERRITE_P200mil_STANDING", "FB1", "Axial")]
    [InlineData("Inductor_THT:L_Axial_L9.5mm_D4.0mm_P12.70mm_Horizontal_Fastron_SMCC", "L1", "Axial")]
    [InlineData("KU-14194HB-RevB-KiCad:INDUCTOR_P15.24mm", "L1", "Axial")]
    [InlineData("c64c-library:RECTIFIER", "CR1", "Axial")]
    [InlineData("KU-14194HB-RevB-KiCad:ZENER_12.7mm", "CR1", "Axial")]
    [InlineData("Issue 4B:Diode", "D1", "Axial")]
    [InlineData("Diode_THT:D_DO-35_SOD27_P12.70mm_Horizontal", "CR1", "DO-35")]
    [InlineData("Diode_THT:D_DO-41_SOD81_P12.70mm_Horizontal", "CR1", "DO-41")]
    [InlineData("Crystal:Crystal_HC18-U_Vertical", "X1", "HC-18")]
    [InlineData("Connector_PinHeader_2.54mm:PinHeader_1x02_P2.54mm_Vertical", "J1", "Pin header")]
    [InlineData("KU-14194HB-RevB-KiCad:CAP_P5.08mm", "C1", "")]
    [InlineData("Issue 4B:Passive_small", "C1", "")]
    [InlineData("KU-14194HB-RevB-KiCad:FUSE", "F1", "")]
    [InlineData("Resistor_SMD:R_0805_2012Metric", "R1", "")]
    [InlineData("Resistor_SMD:R_2512_6332Metric", "R5", "")]
    [InlineData("Resistor_SMD:R_1812_4532Metric_Pad1.30x3.40mm_HandSolder", "R5", "")]
    [InlineData("Resistor_SMD:R_0612_1632Metric", "R5", "")]
    [InlineData("MyLib:R_2010", "R5", "")]
    [InlineData("Resistor_SMD:R_MELF_MMB-0207", "R5", "")]
    [InlineData("Resistor_SMD:R_Shunt_Ohmite_LVK12", "R5", "")]
    [InlineData("MyLib:R_SMD_Thing", "R1", "")]
    [InlineData("MyLib:R_DIN0207_L6.3mm", "R1", "Axial")]
    [InlineData("C64_2010:RESISTOR", "R1", "Axial")]
    [InlineData("MyLib:C_Discrete", "C1", "")]
    [InlineData("MyLib:RCA_Coaxial", "J1", "")]
    [InlineData("MyLib:CONN_COAXIAL_P5mm", "P1", "")]
    [InlineData("MyLib:Biradial_Thing", "C1", "")]
    public void A_part_on_two_pads_is_named_by_its_footprint_and_a_resistor_by_its_reference(string footprint, string reference, string package)
    {
        Assert.Equal(package, KiCadBoardComponents.PackageOf(footprint, Pads(2), reference));
    }

    [Theory]
    [InlineData("Open128:EMI-Filter", 3, "EMI filter")]
    [InlineData("c64c-library:T-FILTER", 3, "EMI filter")]
    [InlineData("Open128:Ferrite_Bead_vertical", 3, "EMI filter")]
    [InlineData("Crystal:Crystal_HC49-U_Horizontal_1EP_style2", 3, "HC-49")]
    [InlineData("KU-14194HB-RevB-KiCad:RP_ARRAY_7WAY", 8, "SIP")]
    [InlineData("Resistor_THT:R_Array_SIP8", 8, "SIP")]
    [InlineData("Resistor_SMD:R_Array_Concave_4x0603", 8, "")]
    [InlineData("Resistor_SMD:R_Array_Convex_4x1206", 8, "")]
    [InlineData("Connector_PinSocket_2.54mm:PinSocket_1x08_P2.54mm_Vertical", 8, "Pin header")]
    [InlineData("KU-14194HB-RevB-KiCad:POTMETER", 3, "")]
    [InlineData("KU-14194HB-RevB-KiCad:CRYSTAL", 3, "")]
    [InlineData("Open128:BRIDGE_RECTIFIER", 4, "")]
    public void Parts_on_more_pads_are_named_by_their_footprint_too(string footprint, int pads, string package)
    {
        Assert.Equal(package, KiCadBoardComponents.PackageOf(footprint, Pads(pads)));
    }

    // Everything guessed is a word the pane offers - a guess it could not show is no guess.
    [Fact]
    public void Every_package_guessed_is_one_the_list_offers()
    {
        foreach (string footprint in new[] { "C_Disc_D5.0mm", "CP_Radial_D5", "R_Axial_DIN", "D_DO-35", "D_DO-41", "Crystal_HC18-U" })
        {
            Assert.Contains(KiCadBoardComponents.PackageOf(footprint, Pads(2)), ComponentVocabulary.Packages);
        }

        Assert.Contains(KiCadBoardComponents.PackageOf("EMI-Filter", Pads(3)), ComponentVocabulary.Packages);
        Assert.Contains(KiCadBoardComponents.PackageOf("Crystal_HC49-U", Pads(3)), ComponentVocabulary.Packages);
        Assert.Contains(KiCadBoardComponents.PackageOf("PinHeader_1x02", Pads(2)), ComponentVocabulary.Packages);
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
    // symbol's, else the footprint's. Listed by category alphabetically (owner request, 2026-11-03), each
    // category in natural label order.
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
                new KiCadBoardComponent("C2", "1 µF", "Capacitor", "Disc", ""),
                new KiCadBoardComponent("C10", "0.1 µF", "Capacitor", "Disc", ""),
                new KiCadBoardComponent("U12", "74LS373", "IC", "DIP-20", "http://www.ti.com/lit/gpn/sn74LS373"),
                new KiCadBoardComponent("Q2", "2SC1815", "Transistor", "TO-92", "https://media.digikey.com/2SC1815.pdf"),
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
    // An import's categories come in alphabetical order (owner request, 2026-11-03: "When adding the
    // categories ... per default sort them alphabetically") - on a new board, and around what a board
    // has already, whose own rows stay where they are.
    // ###########################################################################################
    [Fact]
    public void An_imports_categories_are_added_in_alphabetical_order()
    {
        KiCadFootprintInfo[] footprints =
        [
            Footprint("Q1", "2SC1815", "Package_TO_SOT_THT:TO-92_Inline_Wide", 3),
            Footprint("U1", "6510", "Package_DIP:DIP-20_W7.62mm", 20),
            Footprint("R1", "10k", "Resistor_THT:R_Axial", 2),
            Footprint("Y1", "XTAL", "Crystal:Crystal_HC49", 2),
            Footprint("C1", "1 µF", "Capacitor_THT:C_Disc", 2),
            Footprint("CN1", "Power", "Connector:Molex", 2),
            Footprint("D1", "1N4148", "Diode_THT:D_DO-35", 2),
        ];

        List<KiCadBoardComponent> components = KiCadBoardComponents.Of(footprints, []);

        string[] alphabetical = ["Capacitor", "Connector", "Diode", "IC", "Misc", "Resistor", "Transistor"];

        Assert.Equal(alphabetical, components.Select(component => component.Category));

        BoardData onANewBoard = KiCadBoardComponents.WithComponents(new BoardData(), components, _ => [], out _);

        Assert.Equal(alphabetical, onANewBoard.Components.Select(row => row.Category));

        // A board with a category of its own: the new ones around it, its own row where it was.
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U99", Category = "IC" }] };

        BoardData added = KiCadBoardComponents.WithComponents(board, components, _ => [], out _);

        Assert.Equal(alphabetical, added.Components.Select(row => row.Category).Distinct());
        Assert.Equal(["U1", "U99"], added.Components.Where(row => row.Category == "IC").Select(row => row.BoardLabel));
    }

    // ###########################################################################################
    // Adding them: only the ones the board lacks (by label, any region, capitals ignored), each placed
    // in its category in label order, a new category where the alphabet puts it among the board's; pins
    // for one with a package - never its net - and a "Datasheet for <value>" link for one with a datasheet.
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

        Assert.Equal(["C1", "u1", "U2", "U3"], result.Components.Select(row => row.BoardLabel));
        ComponentEntry u2 = result.Components[2];
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
