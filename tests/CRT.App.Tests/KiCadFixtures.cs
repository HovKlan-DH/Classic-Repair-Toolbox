namespace ClassicRepairToolbox.Tests;

// Minimal but structurally real KiCad file content used by the loader tests. These are hand
// written rather than copied from Assets/Data so each test can isolate one feature, and so the
// suite does not depend on any particular board's contributed data staying unchanged.
internal static class KiCadFixtures
{
    // A board with two nets, one 14-pin-style footprint, a track, a via, an arc and a zone.
    public const string Pcb = """
    (kicad_pcb (version 20221018) (generator pcbnew)
      (net 0 "")
      (net 1 "GND")
      (net 2 "/Sheet1/CLK")

      (footprint "Package_DIP:DIP-14" (layer "F.Cu")
        (at 100 50)
        (fp_text reference "U1" (at 0 -5) (layer "F.SilkS"))
        (pad "1" thru_hole rect (at -3.81 -7.62) (size 1.6 1.6) (layers "*.Cu" "*.Mask") (net 1 "GND"))
        (pad "2" thru_hole oval (at -2.54 -7.62) (size 1.6 1.6) (layers "*.Cu" "*.Mask") (net 2 "/Sheet1/CLK"))
      )

      (segment (start 100 50) (end 110 50) (width 0.25) (layer "F.Cu") (net 1))
      (segment (start 110 50) (end 110 60) (width 0.25) (layer "B.Cu") (net 2))
      (via (at 105 55) (size 0.8) (drill 0.4) (layers "F.Cu" "B.Cu") (net 1))
      (arc (start 120 50) (mid 122 52) (end 124 50) (width 0.25) (layer "F.Cu") (net 2))

      (zone (net 1) (net_name "GND") (layer "B.Cu")
        (polygon (pts (xy 0 0) (xy 10 0) (xy 10 10) (xy 0 10)))
        (filled_polygon (pts (xy 1 1) (xy 9 1) (xy 9 9) (xy 1 9)))
      )
    )
    """;

    // KiCad 7+ writes the reference as a (property "Reference" ...) instead of (fp_text reference ...).
    public const string PcbWithPropertyReference = """
    (kicad_pcb (version 20240108) (generator pcbnew)
      (net 1 "VCC")
      (footprint "R_0805" (layer "F.Cu")
        (at 10 10)
        (property "Reference" "R1")
        (pad "1" smd rect (at -1 0) (size 1 1) (layers "F.Cu") (net 1 "VCC"))
      )
    )
    """;

    // Pre-footprint KiCad used (module ...) even inside a .kicad_pcb file.
    public const string PcbWithLegacyModule = """
    (kicad_pcb (version 20171130)
      (net 1 "GND")
      (module "DIP-8" (layer F.Cu)
        (at 20 20)
        (fp_text reference "U9" (at 0 0) (layer F.SilkS))
        (pad "1" thru_hole circle (at 1 1) (size 1 1) (layers *.Cu) (net 1 "GND"))
      )
    )
    """;

    // A footprint placed on the bottom copper layer. KiCad flips a footprint by rewriting the
    // pads' stored local coordinates, so what the file holds is already the flipped geometry and
    // the loader must not mirror it a second time.
    public const string PcbBottomLayerFootprint = """
    (kicad_pcb (version 20221018)
      (net 1 "GND")
      (footprint "DIP-8" (layer "B.Cu")
        (at 100 100)
        (fp_text reference "U2" (at 0 0) (layer "B.SilkS"))
        (pad "1" thru_hole circle (at 5 3) (size 1 1) (layers "*.Cu") (net 1 "GND"))
      )
    )
    """;

    // Rotated rectangular and oval pads. The footprint sits at 90 degrees, so KiCad writes 90 into
    // each pad's own (at ...) too; pad 3 additionally carries a further 90 of its own, giving 180.
    public const string PcbRotatedPads = """
    (kicad_pcb (version 20221018)
      (net 1 "GND")
      (footprint "Connector" (layer "F.Cu")
        (at 50 50 90)
        (fp_text reference "CN1" (at 0 0) (layer "F.SilkS"))
        (pad "1" smd rect (at 0 0 90) (size 2 0.8) (layers "F.Cu") (net 1 "GND"))
        (pad "2" thru_hole oval (at 2 0 90) (size 2 0.8) (layers "*.Cu") (net 1 "GND"))
        (pad "3" smd roundrect (at 4 0 180) (size 2 0.8) (layers "F.Cu") (net 1 "GND"))
        (pad "4" smd rect (at 6 0) (size 2 0.8) (layers "F.Cu") (net 1 "GND"))
      )
    )
    """;

    // A footprint rotated 90 degrees: pad offsets rotate with it.
    public const string PcbRotatedFootprint = """
    (kicad_pcb (version 20221018)
      (net 1 "GND")
      (footprint "DIP-8" (layer "F.Cu")
        (at 100 100 90)
        (fp_text reference "U3" (at 0 0) (layer "F.SilkS"))
        (pad "1" thru_hole circle (at 10 0) (size 1 1) (layers "*.Cu") (net 1 "GND"))
      )
    )
    """;

    // A hierarchical design: the root sheet names its child both by file and by "Sheetname". KiCad
    // writes the child file into the same folder, so a board folder listing contains both files.
    public const string SchematicRootWithChildSheet = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "aaaaaaaa-0000-0000-0000-000000000001")
      (wire (pts (xy 10 10) (xy 20 10)))
      (sheet
        (at 100 100)
        (uuid "aaaaaaaa-0000-0000-0000-000000000002")
        (property "Sheetname" "Power supply")
        (property "Sheetfile" "child.kicad_sch")
      )
    )
    """;

    // ###########################################################################################
    // A parent sheet whose TWO child sheets are wired together by their sheet pins, under
    // DIFFERENT names - the real C128 topology that produced the CNT1-vs-CNT bug.
    //
    // The joining wire run deliberately carries NO LABEL: on the real board the two pins are simply
    // wired together, and an implementation that looks for the parent label finds nothing. Keep it
    // unlabelled, or this fixture stops testing the thing it exists for.
    // ###########################################################################################
    public const string SchematicParentJoiningTwoChildren = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "dddddddd-0000-0000-0000-000000000001")
      (wire (pts (xy 10 10) (xy 20 10)))
      (wire (pts (xy 20 10) (xy 30 10)))
      (sheet
        (at 5 5)
        (uuid "dddddddd-0000-0000-0000-000000000002")
        (property "Sheetname" "Userport")
        (property "Sheetfile" "userport.kicad_sch")
        (pin "CNT1" bidirectional (at 10 10 180))
      )
      (sheet
        (at 40 5)
        (uuid "dddddddd-0000-0000-0000-000000000003")
        (property "Sheetname" "Serial Bus")
        (property "Sheetfile" "serialbus.kicad_sch")
        (pin "CNT" bidirectional (at 30 10 0))
      )
    )
    """;

    // The userport child: it draws CNT1 and has never heard of "CNT".
    public const string SchematicUserportChild = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "eeeeeeee-0000-0000-0000-000000000001")
      (wire (pts (xy 60 60) (xy 70 60)))
      (hierarchical_label "CNT1" (at 60 60 180))
    )
    """;

    // The serial bus child: same wire, its own name for it.
    public const string SchematicSerialBusChild = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "ffffffff-0000-0000-0000-000000000001")
      (wire (pts (xy 80 80) (xy 90 80)))
      (hierarchical_label "CNT" (at 80 80 0))
    )
    """;

    // The child of SchematicRootWithChildSheet, which on its own knows nothing about that name.
    public const string SchematicChildSheet = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "bbbbbbbb-0000-0000-0000-000000000001")
      (wire (pts (xy 30 30) (xy 40 30)))
      (label "VCC" (at 35 30 0))
    )
    """;

    // Two sheets that reference each other. No file is a root, so a naive "start from the roots"
    // walk would load neither.
    public const string SchematicCycleA = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "cccccccc-0000-0000-0000-000000000001")
      (wire (pts (xy 1 1) (xy 2 1)))
      (sheet
        (at 10 10)
        (uuid "cccccccc-0000-0000-0000-000000000002")
        (property "Sheetname" "B side")
        (property "Sheetfile" "cycleB.kicad_sch")
      )
    )
    """;

    public const string SchematicCycleB = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "dddddddd-0000-0000-0000-000000000001")
      (wire (pts (xy 3 3) (xy 4 3)))
      (sheet
        (at 10 10)
        (uuid "dddddddd-0000-0000-0000-000000000002")
        (property "Sheetname" "A side")
        (property "Sheetfile" "cycleA.kicad_sch")
      )
    )
    """;

    // KiCad 6+ copies each pin's name and type from the symbol onto its pad.
    public const string PcbWithPinFunctions = """
    (kicad_pcb (version 20240108) (generator pcbnew)
      (net 1 "/CPU/~{RESET}")
      (footprint "Package_DIP:DIP-40" (layer "F.Cu")
        (at 10 10)
        (property "Reference" "U8")
        (pad "40" thru_hole circle (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "/CPU/~{RESET}") (pinfunction "~{RES}") (pintype "input"))
        (pad "1" thru_hole rect (at 0 2.54) (size 1.6 1.6) (layers "*.Cu"))
      )
    )
    """;

    // Two footprints with their names, values and datasheets: KiCad 7's properties, and an older
    // file's "module" with its texts.
    public const string PcbWithFootprintNames = """
    (kicad_pcb (version 20240108) (generator pcbnew)
      (footprint "Package_DIP:DIP-20_W7.62mm" (layer "F.Cu")
        (at 10 10)
        (property "Reference" "U12")
        (property "Value" "74LS373")
        (property "Datasheet" "http://www.ti.com/lit/gpn/sn74LS373")
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu"))
      )
      (module KU-14194HB-RevB-KiCad:DIP16 (layer F.Cu)
        (at 30 10)
        (fp_text reference U22 (at 0 0) (layer F.SilkS))
        (fp_text value 4164-2 (at 0 2) (layer F.Fab))
        (pad 1 thru_hole rect (at 0 0) (size 1.6 1.6) (layers *.Cu))
      )
    )
    """;

    // A sheet keeping its library symbols: a two-unit gate with a De Morgan body style, and a symbol
    // derived from it with no pins of its own.
    public const string SchematicWithLibrarySymbols = """
    (kicad_sch (version 20230121) (generator eeschema)
      (lib_symbols
        (symbol "74xx:74LS00"
          (symbol "74LS00_1_1"
            (pin input line (at 0 0 0) (length 2.54) (name "A" (effects (font (size 1 1)))) (number "1" (effects (font (size 1 1)))))
            (pin output inverted (at 5 0 180) (length 2.54) (name "~{Y}") (number "3"))
          )
          (symbol "74LS00_1_2"
            (pin input line (at 0 0 0) (length 2.54) (name "A_DEMORGAN") (number "1"))
          )
          (symbol "74LS00_2_1"
            (pin input line (at 0 0 0) (length 2.54) (name "A") (number "4"))
          )
          (symbol "74LS00_0_1"
            (pin power_in line (at 0 5 270) (length 2.54) hide (name "VCC") (number "14"))
            (pin power_in line (at 0 -5 90) (length 2.54) hide (name "GND") (number "7"))
          )
        )
        (symbol "74xx:74LS01" (extends "74LS00"))
      )
      (symbol (lib_id "74xx:74LS00") (at 30 30 0) (unit 1)
        (property "Reference" "U5" (at 30 28 0))
        (property "Value" "74LS00" (at 30 32 0))
      )
      (symbol (lib_id "74xx:74LS00") (at 60 30 0) (unit 2)
        (property "Reference" "U5" (at 60 28 0))
      )
      (symbol (lib_id "74xx:74LS01") (at 90 30 0) (unit 1)
        (property "Reference" "U6" (at 90 28 0))
      )
    )
    """;

    // A sheet naming its nets with power symbols alone, as the Open128's "Power" sheet does. Today's KiCad
    // names the net by the Value; 9V_UNREG has an older pin whose name differs, and which its PCB uses.
    // PWR_FLAG is flagged as power too, but its pin drives: it names nothing. VEE's pin stands off its
    // origin, and it is turned a quarter - its tip lands at (42.54, 10). The last wire is on no net.
    public const string SchematicWithPowerSymbols = """
    (kicad_sch (version 20230121) (generator eeschema)
      (lib_symbols
        (symbol "power:+5V" (power)
          (symbol "+5V_0_1" (polyline (pts (xy -0.762 1.27) (xy 0 2.54) (xy 0.762 1.27))))
          (symbol "+5V_1_1" (pin power_in line (at 0 0 90) (length 0) hide (name "+5V" (effects (font (size 1 1)))) (number "1")))
        )
        (symbol "power:GND" (power global)
          (symbol "GND_1_1" (pin power_in line (at 0 0 270) (length 0) (hide yes) (name "") (number "1")))
        )
        (symbol "power:PWR_FLAG" (power)
          (symbol "PWR_FLAG_1_1" (pin power_out line (at 0 0 90) (length 0) (name "pwr") (number "1")))
        )
        (symbol "power:9V_UNREG" (power)
          (symbol "9V_UNREG_1_1" (pin power_in line (at 0 0 90) (length 0) hide (name "+9V_UNREG") (number "1")))
        )
        (symbol "power:VEE" (power)
          (symbol "VEE_1_1" (pin power_in line (at 0 -2.54 90) (length 0) hide (name "~") (number "1")))
        )
        (symbol "Device:R"
          (symbol "R_1_1" (pin power_in line (at 0 0 90) (length 0) (name "NOT_A_RAIL") (number "1")))
        )
      )
      (wire (pts (xy 10 10) (xy 10 20)))
      (wire (pts (xy 10 20) (xy 15 20)))
      (wire (pts (xy 20 10) (xy 20 15)))
      (wire (pts (xy 50 50) (xy 60 50)))
      (symbol (lib_id "power:+5V") (at 10 10 0) (property "Reference" "#PWR01") (property "Value" "+5V"))
      (symbol (lib_id "power:GND") (at 20 10 0) (property "Reference" "#PWR02") (property "Value" "GND"))
      (symbol (lib_id "power:PWR_FLAG") (at 30 10 0) (property "Reference" "#FLG01") (property "Value" "PWR_FLAG"))
      (symbol (lib_id "power:9V_UNREG") (at 35 10 0) (property "Reference" "#PWR03") (property "Value" "9V_UNREG"))
      (symbol (lib_id "power:VEE") (at 40 10 90) (property "Reference" "#PWR04") (property "Value" "VEE"))
      (symbol (lib_id "Device:R") (at 70 70 0) (property "Reference" "R1") (property "Value" "10k"))
    )
    """;

    // Where a symbol lands: "Test:L" is lopsided on purpose - a body 0..4 by 0..2 (Y up), a pin coming in
    // from x -2 - so a turn or a flip cannot land on the same rectangle. "Test:Gate" has a part common to
    // its units, two units apart, and a De Morgan body; a hidden pin is far out, and must not count.
    public const string SchematicWithPlacedSymbols = """
    (kicad_sch (version 20230121) (generator eeschema)
      (lib_symbols
        (symbol "Test:L"
          (symbol "L_0_1" (rectangle (start 0 0) (end 4 2)))
          (symbol "L_1_1" (pin input line (at -2 1 0) (length 2) (name "A") (number "1")))
        )
        (symbol "Test:Gate"
          (symbol "Gate_0_1" (circle (center 1 1) (radius 1)))
          (symbol "Gate_1_1"
            (polyline (pts (xy 0 0) (xy 3 0)))
            (pin input line (at 50 50 0) (length 2) hide (name "H1") (number "1"))
            (pin input line (at 60 60 0) (length 2) (hide yes) (name "H2") (number "2"))
          )
          (symbol "Gate_1_2" (arc (start 0 0) (mid 6 6) (end 7 0)))
          (symbol "Gate_2_1" (rectangle (start 10 0) (end 12 2)))
        )
        (symbol "Test:Gate2" (extends "Gate"))
      )
      (symbol (lib_id "Test:L") (at 100 100 0) (property "Reference" "L1"))
      (symbol (lib_id "Test:L") (at 100 100 0) (mirror x) (property "Reference" "L2"))
      (symbol (lib_id "Test:L") (at 100 100 0) (mirror y) (property "Reference" "L3"))
      (symbol (lib_id "Test:L") (at 100 100 90) (property "Reference" "L4"))
      (symbol (lib_id "Test:L") (at 100 100 90) (mirror x) (property "Reference" "L5"))
      (symbol (lib_id "Test:Gate") (at 200 200 0) (unit 1) (property "Reference" "U1"))
      (symbol (lib_id "Test:Gate") (at 200 200 0) (unit 2) (property "Reference" "U1"))
      (symbol (lib_id "Test:Gate") (at 200 200 0) (unit 1) (convert 2) (property "Reference" "U2"))
      (symbol (lib_id "Test:Gate") (at 200 200 0) (unit 1) (body_style 2) (property "Reference" "U3"))
      (symbol (lib_id "Test:Gate2") (at 200 200 0) (unit 2) (property "Reference" "U4"))
      (symbol (lib_id "Test:NotInTheLibrary") (at 300 300 0) (property "Reference" "U5"))
    )
    """;

    public const string Schematic = """
    (kicad_sch (version 20230121) (generator eeschema)
      (uuid "11111111-2222-3333-4444-555555555555")
      (wire (pts (xy 10 10) (xy 20 10)))
      (wire (pts (xy 20 10) (xy 20 20)))
      (label "CLK" (at 15 10 0))
      (global_label "GND" (at 20 20 0))
      (symbol (lib_id "Device:R") (at 30 30 0)
        (property "Reference" "R1" (at 30 28 0))
        (property "Value" "10k" (at 30 32 0))
      )
    )
    """;
}
