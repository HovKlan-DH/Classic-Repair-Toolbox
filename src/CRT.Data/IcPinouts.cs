using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // An IC whose pins' directions CRT holds from its datasheet - its name, the datasheet's name of every pin (index 0 is pin 1)
    // and each pin's direction in the pin list's words (ComponentVocabulary.PinDirections), "" for a pin the datasheet does not
    // settle - and the names it is known by, in capitals (IcPinouts.Named): a number ("6510"), a maker's whole name ("Z8400",
    // "ULN2003"), for a 74-series part "74" and its number ("74123"), and "=4044" for a name matched only exactly.
    // ###########################################################################################
    public sealed record IcPinout(string Name, IReadOnlyList<string> PinNames, IReadOnlyList<string> Directions, IReadOnlyList<string> Aliases)
    {
        public int PinCount => this.PinNames.Count;

        // The case the pin-out is of - "DIP-18", the CX20185's "QFP-44" - which the conversion gives a component that names
        // none (owner request, 2026-11-17 - ConversionPinouts), and which a component's own must be (ChipNames.PackageFits).
        public string Package { get; init; } = string.Empty;

        // For a 74-series number whose families differ in their pins, the families this pin-out is of ("" the plain one,
        // "S", "LS" - ChipNames.SeventyFourFamily): the 7451 and 74S51 have two AND-OR-INVERT gates of 2 x 2 inputs, the 74LS51
        // one of 3 x 3 on pins 11 and 12 too. Null for every family.
        public IReadOnlyList<string>? Families { get; init; }

        public string NameOf(int pin) => pin >= 1 && pin <= this.PinCount ? this.PinNames[pin - 1] : string.Empty;

        public string DirectionOf(int pin) => pin >= 1 && pin <= this.PinCount ? this.Directions[pin - 1] : string.Empty;
    }

    // ###########################################################################################
    // THE OTHER ICs' PIN DIRECTIONS, FROM THEIR DATASHEETS (owner decision, 2026-11-17: "OK to what you suggest for 1-4,
    // including for you to investigate for other ICs like CPU, CIAs etc."): the CPUs, the CIAs and VIAs, the video and sound
    // chips, the ROMs, RAMs and PLAs on the published boards - directions only; the logic chips, whose insides CRT draws, are
    // LogicChips'. DatasheetPins puts them before what KiCad says.
    //
    // TYPED FROM THE MAKERS' DATASHEETS - for a Commodore chip with none published, its specification (the TED's), the C128
    // Programmer's Reference Guide (the MMU, the VDC, the VIC-IIe), the C64 service manual's data page (the 8701), the PLA's
    // die-derived model (the 8721) and, for the C64E's 251715, William Levak's pinout and Frank Buss's replacement proven in
    // that board - and checked against the pin names the 14 published boards already give these chips: every name stands at
    // the datasheet's pin, differing only in spelling (VSS for GND, /CE for /CS, DQ for D) or naming the board's own signal
    // (a PLA's A13 on I7), 2026-11-17. An open-collector output, and a 3-state one, is an "Output"; a data bus and a port pin
    // are "Input/output"; a crystal's and a filter capacitor's pins, which are neither, "Other". A pin no source settles is
    // "?" here, "" for DirectionOf - DatasheetPins then leaves it to the board (the VIC-IIe's pins 11 and 15, the 8721's pin
    // 29, the 251715's 18 to 21, the VIC's oscillator pins).
    //
    // THE C128 250477's AND THE CPC 664's CHIPS, ADDED FOR THEIR PINOUTS (owner request, 2026-11-17: "Can you now draw also
    // this IC? (and add its pins in data migration)", then "investigate the other ICs mentioned here"): the 8568 from the
    // C128/128D service manual (314001-08, its "IC 8568" page - the 250477's own schematic agrees pin for pin; its "Pinout"
    // picture is the 8563's, whose pins differ), the CX20185 from the same manual's read/write amplifier pages and the
    // board's own Sony picture, the 5710 from William Levak's sheet (the board's own "Info about 5710") with its directions
    // traced on the board's schematic, and the CPC's Gate Array from the board's own picture with its directions from the
    // Z80's and the CRTC's sides and Bread80's description. The 4464's names are as the owner's picture spells them (the
    // boards' "Pinout" picture of it), checked against TI's TMS4464 datasheet - the boards' own "4464.PDF" - which writes
    // DQ1-DQ4 for D0-D3 and VDD for VCC.
    //
    // *** THE 64H156 (251828) BY ITS PART NUMBER ALONE. *** Its -01 and -03 have 40 pins and its -02 42 (the 1571 service
    // manual, whose pin assignment and block diagram this is), so "64H156" alone says neither and is no chip here - a drawing
    // of the wrong one numbers every pin past 20 wrongly. The 250477's is a 251828-01 of 40 pins (owner, 2026-11-18: "U105 is
    // 40 pin - I just counted them"), though its schematic numbers the 42-pin one. Its ATN is an output: the manual's block
    // diagram and MAME's emulation drive it (the attention acknowledge), where the manual's table says "input".
    //
    // THE AMIGA 500's AND 1000's CHIPS (owner request, 2026-11-18: "please investigate what kind of components an Amiga 500
    // and Amiga 1000 have, and then support those" - no Amiga board is published yet): the 68000 from Motorola's M68000
    // User's Manual (Fig. 11-1); the custom chips from the A500 service manual (314981-01: Fat Agnus, Denise, Paula, Gary,
    // the 8520), the A500 rev 5/6A/7 schematics (314981-02, whose names the Fat Agnus and Gary take) and the Amiga Hardware
    // Reference Manual's Appendix J (the A1000's original Agnus, and the directions), every pin cross-checked against a second
    // of them or a KiCad remake of the board; Gary's directions from Commodore's A500/A2000 Gary specification, its pins 8
    // and 9 named as the schematics wire them (the specification swaps them); the 27C400 from Atmel's datasheet. The 8375
    // (A500+) moves pins 35 to 57 and is none of them; the Kickstart mask ROMs have no datasheet, so they are matched by no
    // part number - their socket on the rev 6A is wired as the 27C400 is, on the rev 5 otherwise.
    //
    // A COMPONENT IS MATCHED BY ITS NUMBER, as LogicChips does ("6510", "MOS6526", "HD6845AP", "4164-2 DRAM", "Z8400APS",
    // "901227-03", "SN74LS629N"): its "Technical name or value" first, its part number after - and taken only when nothing the
    // board says contradicts the datasheet (For): a category no IC has, a package of another pin count, a pin past the case,
    // a ground or supply named where the datasheet has none (Contradicts).
    // ###########################################################################################
    public static class IcPinouts
    {
        // One code per direction: Input, Output, Both ways, Power, Ground, Not connected, other (X), and "?" for unsettled.
        private static readonly Dictionary<char, string> Codes = new()
        {
            ['I'] = "Input",
            ['O'] = "Output",
            ['B'] = "Input/output",
            ['P'] = "Power",
            ['G'] = "Ground",
            ['N'] = "Not connected",
            ['X'] = "Other",
            ['?'] = string.Empty,
        };

        // The pin-outs several parts share.
        private const string Rom2332 = "A7..A0:I D0..D2:O GND:G D3..D7:O A11:I A10:I /CS1:I CS2:I A9:I A8:I VCC:P";
        private const string Rom2364 = "A7..A0:I D0..D2:O GND:G D3..D7:O A11:I A10:I /CS:I A12:I A9:I A8:I VCC:P";
        private const string Rom23128 = "NC:N A12:I A7..A0:I D0..D2:O GND:G D3..D7:O /CS2:I A10:I /CS1:I A11:I A9:I A8:I A13:I CS3:I VCC:P";
        private const string Cpu6510 = "PHI0:I RDY:I /IRQ:I /NMI:I AEC:I VCC:P A0..A13:O GND:G A14..A15:O P5..P0:B D7..D0:B R/W:O PHI2:O /RES:I";
        private const string VicII = "D6..D0:B /IRQ:O /LP:I /CS:I R/W:I BA:O VDD:P COLOR:O S/LUM:O AEC:O PHI0:O /RAS:O /CAS:O VSS:G PHI_COLOR:I PHI_IN:I " +
                                     "A11:O A0/A8:B A1/A9:B A2/A10:B A3/A11:B A4/A12:B A5/A13:B A6..A10:O D11..D8:I D7:B VCC:P";
        private const string Sid = "CAP1A:X CAP1B:X CAP2A:X CAP2B:X /RES:I PHI2:I R/W:I /CS:I A0..A4:I GND:G D0..D7:B POTY:I POTX:I VCC:P EXT_IN:I AUDIO_OUT:O VDD:P";
        private const string Dram64Kx4 = "/G:I D0:B D1:B /W:I /RAS:I A6:I A5:I A4:I VCC:P A7:I A3:I A2:I A1:I A0:I D2:B /CAS:I D3:B VSS:G";
        private const string Pla82S100 = "FE:I I7..I0:I F7..F4:O GND:G F3..F0:O /CE:I I15..I8:I VCC:P";
        private const string FatAgnusTo40 =
            "DRD13..DRD0:B VCC:P /RESET:I /INT3:O DMAL:I /BLISS:I /BLIT:O /WE:O R/W:I /REGEN:I /AS:I /RAMEN:I RGA8..RGA1:O 28MHz:I XCLK:I /XCLKEN:I " +
            "/CDAC:O 7MHz:O CCKQ:O CCK:O";
        private const string FatAgnusFrom42 =
            "VSS:G DRA0..DRA8:O /LDS:I /UDS:I /CASL:O /CASU:O /RAS1:O /RAS0:O VSS:G A19:I A1..A18:I /LPEN:I /VSYNC:B /CSYNC:O /HSYNC:B VSS:G DRD15:B DRD14:B";
        private const string DeniseTo31 = "D6..D0:B M1H:I M0H:I RGA8..RGA1:I /BURST:O VCC:P R0..R3:O B0..B3:O G0..G3:O";
        private const string DeniseFrom35 = "7M:I CCK:I VSS:G M0V:I M1V:I D15..D7:B";

        // The catalogue - each chip with the names it is known by, in the order Named tries them.
        public static readonly IReadOnlyList<IcPinout> All;

        static IcPinouts()
        {
            var all = new List<IcPinout>();

            // Every pin-out is a DIP's but where `package` names another case.
            void Chip(string name, string[] aliases, string pinout, string? package = null, string[]? families = null)
            {
                IcPinout chip = IcPinouts.Parse(name, pinout, [.. aliases.Select(alias => alias.ToUpperInvariant())]);

                all.Add(chip with { Package = package ?? $"DIP-{chip.PinCount.ToString(CultureInfo.InvariantCulture)}", Families = families });
            }

            // ----- CPUs -----------------------------------------------------------------------------
            Chip("6502", ["6502"], "VSS:G RDY:I PHI1:O /IRQ:I NC:N /NMI:I SYNC:O VCC:P A0..A11:O VSS:G A12..A15:O D7..D0:B R/W:O NC:N NC:N PHI0:I SO:I PHI2:O /RES:I");
            Chip("6510", ["6510", "8500"], Cpu6510);
            Chip("8502", ["8502"], "PHI0:I RDY:I /IRQ:I /NMI:I AEC:I VCC:P A0..A13:O GND:G A14..A15:O P6..P0:B D7..D0:B R/W:O /RES:I");
            Chip("7501", ["7501", "8501"], "PHI0:I RDY:I /IRQ:I AEC:I VCC:P A0..A13:O GND:G A14..A15:O GATE:I P7:B P6:B P4..P0:B D7..D0:B R/W:O /RES:I");
            Chip("Z80", ["Z80", "Z8400", "Z84C00"],
                "A11..A15:O CLK:I D4:B D3:B D5:B D6:B VCC:P D2:B D7:B D0:B D1:B /INT:I /NMI:I /HALT:O /MREQ:O /IORQ:O /RD:O /WR:O /BUSACK:O " +
                "/WAIT:I /BUSREQ:I /RESET:I /M1:O /RFSH:O GND:G A0..A10:O");

            // ----- Interface chips ------------------------------------------------------------------
            Chip("6522", ["6522"], "VSS:G PA0..PA7:B PB0..PB7:B CB1:B CB2:B VCC:P /IRQ:O R/W:I /CS2:I CS1:I PHI2:I D7..D0:B /RES:I RS3..RS0:I CA2:B CA1:I");
            // The Amiga's 8520 is the 6526 pin for pin (the A500 service manual's 8520 page); only its TOD counts in binary.
            Chip("6526", ["6526", "8521", "8520"], "VSS:G PA0..PA7:B PB0..PB7:B /PC:O TOD:I VCC:P /IRQ:O R/W:I /CS:I /FLAG:I PHI2:I D7..D0:B /RES:I RS3..RS0:I SP:B CNT:B");
            Chip("6529", ["6529"], "R/W:I P0..P7:B VSS:G D7..D0:B /CS:I VDD:P");
            Chip("6551", ["6551"], "VSS:G CS0:I /CS1:I /RES:I RXC:B XTAL1:I XTAL2:X /RTS:O /CTS:I TXD:O /DTR:O RXD:I RS0:I RS1:I VCC:P /DCD:I /DSR:I D0..D7:B /IRQ:O PHI2:I R/W:I");
            Chip("8255", ["8255"], "PA3..PA0:B /RD:I /CS:I GND:G A1:I A0:I PC7..PC4:B PC0..PC3:B PB0..PB7:B VCC:P D7..D0:B RESET:I /WR:I PA7..PA4:B");
            Chip("8722", ["8722"],
                "VDD:P /RES:I TA15..TA8:O /CAS1:O /CAS0:O I/O_SEL:O MS1:O MS0:O AEC:I MUX:I A0..A3:I A4/5:I A6/7:I A8..A15:I R/W:I PHI0:I VSS:G D0..D7:B " +
                "/Z80EN:O FSDIR:O /GAME:B /EXROM:B 128/64:O SENSE40:I");

            // The C128DCR's floppy controller and CIA in one (William Levak's names). CK comes from the 64H156's clock-select
            // output, 16MHz goes to its OSC input, WD and WG are wired-OR with its write outputs - traced on the 250477's
            // schematic; TED, phi0, /ACCL and VIA1 PA5 were not, and pins 42 and 43 are connected nowhere on it.
            Chip("5710", ["5710"],
                "SRQ_OUT:O TED:? PHI0:? CK:I /ACCL:? VIA1_PA5:? PHI2:I D7..D4:B VSS:G VCC:P D3..D0:B A15..A12:I A10:I A4..A0:I XTL2:X XTL1:X " +
                "16MHz:O R/W:I /VIA2:O /RAM:O /VIA1:O VSS:G VCC:P WD:O /RD:I /WPRT:I WG:O /INDEX:I NC:N NC:N /RESET:I /IRQ:O /DATA_IN:I SRQ_IN:I /DATA_OUT:O");

            // The CPC's Gate Array: the Z80's data bus, address lines and strobes and the CRTC's syncs come in; the clocks, the
            // RAM's and ROM's strobes, the colours and the composite sync go out.
            Chip("40010", ["40008", "40010"],
                "D5..D7:I CCLK:O /NSYNC:O VDD1:P /NRESET:I B:O DISPEN:I G:O HSYNC:I R:O VSYNC:I /NCPU:O VSS1:G /NCAS:O /NMREQ:I /NIORQ:I /NPHI:O /NM1:I " +
                "/NRD:I READY:O /N244EN:O CK16:I VDD2:P VSS2:G /NROMEN:O A15:I /NRAMRD:O A14:I /NCASAD:O /NINTERRUPT:O /NMWE:O /NRAS:O D0:I VSS3:G D1..D4:I");

            // ----- Video and sound ------------------------------------------------------------------
            Chip("6561", ["6560", "6561"], "NC:N COLOR:O S/LUM:O R/W:I DB11..DB8:I DB7..DB0:B POTX:I POTY:I AUDIO:O VSS:G A0..A13:B PHI1:O PHI2:O OPTION:I OSC2:? OSC1:? VDD:P");
            Chip("6567", ["6567", "6569", "6572", "6573", "8562", "8565"], VicII);
            Chip("8564", ["8564", "8566"],
                "D6..D0:B /IRQ:O /LP:I BA:O /DMAR:? AEC:O /CS:I R/W:I /DMAA:? COLOR:O S/LUM:O 1MHz:O /RAS:O /CAS:O MUX:O /IOACC:I 2MHz:O VSS:G Z80_PHI:O " +
                "K0..K2:O PHI_COLOR:I PHI_IN:I A11:O A0..A5:B A6:O A7..A10:O D11..D8:I D7:B VCC:P");
            Chip("7360", ["7360", "8360"], "A2:B A1:B A0:B VDD:P /CS0:O /CS1:O R/W:B /IRQ:O MUX:O /RAS:O /CAS:O PHI0:O COLOR:O CLK_IN:I K0..K7:I LUM:O VSS:G DB0..DB7:B SND:O BA:O AEC:O A15..A3:B");
            Chip("8563", ["8563"],
                "CCLK:O DCLK:I HSYNC:O CS:I NC:N NC:N /CS:I /RS:I R/W:I D7:B D6:B VSS:G D5..D0:B DISPEN:O VSYNC:O DR/W:O INIT:I /RES:I TST:I /LP:I DA0..DA7:O " +
                "DD0..DD2:B VDD:P DD3..DD7:B I:O B:O G:O R:O /RAS:O /CAS:O");

            // The 8563's successor in the C128DCR: the same chip with /INTR, CSYNC and CV, its pins in other places. INIT is the
            // power-up reset (active low, wired to the board's reset); RES only restarts the scan counters (grounded on the board).
            Chip("8568", ["8568"],
                "/CAS:O CCLK:O DCLK:I HSYNC:O CS:I /CS:I RS:I R/W:I /INTR:O D7:B D6:B VSS:G D5..D0:B CSYNC:O VSYNC:O DR/W:O RES:I /INIT:I TEST:I LPEN:I " +
                "DA0..DA7:O DD0..DD2:B VCC:P DD3..DD7:B I:O B:O G:O R:O CV:O /RAS:O");
            Chip("6581", ["6581", "6582", "8580"], Sid);
            Chip("AY-3-8912", ["AY-3-8912"], "ANALOG_C:O TEST_1:X VCC:P ANALOG_B:O ANALOG_A:O VSS:G IOA7..IOA0:B CLOCK:I /RESET:I A8:I BDIR:I BC2:I BC1:I DA7..DA0:B");
            Chip("6845", ["6845"], "VSS:G /RES:I LPSTB:I MA0..MA13:O DISPTMG:O CUDISP:O VCC:P CLK:I R/W:I E:I RS:I /CS:I D7..D0:B RA4..RA0:O HSYNC:O VSYNC:O");

            // ----- Clocks ---------------------------------------------------------------------------
            Chip("8701", ["8701"], "NC:N VSS:G NC:N NC:N /RESET:I PHI_DOT:O /PAL:I PHI_COLOR:O VSS:G NC:N NC:N VDD:P XTAL_OUT:X XTAL_IN:X VDD:P NC:N");
            Chip("74629", ["74629"], "2FC:I 1FC:I 1RNG:I 1CX1:X 1CX2:X /1EN:I 1Y:O OSC_GND:G GND:G 2Y:O /2EN:I 2CX2:X 2CX1:X 2RNG:I OSC_VCC:P VCC:P");
            Chip("MC4044", ["MC4044", "=4044"], "R:I D1:O V:I PU:I UF:O D2:O GND:G AMP_OUT:O AMP_IN:I DF:O PD:I U2:O U1:O VCC:P");

            // ----- ROMs -----------------------------------------------------------------------------
            Chip("2332", ["2332", "901225", "901460"], Rom2332);
            Chip("2364", ["2364", "901226", "901227", "901486", "390059"], Rom2364);
            Chip("23128", ["23128", "251913"], Rom23128);
            Chip("23256", ["23256"], "NC:N A12:I A7..A0:I D0..D2:O GND:G D3..D7:O /CE:I A10:I /OE:I A11:I A9:I A8:I A13:I A14:I VCC:P");
            Chip("27256", ["27256"], "VPP:P A12:I A7..A0:I O0..O2:O GND:G O3..O7:O /CE:I A10:I /OE:I A11:I A9:I A8:I A13:I A14:I VCC:P");

            // ----- RAMs -----------------------------------------------------------------------------
            Chip("2114", ["2114"], "A6:I A5:I A4:I A3:I A0:I A1:I A2:I /CS:I GND:G /WE:I IO4:B IO3:B IO2:B IO1:B A9:I A8:I A7:I VCC:P");
            Chip("2016", ["2016", "6116", "M58725"], "A7..A0:I D0..D2:B GND:G D3..D7:B /CE:I A10:I /OE:I /WE:I A9:I A8:I VCC:P");
            Chip("4116", ["4116"], "VBB:P DIN:I /WE:I /RAS:I A0:I A2:I A1:I VDD:P VCC:P A5:I A4:I A3:I A6:I DOUT:O /CAS:I VSS:G");
            Chip("4164", ["4164", "4532"], "NC:N DIN:I /WE:I /RAS:I A0:I A2:I A1:I VCC:P A7:I A5:I A4:I A3:I A6:I DOUT:O /CAS:I VSS:G");
            Chip("41464", ["41464", "4464", "4416"], Dram64Kx4);

            // ----- PLAs and the rest ----------------------------------------------------------------
            Chip("82S100", ["82S100", "PLS100", "906114", "251641", "7700-010"], Pla82S100);
            Chip("8721", ["8721"],
                "A15..A10:I VICFIX:I DMAACK:I AEC:I R/W:I /GAME:I /EXROM:I Z80_EN:I Z80_I/O:I 64/128:I I/O_SEL:I ROMBANKHI:I ROMBANKLO:I VMA4:I VMA5:I BA:I " +
                "LORAM:I HIRAM:I VSS:G CHAREN:I VA14:I 128/256:I NC:N SDEN:? /ROML:O /ROMH:O CLRBNK:O /FROM:O /ROM4../ROM1:O /IOCS:O DIR:O DWE:O CASENB:O " +
                "/VIC:O /IOACC:O GWE:O /COLRAM:O /CHAROM:O CLK:I VCC:P");
            Chip("251715", ["251715"],
                "A13..A15:I R/W:I /LORAM:I /HIRAM:I /CHAREN:I NMI:O RESTORE:I /VA14:I /VA15:I /CIA2:O /CHAROM:O /BASIC:O /KERNAL:O VSS:G /COLRAM:O " +
                "/CAS0../CAS3:? /CASRAM:O RAM_R/W:O MA6:O MA3:B MA0:B MA2:B MA4:B MA1:B MA5:B MA7:O NC:N VSS:G /CIA1:O /CAS:I /RAS:I PHI0:I AEC:I BA:I " +
                "/VIC:O VA6:I VA7:I /SID:O /ROML:O /I/O2:O /EXROM:I /GAME:I VSS:G /I/O1:O /ROMH:O A0..A7:B A8..A12:I VDD:P");
            Chip("ULN2003", ["ULN2003", "UPA2003"], "1B:I 2B:I 3B:I 4B:I 5B:I 6B:I 7B:I E:G COM:X 7C:O 6C:O 5C:O 4C:O 3C:O 2C:O 1C:O");

            // The 1541B's and 1571's gate array, its 40-pin revisions only (see the header).
            Chip("64H156-01", ["=64H156-01", "=64H156-03", "=251828-01", "=251828-03"],
                "TEST:I YB0..YB7:B VSS:G STP1:I STP0:I MTR:I A:I DS0:I DS1:I SYNC:O TED:I OE:I ACCL:I OSC:I ATNA:I ATNI:I ATN:O Y0..Y3:O XWR:O VDD:P " +
                "CLR:I PLL:I LOCK:I RW:I Q:O QX:O CK:O B:O SOE:I BYTE:O");

            // The C128DCR's floppy read/write amplifier: a QFP of 44 pins (the 250477's layout draws it square, 11 a side, pin 1
            // by its cut corner). Its "connect a resistor/capacitor" pins are "Other", the head terminals both ways, the digital
            // controls inputs.
            Chip("CX20185", ["CX20185"],
                "DIFF_CONSTANT_B:X DIFF_CONSTANT_A:X DIFF_IN_B:I DIFF_IN_A:I PRE_OUT_B:O PRE_OUT_A:O GAIN_SEL_C:X GAIN_SEL_B:X GAIN_SEL_A:X " +
                "READ_DUMP_B:X READ_DUMP_A:X HEAD0_A:B HEAD0_B:B HEAD1_A:B HEAD1_B:B A.GND:G NC:N WRITE_DUMP_A:X WRITE_DUMP_B:X COMMON_1:X VCC2:P " +
                "COMMON_0:X E.GND:G ERASE_OUT:O W/C_COMP:X W/C_SET:X /WRITE_CURRENT:I /WRITE_DATA:I /WRITE_GATE:I /ERASE_GATE:I /SIDE_1:I POWER_ON:O " +
                "READ_DATA:O /MMVA_CONTROL:I VCC1:P MMVB:X MMVA:X D.GND:G NC:N MMVA_COMP:X COMP_IN_B:I COMP_IN_A:I DIFF_OUT_B:O DIFF_OUT_A:O",
                package: "QFP-44");
            Chip("74123", ["74123"], "1A:I 1B:I /1CLR:I /1Q:O 2Q:O 2CEXT:X 2RCEXT:X GND:G 2A:I 2B:I /2CLR:I /2Q:O 1Q:O 1CEXT:X 1RCEXT:X VCC:P");

            // ----- The Amiga 500's and 1000's (see the header) ---------------------------------------
            Chip("68000", ["68000", "68010", "68HC000"],
                "D4..D0:B /AS:O /UDS:O /LDS:O R//W:O /DTACK:I /BG:O /BGACK:I /BR:I VCC:P CLK:I GND:G /HALT:B /RESET:B /VMA:O E:O /VPA:I /BERR:I " +
                "/IPL2../IPL0:I FC2..FC0:O A1..A20:O VCC:P A21..A23:O GND:G D15..D5:B");

            // The A1000's Agnus (8361 NTSC, 8367 PAL), the Hardware Reference Manual's "original Agnus".
            Chip("8361", ["8361", "8367"],
                "D8..D0:B VCC:P /RES:I /INT3:O DMAL:I /BLS:I /DBR:O /ARW:O RGA8..RGA1:B CCK:I CCKQ:I VSS:G DRA0..DRA8:O /LP:I /VSY:B /CSY:O /HSY:B VSS:G " +
                "D15..D9:B");

            // The A500's Fat Agnus, a PLCC of 84 pins in a socket: the 8370 (NTSC) and 8371 (PAL) with their pin 41 for testing,
            // the 8372 with it choosing the video standard - by the A500 rev 6A schematic's names.
            Chip("8370", ["8370", "8371"], IcPinouts.FatAgnusTo40 + " TEST:I " + IcPinouts.FatAgnusFrom42, package: "PLCC-84");
            Chip("8372", ["8372"], IcPinouts.FatAgnusTo40 + " /NTSC/PAL:I " + IcPinouts.FatAgnusFrom42, package: "PLCC-84");

            // Denise (8362) and Super Denise (8373): pin 32 is no connection on the 8362 by the A500 service manual, a sync input
            // by the Hardware Reference Manual, and wired to Agnus's /CSYNC on the rev 6A board - unsettled; pin 34 is the 8373's
            // /CDAC alone.
            Chip("8362", ["8362"], IcPinouts.DeniseTo31 + " /CSYNC:? /ZD:O NC:N " + IcPinouts.DeniseFrom35);
            Chip("8373", ["8373"], IcPinouts.DeniseTo31 + " /CSYNC:? /ZD:O /CDAC:I " + IcPinouts.DeniseFrom35);

            Chip("8364", ["8364"],
                "D8..D2:B VSS:G D1:B D0:B /RES:I DMAL:O /IPL0../IPL2:O /INT2:I /INT3:I /INT6:I RGA8..RGA1:I VCC:P CCK:I CCKQ:I AUDB:O AUDA:O POT0X:B POT0Y:B " +
                "VSSANA:G POT1X:B POT1Y:B /DKRD:I /DKWD:O DKWE:O TXD:O RXD:I D15..D9:B");

            // Gary: the A500 schematic's names, the directions of Commodore's Gary specification. Pin 40 is a ground by the A500
            // service manual and no connection by the specification - unsettled.
            Chip("5719", ["5719"],
                "GND:G /VPA:O /OEL:O /OEB:O /KBRESET:I VCC:P /MTR:I DKWE:I /DKWD:I /LDS:I /UDS:I R/W:I /AS:I /BGACK:I /BLIT:I /SEL:I VCC:P /REGEN:O " +
                "/BLISS:O /RAMEN:O /ROMEN:O /CLKRD:O /CLKWR:O GND:G /LATCH:O /CDAC:I CCKQ:I CCK:I /OVR:I OVL:I XRDY:I /EXRAM:I A17..A23:I GND:? /RST:O " +
                "/HLT:O /DTACK:O DKWEB:O DKWDB:O MTRON:O MTRX:O VCC:P");

            // The 16-bit EPROM a Kickstart is burnt into for the A500 rev 6A's ROM socket, which is wired as its pins are. Its
            // pin 29 is an output in word mode and an address in byte mode.
            Chip("27C400", ["27C400"],
                "A17:I A7..A0:I /CE:I GND:G /OE:I O0:O O8:O O1:O O9:O O2:O O10:O O3:O O11:O VCC:P O4:O O12:O O5:O O13:O O6:O O14:O O7:O O15/A-1:? GND:G " +
                "/BYTE/VPP:I A16..A8:I");

            // Their RAMs: the A500's chip RAM and the A501's (NEC's uPD41256, TI's TMS4256; TI's TMS44C256, Fujitsu's MB81C4256 -
            // its pin 5 TI's test pin, left open in use, and Fujitsu's no connection).
            Chip("41256", ["41256", "4256"], "A8:I DIN:I /WE:I /RAS:I A0:I A2:I A1:I VCC:P A7:I A5:I A4:I A3:I A6:I DOUT:O /CAS:I GND:G");
            Chip("44256", ["44256", "44C256", "81C4256"], "DQ1:B DQ2:B /WE:I /RAS:I NC:N A0..A3:I VCC:P A4..A8:I /OE:I /CAS:I DQ3:B DQ4:B VSS:G");

            // The A501's and A500+'s clock (OKI's data book): its crystal's pins neither in nor out.
            Chip("6242", ["6242"], "STD.P:O /CS0:I ALE:I A0..A3:I /RD:I GND:G /WR:I D3..D0:B CS1:I XT:X /XT:X VDD:P");

            // The RGB to composite encoder (Motorola's MC1377/D): a pin only for an outside capacitor or resistor is "Other";
            // VB is the 8.2 V it gives out.
            Chip("MC1377", ["1377"],
                "TR:X SYNC:I R:I G:I B:I -YOUT:O VCLAMP:X -YIN:I CVOUT:O CHROMA_IN:I B-Y_CLAMP:X R-Y_CLAMP:X CHROMA_OUT:O VCC:P GND:G VB:O OSC_IN:I OSC_OUT:O " +
                "PHI_M:X NTSC/PAL:I");

            // The serial port's line driver and receiver (TI's and onsemi's datasheets); the receiver's response control takes a
            // capacitor or a resistor.
            Chip("MC1488", ["1488"], "VCC-:P 1A:I 1Y:O 2A:I 2B:I 2Y:O GND:G 3Y:O 3A:I 3B:I 4Y:O 4A:I 4B:I VCC+:P");
            Chip("MC1489", ["1489"], "1A:I 1CONT:X 1Y:O 2A:I 2CONT:X 2Y:O GND:G 3Y:O 3CONT:X 3A:I 4Y:O 4CONT:X 4A:I VCC:P");

            // The audio filter's quad op-amp, which has no ground: its negative supply on pin 11 (TI's LF347 and TL084).
            Chip("LF347", ["LF347", "TL084"], "1OUT:O 1IN-:I 1IN+:I VCC+:P 2IN+:I 2IN-:I 2OUT:O 3OUT:O 3IN-:I 3IN+:I VCC-:P 4IN+:I 4IN-:I 4OUT:O");
            Chip("LM2901", ["LM2901", "LM339"], "1OUT:O 2OUT:O VCC:P 2IN-:I 2IN+:I 1IN-:I 1IN+:I 3IN-:I 3IN+:I 4IN-:I 4IN+:I GND:G 4OUT:O 3OUT:O");

            // The A1000's logic CRT draws no insides of yet (TI's, Motorola's and National's datasheets): every family of these
            // numbers alike - but the 51, whose 74LS51 is another pin-out than the 7451's and 74S51's.
            Chip("74374", ["74374"], "/OE:I 1Q:O 1D:I 2D:I 2Q:O 3Q:O 3D:I 4D:I 4Q:O GND:G CLK:I 5Q:O 5D:I 6D:I 6Q:O 7Q:O 7D:I 8D:I 8Q:O VCC:P");
            Chip("74399", ["74399"], "S:I Qa:O I0a:I I1a:I I1b:I I0b:I Qb:O GND:G CP:I Qc:O I0c:I I1c:I I1d:I I0d:I Qd:O VCC:P");
            Chip("74393", ["74393"], "1A:I 1CLR:I 1QA:O 1QB:O 1QC:O 1QD:O GND:G 2QD:O 2QC:O 2QB:O 2QA:O 2CLR:I 2A:I VCC:P");
            Chip("7451", ["7451"], "1A:I 2A:I 2B:I 2C:I 2D:I 2Y:O GND:G 1Y:O 1C:I 1D:I NC:N NC:N 1B:I VCC:P", families: ["", "S"]);
            Chip("74LS51", ["7451"], "1A:I 2A:I 2B:I 2C:I 2D:I 2Y:O GND:G 1Y:O 1D:I 1E:I 1F:I 1B:I 1C:I VCC:P", families: ["LS"]);

            IcPinouts.All = all;
        }

        // ###########################################################################################
        // *** ANOTHER CHIP OF A FAMILY IS NO CPU (code review, 2026-11-17): a name matches whatever letters follow its number,
        // and only the first word counts - so "Z80 PIO", "Z80CTC" and the 1551's "6510T", chips of other pin-outs, read as
        // the Z80 and the 6510, and "Fill pins from datasheet" wrote a CPU's pins onto them.
        // ###########################################################################################
        private static readonly (string Chip, string[] Others)[] OtherChipsOfAFamily =
        [
            ("Z80", ["PIO", "CTC", "SIO", "DMA", "DART", "KIO"]),
            ("6510", ["6510T"]),

            // The A3000's 2 MB Agnus, whose pins no source here checks against the A500's.
            ("8372", ["8372B"]),
        ];

        // ###########################################################################################
        // The chip `text` names, or null for none CRT holds. Only its first word counts (ChipNames.FirstWord - "µPD4164C" is
        // "UPD4164C"). A name matches a chip's when it is that name, or that name and then a "-", a "/" or a letter (the
        // grade, the package: "4164-2", "7360/8360", "6551A"); a number may also follow a maker's letters ("MOS6526",
        // "HD6845AP"). A 74-series part is read by ChipNames.SeventyFourNumber.
        // ###########################################################################################
        public static IcPinout? Named(string? text)
        {
            string word = ChipNames.FirstWord(text);

            if (word.Length == 0)
            {
                return null;
            }

            string? ttl = ChipNames.SeventyFourNumber(word) is { } number ? "74" + number : null;
            string bare = IcPinouts.WithoutMakersLetters(word);


            foreach ((string alias, IcPinout chip) in IcPinouts.All.SelectMany(chip => chip.Aliases.Select(alias => (alias, chip))))
            {
                // "=4044" only exactly as it stands (code review, 2026-11-17): after a maker's letters, or with a suffix after it
                // ("CD4044BE", "4044BE"), it is another chip - the CMOS 4044, a latch.
                if (alias[0] == '=')
                {
                    if (word == alias[1..])
                    {
                        return chip;
                    }

                    continue;
                }

                // A 74-series number of another family than this pin-out's is the next entry's (the 74LS51's is not the 7451's).
                if (alias == ttl && chip.Families is { } families && !families.Contains(ChipNames.SeventyFourFamily(word) ?? string.Empty))
                {
                    continue;
                }

                if (alias == ttl || IcPinouts.Starts(word, alias) || (char.IsAsciiDigit(alias[0]) && IcPinouts.Starts(bare, alias)))
                {
                    return IcPinouts.IsAnotherOfItsFamily(chip, text) ? null : chip;
                }
            }

            return null;
        }

        // Whether `text`, read as `chip`, names another chip of its family - each word of it looked at on its own, so "Z80PIO"
        // and "Z80A CTC" alike, and never letters running across two words ("Z80 BUS IO" is no "SIO" - code review,
        // 2026-11-17).
        private static bool IsAnotherOfItsFamily(IcPinout chip, string? text)
        {
            List<string> words = [.. (text ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(ChipNames.FirstWord)];

            return IcPinouts.OtherChipsOfAFamily.Any(family =>
                family.Chip == chip.Name && family.Others.Any(other => words.Any(word => word.Contains(other, StringComparison.Ordinal))));
        }

        // ###########################################################################################
        // The chip whose datasheet directions go with `entry` and its `pins` (of one region) - null when CRT holds none, or the
        // board says something the datasheet does not (see the header). A package left empty is no contradiction.
        // ###########################################################################################
        public static IcPinout? For(ComponentEntry entry, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(entry);
            ArgumentNullException.ThrowIfNull(pins);

            if (ChipNames.IsNoChip(entry.Category) ||
                (IcPinouts.Named(entry.TechnicalNameOrValue) ?? IcPinouts.Named(entry.PartNumber)) is not { } chip)
            {
                return null;
            }

            if (!ChipNames.PackageFits(entry, chip.Package))
            {
                return null;
            }

            foreach (ComponentPinEntry pin in pins)
            {
                if (ComponentVocabulary.PinNumberOf(pin.Pin) is not int number)
                {
                    continue;
                }

                if (number < 1 || number > chip.PinCount || IcPinouts.Contradicts(pin, chip.DirectionOf(number)))
                {
                    return null;
                }
            }

            return chip;
        }

        // ###########################################################################################
        // Whether the name the board gives a pin is against the datasheet's `datasheet` direction of it - as LogicChips.For
        // judges a logic chip's (code review, 2026-11-17: a ground named at a supply pin, and a supply at a ground pin, passed
        // until then): a ground named where the datasheet has no ground, a supply named where it has no supply. A pin the
        // datasheet leaves unconnected or does not settle cannot be said against - a ROM's pin 1 is wired to +5V.
        //
        // *** A DIRECTION TYPED IS NO JUDGEMENT OF THE DATASHEET. *** Measured on the shipped boards the same day: the
        // 250469's KiCad project types its ROM's /CS2 and CS3, strapped to the rails, as "Ground" and "Power", and the
        // 250407's types the SID's AUDIO OUT "Power" - the datasheet is right about all three, and judging the typed direction
        // threw it away for the whole chip.
        // ###########################################################################################
        private static bool Contradicts(ComponentPinEntry pin, string datasheet)
        {
            if (datasheet is "" or "Not connected")
            {
                return false;
            }

            string name = (pin.Name ?? string.Empty).Trim();

            return (ChipNames.IsGround(name) && datasheet != "Ground") ||
                   (ChipNames.IsSupply(name) && datasheet != "Power");
        }

        // ---- reading the pin-outs -------------------------------------------------------------------

        // ###########################################################################################
        // A pin-out written as "NAME:code" words, pin 1 first - a run of numbered pins as one word, "A0..A13:O" or "D7..D0:B".
        // A word that cannot be read is a typing error, said at once.
        // ###########################################################################################
        private static IcPinout Parse(string chip, string pinout, IReadOnlyList<string> aliases)
        {
            var names = new List<string>();
            var directions = new List<string>();

            foreach (string word in pinout.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = word.LastIndexOf(':');

                if (colon <= 0 || colon != word.Length - 2 || !IcPinouts.Codes.TryGetValue(word[^1], out string? direction))
                {
                    throw new ArgumentException($"{chip}: cannot read \"{word}\".", nameof(pinout));
                }

                // A name's "_" is a space the words of a pin-out cannot hold ("PHI_COLOR" is "PHI COLOR").
                foreach (string name in IcPinouts.Run(chip, word[..colon]))
                {
                    names.Add(name.Replace('_', ' '));
                    directions.Add(direction);
                }
            }

            return new IcPinout(chip, names, directions, aliases);
        }

        // "A0..A3" is A0, A1, A2, A3; "D7..D5" is D7, D6, D5; any other name is itself.
        private static IEnumerable<string> Run(string chip, string name)
        {
            int dots = name.IndexOf("..", StringComparison.Ordinal);

            if (dots < 0)
            {
                return [name];
            }

            Match first = Regex.Match(name[..dots], @"^(?<prefix>\D*)(?<number>\d+)$");
            Match last = Regex.Match(name[(dots + 2)..], @"^(?<prefix>\D*)(?<number>\d+)$");

            if (!first.Success || !last.Success || first.Groups["prefix"].Value != last.Groups["prefix"].Value)
            {
                throw new ArgumentException($"{chip}: cannot read the run \"{name}\".", nameof(name));
            }

            string prefix = first.Groups["prefix"].Value;
            int from = int.Parse(first.Groups["number"].Value, CultureInfo.InvariantCulture);
            int to = int.Parse(last.Groups["number"].Value, CultureInfo.InvariantCulture);
            int step = to >= from ? 1 : -1;

            return Enumerable.Range(0, Math.Abs(to - from) + 1).Select(offset => prefix + (from + (offset * step)).ToString(CultureInfo.InvariantCulture));
        }

        // `word` is `alias`, or `alias` and then a "-", a "/" or a letter.
        private static bool Starts(string word, string alias) =>
            word.StartsWith(alias, StringComparison.Ordinal) &&
            (word.Length == alias.Length || word[alias.Length] is '-' or '/' || char.IsAsciiLetter(word[alias.Length]));

        // A part number without the maker's letters before it: "MOS6526" is "6526", "UPD4164C" "4164C" (a micro sign is a "U"
        // by then - ChipNames.FirstWord).
        private static string WithoutMakersLetters(string word)
        {
            int start = 0;

            while (start < word.Length && char.IsAsciiLetter(word[start]))
            {
                start++;
            }

            return word[start..];
        }
    }
}
