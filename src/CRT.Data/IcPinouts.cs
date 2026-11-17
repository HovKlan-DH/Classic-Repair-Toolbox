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
        private const string Dram64Kx4 = "/OE:I DQ1:B DQ2:B /WE:I /RAS:I A6:I A5:I A4:I VCC:P A7:I A3:I A2:I A1:I A0:I DQ3:B /CAS:I DQ4:B VSS:G";
        private const string Pla82S100 = "FE:I I7..I0:I F7..F4:O GND:G F3..F0:O /CE:I I15..I8:I VCC:P";

        // The catalogue - each chip with the names it is known by, in the order Named tries them.
        public static readonly IReadOnlyList<IcPinout> All;

        static IcPinouts()
        {
            var all = new List<IcPinout>();

            void Chip(string name, string[] aliases, string pinout) =>
                all.Add(IcPinouts.Parse(name, pinout, [.. aliases.Select(alias => alias.ToUpperInvariant())]));

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
            Chip("6526", ["6526", "8521"], "VSS:G PA0..PA7:B PB0..PB7:B /PC:O TOD:I VCC:P /IRQ:O R/W:I /CS:I /FLAG:I PHI2:I D7..D0:B /RES:I RS3..RS0:I SP:B CNT:B");
            Chip("6529", ["6529"], "R/W:I P0..P7:B VSS:G D7..D0:B /CS:I VDD:P");
            Chip("6551", ["6551"], "VSS:G CS0:I /CS1:I /RES:I RXC:B XTAL1:I XTAL2:X /RTS:O /CTS:I TXD:O /DTR:O RXD:I RS0:I RS1:I VCC:P /DCD:I /DSR:I D0..D7:B /IRQ:O PHI2:I R/W:I");
            Chip("8255", ["8255"], "PA3..PA0:B /RD:I /CS:I GND:G A1:I A0:I PC7..PC4:B PC0..PC3:B PB0..PB7:B VCC:P D7..D0:B RESET:I /WR:I PA7..PA4:B");
            Chip("8722", ["8722"],
                "VDD:P /RES:I TA15..TA8:O /CAS1:O /CAS0:O I/O_SEL:O MS1:O MS0:O AEC:I MUX:I A0..A3:I A4/5:I A6/7:I A8..A15:I R/W:I PHI0:I VSS:G D0..D7:B " +
                "/Z80EN:O FSDIR:O /GAME:B /EXROM:B 128/64:O SENSE40:I");

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
            Chip("74123", ["74123"], "1A:I 1B:I /1CLR:I /1Q:O 2Q:O 2CEXT:X 2RCEXT:X GND:G 2A:I 2B:I /2CLR:I /2Q:O 1Q:O 1CEXT:X 1RCEXT:X VCC:P");

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

            if (!ChipNames.PackageFits(entry, chip.PinCount))
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
