using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What a unit of a chip does. AND, OR, XOR and a buffer (one input) are gates - NAND, NOR, XNOR and an inverter are
    // these with their output inverted (LogicUnit.Inverted); a 2-to-1 multiplexer (the 157's) and an n-to-1 selector (the
    // 153's); an analog switch (the 4066's); a decoder; a D latch, an edge-triggered D flip-flop and an up/down counter; a
    // bus transceiver (the 245's); and a 555 timer.
    // ###########################################################################################
    public enum LogicFunction
    {
        And,
        Or,
        Xor,
        Buffer,
        Mux,
        Selector,
        Switch,
        Decoder,
        Latch,
        FlipFlop,
        Counter,
        Transceiver,
        Timer,
    }

    // How a chip's outputs drive their pins, as the datasheet says.
    public enum LogicOutputStage
    {
        TotemPole,
        OpenCollector,
        ThreeState,
    }

    // ###########################################################################################
    // What a pin is to its unit. An input of a gate, a data input (a selector's, numbered by Index; a decoder's or
    // counter's, weighted by Index) or one side of a switch or transceiver (Index 0 and 1); an output (a decoder's
    // numbered, a counter's weighted - and a flip-flop's Q-bar is its output with ActiveLow); a select (weighted); an
    // enable; a clock (a latch's enable too - the datasheets' "C"); a set and a reset; a counter's load, up, down, carry and
    // borrow; a switch's control and a timer's control voltage; a timer's trigger, threshold and discharge; a
    // transceiver's direction.
    // ###########################################################################################
    public enum TerminalRole
    {
        Input,
        Output,
        Select,
        Enable,
        Clock,
        Set,
        Reset,
        Load,
        CountUp,
        CountDown,
        Carry,
        Borrow,
        Control,
        Trigger,
        Threshold,
        Discharge,
        Direction,
    }

    // How LogicChipGeometry lays a chip's insides out - by what its units are and where their pins are.
    public enum LogicLayout
    {
        Gates,
        Buffers,
        Transceivers,
        Switches,
        FlipFlops,
        Columns,
        Blocks,
        Timers,
    }

    // A pin of a unit: what it is to it, active low or not (an output with ActiveLow is the inverted one), and its number
    // or weight.
    public sealed record LogicTerminal(TerminalRole Role, int Pin, bool ActiveLow = false, int Index = 0);

    // ###########################################################################################
    // One gate, multiplexer, switch, decoder, latch, flip-flop, counter, transceiver or timer of a chip: what it does,
    // whether its output is inverted (a gate's), and its pins - each once, in the datasheet's order.
    // ###########################################################################################
    public sealed record LogicUnit(LogicFunction Function, bool Inverted, IReadOnlyList<LogicTerminal> Terminals)
    {
        // Every pin it uses.
        public IEnumerable<int> Pins => this.Terminals.Select(terminal => terminal.Pin).Distinct();

        // Its input pins in order - a gate's inputs, a multiplexer's A then B.
        public IReadOnlyList<int> Inputs => this.Of(TerminalRole.Input).Select(terminal => terminal.Pin).ToList();

        // Its (first) output pin.
        public int Output => this.Of(TerminalRole.Output).First().Pin;

        // Its pins of one role, in order.
        public IEnumerable<LogicTerminal> Of(TerminalRole role) => this.Terminals.Where(terminal => terminal.Role == role);
    }

    // ###########################################################################################
    // A chip whose insides CRT draws: its family ("74", "40" or "55") and number, its pins' names as the datasheet's
    // connection diagram writes them (index 0 is pin 1), its units, how its outputs drive, Schmitt-trigger inputs, and how
    // its insides are laid out. What it is in words is not written under the drawing (owner request, 2026-11-16: "the name
    // is already given elsewhere (e.g. in header line)").
    // ###########################################################################################
    public sealed record LogicChip(
        string Family,
        string Number,
        IReadOnlyList<string> PinNames,
        IReadOnlyList<LogicUnit> Units,
        LogicOutputStage Outputs,
        bool Schmitt,
        LogicLayout Layout = LogicLayout.Gates)
    {
        public int PinCount => this.PinNames.Count;

        // The case its pin-out is of: every logic chip's is a DIP's, which its SOIC and the like share (ChipNames.PackageFits).
        public string Package => $"DIP-{this.PinCount.ToString(System.Globalization.CultureInfo.InvariantCulture)}";

        // How the chip is known: "74LS08" is a "7408", CMOS and the timers keep their own numbers.
        public string Name => this.Family == "74" ? "74" + this.Number : this.Number;

        // The ground pin (GND, VSS) and the supply pin (VCC, VDD).
        public int Ground => this.PinOf("GND", "VSS");

        public int Supply => this.PinOf("VCC", "VDD");

        // ###########################################################################################
        // The datasheet's name of `pin` as the board data writes it - "" for none. AN ACTIVE-LOW PIN - one its units use active
        // low: an enable, a clear, a decoder's output, a Q-bar - is written with "/" before it, as the datasheet draws a bar over
        // it (owner decision, 2026-11-17: "use "/" everywhere"): "/G2A", "/OE", "/1CLR", "/Y0", and a Q-bar's "1QN" as "/1Q".
        // So a 74125's enable is "/1OE" and a 74126's "1OE". PinNames keep the datasheet models' own spelling.
        // ###########################################################################################
        public string NameOf(int pin)
        {
            if (pin < 1 || pin > this.PinNames.Count)
            {
                return string.Empty;
            }

            string name = this.PinNames[pin - 1];
            List<LogicTerminal> terminals = this.Units.SelectMany(unit => unit.Terminals).Where(terminal => terminal.Pin == pin).ToList();

            if (!terminals.Any(terminal => terminal.ActiveLow))
            {
                return name;
            }

            bool qBar = name.EndsWith("QN", StringComparison.Ordinal) && terminals.Any(terminal => terminal.Role == TerminalRole.Output && terminal.ActiveLow);

            return "/" + (qBar ? name[..^1] : name);
        }

        // ###########################################################################################
        // What `pin` is as the datasheet says it, in the pin list's words (ComponentVocabulary.PinDirections): the ground and
        // the supply "Ground" and "Power"; a pin no unit uses "Not connected"; a transceiver's A or B, or a switch's side, which
        // passes a signal either way, "Input/output"; what a unit drives - an output, a carry, a borrow, a timer's discharge -
        // "Output"; every other terminal - an input, a select, an enable, a clock, a control - "Input".
        // ###########################################################################################
        public string DirectionOf(int pin)
        {
            if (pin == this.Ground)
            {
                return "Ground";
            }

            if (pin == this.Supply)
            {
                return "Power";
            }

            List<(LogicUnit Unit, LogicTerminal Terminal)> uses = this.Units
                .SelectMany(unit => unit.Terminals.Where(terminal => terminal.Pin == pin).Select(terminal => (unit, terminal)))
                .ToList();

            return uses.Count == 0 ? "Not connected"
                : uses.Any(use => use.Unit.Function is LogicFunction.Transceiver or LogicFunction.Switch && use.Terminal.Role == TerminalRole.Input) ? "Input/output"
                : uses.Any(use => use.Terminal.Role is TerminalRole.Output or TerminalRole.Carry or TerminalRole.Borrow or TerminalRole.Discharge) ? "Output"
                : "Input";
        }

        private int PinOf(params string[] names) =>
            this.PinNames.Select((name, index) => (name, index)).First(pin => names.Contains(pin.name)).index + 1;
    }

    // ###########################################################################################
    // THE CHIPS WHOSE INSIDES CRT DRAWS (owner request, 2026-11-16: "is it then also possible to visualize its internals for
    // e.g. its logic cirquits ... the very simple chips like AND, OR, XOR, NAND etc."; "OK, do it" to the proposal: the gates
    // and the 157/257 multiplexers first, drawn in place of the plain case, the pins colored by their expected reading - by
    // their direction since 2026-11-17 (the app's PinDirectionKinds); and
    // the same day "OK, now go ahead and make the rest": the buffers, the 4066's switches, the decoders, the 153, the
    // latches and flip-flops, the 193 counter, the 245 transceiver and the 555 and 556 timers - every other logic chip on the
    // published boards in November 2026 but the 74629's oscillators, the 74123's monostables and the 4044 - the MC4044 phase
    // detector a C64 board names so, whose directions IcPinouts holds).
    //
    // TYPED IN FROM THE DATASHEETS' CONNECTION DIAGRAMS, never read from a board: each pin-out is written as the datasheet
    // draws it, pin 1 first, the very strings of the MiniPro tests' datasheet models (tests/CRT.App.Tests/
    // LogicIcReferenceModels.cs) wherever a model exists - and LogicChipsTests holds every chip with a model against it:
    // every input combination of a combinational chip, the 4066's switches both ways, and a sequential chip through a long
    // run of pin changes - so a wrong gate, a swapped pin or a forgotten bubble fails the suite. A gate chip with no model
    // shares its pin-out with one that has (the 7409 is the 7408's, open collector), and the 74175, 74273, 74245 and 555 -
    // no model, no sibling - are spelt out pin by pin there. A gate's units come from the names as the models' Gates helper
    // reads them: "<n>Y" is an output and the other "<n>..." its inputs; a plain "Y" takes every other input.
    //
    // A COMPONENT IS MATCHED BY ITS NUMBER, whatever maker's prefix, logic family or suffix it is written with ("74LS08",
    // "SN74HCT08N", "7408", "HD74LS08P", "MC14011BCP", "CD4011BE", "NE555P", "LM556CN") - its "Technical name or value"
    // first, its part number after - and drawn only when nothing the board says contradicts the datasheet (For): a category
    // no chip has (ChipNames.IsNoChip - a "5400pF" capacitor is no 7400), a package of another pin count, a pin past the
    // case, a ground or supply pin NAMED where the datasheet has none. A wrong picture is worse than none - but a direction
    // typed is no judgement of the chip (code review, 2026-11-17): KiCad types inputs strapped to a rail "Power".
    // ###########################################################################################
    public static class LogicChips
    {
        // The pin-outs several parts share - the MiniPro models' strings, and CMOS's.
        private const string Quad2Input = "1A 1B 1Y 2A 2B 2Y GND 3Y 3A 3B 4Y 4A 4B VCC";       // 00 03 08 09 26 32 37 38 86 132 136
        private const string Quad2InputNor = "1Y 1A 1B 2Y 2A 2B GND 3A 3B 3Y 4A 4B 4Y VCC";    // 02 28 33
        private const string Hex1Input = "1A 1Y 2A 2Y 3A 3Y GND 4Y 4A 5Y 5A 6Y 6A VCC";        // 04 05 06 07 14 16 17
        private const string Triple3Input = "1A 1B 2A 2B 2C 2Y GND 3Y 3A 3B 3C 1Y 1C VCC";     // 10 11 12 15 27
        private const string Dual4Input = "1A 1B NC 1C 1D 1Y GND 2Y 2A 2B NC 2C 2D VCC";       // 20 21 22 40
        private const string Quad2InputSplit = "1A 1B 1Y 2Y 2A 2B GND 3A 3B 3Y 4Y 4A 4B VCC";  // 266
        private const string Quad2To1 = "S 1A 1B 1Y 2A 2B 2Y GND 3Y 3B 3A 4Y 4B 4A G VCC";     // 157 158 257 258
        private const string QuadBuffer = "1OE 1A 1Y 2OE 2A 2Y GND 3Y 3A 3OE 4Y 4A 4OE VCC";   // 125 126
        private const string Octal = "1G 1A1 2Y4 1A2 2Y3 1A3 2Y2 1A4 2Y1 GND 2A1 1Y4 2A2 1Y3 2A3 1Y2 2A4 1Y1 2G VCC"; // 240 244
        private const string OctalLatch = "OE 1Q 1D 2D 2Q 3Q 3D 4D 4Q GND LE 5Q 5D 6D 6Q 7Q 7D 8D 8Q VCC";           // 373
        private const string OctalFlipFlop = "CLR 1Q 1D 2D 2Q 3Q 3D 4D 4Q GND CLK 5Q 5D 6D 6Q 7Q 7D 8D 8Q VCC";       // 273

        private const string CmosQuad2Input = "1A 1B 1Y 2Y 2A 2B VSS 3A 3B 3Y 4Y 4A 4B VDD";   // 4001 4011 4070 4071 4077 4081 4093
        private const string CmosHex1Input = "1A 1Y 2A 2Y 3A 3Y VSS 4Y 4A 5Y 5A 6Y 6A VDD";    // 4069

        private const LogicOutputStage TotemPole = LogicOutputStage.TotemPole;
        private const LogicOutputStage OpenCollector = LogicOutputStage.OpenCollector;
        private const LogicOutputStage ThreeState = LogicOutputStage.ThreeState;

        public static readonly IReadOnlyList<LogicChip> All =
        [
            // ----- 74 series: gates --------------------------------------------------------------
            Gates("74", "00", Quad2Input, LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "02", Quad2InputNor, LogicFunction.Or, inverted: true, TotemPole),
            Gates("74", "03", Quad2Input, LogicFunction.And, inverted: true, OpenCollector),
            Gates("74", "04", Hex1Input, LogicFunction.Buffer, inverted: true, TotemPole),
            Gates("74", "05", Hex1Input, LogicFunction.Buffer, inverted: true, OpenCollector),
            Gates("74", "06", Hex1Input, LogicFunction.Buffer, inverted: true, OpenCollector),
            Gates("74", "07", Hex1Input, LogicFunction.Buffer, inverted: false, OpenCollector),
            Gates("74", "08", Quad2Input, LogicFunction.And, inverted: false, TotemPole),
            Gates("74", "09", Quad2Input, LogicFunction.And, inverted: false, OpenCollector),
            Gates("74", "10", Triple3Input, LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "11", Triple3Input, LogicFunction.And, inverted: false, TotemPole),
            Gates("74", "12", Triple3Input, LogicFunction.And, inverted: true, OpenCollector),
            Gates("74", "14", Hex1Input, LogicFunction.Buffer, inverted: true, TotemPole, schmitt: true),
            Gates("74", "15", Triple3Input, LogicFunction.And, inverted: false, OpenCollector),
            Gates("74", "16", Hex1Input, LogicFunction.Buffer, inverted: true, OpenCollector),
            Gates("74", "17", Hex1Input, LogicFunction.Buffer, inverted: false, OpenCollector),
            Gates("74", "20", Dual4Input, LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "21", Dual4Input, LogicFunction.And, inverted: false, TotemPole),
            Gates("74", "22", Dual4Input, LogicFunction.And, inverted: true, OpenCollector),
            Gates("74", "26", Quad2Input, LogicFunction.And, inverted: true, OpenCollector),
            Gates("74", "27", Triple3Input, LogicFunction.Or, inverted: true, TotemPole),
            Gates("74", "28", Quad2InputNor, LogicFunction.Or, inverted: true, TotemPole),
            Gates("74", "30", "A B C D E F GND Y NC NC G H NC VCC", LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "32", Quad2Input, LogicFunction.Or, inverted: false, TotemPole),
            Gates("74", "33", Quad2InputNor, LogicFunction.Or, inverted: true, OpenCollector),
            Gates("74", "37", Quad2Input, LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "38", Quad2Input, LogicFunction.And, inverted: true, OpenCollector),
            Gates("74", "40", Dual4Input, LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "86", Quad2Input, LogicFunction.Xor, inverted: false, TotemPole),
            Gates("74", "132", Quad2Input, LogicFunction.And, inverted: true, TotemPole, schmitt: true),
            Gates("74", "133", "A B C D E F G GND Y H I J K L M VCC", LogicFunction.And, inverted: true, TotemPole),
            Gates("74", "136", Quad2Input, LogicFunction.Xor, inverted: false, OpenCollector),
            Gates("74", "266", Quad2InputSplit, LogicFunction.Xor, inverted: true, OpenCollector),

            // ----- 74 series: buffers and the transceiver ----------------------------------------------
            QuadBuffers("125", enableActiveLow: true),
            QuadBuffers("126", enableActiveLow: false),
            OctalBuffers("240", inverted: true),
            OctalBuffers("244", inverted: false),
            Transceiver245(),

            // ----- 74 series: multiplexers and decoders -------------------------------------------------
            Selector153(),
            Multiplexers("74", "157", inverted: false, TotemPole),
            Multiplexers("74", "158", inverted: true, TotemPole),
            Multiplexers("74", "257", inverted: false, ThreeState),
            Multiplexers("74", "258", inverted: true, ThreeState),
            Decoder138(),
            Decoder139(),
            Decoder145(),

            // ----- 74 series: flip-flops, latches and the counter -----------------------------------------
            FlipFlops74(),
            FlipFlops175(),
            OctalFlipFlops273(),
            OctalLatches373(),
            Counter193(),

            // ----- CMOS 4000 series ---------------------------------------------------------------------------
            Gates("40", "4001", CmosQuad2Input, LogicFunction.Or, inverted: true, TotemPole),
            Gates("40", "4011", CmosQuad2Input, LogicFunction.And, inverted: true, TotemPole),
            Switches4066(),
            Gates("40", "4069", CmosHex1Input, LogicFunction.Buffer, inverted: true, TotemPole),
            Gates("40", "4070", CmosQuad2Input, LogicFunction.Xor, inverted: false, TotemPole),
            Gates("40", "4071", CmosQuad2Input, LogicFunction.Or, inverted: false, TotemPole),
            Gates("40", "4077", CmosQuad2Input, LogicFunction.Xor, inverted: true, TotemPole),
            Gates("40", "4081", CmosQuad2Input, LogicFunction.And, inverted: false, TotemPole),
            Gates("40", "4093", CmosQuad2Input, LogicFunction.And, inverted: true, TotemPole, schmitt: true),

            // ----- timers -----------------------------------------------------------------------------------------
            Timer555(),
            Timers556(),
        ];

        private const RegexOptions Words = RegexOptions.CultureInvariant | RegexOptions.IgnoreCase;

        // ###########################################################################################
        // A 74-series name is ChipNames.SeventyFourNumber's - "SN74LS08N".
        // CMOS: CD, HEF, MC1, TC, MM, HCF or none, the 4-digit number, letters after it - "MC14011BCP", "CD4011BE".
        // A timer: NE, SE, SA, LM, UA, TLC, TS, KA or none, 555 or 556, a suffix that starts with a letter - "NE555P" (a
        // micro sign or Greek mu is read as the "U" of "UA555" - ChipNames.FirstWord).
        // ###########################################################################################
        private static readonly Regex FourThousand = new(
            @"^(?:CD|HEF|MC1|TC|MM|HCF)?(?<number>40\d\d)(?:[A-Z]{1,4}\d?)?$",
            Words);

        private static readonly Regex Timer = new(
            @"^(?:NE|SE|SA|LM|UA|TLC|TS|KA)?(?<number>55[56])(?:[A-Z][A-Z0-9-]*)?$",
            Words);

        // The chip `text` names - "74LS08", "SN74HCT08N", "CD4011BE", "NE555P" - or null for none CRT draws. Only its first
        // word counts.
        public static LogicChip? Named(string? text)
        {
            string word = ChipNames.FirstWord(text);

            if (word.Length == 0)
            {
                return null;
            }

            // A 74-series part carrying a CMOS number - "74HC4066" - is that CMOS chip, pin for pin (code review, 2026-11-17).
            if (ChipNames.SeventyFourNumber(word) is { } ttl)
            {
                return LogicChips.Find("74", ttl) ?? (ttl.Length == 4 && ttl.StartsWith("40", StringComparison.Ordinal) ? LogicChips.Find("40", ttl) : null);
            }

            if (FourThousand.Match(word) is { Success: true } cmos)
            {
                return LogicChips.Find("40", cmos.Groups["number"].Value);
            }

            if (Timer.Match(word) is { Success: true } timer)
            {
                return LogicChips.Find("55", timer.Groups["number"].Value);
            }

            return null;
        }

        // ###########################################################################################
        // The chip whose insides are drawn for `entry` with its `pins` (of the region shown) - null when CRT draws none: no
        // chip it knows by that number, or the board saying something the datasheet does not (see the header). A package
        // left empty is no contradiction: the chip says its own pin count.
        // ###########################################################################################
        public static LogicChip? For(ComponentEntry entry, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(entry);
            ArgumentNullException.ThrowIfNull(pins);

            if (ChipNames.IsNoChip(entry.Category) ||
                (LogicChips.Named(entry.TechnicalNameOrValue) ?? LogicChips.Named(entry.PartNumber)) is not { } chip)
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

                // The pin's NAME is judged, never a direction typed (code review, 2026-11-17 - as IcPinouts judges): KiCad types
                // an input strapped to a rail "Power" or "Ground", and the chip is the datasheet's all the same.
                string name = (pin.Name ?? string.Empty).Trim();

                if (number < 1 ||
                    number > chip.PinCount ||
                    (ChipNames.IsGround(name) && number != chip.Ground) ||
                    (ChipNames.IsSupply(name) && number != chip.Supply))
                {
                    return null;
                }
            }

            return chip;
        }

        private static LogicChip? Find(string family, string number)
        {
            // "7408" and "74LS08" are one chip, "08" read of both; a number with a leading zero CRT has none of ("74008") is
            // none (code review, 2026-11-17: the zeros were taken off and "74LS008" drawn as a 7408).
            return LogicChips.All.FirstOrDefault(chip => chip.Family == family && chip.Number == number);
        }

        // ---- building the units from the datasheets' names -------------------------------------------

        // A pin-out's names, and the pin a name is - a name the pin-out does not have is a typing error, said at once.
        private sealed class Pinout(string names)
        {
            public string[] Names { get; } = names.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            public int this[string name]
            {
                get
                {
                    int index = Array.IndexOf(this.Names, name);

                    return index >= 0 ? index + 1 : throw new ArgumentException($"No pin named {name} in \"{string.Join(' ', this.Names)}\".", nameof(name));
                }
            }
        }

        private static string N(int number) => number.ToString(CultureInfo.InvariantCulture);

        private static LogicTerminal Term(TerminalRole role, int pin, bool activeLow = false, int index = 0) => new(role, pin, activeLow, index);

        private static LogicChip Chip(string family, string number, Pinout pinout, IEnumerable<LogicUnit> units, LogicOutputStage stage, LogicLayout layout, bool schmitt = false) =>
            new(family, number, pinout.Names, units.ToList(), stage, schmitt, layout);

        // ###########################################################################################
        // A chip of identical gates from its datasheet pin-out: every pin named "<n>Y" (or plain "Y") is an output, and its
        // inputs are the other pins named "<n>..." (for "Y": every pin that is no supply, no NC and no output), in pin order -
        // the models' Gates helper's reading of the same strings.
        // ###########################################################################################
        private static LogicChip Gates(
            string family, string number, string pinout, LogicFunction function, bool inverted, LogicOutputStage stage, bool schmitt = false)
        {
            string[] names = pinout.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            List<LogicUnit> units = names
                .Select((name, index) => (name, pin: index + 1))
                .Where(pin => pin.name.EndsWith('Y'))
                .Select(output =>
                {
                    string prefix = output.name[..^1];

                    List<LogicTerminal> terminals = names
                        .Select((name, index) => (name, pin: index + 1))
                        .Where(pin => pin.name != output.name &&
                                      pin.name.StartsWith(prefix, StringComparison.Ordinal) &&
                                      !pin.name.EndsWith('Y') &&
                                      !LogicChips.IsSupplyOrNc(pin.name))
                        .Select(pin => Term(TerminalRole.Input, pin.pin))
                        .Append(Term(TerminalRole.Output, output.pin))
                        .ToList();

                    return new LogicUnit(function, inverted, terminals);
                })
                .ToList();

            return new LogicChip(family, number, names, units, stage, schmitt);
        }

        // Four 2-to-1 multiplexers sharing the select S and the enable G (the 157's family): nA picked while S is low, nB
        // while it is high, onto nY; G high turns them off.
        private static LogicChip Multiplexers(string family, string number, bool inverted, LogicOutputStage stage)
        {
            var pins = new Pinout(Quad2To1);

            return Chip(family, number, pins, Enumerable.Range(1, 4).Select(n => new LogicUnit(LogicFunction.Mux, inverted,
            [
                Term(TerminalRole.Input, pins[N(n) + "A"], index: 0),
                Term(TerminalRole.Input, pins[N(n) + "B"], index: 1),
                Term(TerminalRole.Output, pins[N(n) + "Y"]),
                Term(TerminalRole.Select, pins["S"], index: 1),
                Term(TerminalRole.Enable, pins["G"], activeLow: true),
            ])), stage, LogicLayout.Gates);
        }

        // The 125 and 126: four 3-state buffers, each with its own enable - active low on the 125, high on the 126.
        private static LogicChip QuadBuffers(string number, bool enableActiveLow)
        {
            var pins = new Pinout(QuadBuffer);

            return Chip("74", number, pins, Enumerable.Range(1, 4).Select(n => new LogicUnit(LogicFunction.Buffer, false,
            [
                Term(TerminalRole.Input, pins[N(n) + "A"]),
                Term(TerminalRole.Output, pins[N(n) + "Y"]),
                Term(TerminalRole.Enable, pins[N(n) + "OE"], activeLow: enableActiveLow),
            ])), ThreeState, LogicLayout.Gates);
        }

        // The 240 and 244: eight 3-state buffers in two groups of four, each group turned on by its G (active low) - the
        // first group from the left pins to the right, the second the other way; the 240's invert.
        private static LogicChip OctalBuffers(string number, bool inverted)
        {
            var pins = new Pinout(Octal);

            return Chip("74", number, pins, new[] { 1, 2 }.SelectMany(group => Enumerable.Range(1, 4).Select(n => new LogicUnit(LogicFunction.Buffer, inverted,
            [
                Term(TerminalRole.Input, pins[N(group) + "A" + N(n)]),
                Term(TerminalRole.Output, pins[N(group) + "Y" + N(n)]),
                Term(TerminalRole.Enable, pins[N(group) + "G"], activeLow: true),
            ]))), ThreeState, LogicLayout.Buffers);
        }

        // ###########################################################################################
        // The 245: eight bus transceivers between An and Bn. OE (active low) turns them on, DIR says which way: high from A
        // to B, low from B to A.
        // ###########################################################################################
        private static LogicChip Transceiver245()
        {
            var pins = new Pinout("DIR A1 A2 A3 A4 A5 A6 A7 A8 GND B8 B7 B6 B5 B4 B3 B2 B1 OE VCC");

            return Chip("74", "245", pins, Enumerable.Range(1, 8).Select(n => new LogicUnit(LogicFunction.Transceiver, false,
            [
                Term(TerminalRole.Input, pins["A" + N(n)], index: 0),
                Term(TerminalRole.Input, pins["B" + N(n)], index: 1),
                Term(TerminalRole.Direction, pins["DIR"]),
                Term(TerminalRole.Enable, pins["OE"], activeLow: true),
            ])), ThreeState, LogicLayout.Transceivers);
        }

        // ###########################################################################################
        // The 153: two 4-to-1 selectors sharing the selects A (weight 1) and B (weight 2), each with its strobe nG (active
        // low - high holds its output low) and data inputs nC0 to nC3.
        // ###########################################################################################
        private static LogicChip Selector153()
        {
            var pins = new Pinout("1G B 1C3 1C2 1C1 1C0 1Y GND 2Y 2C0 2C1 2C2 2C3 A 2G VCC");

            return Chip("74", "153", pins, new[] { 1, 2 }.Select(n => new LogicUnit(LogicFunction.Selector, false,
                Enumerable.Range(0, 4).Select(data => Term(TerminalRole.Input, pins[N(n) + "C" + N(data)], index: data))
                    .Append(Term(TerminalRole.Output, pins[N(n) + "Y"]))
                    .Append(Term(TerminalRole.Select, pins["A"], index: 1))
                    .Append(Term(TerminalRole.Select, pins["B"], index: 2))
                    .Append(Term(TerminalRole.Enable, pins[N(n) + "G"], activeLow: true))
                    .ToList())), TotemPole, LogicLayout.Gates);
        }

        // The 138: one 3-to-8 decoder - A, B, C weighted 1, 2, 4 - on while G1 is high and G2A, G2B low; the output picked
        // goes low.
        private static LogicChip Decoder138()
        {
            var pins = new Pinout("A B C G2A G2B G1 Y7 GND Y6 Y5 Y4 Y3 Y2 Y1 Y0 VCC");

            return Chip("74", "138", pins,
            [
                new LogicUnit(LogicFunction.Decoder, false,
                    new[] { ("A", 1), ("B", 2), ("C", 4) }.Select(input => Term(TerminalRole.Input, pins[input.Item1], index: input.Item2))
                        .Append(Term(TerminalRole.Enable, pins["G2A"], activeLow: true))
                        .Append(Term(TerminalRole.Enable, pins["G2B"], activeLow: true))
                        .Append(Term(TerminalRole.Enable, pins["G1"]))
                        .Concat(Enumerable.Range(0, 8).Select(output => Term(TerminalRole.Output, pins["Y" + N(output)], activeLow: true, index: output)))
                        .ToList()),
            ], TotemPole, LogicLayout.Blocks);
        }

        // The 139: two 2-to-4 decoders, each on while its G is low, the output picked going low.
        private static LogicChip Decoder139()
        {
            var pins = new Pinout("1G 1A 1B 1Y0 1Y1 1Y2 1Y3 GND 2Y3 2Y2 2Y1 2Y0 2B 2A 2G VCC");

            return Chip("74", "139", pins, new[] { 1, 2 }.Select(n => new LogicUnit(LogicFunction.Decoder, false,
                new[] { Term(TerminalRole.Enable, pins[N(n) + "G"], activeLow: true), Term(TerminalRole.Input, pins[N(n) + "A"], index: 1), Term(TerminalRole.Input, pins[N(n) + "B"], index: 2) }
                    .Concat(Enumerable.Range(0, 4).Select(output => Term(TerminalRole.Output, pins[N(n) + "Y" + N(output)], activeLow: true, index: output)))
                    .ToList())), TotemPole, LogicLayout.Blocks);
        }

        // The 145: a BCD-to-decimal decoder - A to D weighted 1 to 8 - pulling the output picked low (open collector); the
        // codes 10 to 15 pull none.
        private static LogicChip Decoder145()
        {
            var pins = new Pinout("Y0 Y1 Y2 Y3 Y4 Y5 Y6 GND Y7 Y8 Y9 D C B A VCC");

            return Chip("74", "145", pins,
            [
                new LogicUnit(LogicFunction.Decoder, false,
                    new[] { ("A", 1), ("B", 2), ("C", 4), ("D", 8) }.Select(input => Term(TerminalRole.Input, pins[input.Item1], index: input.Item2))
                        .Concat(Enumerable.Range(0, 10).Select(output => Term(TerminalRole.Output, pins["Y" + N(output)], activeLow: true, index: output)))
                        .ToList()),
            ], OpenCollector, LogicLayout.Blocks);
        }

        // ###########################################################################################
        // The 7474: two D flip-flops, each taking D on the rising edge of its CLK; PRE and CLR (active low) set and clear it
        // at once. QN is the inverted output.
        // ###########################################################################################
        private static LogicChip FlipFlops74()
        {
            var pins = new Pinout("1CLR 1D 1CLK 1PRE 1Q 1QN GND 2QN 2Q 2PRE 2CLK 2D 2CLR VCC");

            return Chip("74", "74", pins, new[] { 1, 2 }.Select(n => new LogicUnit(LogicFunction.FlipFlop, false,
            [
                Term(TerminalRole.Reset, pins[N(n) + "CLR"], activeLow: true),
                Term(TerminalRole.Input, pins[N(n) + "D"]),
                Term(TerminalRole.Clock, pins[N(n) + "CLK"]),
                Term(TerminalRole.Set, pins[N(n) + "PRE"], activeLow: true),
                Term(TerminalRole.Output, pins[N(n) + "Q"]),
                Term(TerminalRole.Output, pins[N(n) + "QN"], activeLow: true),
            ])), TotemPole, LogicLayout.FlipFlops);
        }

        // The 74175: four D flip-flops sharing CLK (rising edge) and CLR (active low), each with Q and its inverted QN.
        private static LogicChip FlipFlops175()
        {
            var pins = new Pinout("CLR 1Q 1QN 1D 2D 2QN 2Q GND CLK 3Q 3QN 3D 4D 4QN 4Q VCC");

            return Chip("74", "175", pins, Enumerable.Range(1, 4).Select(n => new LogicUnit(LogicFunction.FlipFlop, false,
            [
                Term(TerminalRole.Input, pins[N(n) + "D"]),
                Term(TerminalRole.Output, pins[N(n) + "Q"]),
                Term(TerminalRole.Output, pins[N(n) + "QN"], activeLow: true),
                Term(TerminalRole.Clock, pins["CLK"]),
                Term(TerminalRole.Reset, pins["CLR"], activeLow: true),
            ])), TotemPole, LogicLayout.Columns);
        }

        // The 74273: eight D flip-flops sharing CLK (rising edge) and CLR (active low).
        private static LogicChip OctalFlipFlops273()
        {
            var pins = new Pinout(OctalFlipFlop);

            return Chip("74", "273", pins, Enumerable.Range(1, 8).Select(n => new LogicUnit(LogicFunction.FlipFlop, false,
            [
                Term(TerminalRole.Input, pins[N(n) + "D"]),
                Term(TerminalRole.Output, pins[N(n) + "Q"]),
                Term(TerminalRole.Clock, pins["CLK"]),
                Term(TerminalRole.Reset, pins["CLR"], activeLow: true),
            ])), TotemPole, LogicLayout.Columns);
        }

        // The 74373: eight D latches - each Q follows its D while LE is high and holds while it is low - behind 3-state
        // outputs OE (active low) turns on.
        private static LogicChip OctalLatches373()
        {
            var pins = new Pinout(OctalLatch);

            return Chip("74", "373", pins, Enumerable.Range(1, 8).Select(n => new LogicUnit(LogicFunction.Latch, false,
            [
                Term(TerminalRole.Input, pins[N(n) + "D"]),
                Term(TerminalRole.Output, pins[N(n) + "Q"]),
                Term(TerminalRole.Clock, pins["LE"]),
                Term(TerminalRole.Enable, pins["OE"], activeLow: true),
            ])), ThreeState, LogicLayout.Columns);
        }

        // ###########################################################################################
        // The 74193: a 4-bit up/down counter - QA to QD weighted 1 to 8 - counting up on UP's rising edge and down on
        // DOWN's; CLR (active high) clears it and LOAD (active low) loads A to D; CO goes low at 15 while UP is low, BO at 0
        // while DOWN is low.
        // ###########################################################################################
        private static LogicChip Counter193()
        {
            var pins = new Pinout("B QB QA DOWN UP QC QD GND D C LOAD CO BO CLR A VCC");
            (string Name, int Weight)[] bits = [("A", 1), ("B", 2), ("C", 4), ("D", 8)];

            return Chip("74", "193", pins,
            [
                new LogicUnit(LogicFunction.Counter, false,
                    bits.Select(bit => Term(TerminalRole.Input, pins[bit.Name], index: bit.Weight))
                        .Concat(bits.Select(bit => Term(TerminalRole.Output, pins["Q" + bit.Name], index: bit.Weight)))
                        .Append(Term(TerminalRole.CountUp, pins["UP"]))
                        .Append(Term(TerminalRole.CountDown, pins["DOWN"]))
                        .Append(Term(TerminalRole.Load, pins["LOAD"], activeLow: true))
                        .Append(Term(TerminalRole.Reset, pins["CLR"]))
                        .Append(Term(TerminalRole.Carry, pins["CO"], activeLow: true))
                        .Append(Term(TerminalRole.Borrow, pins["BO"], activeLow: true))
                        .ToList()),
            ], TotemPole, LogicLayout.Blocks);
        }

        // ###########################################################################################
        // The 4066: four analog switches, each closed between nA and nB while its control nC is high - the model's names
        // (the datasheet's "signal in/out" and "control"), the CMOS supply's.
        // ###########################################################################################
        private static LogicChip Switches4066()
        {
            var pins = new Pinout("1A 1B 2B 2A 2C 3C VSS 3A 3B 4B 4A 4C 1C VDD");

            return Chip("40", "4066", pins, Enumerable.Range(1, 4).Select(n => new LogicUnit(LogicFunction.Switch, false,
            [
                Term(TerminalRole.Input, pins[N(n) + "A"], index: 0),
                Term(TerminalRole.Input, pins[N(n) + "B"], index: 1),
                Term(TerminalRole.Control, pins[N(n) + "C"]),
            ])), TotemPole, LogicLayout.Switches);
        }

        // ###########################################################################################
        // The 555 and 556: a timer each - RESET low wins and turns the output low; else TRIG below a third of the supply
        // sets it (output high, discharge off); else THRES above CONT (two thirds of the supply, unless CONT is driven)
        // turns it low again; else it holds. DISCH is an open collector, on while the output is low.
        // ###########################################################################################
        private static LogicUnit TimerUnit(Pinout pins, string prefix, string trigger, string threshold, string control, string reset, string output, string discharge) =>
            new(LogicFunction.Timer, false,
            [
                Term(TerminalRole.Trigger, pins[prefix + trigger], activeLow: true),
                Term(TerminalRole.Threshold, pins[prefix + threshold]),
                Term(TerminalRole.Control, pins[prefix + control]),
                Term(TerminalRole.Reset, pins[prefix + reset], activeLow: true),
                Term(TerminalRole.Output, pins[prefix + output]),
                Term(TerminalRole.Discharge, pins[prefix + discharge]),
            ]);

        private static LogicChip Timer555()
        {
            var pins = new Pinout("GND TRIG OUT RESET CONT THRES DISCH VCC");

            return Chip("55", "555", pins, [TimerUnit(pins, string.Empty, "TRIG", "THRES", "CONT", "RESET", "OUT", "DISCH")], TotemPole, LogicLayout.Timers);
        }

        private static LogicChip Timers556()
        {
            var pins = new Pinout("1DIS 1THR 1CON 1RES 1OUT 1TRG GND 2TRG 2OUT 2RES 2CON 2THR 2DIS VCC");

            return Chip("55", "556", pins, new[] { 1, 2 }.Select(n => TimerUnit(pins, N(n), "TRG", "THR", "CON", "RES", "OUT", "DIS")), TotemPole, LogicLayout.Timers);
        }

        private static bool IsSupplyOrNc(string name) => name is "GND" or "VCC" or "VSS" or "VDD" or "NC";
    }
}
