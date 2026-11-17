using Handlers.DataHandling;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// THE CHIPS CRT DRAWS WITH THEIR INSIDES (LogicChips; owner request, 2026-11-16). Every chip with a datasheet model of the
// MiniPro tests (LogicIcReferenceModels - typed in from the datasheets, never from the catalogue) is held against it: a
// combinational chip for EVERY input combination, the 4066's switches both ways, and a sequential chip (the 7474, 74373,
// 74193 and 556) through a long run of pin changes, its outputs compared after each - so a wrong gate, a swapped pin, a
// forgotten inversion or a clock taken for a data input fails here, before anybody is shown a wrong picture. A gate chip
// with no model must share its pin-out with one that has; the four other chips with none are spelt out pin by pin from
// their datasheets. Then which component is which chip: by its number whatever it is written with, and never when the
// board says something the datasheet does not.
// ###########################################################################################
public sealed class LogicChipsTests
{
    private static LogicChip Chip(string name) => LogicChips.Named(name)!;

    // The model of a 74 chip is "74ls<number>", of the others their number.
    private static LogicIcModel? ModelOf(LogicChip chip) =>
        LogicIcReferenceModels.For(chip.Family == "74" ? "74ls" + chip.Number : chip.Number);

    // What a good chip drives on an output for its logic level, as the model writes it.
    private static ChipDrive DriveOf(bool? level, LogicOutputStage stage) => stage switch
    {
        LogicOutputStage.TotemPole => level!.Value ? ChipDrive.High : ChipDrive.Low,
        LogicOutputStage.OpenCollector => level!.Value ? ChipDrive.None : ChipDrive.Low,
        _ => level is null ? ChipDrive.None : level.Value ? ChipDrive.High : ChipDrive.Low,
    };

    private static bool IsCombinational(LogicChip chip) =>
        chip.Units.All(unit => unit.Function is LogicFunction.And or LogicFunction.Or or LogicFunction.Xor or LogicFunction.Buffer
            or LogicFunction.Mux or LogicFunction.Selector or LogicFunction.Decoder);

    // A model's names, the supply's as the 4000 series' datasheets write them.
    private static IEnumerable<string> NamesOf(LogicChip chip, LogicIcModel model) =>
        model.Pins.Select(pin => chip.Family == "40" ? pin.Name switch { "GND" => "VSS", "VCC" => "VDD", _ => pin.Name } : pin.Name);

    // ###########################################################################################
    // What a combinational unit's outputs are for the levels on the chip's pins (`level`, by pin number) - null for a
    // 3-state output turned off - by what its terminals are: every enable active (or the unit off), a gate of its inputs,
    // a multiplexer's or selector's input picked by its selects' weights (held low while off), a decoder's output whose
    // number its inputs' weights add up to (none while off) - each inverted as the unit or the output says.
    // ###########################################################################################
    private static Dictionary<int, bool?> OutputsOf(LogicUnit unit, Func<int, bool> level, LogicOutputStage stage)
    {
        bool Active(LogicTerminal terminal) => level(terminal.Pin) != terminal.ActiveLow;
        int Weight(TerminalRole role) => unit.Of(role).Where(terminal => level(terminal.Pin)).Sum(terminal => terminal.Index);

        bool enabled = unit.Of(TerminalRole.Enable).All(Active);
        IReadOnlyList<int> inputs = unit.Inputs;

        if (unit.Function == LogicFunction.Decoder)
        {
            int picked = Weight(TerminalRole.Input);

            return unit.Of(TerminalRole.Output).ToDictionary(output => output.Pin, output => (bool?)((enabled && picked == output.Index) != output.ActiveLow));
        }

        if (!enabled && stage == LogicOutputStage.ThreeState)
        {
            return new Dictionary<int, bool?> { [unit.Output] = null };
        }

        bool value = unit.Function switch
        {
            LogicFunction.And => inputs.All(level),
            LogicFunction.Or => inputs.Any(level),
            LogicFunction.Xor => inputs.Count(level) % 2 == 1,
            LogicFunction.Buffer => level(inputs[0]),
            _ => enabled && level(unit.Of(TerminalRole.Input).Single(input => input.Index == Weight(TerminalRole.Select)).Pin),
        };

        return new Dictionary<int, bool?> { [unit.Output] = value != unit.Inverted };
    }

    // ###########################################################################################
    // Every combinational chip with a model: its pins as the datasheet names them, its supply where the model has it, the
    // pins it drives its units' outputs and the pins the model reads its units' inputs - and for every combination of
    // those inputs, each output what the datasheet's function table says.
    // ###########################################################################################
    [Fact]
    public void Every_combinational_chip_with_a_datasheet_model_behaves_as_the_datasheet_says_for_every_input_combination()
    {
        List<(LogicChip Chip, LogicIcModel Model)> modelled = LogicChips.All
            .Where(LogicChipsTests.IsCombinational)
            .Select(chip => (chip, model: ModelOf(chip)))
            .Where(pair => pair.model is not null)
            .Select(pair => (pair.chip, pair.model!))
            .ToList();

        // The parts on the published boards in November 2026 all have one.
        Assert.Superset(
            new HashSet<string>
            {
                "7400", "7402", "7403", "7404", "7406", "7407", "7408", "7414", "7427", "7432", "7486", "74125", "74132", "74133",
                "74138", "74139", "74145", "74153", "74157", "74244", "74257", "74258",
            },
            modelled.Select(pair => pair.Chip.Name).ToHashSet());

        foreach ((LogicChip chip, LogicIcModel model) in modelled)
        {
            Assert.Equal(LogicChipsTests.NamesOf(chip, model), chip.PinNames);
            Assert.Equal(model.Pins.Single(pin => pin.Kind == ModelPinKind.Ground).Number, chip.Ground);
            Assert.Equal(model.Pins.Single(pin => pin.Kind == ModelPinKind.Supply).Number, chip.Supply);

            List<int> outputs = chip.Units.SelectMany(unit => unit.Of(TerminalRole.Output)).Select(terminal => terminal.Pin).Order().ToList();
            Assert.Equal(model.ChipDriven.Select(pin => pin.Number).Order(), outputs);

            List<int> inputs = model.Inputs.Select(pin => pin.Number).Order().ToList();
            Assert.Equal(inputs, chip.Units.SelectMany(unit => unit.Pins).Except(outputs).Distinct().Order());

            for (int combination = 0; combination < 1 << inputs.Count; combination++)
            {
                var levels = new Dictionary<int, bool>();

                for (int bit = 0; bit < inputs.Count; bit++)
                {
                    levels[inputs[bit]] = ((combination >> bit) & 1) == 1;
                }

                Dictionary<int, ChipDrive> expected = model.Respond(name => levels.TryGetValue(model.Pin(name).Number, out bool level) ? level : null);

                foreach (LogicUnit unit in chip.Units)
                {
                    foreach ((int pin, bool? level) in LogicChipsTests.OutputsOf(unit, pin => levels[pin], chip.Outputs))
                    {
                        ChipDrive drawn = DriveOf(level, chip.Outputs);

                        Assert.True(
                            expected[pin] == drawn,
                            $"{chip.Name} pin {pin} with {string.Join(" ", levels.OrderBy(pair => pair.Key).Select(pair => $"{pair.Key}={(pair.Value ? 1 : 0)}"))}: the datasheet says {expected[pin]}, the drawing {drawn}");
                    }
                }
            }
        }
    }

    // ###########################################################################################
    // The 4066's switches: each between the two pins the datasheet says, closed by the control it says - a level on one
    // side reaches the other while the control is high, and nothing does while it is low.
    // ###########################################################################################
    [Fact]
    public void Each_4066_switch_carries_a_level_between_its_two_pins_while_its_control_is_high()
    {
        LogicChip chip = Chip("CD4066BE");
        LogicIcModel model = ModelOf(chip)!;

        Assert.Equal(LogicChipsTests.NamesOf(chip, model), chip.PinNames);
        Assert.Equal(4, chip.Units.Count);

        foreach (LogicUnit unit in chip.Units)
        {
            Assert.Equal(LogicFunction.Switch, unit.Function);
            int control = unit.Of(TerminalRole.Control).Single().Pin;

            foreach ((int from, int to) in new[] { (unit.Inputs[0], unit.Inputs[1]), (unit.Inputs[1], unit.Inputs[0]) })
            {
                foreach (bool closed in new[] { false, true })
                {
                    foreach (bool level in new[] { false, true })
                    {
                        // Every other control low, so no other switch carries anything.
                        var socket = chip.Units.Select(other => other.Of(TerminalRole.Control).Single().Pin).ToDictionary(pin => pin, _ => (bool?)false);
                        socket[control] = closed;
                        socket[from] = level;

                        Dictionary<int, ChipDrive> drives = model.Respond(name => socket.TryGetValue(model.Pin(name).Number, out bool? value) ? value : null);

                        Assert.Equal(closed ? (level ? ChipDrive.High : ChipDrive.Low) : ChipDrive.None, drives[to]);
                    }
                }
            }
        }
    }

    // ###########################################################################################
    // A sequential unit as its terminals say it works - the test's own reading of the datasheets, so a terminal given the
    // wrong role (a clock taken for D, a set for a reset) disagrees with the model: a flip-flop's set and reset at once, its
    // D on the clock's rising edge; a latch following D while its clock (LE) is high, behind its enable; a counter cleared,
    // loaded, or counting on an edge of UP or DOWN while the other is high (else what follows is not known); a timer reset,
    // set by its trigger, reset by its threshold.
    // ###########################################################################################
    private sealed class UnitSimulation(LogicUnit unit)
    {
        private int thisState;

        public bool Known { get; private set; } = true;

        private bool Active(Func<int, bool> level, TerminalRole role) =>
            unit.Of(role).Any(terminal => level(terminal.Pin) != terminal.ActiveLow);

        private int Pin(TerminalRole role) => unit.Of(role).Single().Pin;

        public void Step(Func<int, bool> level, int? changed)
        {
            switch (unit.Function)
            {
                case LogicFunction.FlipFlop:
                {
                    bool set = this.Active(level, TerminalRole.Set);
                    bool reset = this.Active(level, TerminalRole.Reset);

                    if (reset && !set)
                    {
                        this.thisState = 0;
                    }
                    else if (set && !reset)
                    {
                        this.thisState = 1;
                    }
                    else if (!set && !reset && changed == this.Pin(TerminalRole.Clock) && level(this.Pin(TerminalRole.Clock)))
                    {
                        this.thisState = level(unit.Inputs[0]) ? 1 : 0;
                    }

                    break;
                }

                case LogicFunction.Latch:
                    if (level(this.Pin(TerminalRole.Clock)))
                    {
                        this.thisState = level(unit.Inputs[0]) ? 1 : 0;
                    }

                    break;

                case LogicFunction.Counter:
                {
                    int up = this.Pin(TerminalRole.CountUp);
                    int down = this.Pin(TerminalRole.CountDown);

                    if (this.Active(level, TerminalRole.Reset))
                    {
                        (this.thisState, this.Known) = (0, true);
                    }
                    else if (this.Active(level, TerminalRole.Load))
                    {
                        (this.thisState, this.Known) = (unit.Of(TerminalRole.Input).Where(terminal => level(terminal.Pin)).Sum(terminal => terminal.Index), true);
                    }
                    else if ((changed == up || changed == down) && level(changed.Value))
                    {
                        if (!level(changed == up ? down : up))
                        {
                            this.Known = false;
                        }
                        else
                        {
                            this.thisState = (this.thisState + (changed == up ? 1 : 15)) % 16;
                        }
                    }

                    break;
                }

                case LogicFunction.Timer:
                    if (this.Active(level, TerminalRole.Reset))
                    {
                        this.thisState = 0;
                    }
                    else if (this.Active(level, TerminalRole.Trigger))
                    {
                        this.thisState = 1;
                    }
                    else if (level(this.Pin(TerminalRole.Threshold)))
                    {
                        this.thisState = 0;
                    }

                    break;
            }
        }

        // Each output pin's level (null while a 3-state output is off), with how the chip drives it.
        public IEnumerable<(int Pin, bool? Level, LogicOutputStage Stage)> Outputs(Func<int, bool> level, LogicOutputStage stage)
        {
            switch (unit.Function)
            {
                case LogicFunction.FlipFlop:
                {
                    bool set = this.Active(level, TerminalRole.Set);
                    bool reset = this.Active(level, TerminalRole.Reset);

                    foreach (LogicTerminal output in unit.Of(TerminalRole.Output))
                    {
                        yield return (output.Pin, output.ActiveLow ? reset || (!set && this.thisState == 0) : set || (!reset && this.thisState == 1), stage);
                    }

                    break;
                }

                case LogicFunction.Latch:
                    yield return (unit.Output, this.Active(level, TerminalRole.Enable) ? this.thisState == 1 : null, stage);
                    break;

                case LogicFunction.Counter:
                    foreach (LogicTerminal output in unit.Of(TerminalRole.Output))
                    {
                        yield return (output.Pin, (this.thisState & output.Index) != 0, stage);
                    }

                    yield return (this.Pin(TerminalRole.Carry), !(this.thisState == 15 && !level(this.Pin(TerminalRole.CountUp))), stage);
                    yield return (this.Pin(TerminalRole.Borrow), !(this.thisState == 0 && !level(this.Pin(TerminalRole.CountDown))), stage);
                    break;

                case LogicFunction.Timer:
                {
                    bool high = !this.Active(level, TerminalRole.Reset) && (this.Active(level, TerminalRole.Trigger) || this.thisState == 1);

                    yield return (unit.Output, high, LogicOutputStage.TotemPole);
                    yield return (this.Pin(TerminalRole.Discharge), high, LogicOutputStage.OpenCollector);
                    break;
                }
            }
        }
    }

    // ###########################################################################################
    // Every sequential chip with a model: from a random start, thousands of single pin changes, each stepping the model and
    // the units, and after each every output compared - a run whose next state the datasheet does not say starts again.
    // ###########################################################################################
    [Theory]
    [InlineData("7474")]
    [InlineData("74373")]
    [InlineData("74193")]
    [InlineData("556")]
    public void A_sequential_chip_behaves_as_its_datasheet_model_through_a_long_run_of_pin_changes(string name)
    {
        LogicChip chip = LogicChips.All.Single(candidate => candidate.Name == name);
        LogicIcModel model = ModelOf(chip)!;

        Assert.True(model.IsSequential);
        Assert.Equal(model.Pins.Select(pin => pin.Name), chip.PinNames);

        List<int> inputs = model.Inputs.Select(pin => pin.Number).ToList();
        HashSet<int> outputs = model.ChipDriven.Select(pin => pin.Number).ToHashSet();
        Assert.Equal(outputs, chip.Units.SelectMany(unit => unit.Of(TerminalRole.Output).Concat(unit.Of(TerminalRole.Carry)).Concat(unit.Of(TerminalRole.Borrow)).Concat(unit.Of(TerminalRole.Discharge))).Select(terminal => terminal.Pin).ToHashSet());
        Assert.Superset(inputs.ToHashSet(), chip.Units.SelectMany(unit => unit.Pins).Except(outputs).Except(chip.Units.SelectMany(unit => unit.Of(TerminalRole.Control)).Select(terminal => terminal.Pin)).ToHashSet());

        // The same run every time, a different one per chip.
        var random = new Random(name.Aggregate(17, (hash, character) => (hash * 31) + character));
        var levels = new Dictionary<int, bool>();
        int state = 0;
        List<UnitSimulation> units = [];
        int compared = 0;

        bool Level(int pin) => levels[pin];
        bool LevelNamed(string pinName) => levels[model.Pin(pinName).Number];

        void Start()
        {
            foreach (int pin in inputs)
            {
                levels[pin] = random.Next(2) == 1;
            }

            state = model.Step!(0, LevelNamed, null);
            units = chip.Units.Select(unit => new UnitSimulation(unit)).ToList();
            units.ForEach(unit => unit.Step(Level, null));
        }

        Start();

        for (int step = 0; step < 4000; step++)
        {
            int pin = inputs[random.Next(inputs.Count)];
            levels[pin] = !levels[pin];
            state = model.Step!(state, LevelNamed, model.Pins[pin - 1].Name);
            units.ForEach(unit => unit.Step(Level, pin));

            if (state < 0 || units.Any(unit => !unit.Known))
            {
                Start();
                continue;
            }

            Dictionary<int, ChipDrive> expected = model.Respond(pinName => levels.TryGetValue(model.Pin(pinName).Number, out bool level) ? level : null, state);

            foreach ((int output, bool? level, LogicOutputStage stage) in units.SelectMany(unit => unit.Outputs(Level, chip.Outputs)))
            {
                ChipDrive drawn = DriveOf(level, stage);
                Assert.True(expected[output] == drawn, $"{name} pin {output} after step {step}: the datasheet says {expected[output]}, the drawing {drawn}");
                compared++;
            }
        }

        Assert.True(compared > 1000, $"{name}: only {compared} outputs compared");
    }

    // ###########################################################################################
    // A CHIP WITH NO MODEL: a gate chip shares its pin-out with one that has - the same units on the same pins, the same
    // supply and the same pins unconnected - only what each gate does and how its outputs drive differ (the 7409 is the 7408
    // open collector, the 4011 the 74266's pins as a NAND). The list is written out, so one added without a model is a
    // decision; the four that are no gate chip are spelt out below.
    // ###########################################################################################
    [Fact]
    public void Every_chip_without_a_model_has_the_pin_out_of_one_that_has_or_is_spelt_out_below()
    {
        static string ShapeOf(LogicChip chip) =>
            $"{chip.Ground}/{chip.Supply}/" +
            string.Join(",", Enumerable.Range(1, chip.PinCount).Where(pin => chip.NameOf(pin) == "NC")) + "/" +
            string.Join(";", chip.Units.Select(unit => $"{string.Join(",", unit.Inputs)}>{unit.Output}"));

        HashSet<string> modelledShapes = LogicChips.All.Where(chip => chip.Layout == LogicLayout.Gates && ModelOf(chip) is not null).Select(ShapeOf).ToHashSet();
        List<LogicChip> unmodelled = LogicChips.All.Where(chip => ModelOf(chip) is null).ToList();

        Assert.Equal(
            ["7405", "7409", "7412", "7415", "7416", "7417", "7422", "7426", "7428", "7433", "7437", "7438", "7440", "74136",
             "74245", "74175", "74273", "4001", "4011", "4069", "4070", "4071", "4077", "4081", "4093", "555"],
            unmodelled.Select(chip => chip.Name));

        Assert.All(unmodelled.Where(chip => chip.Layout == LogicLayout.Gates), chip => Assert.Contains(ShapeOf(chip), modelledShapes));
        Assert.Equal(["74245", "74175", "74273", "555"], unmodelled.Where(chip => chip.Layout != LogicLayout.Gates).Select(chip => chip.Name));
    }

    // The pins of a unit's terminals of a role, in order.
    private static int[] PinsOf(LogicUnit unit, TerminalRole role) => [.. unit.Of(role).Select(terminal => terminal.Pin)];

    // The 74175's datasheet: CLR 1, CLK 9; flip-flop n's D, Q and Q-bar.
    [Fact]
    public void The_74175_is_four_flip_flops_sharing_pin_9s_clock_and_pin_1s_clear()
    {
        LogicChip chip = Chip("74LS175");
        int[][] flipFlops = [[4, 2, 3], [5, 7, 6], [12, 10, 11], [13, 15, 14]];

        Assert.Equal((8, 16), (chip.Ground, chip.Supply));
        Assert.Equal(4, chip.Units.Count);

        for (int n = 0; n < 4; n++)
        {
            LogicUnit unit = chip.Units[n];
            Assert.Equal(LogicFunction.FlipFlop, unit.Function);
            Assert.Equal([flipFlops[n][0]], PinsOf(unit, TerminalRole.Input));
            Assert.Equal([flipFlops[n][1], flipFlops[n][2]], PinsOf(unit, TerminalRole.Output));
            Assert.Equal([false, true], unit.Of(TerminalRole.Output).Select(terminal => terminal.ActiveLow));
            Assert.Equal([9], PinsOf(unit, TerminalRole.Clock));
            Assert.Equal(new LogicTerminal(TerminalRole.Reset, 1, ActiveLow: true), unit.Of(TerminalRole.Reset).Single());
        }
    }

    // The 74273's datasheet: CLR 1, CLK 11; flip-flop n's D and Q.
    [Fact]
    public void The_74273_is_eight_flip_flops_sharing_pin_11s_clock_and_pin_1s_clear()
    {
        LogicChip chip = Chip("74LS273");
        int[][] flipFlops = [[3, 2], [4, 5], [7, 6], [8, 9], [13, 12], [14, 15], [17, 16], [18, 19]];

        Assert.Equal((10, 20), (chip.Ground, chip.Supply));
        Assert.Equal(8, chip.Units.Count);

        for (int n = 0; n < 8; n++)
        {
            LogicUnit unit = chip.Units[n];
            Assert.Equal([flipFlops[n][0]], PinsOf(unit, TerminalRole.Input));
            Assert.Equal([flipFlops[n][1]], PinsOf(unit, TerminalRole.Output));
            Assert.Equal([11], PinsOf(unit, TerminalRole.Clock));
            Assert.Equal(new LogicTerminal(TerminalRole.Reset, 1, ActiveLow: true), unit.Of(TerminalRole.Reset).Single());
        }
    }

    // The 74245's datasheet: DIR 1, OE 19 (active low); An on pin n+1, Bn on pin 19-n.
    [Fact]
    public void The_74245_is_eight_transceivers_between_An_and_Bn_sharing_DIR_and_the_enable()
    {
        LogicChip chip = Chip("SN74LS245N");

        Assert.Equal((10, 20), (chip.Ground, chip.Supply));
        Assert.Equal(8, chip.Units.Count);

        for (int n = 1; n <= 8; n++)
        {
            LogicUnit unit = chip.Units[n - 1];
            Assert.Equal(LogicFunction.Transceiver, unit.Function);
            Assert.Equal([n + 1, 19 - n], unit.Inputs);
            Assert.Equal([1], PinsOf(unit, TerminalRole.Direction));
            Assert.Equal(new LogicTerminal(TerminalRole.Enable, 19, ActiveLow: true), unit.Of(TerminalRole.Enable).Single());
        }
    }

    // The 555's datasheet: GND 1, TRIG 2 (active low), OUT 3, RESET 4 (active low), CONT 5, THRES 6, DISCH 7, VCC 8.
    [Fact]
    public void The_555_is_one_timer_on_the_pins_its_datasheet_says()
    {
        LogicChip chip = Chip("NE555P");
        LogicUnit timer = chip.Units.Single();

        Assert.Equal((1, 8), (chip.Ground, chip.Supply));
        Assert.Equal(LogicFunction.Timer, timer.Function);
        Assert.Equal(
            [
                new LogicTerminal(TerminalRole.Trigger, 2, ActiveLow: true),
                new LogicTerminal(TerminalRole.Threshold, 6),
                new LogicTerminal(TerminalRole.Control, 5),
                new LogicTerminal(TerminalRole.Reset, 4, ActiveLow: true),
                new LogicTerminal(TerminalRole.Output, 3),
                new LogicTerminal(TerminalRole.Discharge, 7),
            ],
            timer.Terminals);
    }

    // Every pin is one thing: the supply, not connected, or a unit's - and no pin is one unit's output and another's input.
    [Fact]
    public void Every_pin_of_every_chip_is_its_supply_unconnected_or_a_units()
    {
        Assert.All(LogicChips.All, chip =>
        {
            var used = chip.Units.SelectMany(unit => unit.Pins).ToHashSet();

            for (int pin = 1; pin <= chip.PinCount; pin++)
            {
                int kinds = (pin == chip.Ground || pin == chip.Supply ? 1 : 0) + (chip.NameOf(pin) == "NC" ? 1 : 0) + (used.Contains(pin) ? 1 : 0);
                Assert.True(kinds == 1, $"{chip.Name} pin {pin}");
            }

            List<int> outputs = chip.Units.SelectMany(unit => unit.Of(TerminalRole.Output)).Select(terminal => terminal.Pin).ToList();
            Assert.Equal(outputs.Count, outputs.Distinct().Count());
            Assert.Empty(chip.Units.SelectMany(unit => unit.Terminals).Where(terminal => terminal.Role != TerminalRole.Output).Select(terminal => terminal.Pin).Intersect(outputs));
            Assert.All(chip.Units, unit => Assert.Equal(unit.Terminals.Count, unit.Pins.Count()));
        });
    }

    // ---- which chip -----------------------------------------------------------------------------

    [Theory]
    [InlineData("74LS08", "7408")]
    [InlineData("74ls08", "7408")]
    [InlineData("7408", "7408")]
    [InlineData("SN74LS08N", "7408")]
    [InlineData("HD74LS08P", "7408")]
    [InlineData("M74LS08P", "7408")]
    [InlineData("74HCT08", "7408")]
    [InlineData("74F32", "7432")]
    [InlineData("74HCU04", "7404")]
    [InlineData("DM74LS257BN", "74257")]
    [InlineData("74LS133", "74133")]
    [InlineData("7400", "7400")]
    [InlineData("74LS08 (quad AND)", "7408")]
    [InlineData("MC14011BCP", "4011")]
    [InlineData("CD4011BE", "4011")]
    [InlineData("4011", "4011")]
    [InlineData("CD4066BE", "4066")]
    [InlineData("4066", "4066")]
    [InlineData("SN74LS245N", "74245")]
    [InlineData("74LS373N", "74373")]
    [InlineData("74HC138", "74138")]
    [InlineData("74LS74", "7474")]
    [InlineData("NE555P", "555")]
    [InlineData("LM555CN", "555")]
    [InlineData("TLC555CP", "555")]
    [InlineData("555", "555")]
    [InlineData("NE556N", "556")]
    // A micro sign or a Greek mu - "µA555", a maker's "uA" - is read as the "U" it stands for (code review, 2026-11-17).
    [InlineData("µA555", "555")]
    [InlineData("μA556", "556")]
    [InlineData("µPD74LS08", "7408")]
    // A 74-series part carrying a CMOS number is that CMOS chip, pin for pin (code review, 2026-11-17).
    [InlineData("74HC4066", "4066")]
    [InlineData("SN74HCT4066N", "4066")]
    public void A_chip_is_known_by_its_number_whatever_it_is_written_with(string text, string chip) =>
        Assert.Equal(chip, LogicChips.Named(text)?.Name);

    // Nothing else is taken for one: a CPU, a DRAM (4164 is no CMOS part), a ROM, a regulator, an op-amp, a 74 part CRT
    // does not draw (the 74629's oscillators, the 74123's monostables), the CMOS 4044 (its pin-out not checked against a
    // datasheet), a number that only starts like a timer's - and a number with a leading zero no part has (code review,
    // 2026-11-17: "74LS008" was drawn as a 7408).
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("6510")]
    [InlineData("4164")]
    [InlineData("41464")]
    [InlineData("40022")]
    [InlineData("7805")]
    [InlineData("LM741")]
    [InlineData("74LS741")]
    [InlineData("74LS629")]
    [InlineData("74LS123")]
    [InlineData("CD4044BE")]
    [InlineData("5555")]
    [InlineData("NE5532")]
    [InlineData("74")]
    [InlineData("AND")]
    [InlineData("74LS008")]
    [InlineData("74008")]
    public void Nothing_else_is_taken_for_a_chip(string? text) =>
        Assert.Null(LogicChips.Named(text));

    private static ComponentEntry Component(string technical, string package = "DIP-14", string partNumber = "") =>
        new() { BoardLabel = "U1", TechnicalNameOrValue = technical, PartNumber = partNumber, Category = "IC", Package = package };

    private static ComponentPinEntry Pin(int number, string name, string direction = "") =>
        new() { BoardLabel = "U1", Pin = number.ToString(CultureInfo.InvariantCulture), Name = name, Direction = direction };

    // Its technical name first, its part number after; a package left empty is no contradiction.
    [Fact]
    public void A_component_is_its_chip_by_technical_name_or_part_number_with_or_without_a_package()
    {
        Assert.Equal("7408", LogicChips.For(Component("74LS08"), [])?.Name);
        Assert.Equal("7408", LogicChips.For(Component("74LS08", package: ""), [])?.Name);
        Assert.Equal("7408", LogicChips.For(Component("AND", partNumber: "SN74LS08N"), [])?.Name);
        Assert.Equal("74257", LogicChips.For(Component("74LS257", package: "DIP-16"), [])?.Name);
        Assert.Equal("555", LogicChips.For(Component("NE555", package: "DIP-8"), [])?.Name);
        Assert.Null(LogicChips.For(Component("6510", package: "DIP-40"), []));
    }

    // ###########################################################################################
    // A WRONG PICTURE IS WORSE THAN NONE: a package of another pin count or kind, a pin past the case, a ground or supply
    // named where the datasheet has none, and the chip is not drawn. The board's names for the pins otherwise say nothing
    // against it (they are its signals, "BA", "/A/B") - nor a name saying the rail a pin is strapped to ("G - GND"): the
    // names are ChipNames', whole names only - nor a DIRECTION typed (code review, 2026-11-17: KiCad types an input strapped
    // to a rail "Power" or "Ground"; a pin typed "Power" was taken against the datasheet until then).
    // ###########################################################################################
    [Fact]
    public void A_component_whose_board_data_contradicts_the_datasheet_is_not_drawn()
    {
        Assert.Null(LogicChips.For(Component("74LS08", package: "DIP-16"), []));
        Assert.Null(LogicChips.For(Component("74LS08", package: "SIP-14"), []));
        Assert.Null(LogicChips.For(Component("74LS08"), [Pin(15, "")]));
        Assert.Null(LogicChips.For(Component("74LS08"), [Pin(8, "GND")]));
        Assert.Null(LogicChips.For(Component("74LS08"), [Pin(7, "VCC")]));
        Assert.Null(LogicChips.For(Component("74LS08"), [Pin(1, "+5V")]));
        Assert.Null(LogicChips.For(Component("74LS08"), [Pin(1, "+12V")]));
        Assert.NotNull(LogicChips.For(Component("74LS08"), [Pin(3, "BA", "Power")]));
        Assert.NotNull(LogicChips.For(Component("74LS08"), [Pin(1, "1A", "Ground")]));
        Assert.NotNull(LogicChips.For(Component("74LS08"), [Pin(1, "1A - GND")]));

        // A capacitor or resistor whose value reads as a chip's number is no chip (code review, 2026-11-17).
        Assert.Null(LogicChips.For(new ComponentEntry { BoardLabel = "C9", TechnicalNameOrValue = "5400pF", Category = "Capacitor" }, []));
        Assert.Null(LogicChips.For(new ComponentEntry { BoardLabel = "R9", TechnicalNameOrValue = "555", Category = "Resistor" }, []));
        Assert.NotNull(LogicChips.For(new ComponentEntry { BoardLabel = "U9", TechnicalNameOrValue = "5400" }, []));
        Assert.Null(LogicChips.For(Component("NE556", package: "DIP-8"), []));

        Assert.NotNull(LogicChips.For(Component("74LS08"), [Pin(7, "GND", "Power"), Pin(14, "VCC", "Power"), Pin(3, "BA", "Output")]));
        Assert.NotNull(LogicChips.For(Component("CD4011BE"), [Pin(7, "VSS"), Pin(14, "VDD")]));
        Assert.NotNull(LogicChips.For(Component("74LS257", package: "DIP-16"), [Pin(1, "/A/B"), Pin(15, "/G"), Pin(8, "Ground")]));
    }

    // ---- the directions -------------------------------------------------------------------------

    // ###########################################################################################
    // The datasheet's direction of every pin, in the pin list's words: what a unit drives is an output, a transceiver's A and
    // B and a switch's sides pass a signal either way, every other terminal is an input - and the ground, the supply and a
    // pin no unit uses are what they are.
    // ###########################################################################################
    [Theory]
    [InlineData("74LS08", "Input Input Output Input Input Output Ground Output Input Input Output Input Input Power")]
    [InlineData("74LS245", "Input Input/output Input/output Input/output Input/output Input/output Input/output Input/output Input/output Ground Input/output Input/output Input/output Input/output Input/output Input/output Input/output Input/output Input Power")]
    [InlineData("4066", "Input/output Input/output Input/output Input/output Input Input Ground Input/output Input/output Input/output Input/output Input Input Power")]
    [InlineData("74LS193", "Input Output Output Input Input Output Output Ground Input Input Input Output Output Input Input Power")]
    [InlineData("556", "Output Input Input Input Output Input Ground Input Output Input Input Input Output Power")]
    [InlineData("74LS133", "Input Input Input Input Input Input Input Ground Output Input Input Input Input Input Input Power")]
    public void The_datasheet_says_each_pins_direction(string name, string directions)
    {
        LogicChip chip = Chip(name);

        Assert.Equal(directions.Split(' '), Enumerable.Range(1, chip.PinCount).Select(chip.DirectionOf));
        Assert.All(Enumerable.Range(1, chip.PinCount), pin => Assert.Contains(chip.DirectionOf(pin), ComponentVocabulary.PinDirections));
    }

    // ###########################################################################################
    // AN ACTIVE-LOW PIN'S NAME HAS A "/" BEFORE IT (owner decision, 2026-11-17: "use "/" everywhere"), as the datasheet draws a
    // bar over it and the published boards write theirs ("/IRQ", "/RES") - where "Fill pins from datasheet" and the drawing take
    // a logic chip's names from. A Q-bar's "QN" is "/Q"; a pin used active high keeps its name, so the 74125's enable is "/1OE"
    // and the 74126's "1OE". The datasheet models' own spelling stays in PinNames.
    // ###########################################################################################
    [Theory]
    [InlineData("74LS138", "A B C /G2A /G2B G1 /Y7 GND /Y6 /Y5 /Y4 /Y3 /Y2 /Y1 /Y0 VCC")]
    [InlineData("74LS74", "/1CLR 1D 1CLK /1PRE 1Q /1Q GND /2Q 2Q /2PRE 2CLK 2D /2CLR VCC")]
    [InlineData("74LS373", "/OE 1Q 1D 2D 2Q 3Q 3D 4D 4Q GND LE 5Q 5D 6D 6Q 7Q 7D 8D 8Q VCC")]
    [InlineData("74LS08", "1A 1B 1Y 2A 2B 2Y GND 3Y 3A 3B 4Y 4A 4B VCC")]
    public void An_active_low_pin_is_named_with_a_slash_before_it(string name, string names)
    {
        LogicChip chip = Chip(name);

        Assert.Equal(names.Split(' '), Enumerable.Range(1, chip.PinCount).Select(chip.NameOf));
        Assert.Equal("1QN", Chip("74LS74").PinNames[5]);
    }

    [Fact]
    public void The_74125s_enable_is_active_low_and_the_74126s_active_high()
    {
        Assert.Equal("/1OE", Chip("74LS125").NameOf(1));
        Assert.Equal("1OE", Chip("74LS126").NameOf(1));
    }

    // A pin no unit uses is not connected - the 7420's pins 3 and 11.
    [Fact]
    public void A_pin_no_unit_uses_is_not_connected()
    {
        LogicChip chip = Chip("74LS20");

        Assert.Equal(["Not connected", "Not connected"], new[] { 3, 11 }.Select(chip.DirectionOf));
    }

    // A multiplexer's enable off: the 157 holds its outputs low, the 158 high, the 257 lets go - and its select picks A
    // while low (spelt out here, as the model's test above is about every combination at once).
    [Fact]
    public void A_multiplexer_picks_A_while_its_select_is_low_and_its_enable_decides_the_rest()
    {
        static bool? Out(string chip, bool select, bool enable, bool a, bool b)
        {
            LogicChip found = Chip(chip);
            var levels = new Dictionary<int, bool> { [1] = select, [15] = enable, [2] = a, [3] = b };

            return LogicChipsTests.OutputsOf(found.Units[0], pin => levels[pin], found.Outputs)[4];
        }

        Assert.Equal(true, Out("74LS157", select: false, enable: false, a: true, b: false));
        Assert.Equal(false, Out("74LS157", select: true, enable: false, a: true, b: false));
        Assert.Equal(false, Out("74LS157", select: false, enable: true, a: true, b: true));
        Assert.Equal(true, Out("74LS158", select: false, enable: true, a: false, b: false));
        Assert.Null(Out("74LS257", select: false, enable: true, a: true, b: true));
        Assert.Equal(false, Out("74LS258", select: true, enable: false, a: false, b: true));
    }
}
