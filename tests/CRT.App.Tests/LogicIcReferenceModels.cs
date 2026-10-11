namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The DATASHEET side of the MiniPro IC tests (2026-10-11).
//
// Every logic part the catalogue can run - the 74-series parts, the 4066 switch and the 556
// timer - typed in from its datasheet's connection diagram and function table, and NOT from the
// catalogue's vectors. Each pin-out is written the way a datasheet draws it, pin 1 first, so a
// reviewer can hold a line against the datasheet:
//
//     "1A 1B 1Y 2A 2B 2Y GND 3Y 3A 3B 4Y 4A 4B VCC"
//
// GND and VCC are the supply, NC is not connected, a name a rule is given for is a pin the CHIP
// can drive, and every other name is an input (or a pin the test must leave open). The
// IcTestCatalogueShippedDataTests hold each catalogue entry against its model here: a vector that
// disagrees with the datasheet would make a GOOD chip fail, and one where the T48 drives a pin the
// chip drives, or puts the supply on the wrong pin, could damage the chip in the socket.
//
// A model says, for every pin the chip can drive, WHEN it drives it - given what the T48 drives
// (a DriveRule). That is what lets a part whose pins swap between driven and read (the 4066's two
// directions) be checked vector by vector and step by step. A SEQUENTIAL part (flip-flop, latch,
// counter, timer) also keeps a state, moved one pin change at a time by its StepRule - which is
// what lets its test be played in every order the T48 might change the pins in.
//
// The open-collector and 3-state rules follow the T48's own test (minipro's t48.c): a pin nothing
// drives is read with a weak pull-up and then a weak pull-down, and is expected as "Z".
// ###########################################################################################
internal enum ModelPinKind { Ground, Supply, NotConnected, LeftOpen, Input, Output, Bidirectional }

internal enum OutputStage { TotemPole, OpenCollector, ThreeState }

// What the chip does to one of its pins: nothing, drives it low or high, or drives it at a level
// that is not known - an input it depends on is floating, or its state is undefined (never allowed
// on a pin the vector reads).
internal enum ChipDrive { None, Low, High, Unknown }

// When and how the chip drives a pin. "socket" gives the level the T48 drives on a pin (by name),
// or null when the T48 does not drive it; "state" is a sequential part's state (0 otherwise).
internal delegate ChipDrive DriveRule(Func<string, bool?> socket, int state);

// How a sequential part's state (0 .. States-1) moves when ONE input pin changes: "changed" is that
// pin and level(changed) its new level (true: it rose), or "changed" is null to settle the state to
// the levels as they are, with no edge (power-up). -1 means the datasheet does not say what happens.
internal delegate int StepRule(int state, Func<string, bool> level, string? changed);

// A set of input combinations a test must hold, each exactly once: every level of the Driven
// pins, with nothing else driven.
internal sealed record CombinationBlock(IReadOnlyList<string> Driven);

internal sealed class ModelPin
{
    public required int Number { get; init; }
    public required string Name { get; init; }
    public required ModelPinKind Kind { get; init; }

    // Output and Bidirectional pins only.
    public DriveRule? Drive { get; init; }
}

internal sealed class LogicIcModel
{
    private LogicIcModel(
        string id, IReadOnlyList<ModelPin> pins, string coverage, IReadOnlyList<CombinationBlock> blocks,
        int states = 1, StepRule? step = null)
    {
        this.Id = id;
        this.Pins = pins;
        this.Coverage = coverage;
        this.Blocks = blocks;
        this.States = states;
        this.Step = step;
    }

    // The catalogue entry this model checks (its "id", compared without regard to capitals).
    public string Id { get; }

    // Index 0 is pin 1.
    public IReadOnlyList<ModelPin> Pins { get; }

    // What the entry may honestly claim: "exhaustive" when its vectors are exactly the Blocks'
    // combinations, "functional" for a sequence (a sequential part's test).
    public string Coverage { get; }

    public IReadOnlyList<CombinationBlock> Blocks { get; }

    // A sequential part's number of states, and how they move; 1 and null for any other part.
    public int States { get; }

    public StepRule? Step { get; }

    public bool IsSequential => this.Step is not null;

    public IEnumerable<ModelPin> Inputs => this.Pins.Where(p => p.Kind == ModelPinKind.Input);

    // The pins the chip can drive: outputs, and pins that are inputs at other times.
    public IEnumerable<ModelPin> ChipDriven => this.Pins.Where(p => p.Drive is not null);

    public bool SwapsPinRoles => this.Pins.Any(p => p.Kind == ModelPinKind.Bidirectional);

    public ModelPin Pin(string name) => this.Pins.Single(p => p.Name == name);

    // What a GOOD chip drives on each pin it can drive, for what the T48 drives (by pin number).
    public Dictionary<int, ChipDrive> Respond(Func<string, bool?> socket, int state = 0) =>
        this.ChipDriven.ToDictionary(p => p.Number, p => p.Drive!(socket, state));

    // The T48 side of a vector, by pin name: "0"/"1" is a level it drives, anything else is not.
    public Func<string, bool?> SocketOf(string vector) => name => vector[this.Pin(name).Number - 1] switch
    {
        '0' => false,
        '1' => true,
        _ => null,
    };

    // A vector's input combination: the level of every chip pin the T48 drives, "." for the
    // others (a not-connected pin held at a level is no input).
    public string KeyOf(string vector) =>
        new string(vector.Select((c, k) =>
            c is '0' or '1' && this.Pins[k].Kind != ModelPinKind.NotConnected ? c : '.').ToArray());

    public IEnumerable<string> Combinations(CombinationBlock block)
    {
        int count = block.Driven.Count;
        for (int value = 0; value < 1 << count; value++)
        {
            char[] key = Enumerable.Repeat('.', this.Pins.Count).ToArray();
            for (int k = 0; k < count; k++)
                key[this.Pin(block.Driven[k]).Number - 1] = ((value >> (count - 1 - k)) & 1) == 1 ? '1' : '0';
            yield return new string(key);
        }
    }

    // A part whose outputs each have their own rule (decoders, multiplexers, buffers): the inputs
    // are always driven, the outputs always read, and the test holds every input combination.
    public static LogicIcModel Define(
        string id, string pinout, OutputStage stage,
        params (string Output, Func<Func<string, bool>, bool?> Rule)[] outputs)
    {
        var rules = outputs.ToDictionary(
            o => o.Output, o => (DriveRule)((socket, _) => Evaluate(stage, () => o.Rule(Need(socket)))), StringComparer.Ordinal);
        List<ModelPin> pins = Build(id, pinout, rules, ModelPinKind.Output);
        string[] inputs = pins.Where(p => p.Kind == ModelPinKind.Input).Select(p => p.Name).ToArray();
        return new LogicIcModel(id, pins, "exhaustive", new[] { new CombinationBlock(inputs) });
    }

    // A part made of identical gates: every pin named "<n>Y" (or plain "Y") is an output, and its
    // inputs are the other pins named "<n>..." (for "Y": every input), in pin order.
    public static LogicIcModel Gates(string id, string pinout, OutputStage stage, Func<bool[], bool> gate)
    {
        string[] names = pinout.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var outputs = names
            .Where(n => n.EndsWith('Y'))
            .Select(y =>
            {
                string prefix = y[..^1];
                string[] gateInputs = names
                    .Where(n => n != y && n.StartsWith(prefix, StringComparison.Ordinal)
                                && n is not ("GND" or "VCC" or "NC") && !n.EndsWith('Y'))
                    .ToArray();
                return (y, (Func<Func<string, bool>, bool?>)(level => gate(gateInputs.Select(level).ToArray())));
            })
            .ToArray();
        return Define(id, pinout, stage, outputs);
    }

    // A part whose named pins are inputs at some times and driven by the chip at others.
    public static LogicIcModel Bidirectional(
        string id, string pinout, string coverage, IReadOnlyList<CombinationBlock> blocks,
        params (string Pin, DriveRule Drive)[] pins)
    {
        var rules = pins.ToDictionary(p => p.Pin, p => p.Drive, StringComparer.Ordinal);
        return new LogicIcModel(id, Build(id, pinout, rules, ModelPinKind.Bidirectional), coverage, blocks);
    }

    // A part that keeps a state: its outputs are rules of the state and the input levels, and its
    // test is a sequence ("functional"). "leftOpen" names pins every vector must leave open.
    public static LogicIcModel Sequential(
        string id, string pinout, int states, StepRule step, string[] leftOpen,
        params (string Output, OutputStage Stage, Func<int, Func<string, bool>, bool?> Rule)[] outputs)
    {
        var rules = outputs.ToDictionary(
            o => o.Output,
            o => (DriveRule)((socket, state) =>
                state < 0 ? ChipDrive.Unknown : Evaluate(o.Stage, () => o.Rule(state, Need(socket)))),
            StringComparer.Ordinal);
        return new LogicIcModel(id, Build(id, pinout, rules, ModelPinKind.Output, leftOpen), "functional",
            Array.Empty<CombinationBlock>(), states, step);
    }

    private static List<ModelPin> Build(
        string id, string pinout, Dictionary<string, DriveRule> rules, ModelPinKind ruled, string[]? leftOpen = null)
    {
        string[] names = pinout.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        string? unknown = rules.Keys.Concat(leftOpen ?? Array.Empty<string>()).FirstOrDefault(k => !names.Contains(k));
        if (unknown is not null)
            throw new ArgumentException($"{id}: pin {unknown} is not in the pin-out");

        return names.Select((name, i) => new ModelPin
        {
            Number = i + 1,
            Name = name,
            Kind = name switch
            {
                "GND" => ModelPinKind.Ground,
                "VCC" => ModelPinKind.Supply,
                "NC" => ModelPinKind.NotConnected,
                _ when leftOpen?.Contains(name) == true => ModelPinKind.LeftOpen,
                _ when rules.ContainsKey(name) => ruled,
                _ => ModelPinKind.Input,
            },
            Drive = rules.GetValueOrDefault(name),
        }).ToList();
    }

    // The levels a rule reads; an input the T48 does not drive makes the result Unknown.
    private static Func<string, bool> Need(Func<string, bool?> socket) =>
        name => socket(name) ?? throw new FloatingInputException();

    // An output's datasheet level (null = a disabled 3-state output) as what the chip drives.
    private static ChipDrive Evaluate(OutputStage stage, Func<bool?> rule)
    {
        bool? level;
        try
        {
            level = rule();
        }
        catch (FloatingInputException)
        {
            return ChipDrive.Unknown;
        }

        return stage switch
        {
            OutputStage.TotemPole => level!.Value ? ChipDrive.High : ChipDrive.Low,
            OutputStage.OpenCollector => level!.Value ? ChipDrive.None : ChipDrive.Low,
            _ => level is null ? ChipDrive.None : level.Value ? ChipDrive.High : ChipDrive.Low,
        };
    }

    private sealed class FloatingInputException : Exception;
}

internal static class LogicIcReferenceModels
{
    private const OutputStage TotemPole = OutputStage.TotemPole;
    private const OutputStage OpenCollector = OutputStage.OpenCollector;
    private const OutputStage ThreeState = OutputStage.ThreeState;

    // The layouts several parts share.
    private const string Quad2Input = "1A 1B 1Y 2A 2B 2Y GND 3Y 3A 3B 4Y 4A 4B VCC";       // 00 03 08 32 86 132
    private const string Quad2InputNor = "1Y 1A 1B 2Y 2A 2B GND 3A 3B 3Y 4A 4B 4Y VCC";    // 02
    private const string Hex1Input = "1A 1Y 2A 2Y 3A 3Y GND 4Y 4A 5Y 5A 6Y 6A VCC";        // 04 06 07 14
    private const string Triple3Input = "1A 1B 2A 2B 2C 2Y GND 3Y 3A 3B 3C 1Y 1C VCC";     // 10 11 27
    private const string Dual4Input = "1A 1B NC 1C 1D 1Y GND 2Y 2A 2B NC 2C 2D VCC";       // 20 21
    private const string QuadBuffer = "1OE 1A 1Y 2OE 2A 2Y GND 3Y 3A 3OE 4Y 4A 4OE VCC";   // 125 126
    private const string Quad2To1 = "S 1A 1B 1Y 2A 2B 2Y GND 3Y 3B 3A 4Y 4B 4A G VCC";     // 157 158 257 258
    private const string Octal = "1G 1A1 2Y4 1A2 2Y3 1A3 2Y2 1A4 2Y1 GND 2A1 1Y4 2A2 1Y3 2A3 1Y2 2A4 1Y1 2G VCC"; // 240 244

    private static bool Nand(bool[] a) => !a.All(x => x);
    private static bool And(bool[] a) => a.All(x => x);
    private static bool Nor(bool[] a) => !a.Any(x => x);
    private static bool Or(bool[] a) => a.Any(x => x);

    // Binary value of the named inputs, the FIRST name being the most significant bit.
    private static int Value(Func<string, bool> level, params string[] msbFirst) =>
        msbFirst.Aggregate(0, (sum, name) => sum * 2 + (level(name) ? 1 : 0));

    public static readonly IReadOnlyList<LogicIcModel> All = new[]
    {
        // ----- gates -----------------------------------------------------------------------
        LogicIcModel.Gates("74ls00", Quad2Input, TotemPole, Nand),
        LogicIcModel.Gates("74ls02", Quad2InputNor, TotemPole, Nor),
        LogicIcModel.Gates("74ls03", Quad2Input, OpenCollector, Nand),
        LogicIcModel.Gates("74ls04", Hex1Input, TotemPole, a => !a[0]),
        LogicIcModel.Gates("74ls06", Hex1Input, OpenCollector, a => !a[0]),
        LogicIcModel.Gates("74ls07", Hex1Input, OpenCollector, a => a[0]),
        LogicIcModel.Gates("74ls08", Quad2Input, TotemPole, And),
        LogicIcModel.Gates("74ls10", Triple3Input, TotemPole, Nand),
        LogicIcModel.Gates("74ls11", Triple3Input, TotemPole, And),
        LogicIcModel.Gates("74ls14", Hex1Input, TotemPole, a => !a[0]),
        LogicIcModel.Gates("74ls20", Dual4Input, TotemPole, Nand),
        LogicIcModel.Gates("74ls21", Dual4Input, TotemPole, And),
        LogicIcModel.Gates("74ls27", Triple3Input, TotemPole, Nor),
        LogicIcModel.Gates("74ls30", "A B C D E F GND Y NC NC G H NC VCC", TotemPole, Nand),
        LogicIcModel.Gates("74ls32", Quad2Input, TotemPole, Or),
        LogicIcModel.Gates("74ls86", Quad2Input, TotemPole, a => a[0] != a[1]),
        LogicIcModel.Gates("74ls132", Quad2Input, TotemPole, Nand),
        LogicIcModel.Gates("74ls133", "A B C D E F G GND Y H I J K L M VCC", TotemPole, Nand),
        LogicIcModel.Gates("74ls266", "1A 1B 1Y 2Y 2A 2B GND 3A 3B 3Y 4Y 4A 4B VCC", OpenCollector,
            a => a[0] == a[1]),

        // ----- 3-state buffers ---------------------------------------------------------------
        LogicIcModel.Define("74ls125", QuadBuffer, ThreeState,
            Enumerable.Range(1, 4).Select(n => ($"{n}Y",
                (Func<Func<string, bool>, bool?>)(p => p($"{n}OE") ? null : p($"{n}A")))).ToArray()),
        LogicIcModel.Define("74ls126", QuadBuffer, ThreeState,
            Enumerable.Range(1, 4).Select(n => ($"{n}Y",
                (Func<Func<string, bool>, bool?>)(p => p($"{n}OE") ? p($"{n}A") : null))).ToArray()),
        LogicIcModel.Define("74ls240", Octal, ThreeState, OctalOutputs(inverting: true)),
        LogicIcModel.Define("74ls244", Octal, ThreeState, OctalOutputs(inverting: false)),
        LogicIcModel.Define("74ls365", "G1 A1 Y1 A2 Y2 A3 Y3 GND Y4 A4 Y5 A5 Y6 A6 G2 VCC", ThreeState,
            Enumerable.Range(1, 6).Select(n => ($"Y{n}",
                (Func<Func<string, bool>, bool?>)(p => p("G1") || p("G2") ? null : p($"A{n}")))).ToArray()),
        LogicIcModel.Define("74ls367", "1G 1A1 1Y1 1A2 1Y2 1A3 1Y3 GND 1Y4 1A4 2Y1 2A1 2Y2 2A2 2G VCC", ThreeState,
            new[] { "1Y1", "1Y2", "1Y3", "1Y4", "2Y1", "2Y2" }.Select(y => (y,
                (Func<Func<string, bool>, bool?>)(p => p($"{y[0]}G") ? null : p($"{y[0]}A{y[2]}")))).ToArray()),

        // ----- decoders ------------------------------------------------------------------------
        LogicIcModel.Define("74ls138", "A B C G2A G2B G1 Y7 GND Y6 Y5 Y4 Y3 Y2 Y1 Y0 VCC", TotemPole,
            Enumerable.Range(0, 8).Select(i => ($"Y{i}",
                (Func<Func<string, bool>, bool?>)(p =>
                    !(p("G1") && !p("G2A") && !p("G2B") && Value(p, "C", "B", "A") == i)))).ToArray()),
        LogicIcModel.Define("74ls139", "1G 1A 1B 1Y0 1Y1 1Y2 1Y3 GND 2Y3 2Y2 2Y1 2Y0 2B 2A 2G VCC", TotemPole,
            new[] { 1, 2 }.SelectMany(n => Enumerable.Range(0, 4).Select(i => ($"{n}Y{i}",
                (Func<Func<string, bool>, bool?>)(p =>
                    !(!p($"{n}G") && Value(p, $"{n}B", $"{n}A") == i))))).ToArray()),
        // 0-9 pull one output low; the invalid codes 10-15 leave every output off.
        LogicIcModel.Define("74ls145", "Y0 Y1 Y2 Y3 Y4 Y5 Y6 GND Y7 Y8 Y9 D C B A VCC", OpenCollector,
            Enumerable.Range(0, 10).Select(i => ($"Y{i}",
                (Func<Func<string, bool>, bool?>)(p => Value(p, "D", "C", "B", "A") != i))).ToArray()),

        // ----- multiplexers ----------------------------------------------------------------------
        // The strobe G is active low: high forces Y low.
        LogicIcModel.Define("74ls153", "1G B 1C3 1C2 1C1 1C0 1Y GND 2Y 2C0 2C1 2C2 2C3 A 2G VCC", TotemPole,
            new[] { 1, 2 }.Select(n => ($"{n}Y",
                (Func<Func<string, bool>, bool?>)(p =>
                    !p($"{n}G") && p($"{n}C{Value(p, "B", "A")}")))).ToArray()),
        LogicIcModel.Define("74ls157", Quad2To1, TotemPole, Quad2To1Outputs(inverting: false, threeState: false)),
        LogicIcModel.Define("74ls158", Quad2To1, TotemPole, Quad2To1Outputs(inverting: true, threeState: false)),
        LogicIcModel.Define("74ls257", Quad2To1, ThreeState, Quad2To1Outputs(inverting: false, threeState: true)),
        LogicIcModel.Define("74ls258", Quad2To1, ThreeState, Quad2To1Outputs(inverting: true, threeState: true)),

        // ----- parts whose pins swap between input and output -------------------------------------
        // CD4066B: the switch between nA and nB is closed while its control nC is high. A closed
        // switch carries the level the T48 drives on one side to the other; the chip drives nothing
        // itself. Every combination of the controls and the driven side, driving each side in turn.
        LogicIcModel.Bidirectional("4066", "1A 1B 2B 2A 2C 3C GND 3A 3B 4B 4A 4C 1C VCC", "exhaustive",
            new[]
            {
                Block("1A 2A 3A 4A 1C 2C 3C 4C"),
                Block("1B 2B 3B 4B 1C 2C 3C 4C"),
            },
            Enumerable.Range(1, 4).SelectMany(n => new[]
            {
                ($"{n}A", Through($"{n}C", $"{n}B")),
                ($"{n}B", Through($"{n}C", $"{n}A")),
            }).ToArray()),

        // ----- sequential parts: a state kept from one vector to the next -------------------------
        // 74LS74: two D flip-flops, each taking D on the RISING edge of CLK. PRE and CLR (active low)
        // act at once and outrank the clock; both low drives Q and QN high.
        LogicIcModel.Sequential("74ls74", "1CLR 1D 1CLK 1PRE 1Q 1QN GND 2QN 2Q 2PRE 2CLK 2D 2CLR VCC",
            states: 4, step: FlipFlops74, leftOpen: Array.Empty<string>(),
            new[] { 1, 2 }.SelectMany(n => new[]
            {
                ($"{n}Q", TotemPole, (Func<int, Func<string, bool>, bool?>)((s, p) =>
                    !p($"{n}PRE") || (p($"{n}CLR") && Bit(s, n)))),
                ($"{n}QN", TotemPole, (Func<int, Func<string, bool>, bool?>)((s, p) =>
                    !p($"{n}CLR") || (p($"{n}PRE") && !Bit(s, n)))),
            }).ToArray()),

        // 74LS373: eight latches; while LE is high each Q follows its D, while LE is low it holds.
        // OE high turns the outputs off - the latches keep working behind them.
        LogicIcModel.Sequential("74ls373", "OE 1Q 1D 2D 2Q 3Q 3D 4D 4Q GND LE 5Q 5D 6D 6Q 7Q 7D 8D 8Q VCC",
            states: 256, leftOpen: Array.Empty<string>(),
            step: (state, level, _) => level("LE")
                ? Enumerable.Range(1, 8).Sum(n => level($"{n}D") ? 1 << (n - 1) : 0)
                : state,
            outputs: Enumerable.Range(1, 8).Select(n => ($"{n}Q", ThreeState,
                (Func<int, Func<string, bool>, bool?>)((s, p) => p("OE") ? null : Bit(s, n)))).ToArray()),

        // 74LS193: CLR (high) clears and LOAD (low) loads at once, outranking the count inputs. It
        // counts up on a rising UP while DOWN is high, down on a rising DOWN while UP is high. CO
        // (low) is "15 while UP is low", BO (low) "0 while DOWN is low".
        LogicIcModel.Sequential("74ls193", "B QB QA DOWN UP QC QD GND D C LOAD CO BO CLR A VCC",
            states: 16, step: Counter193, leftOpen: Array.Empty<string>(),
            ("QA", TotemPole, (s, p) => Bit(s, 1)),
            ("QB", TotemPole, (s, p) => Bit(s, 2)),
            ("QC", TotemPole, (s, p) => Bit(s, 3)),
            ("QD", TotemPole, (s, p) => Bit(s, 4)),
            ("CO", TotemPole, (s, p) => !(s == 15 && !p("UP"))),
            ("BO", TotemPole, (s, p) => !(s == 0 && !p("DOWN")))),

        // 556: two 555 timers, tested with their inputs at 0 V and 5 V and CONTROL left open (the
        // thresholds are then 1/3 and 2/3 of the supply). RESET low wins; then TRIGGER low sets
        // (output high, discharge off); then THRESHOLD high resets; otherwise the timer holds.
        LogicIcModel.Sequential("556", "1DIS 1THR 1CON 1RES 1OUT 1TRG GND 2TRG 2OUT 2RES 2CON 2THR 2DIS VCC",
            states: 4, step: Timers556, leftOpen: new[] { "1CON", "2CON" },
            new[] { 1, 2 }.SelectMany(n => new[]
            {
                ($"{n}OUT", TotemPole, (Func<int, Func<string, bool>, bool?>)((s, p) => TimerHigh(s, p, n))),
                ($"{n}DIS", OpenCollector, (Func<int, Func<string, bool>, bool?>)((s, p) => TimerHigh(s, p, n))),
            }).ToArray()),
    };

    public static LogicIcModel? For(string catalogueId) =>
        LogicIcReferenceModels.All.FirstOrDefault(m => string.Equals(m.Id, catalogueId, StringComparison.OrdinalIgnoreCase));

    private static CombinationBlock Block(string driven) => new(driven.Split(' '));

    // Bit n (1 = the lowest) of a sequential part's state.
    private static bool Bit(int state, int n) => (state >> (n - 1) & 1) == 1;

    private static int WithBit(int state, int n, bool level) => level ? state | 1 << (n - 1) : state & ~(1 << (n - 1));

    private static int FlipFlops74(int state, Func<string, bool> level, string? changed)
    {
        foreach (int n in new[] { 1, 2 })
        {
            bool clr = level($"{n}CLR"), pre = level($"{n}PRE");
            if (!clr && pre)
                state = WithBit(state, n, false);
            else if (clr && !pre)
                state = WithBit(state, n, true);
            else if (clr && pre && changed == $"{n}CLK" && level(changed))
                state = WithBit(state, n, level($"{n}D"));
        }
        return state;
    }

    // An edge on UP or DOWN while the OTHER is low is not in the datasheet's table: undefined.
    private static int Counter193(int state, Func<string, bool> level, string? changed)
    {
        if (level("CLR"))
            return 0;
        if (!level("LOAD"))
            return (level("A") ? 1 : 0) | (level("B") ? 2 : 0) | (level("C") ? 4 : 0) | (level("D") ? 8 : 0);
        if (changed is not ("UP" or "DOWN") || !level(changed))
            return state;
        if (state < 0 || !level(changed == "UP" ? "DOWN" : "UP"))
            return -1;
        return (state + (changed == "UP" ? 1 : 15)) % 16;
    }

    private static int Timers556(int state, Func<string, bool> level, string? changed)
    {
        foreach (int n in new[] { 1, 2 })
        {
            if (!level($"{n}RES"))
                state = WithBit(state, n, false);
            else if (!level($"{n}TRG"))
                state = WithBit(state, n, true);
            else if (level($"{n}THR"))
                state = WithBit(state, n, false);
        }
        return state;
    }

    // A 556 timer's output (the discharge transistor is off exactly when the output is high).
    private static bool TimerHigh(int state, Func<string, bool> level, int n) =>
        level($"{n}RES") && (!level($"{n}TRG") || Bit(state, n));

    // A 4066 switch side: while the control is high, the level the T48 drives on the other side.
    private static DriveRule Through(string control, string otherSide) => (socket, _) => socket(control) switch
    {
        null => ChipDrive.Unknown,
        false => ChipDrive.None,
        true => socket(otherSide) switch
        {
            null => ChipDrive.None,
            true => ChipDrive.High,
            false => ChipDrive.Low,
        },
    };

    // 74LS240/244: the 1G (2G) enable, active low, gates the 1Ax -> 1Yx (2Ax -> 2Yx) buffers.
    private static (string, Func<Func<string, bool>, bool?>)[] OctalOutputs(bool inverting) =>
        new[] { 1, 2 }.SelectMany(g => Enumerable.Range(1, 4).Select(n => ($"{g}Y{n}",
            (Func<Func<string, bool>, bool?>)(p => p($"{g}G") ? null : p($"{g}A{n}") != inverting)))).ToArray();

    // 74LS157/158/257/258: S low picks A, high picks B. G high forces Y low (157), high (158), or
    // turns the outputs off (257/258).
    private static (string, Func<Func<string, bool>, bool?>)[] Quad2To1Outputs(bool inverting, bool threeState) =>
        Enumerable.Range(1, 4).Select(n => ($"{n}Y", (Func<Func<string, bool>, bool?>)(p =>
        {
            if (p("G"))
                return threeState ? null : inverting;
            bool picked = p("S") ? p($"{n}B") : p($"{n}A");
            return picked != inverting;
        }))).ToArray();
}
