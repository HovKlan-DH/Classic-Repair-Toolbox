using System.Text;
using System.Xml.Linq;
using Handlers.IcTesting;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A T48 with a SIMULATED chip in its socket (2026-10-11).
//
// MockMiniproRunner, behind the IC test panel's demo mode, answers "Logic test successful"
// whatever the vectors say - it proves the panel, not the test. This runner proves the test: it
// reads the logicic XML CRT wrote, plays the T48's logic test against a chip built from its
// datasheet model (LogicIcReferenceModels), optionally with a fault, and prints what minipro
// prints, so CRT's own parser reads the verdict.
//
// The test is the one in the minipro build CRT bundles (KevinWelton/minipro-win c87e4f5980,
// t48.c, t48_logic_ic_test): the whole list of vectors is run twice, first with a weak pull-up on
// every pin a vector reads (L, H, Z), then with a weak pull-down. L must read 0 in both passes, H
// 1 in both, Z 1 with the pull-up and 0 with the pull-down. One error is counted per pin per
// vector, and the grid prints every vector, a failing pin marked "-". A sequential chip keeps its
// state from vector to vector and from the first pass into the second (ChipInSocket).
//
// What a real socket would do with a vector the datasheet does not allow - the T48 driving a pin
// a GOOD chip drives, power on the wrong pin, an input left floating, a read pin whose level
// depends on one - is not simulated: it throws, so a test can never pass on a vector that would
// be unsafe or undefined in a real socket. (A FAULTY chip may fight the T48; the T48 wins on the
// pins it drives, and those are not read.)
// ###########################################################################################
internal sealed class SimulatedT48Runner : IMiniproRunner
{
    private readonly LogicIcModel chip;
    private readonly ChipFault? fault;

    public SimulatedT48Runner(LogicIcModel chip, ChipFault? fault = null)
    {
        this.chip = chip;
        this.fault = fault;
    }

    public string? ResolveBinary() => "minipro (simulated T48)";

    public Task<MiniproRunResult> RunAsync(IReadOnlyList<string> args, IProgress<string>? output, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        string device = args[args.ToList().IndexOf("-p") + 1];
        string xmlPath = args[args.ToList().IndexOf("--logicic") + 1];
        if (!args.Contains("-T"))
            throw new InvalidOperationException("CRT did not ask for a logic test (-T)");

        // minipro finds the part by comparing -p with the <ic name>, ignoring capitals.
        XElement? ic = XDocument.Load(xmlPath).Descendants("ic")
            .FirstOrDefault(e => string.Equals((string?)e.Attribute("name"), device, StringComparison.OrdinalIgnoreCase));
        if (ic is null)
            return Task.FromResult(new MiniproRunResult(true, 1, "", $"Device {device} not found!\n"));

        int pins = int.Parse((string)ic.Attribute("pins")!);
        if (pins != this.chip.Pins.Count)
            throw new InvalidOperationException($"the XML says {pins} pins, the chip in the socket has {this.chip.Pins.Count}");

        List<string> vectors = ic.Elements("vector")
            .Select(v => string.Concat(((string)v).Where(c => !char.IsWhiteSpace(c))))
            .ToList();
        for (int i = 0; i < vectors.Count; i++)
            this.Validate(vectors[i], i);

        // The two passes, in the T48's order; a sequential chip carries its state across both.
        ChipInSocket? socket = null;
        List<bool[]> withPullUp = this.Pass(vectors, pullUp: true, ref socket);
        List<bool[]> withPullDown = this.Pass(vectors, pullUp: false, ref socket);

        var stdout = new StringBuilder();
        stdout.Append("      ");
        for (int pin = 1; pin <= pins; pin++)
            stdout.Append($"{pin,-3}");
        stdout.Append('\n');

        int errors = 0;
        for (int i = 0; i < vectors.Count; i++)
        {
            string vector = vectors[i];
            stdout.Append($"{i:0000}: ");
            for (int k = 0; k < pins; k++)
            {
                char want = vector[k];
                bool err = want switch
                {
                    'L' => withPullUp[i][k] || withPullDown[i][k],
                    'H' => !withPullUp[i][k] || !withPullDown[i][k],
                    'Z' => !withPullUp[i][k] || withPullDown[i][k],
                    _ => false,
                };
                errors += err ? 1 : 0;
                stdout.Append(err ? "\u001b[0;91m" : "\u001b[0m").Append(want).Append(err ? '-' : ' ').Append(' ');
            }
            stdout.Append("\u001b[0m\n");
        }

        string stderr = errors > 0
            ? $"Logic test failed: {errors} errors encountered.\n"
            : "Logic test successful.\n";

        foreach (string line in stdout.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries))
            output?.Report(line);

        return Task.FromResult(new MiniproRunResult(true, errors > 0 ? 1 : 0, stdout.ToString(), stderr));
    }

    private List<bool[]> Pass(List<string> vectors, bool pullUp, ref ChipInSocket? socket)
    {
        var reads = new List<bool[]>();
        for (int i = 0; i < vectors.Count; i++)
        {
            string vector = vectors[i];
            Dictionary<int, ChipDrive> drives;
            if (this.chip.IsSequential)
            {
                // Whatever a real chip wakes up holding, the shipped-data tests prove it does not
                // matter; this one wakes up holding 0.
                if (socket is null)
                    socket = ChipInSocket.PowerUp(this.chip, vector, heldBefore: 0, PulseShape.HighThenLow, this.fault);
                else
                    socket.Play(vector, PulseShape.HighThenLow);
                drives = socket.Drives();
            }
            else
            {
                drives = this.Combinational(vector, i);
            }
            reads.Add(this.Read(vector, i, drives, pullUp));
        }
        return reads;
    }

    // The characters a vector may put on each pin.
    private void Validate(string vector, int index)
    {
        if (vector.Length != this.chip.Pins.Count)
            throw new InvalidOperationException($"vector {index} has {vector.Length} pins, not {this.chip.Pins.Count}");

        foreach (ModelPin pin in this.chip.Pins)
        {
            char c = vector[pin.Number - 1];
            bool supplyChar = c is 'G' or 'V';
            bool wrong = pin.Kind switch
            {
                ModelPinKind.Ground => c != 'G',
                ModelPinKind.Supply => c != 'V',
                ModelPinKind.Input => !(c is '0' or '1' || (c == 'C' && this.chip.IsSequential)),
                ModelPinKind.Output => c is '0' or '1' or 'C' || supplyChar,
                ModelPinKind.LeftOpen => c != 'X',
                _ => supplyChar || c == 'C',
            };
            if (wrong)
                throw new InvalidOperationException($"vector {index}: '{c}' on pin {pin.Number} ({pin.Name}, {pin.Kind})");
        }
    }

    // What a combinational chip drives for one vector.
    private Dictionary<int, ChipDrive> Combinational(string vector, int index)
    {
        Func<string, bool?> socket = this.chip.SocketOf(vector);

        // A GOOD chip must never drive a pin the T48 drives.
        foreach (var (number, drive) in this.chip.Respond(socket))
        {
            if (drive != ChipDrive.None && vector[number - 1] is '0' or '1')
                throw new InvalidOperationException($"vector {index}: the T48 and the chip both drive pin {number}");
        }

        // The chip as it is: an input it sees at a fixed level, an output stage that is broken.
        Func<string, bool?> seen = name =>
            this.fault is { Kind: ChipFaultKind.InputStuckLow or ChipFaultKind.InputStuckHigh } f
            && f.Pin == this.chip.Pin(name).Number && socket(name) is not null
                ? f.Kind == ChipFaultKind.InputStuckHigh
                : socket(name);
        return ChipFault.ApplyTo(this.fault, this.chip, this.chip.Respond(seen));
    }

    // What every pin reads, with the weak pull-up or the weak pull-down on the pins the vector reads.
    private bool[] Read(string vector, int index, Dictionary<int, ChipDrive> drives, bool pullUp)
    {
        var read = new bool[vector.Length];
        foreach (ModelPin pin in this.chip.Pins)
        {
            if (vector[pin.Number - 1] is not ('L' or 'H' or 'Z'))
                continue;

            ChipDrive drive = drives.GetValueOrDefault(pin.Number, ChipDrive.None);
            read[pin.Number - 1] = drive switch
            {
                ChipDrive.High => true,
                ChipDrive.Low => false,
                ChipDrive.Unknown => throw new InvalidOperationException(
                    $"vector {index}: pin {pin.Number} ({pin.Name}) is read, but its level is not known (a floating input, or an undefined state)"),
                // Nothing drives the pin: it follows the T48's pull - unless it leaks harder.
                _ => this.fault is { Kind: ChipFaultKind.PinLeaksLow or ChipFaultKind.PinLeaksHigh } leak && leak.Pin == pin.Number
                    ? leak.Kind == ChipFaultKind.PinLeaksHigh
                    : pullUp,
            };
        }
        return read;
    }
}

internal enum ChipFaultKind
{
    OutputStuckLow, OutputStuckHigh, OutputOpen,
    InputStuckLow, InputStuckHigh,
    PinLeaksLow, PinLeaksHigh,
}

// One broken pin: an output stuck at a level or not driving at all (a blown output stage or a
// broken bond wire); an input the chip sees at a fixed level whatever the socket drives; or a pin
// leaking to ground or the supply harder than the T48's weak pull, which only shows when nothing
// drives the pin. On a pin the chip drives only some of the time, a stuck output shows while it
// drives.
internal sealed record ChipFault(ChipFaultKind Kind, int Pin)
{
    // What a chip with this fault drives, given what the good chip would.
    public static Dictionary<int, ChipDrive> ApplyTo(ChipFault? fault, LogicIcModel model, Dictionary<int, ChipDrive> drives)
    {
        if (fault is not null && drives.TryGetValue(fault.Pin, out ChipDrive good))
        {
            bool alwaysDriven = model.Pins[fault.Pin - 1].Kind == ModelPinKind.Output;
            drives[fault.Pin] = fault.Kind switch
            {
                ChipFaultKind.OutputStuckLow when alwaysDriven || good != ChipDrive.None => ChipDrive.Low,
                ChipFaultKind.OutputStuckHigh when alwaysDriven || good != ChipDrive.None => ChipDrive.High,
                ChipFaultKind.OutputOpen => ChipDrive.None,
                _ => good,
            };
        }
        return drives;
    }
}
