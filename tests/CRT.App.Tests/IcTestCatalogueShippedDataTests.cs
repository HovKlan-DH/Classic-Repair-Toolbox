using System.Text.RegularExpressions;
using Handlers.DataHandling;
using Handlers.IcTesting;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The MiniPro IC tests CRT ships, held against the datasheets (2026-10-11).
//
// The catalogue in Assets/Data reaches every CRT through the data sync - released 2.5.0 reads it
// too - and a test in it is run on a real chip in a real socket. Nobody can try each part on
// hardware before it ships, so everything that CAN be checked without hardware is checked here:
//
//   - it is safe to put the chip in the socket: the supply is on the datasheet's supply pins,
//     every input is driven, and the T48 never drives a pin the chip drives - in any vector, nor
//     on the way from one vector to the next (for a part whose pins swap between input and output);
//   - every vector expects what the datasheet's function table says, so a good chip passes - for
//     a SEQUENTIAL part (flip-flop, latch, counter, timer) whatever it held at power-up, whatever
//     shape the T48's clock pulse has, and in whatever order the pins of a vector change;
//   - the test holds every input combination its coverage claims, once (a sequential part's test
//     claims "functional" and must take the chip through every state it has);
//   - run through CRT's own IcTestService and minipro output parser, against a T48 with a
//     SIMULATED chip in the socket (SimulatedT48Runner), a good chip passes, and a chip with any
//     one broken pin fails.
//
// The datasheet side is LogicIcReferenceModels, typed in from the datasheets and not from these
// vectors. A runnable logic part without a model fails here, so a new test cannot ship unchecked.
// ###########################################################################################
public sealed class IcTestCatalogueShippedDataTests
{
    private static readonly Lazy<List<IcTestEntry>> Shipped =
        new(() => IcTestCatalogue.ReadEntries(IcTestCatalogueShippedDataTests.CatalogueFolder()));

    public static TheoryData<string> RunnableLogicParts()
    {
        var data = new TheoryData<string>();
        foreach (IcTestEntry entry in Shipped.Value.Where(IsRunnableLogicPart))
            data.Add(entry.Id);
        return data;
    }

    // ------------------------------------------------------------------- the catalogue

    [Fact]
    public void Every_file_in_the_catalogue_is_read_as_an_entry()
    {
        // The loader skips a file it cannot parse rather than lose the whole catalogue - so a
        // malformed new file would simply be missing in CRT, with nothing said.
        int files = Directory.EnumerateFiles(CatalogueFolder(), "*.json")
            .Count(f => !string.Equals(Path.GetFileName(f), "index.json", StringComparison.OrdinalIgnoreCase));

        Assert.Equal(files, Shipped.Value.Count);
    }

    [Fact]
    public void Every_runnable_logic_part_has_a_datasheet_model_and_every_model_a_part()
    {
        var parts = Shipped.Value.Where(IsRunnableLogicPart)
            .Select(e => e.Id.ToLowerInvariant()).OrderBy(id => id).ToList();
        var models = LogicIcReferenceModels.All
            .Select(m => m.Id.ToLowerInvariant()).OrderBy(id => id).ToList();

        Assert.Equal(models, parts);
    }

    [Fact]
    public void No_two_entries_claim_the_same_technical_name()
    {
        // The first entry naming a part wins, so a second one would be dead without anybody noticing.
        var claimed = Shipped.Value
            .SelectMany(e => e.Match.Select(m => (Name: m.Trim().ToUpperInvariant(), e.Id)))
            .GroupBy(x => x.Name)
            .Where(g => g.Select(x => x.Id).Distinct().Count() > 1)
            .Select(g => g.Key)
            .ToList();

        Assert.Empty(claimed);
    }

    [Fact]
    public void Every_device_name_reaches_minipro_as_it_is_written()
    {
        // CRT writes the name into the XML's <ic name="..."> without escaping it and passes the
        // same text to "-p"; minipro compares the two raw (newer builds also split a name at ",").
        foreach (IcTestEntry entry in Shipped.Value)
        {
            string name = string.IsNullOrWhiteSpace(entry.DeviceName) ? entry.Id : entry.DeviceName;
            Assert.Matches(new Regex("^[A-Za-z0-9 ()-]+$"), name);
        }
    }

    [Theory]
    [InlineData("7406", "74ls06")]
    [InlineData("7407", "74ls07")]
    [InlineData("7402", "74ls02")]
    [InlineData("74F32", "74ls32")]
    [InlineData("74HCU04", "74ls04")]
    [InlineData("74LS132", "74ls132")]
    [InlineData("74LS133", "74ls133")]
    [InlineData("74LS145", "74ls145")]
    [InlineData("74LS153", "74ls153")]
    [InlineData("4066", "4066")]
    [InlineData("74LS74", "74ls74")]
    [InlineData("74LS373", "74ls373")]
    [InlineData("74LS193", "74ls193")]
    [InlineData("556", "556")]
    public void A_technical_name_the_boards_use_finds_its_test(string technicalName, string expectedId)
    {
        // These are spelled as the boards' "Technical name or value" spells them. A family
        // variant is the same pin-out and function table as the part whose test it runs.
        var index = IcTestCatalogue.IndexByMatch(Shipped.Value);

        IcTestEntry? entry = IcTestCatalogue.Find(index, technicalName);

        Assert.NotNull(entry);
        Assert.Equal(expectedId, entry!.Id, ignoreCase: true);
    }

    // ------------------------------------------------------------- against the datasheet

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public void Every_vector_is_safe_to_put_in_the_socket(string id)
    {
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);

        Assert.Equal(model.Pins.Count, entry.PinCount);
        Assert.Equal(5, entry.Voltage);
        for (int i = 0; i < entry.Vectors!.Count; i++)
        {
            string vector = entry.Vectors[i];
            Assert.True(vector.Length == entry.PinCount, $"vector {i} has {vector.Length} pins");
            foreach (ModelPin pin in model.Pins)
            {
                char c = vector[pin.Number - 1];
                string where = $"vector {i}, pin {pin.Number} ({pin.Name}) is '{c}'";
                switch (pin.Kind)
                {
                    case ModelPinKind.Ground:
                        Assert.True(c == 'G', where);
                        break;
                    case ModelPinKind.Supply:
                        Assert.True(c == 'V', where);
                        break;
                    case ModelPinKind.Input:
                        Assert.True(c is '0' or '1' || (c == 'C' && model.IsSequential),
                            where + " - every input is driven (a clock pulse only on a sequential part)");
                        break;
                    case ModelPinKind.LeftOpen:
                        Assert.True(c == 'X', where + " - this pin must be left open");
                        break;
                    case ModelPinKind.Output:
                        // A sequential test may leave an output unread ("X") where its level would
                        // depend on how the T48 shapes a clock pulse (the 74LS193's CO and BO).
                        Assert.True(c is 'L' or 'H' or 'Z' || (c == 'X' && model.IsSequential),
                            where + " - an output is only ever read");
                        break;
                    case ModelPinKind.Bidirectional:
                        Assert.True(c is '0' or '1' or 'L' or 'H' or 'Z' or 'X', where);
                        break;
                    default:
                        Assert.True(c is '0' or '1' or 'X', where);
                        break;
                }
            }

            foreach (var (number, drive) in model.Respond(model.SocketOf(vector)))
            {
                Assert.False(drive != ChipDrive.None && vector[number - 1] is '0' or '1',
                    $"vector {i} ({vector}): the T48 and the chip both drive pin {number}");
            }
        }
    }

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public void Every_step_from_one_vector_to_the_next_is_safe(string id)
    {
        // The T48 does not change every pin at the same instant, and the chip answers a changed
        // input at once. So on the way from one vector to the next - and from the last back to the
        // first, where the second pass starts - every MIX of the two vectors' pins must be safe:
        // nowhere may the T48 drive a pin the chip could be driving. A part whose pins never swap
        // roles is safe by the per-vector check (the T48 only ever drives its inputs).
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);
        if (!model.SwapsPinRoles)
            return;

        List<string> vectors = entry.Vectors!;
        for (int i = 0; i < vectors.Count; i++)
        {
            string from = vectors[i];
            string to = vectors[(i + 1) % vectors.Count];
            int[] changing = Enumerable.Range(0, from.Length)
                .Where(k => Driven(from[k]) != Driven(to[k]))
                .ToArray();
            Assert.True(changing.Length <= 20, $"step {i}: {changing.Length} pins change at once");

            for (int mask = 0; mask < 1 << changing.Length; mask++)
            {
                char[] mix = from.ToCharArray();
                for (int k = 0; k < changing.Length; k++)
                {
                    if ((mask >> k & 1) == 1)
                        mix[changing[k]] = to[changing[k]];
                }

                string between = new(mix);
                foreach (var (number, drive) in model.Respond(model.SocketOf(between)))
                {
                    Assert.False(drive != ChipDrive.None && between[number - 1] is '0' or '1',
                        $"from vector {i} ({from}) to the next ({to}): at {between} the T48 and the chip could both drive pin {number}");
                }
            }
        }

        static char Driven(char c) => c is '0' or '1' ? c : '.';
    }

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public void Every_vector_expects_what_the_datasheet_says(string id)
    {
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);
        if (model.IsSequential)
        {
            PlayEveryWay(entry, model);
            return;
        }

        for (int i = 0; i < entry.Vectors!.Count; i++)
        {
            string vector = entry.Vectors[i];
            Dictionary<int, ChipDrive> drives = model.Respond(model.SocketOf(vector));

            foreach (ModelPin pin in model.Pins.Where(p => vector[p.Number - 1] is 'L' or 'H' or 'Z'))
            {
                ChipDrive drive = drives.GetValueOrDefault(pin.Number, ChipDrive.None);
                Assert.True(drive != ChipDrive.Unknown,
                    $"vector {i} ({vector}), pin {pin.Number} ({pin.Name}) is read, but its level depends on an input the vector leaves floating");

                char expected = drive switch
                {
                    ChipDrive.High => 'H',
                    ChipDrive.Low => 'L',
                    _ => 'Z',
                };
                char actual = vector[pin.Number - 1];
                Assert.True(expected == actual,
                    $"vector {i} ({vector}), pin {pin.Number} ({pin.Name}): the datasheet says '{expected}', the catalogue '{actual}'");
            }
        }
    }

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public void The_test_holds_every_combination_its_coverage_claims_once(string id)
    {
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);

        Assert.Equal(model.IsSequential ? "logic-sequential" : "logic-combinational", entry.Kind);
        if (model.IsSequential)
        {
            // A sequence cannot hold every history, so it claims "functional" - and must at least
            // take the chip through every state it has.
            var socket = ChipInSocket.PowerUp(model, entry.Vectors![0], heldBefore: 0, PulseShape.HighThenLow);
            var visited = new HashSet<int> { socket.State };
            foreach (string vector in entry.Vectors.Skip(1))
            {
                socket.Play(vector, PulseShape.HighThenLow);
                visited.Add(socket.State);
            }

            Assert.Equal("functional", entry.Coverage);
            Assert.Equal(model.Inputs.Count(), entry.Inputs);
            Assert.Equal(Enumerable.Range(0, model.States), visited.Order());
            return;
        }

        var keys = entry.Vectors!.Select(model.KeyOf).ToList();
        var held = keys.ToHashSet();
        var wanted = model.Blocks.SelectMany(model.Combinations).ToList();

        var missing = wanted.Where(w => !held.Contains(w)).ToList();

        Assert.Equal(model.Coverage, entry.Coverage);
        Assert.Equal(keys.Count, held.Count);
        Assert.True(missing.Count == 0, $"{missing.Count} combinations are not tested, e.g. {missing.FirstOrDefault()}");
        if (model.Coverage == "exhaustive")
            Assert.Equal(wanted.Count, keys.Count);
        if (!model.SwapsPinRoles)
            Assert.Equal(model.Inputs.Count(), entry.Inputs);
    }

    // ------------------------------------------- through CRT's own pipeline, simulated T48

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public async Task A_good_chip_passes(string id)
    {
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);

        IcTestResult result = await RunAsync(entry, new SimulatedT48Runner(model));

        Assert.Equal(TestOutcome.Pass, result.Outcome);
        Assert.Equal(entry.Vectors!.Count, result.TotalVectors);
        Assert.Empty(result.FailingPins);
    }

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public async Task A_broken_output_fails_and_is_named(string id)
    {
        // Stuck low, stuck high, or not driving at all - each must fail on THAT pin, or the test
        // would pass a chip with a dead gate.
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);

        foreach (ModelPin output in model.ChipDriven)
        {
            foreach (ChipFaultKind kind in new[] { ChipFaultKind.OutputStuckLow, ChipFaultKind.OutputStuckHigh, ChipFaultKind.OutputOpen })
            {
                IcTestResult result = await RunAsync(entry, new SimulatedT48Runner(model, new ChipFault(kind, output.Number)));

                Assert.True(result.Outcome == TestOutcome.Fail, $"{kind} on pin {output.Number} ({output.Name}) passed");
                Assert.Contains(output.Number, result.FailingPins);
            }
        }
    }

    [Theory]
    [MemberData(nameof(RunnableLogicParts))]
    public async Task A_stuck_input_fails(string id)
    {
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);

        foreach (ModelPin input in model.Pins.Where(p => p.Kind is ModelPinKind.Input or ModelPinKind.Bidirectional))
        {
            foreach (ChipFaultKind kind in new[] { ChipFaultKind.InputStuckLow, ChipFaultKind.InputStuckHigh })
            {
                IcTestResult result = await RunAsync(entry, new SimulatedT48Runner(model, new ChipFault(kind, input.Number)));

                Assert.True(result.Outcome == TestOutcome.Fail, $"{kind} on pin {input.Number} ({input.Name}) passed");
            }
        }
    }

    [Theory]
    [InlineData("4066")]
    public async Task A_leaking_signal_pin_fails_on_either_side(string id)
    {
        // Each signal pin is read at least once with nothing driving it, so a pin that leaks to
        // ground or the supply harder than the T48's weak pull is caught - on BOTH sides of a
        // switch, not only the side the T48 reads in the first direction (2026-10-11: the
        // one-direction 4066 test never read pins 1, 4, 8 and 11).
        (IcTestEntry entry, LogicIcModel model) = EntryAndModel(id);

        foreach (ModelPin pin in model.Pins.Where(p => p.Kind == ModelPinKind.Bidirectional))
        {
            foreach (ChipFaultKind kind in new[] { ChipFaultKind.PinLeaksLow, ChipFaultKind.PinLeaksHigh })
            {
                IcTestResult result = await RunAsync(entry, new SimulatedT48Runner(model, new ChipFault(kind, pin.Number)));

                Assert.True(result.Outcome == TestOutcome.Fail, $"{kind} on pin {pin.Number} ({pin.Name}) passed");
                Assert.Contains(pin.Number, result.FailingPins);
            }
        }
    }

    // --------------------------------------------------------------------------- helpers

    private static bool IsRunnableLogicPart(IcTestEntry entry) =>
        entry.IsTestable && entry.Kind is "logic-combinational" or "logic-sequential";

    // A sequential part's test must read the same however the T48 does what the datasheet cannot
    // tell us: whatever the chip held at power-up, whatever shape a clock pulse has, and in whatever
    // order the pins of one vector change. It must also never lean on behaviour the datasheet leaves
    // undefined, and the second pass (pull-down) must start exactly where the first did.
    private static void PlayEveryWay(IcTestEntry entry, LogicIcModel model)
    {
        List<string> vectors = entry.Vectors!;
        foreach (PulseShape shape in Enum.GetValues<PulseShape>())
        {
            var woken = Enumerable.Range(0, model.States)
                .Select(held => ChipInSocket.PowerUp(model, vectors[0], held, shape))
                .ToList();
            foreach (ChipInSocket chip in woken)
                AssertReads(model, vectors[0], $"vector 0 at power-up ({shape})", chip);
            Assert.True(woken.Select(c => c.State).Distinct().Count() == 1,
                $"vector 0 does not set a known state whatever the chip held at power-up ({shape})");

            ChipInSocket first = woken[0];
            ChipInSocket at = first;
            for (int i = 1; i <= vectors.Count; i++)
            {
                string vector = vectors[i % vectors.Count];
                string where = i == vectors.Count ? $"vector 0 again, starting the second pass ({shape})" : $"vector {i} ({shape})";
                List<PinEvent> events = at.EventsOf(vector, shape);
                Assert.True(events.Count <= 8, $"{where}: {events.Count} pins change at once");

                ChipInSocket? settled = null;
                foreach (IReadOnlyList<PinEvent> order in Orders(events))
                {
                    ChipInSocket chip = at.Clone();
                    foreach (PinEvent e in order)
                        chip.Play(e);

                    Assert.True(chip.State >= 0, $"{where} ({vector}): relies on behaviour the datasheet leaves undefined");
                    Assert.True(settled is null || chip.State == settled.State,
                        $"{where} ({vector}): the state depends on the order the T48 changes the pins in");
                    AssertReads(model, vector, where, chip);
                    settled ??= chip;
                }
                at = settled!;
            }

            Assert.True(at.State == first.State, $"the second pass does not start where the first did ({shape})");
        }
    }

    private static void AssertReads(LogicIcModel model, string vector, string where, ChipInSocket chip)
    {
        Dictionary<int, ChipDrive> drives = chip.Drives();
        foreach (ModelPin pin in model.Pins.Where(p => vector[p.Number - 1] is 'L' or 'H' or 'Z'))
        {
            ChipDrive drive = drives.GetValueOrDefault(pin.Number, ChipDrive.None);
            char expected = drive switch
            {
                ChipDrive.High => 'H',
                ChipDrive.Low => 'L',
                ChipDrive.None => 'Z',
                _ => '?',
            };
            Assert.True(expected == vector[pin.Number - 1],
                $"{where} ({vector}), pin {pin.Number} ({pin.Name}): the datasheet says '{expected}', the catalogue '{vector[pin.Number - 1]}'");
        }
    }

    // Every order of a short list.
    private static IEnumerable<IReadOnlyList<T>> Orders<T>(IReadOnlyList<T> items)
    {
        if (items.Count <= 1)
        {
            yield return items;
            yield break;
        }

        for (int k = 0; k < items.Count; k++)
        {
            var rest = items.Where((_, j) => j != k).ToList();
            foreach (IReadOnlyList<T> tail in Orders(rest))
                yield return new[] { items[k] }.Concat(tail).ToList();
        }
    }

    private static (IcTestEntry Entry, LogicIcModel Model) EntryAndModel(string id)
    {
        IcTestEntry entry = Shipped.Value.Single(e => e.Id == id);
        LogicIcModel? model = LogicIcReferenceModels.For(id);
        Assert.NotNull(model);
        Assert.NotNull(entry.Vectors);
        return (entry, model!);
    }

    private static Task<IcTestResult> RunAsync(IcTestEntry entry, IMiniproRunner runner) =>
        new IcTestService(runner).RunAsync(entry, mode: null, output: null, ct: CancellationToken.None);

    // The catalogue as it ships: Assets/Data plus the folder CRT reads it from.
    private static string CatalogueFolder()
    {
        string? folder = AppContext.BaseDirectory;
        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        string relative = DataTreeUsage.MiniProTestsFolder.Replace('/', Path.DirectorySeparatorChar);
        return Path.Combine(folder!, "Assets", "Data", relative, "Catalogue");
    }
}
