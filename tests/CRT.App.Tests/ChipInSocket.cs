namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A SEQUENTIAL chip powered in the T48's socket (2026-10-11): what the T48 drives on each input,
// and the state the chip keeps from one vector to the next. A vector is played one pin change at
// a time, a "C" as a clock pulse.
//
// What the T48 does with a "C" is its closed firmware's business (minipro's t48.c only says "the
// pin marked as C" is pulsed before the pins are read). The manufacturer's own 74193 test in
// minipro's logicic.xml only reads right if the pin is driven high and then low, and every pin is
// held between vectors - so that is what the simulated T48 does (PulseShape.HighThenLow). The
// shipped-data tests play every sequence with the other plausible shapes too, and with the pin
// changes of a vector in every order, so a test cannot depend on which one the firmware does.
// ###########################################################################################
internal enum PulseShape { HighThenLow, LowHighLow, LowThenHigh }

// One thing a vector does to an input: a new level, or a pulse (several levels in a row).
internal sealed record PinEvent(string Pin, IReadOnlyList<bool> Levels, bool IsPulse);

internal sealed class ChipInSocket
{
    private readonly LogicIcModel model;
    private readonly ChipFault? fault;
    private readonly Dictionary<string, bool> socket;

    private ChipInSocket(LogicIcModel model, ChipFault? fault, Dictionary<string, bool> socket, int state)
    {
        this.model = model;
        this.fault = fault;
        this.socket = socket;
        this.State = state;
    }

    public int State { get; private set; }

    // Powering up into the first vector: every input is already at its level (a "C" low) and the
    // state settles from whatever the chip woke up holding - then the vector's pulses.
    public static ChipInSocket PowerUp(LogicIcModel model, string vector, int heldBefore, PulseShape shape, ChipFault? fault = null)
    {
        var socket = model.Inputs.ToDictionary(p => p.Name, p => vector[p.Number - 1] == '1', StringComparer.Ordinal);
        var chip = new ChipInSocket(model, fault, socket, heldBefore);
        chip.State = model.Step!(heldBefore, chip.Seen, null);
        foreach (PinEvent pulse in chip.EventsOf(vector, shape).Where(e => e.IsPulse))
            chip.Play(pulse);
        return chip;
    }

    public ChipInSocket Clone() =>
        new(this.model, this.fault, new Dictionary<string, bool>(this.socket, StringComparer.Ordinal), this.State);

    // Everything a vector changes, in the order the simulated T48 does it: the new levels in pin
    // order, then the pulses in pin order.
    public List<PinEvent> EventsOf(string vector, PulseShape shape)
    {
        var changes = new List<PinEvent>();
        var pulses = new List<PinEvent>();
        foreach (ModelPin pin in this.model.Inputs)
        {
            char c = vector[pin.Number - 1];
            if (c == 'C')
            {
                pulses.Add(new PinEvent(pin.Name, shape switch
                {
                    PulseShape.HighThenLow => new[] { true, false },
                    PulseShape.LowHighLow => new[] { false, true, false },
                    _ => new[] { false, true },
                }, IsPulse: true));
            }
            else if ((c == '1') != this.socket[pin.Name])
            {
                changes.Add(new PinEvent(pin.Name, new[] { c == '1' }, IsPulse: false));
            }
        }
        return changes.Concat(pulses).ToList();
    }

    public void Play(string vector, PulseShape shape)
    {
        foreach (PinEvent e in this.EventsOf(vector, shape))
            this.Play(e);
    }

    public void Play(PinEvent e)
    {
        foreach (bool level in e.Levels)
            this.Set(e.Pin, level);
    }

    // What the chip drives on each pin it can drive, a broken output stage included.
    public Dictionary<int, ChipDrive> Drives() =>
        ChipFault.ApplyTo(this.fault, this.model, this.model.Respond(name => this.SeenOrNull(name), this.State));

    // The level the chip sees on an input: the socket's, unless the input is stuck.
    private bool Seen(string name) =>
        this.fault is { Kind: ChipFaultKind.InputStuckLow or ChipFaultKind.InputStuckHigh } f
        && f.Pin == this.model.Pin(name).Number
            ? f.Kind == ChipFaultKind.InputStuckHigh
            : this.socket[name];

    private bool? SeenOrNull(string name) => this.socket.ContainsKey(name) ? this.Seen(name) : null;

    private void Set(string pin, bool level)
    {
        if (this.socket[pin] == level)
            return;

        bool before = this.Seen(pin);
        this.socket[pin] = level;
        if (this.Seen(pin) == before)
            return;

        // A GOOD chip's test must never reach what the datasheet leaves undefined (-1 stays, and
        // the shipped-data tests fail on it). A FAULTY chip can be driven there by its fault; it
        // is taken to hold its state - the 74LS193 counts only while both count inputs are high.
        int next = this.model.Step!(this.State, this.Seen, pin);
        this.State = next < 0 && this.fault is not null ? this.State : next;
    }
}
