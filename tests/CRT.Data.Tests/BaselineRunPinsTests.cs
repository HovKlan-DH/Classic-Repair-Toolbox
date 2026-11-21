using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BaselineRunPins - "Capture from oscilloscope pin-by-pin" (owner requests, 2026-11-18: "how I easily can take an entire
// oscilloscope baseline in one go - e.g. for updating it", and "Maybe the user can easily select the pins in scope?"; 2026-11-20,
// a window of its own): which pins a run can go round, the quick picks of them, the steps - every baseline of a pin, else one new
// one - and going round the steps picked.
// ###########################################################################################
public sealed class BaselineRunPinsTests
{
    private static readonly ComponentPinEntry[] Pinout =
    [
        new() { Pin = "1", Name = "1A", Direction = "Input" },
        new() { Pin = "3", Name = "1Y", Direction = "Output" },
        new() { Pin = "7", Name = "GND", Direction = "Ground" },
        new() { Pin = "14", Name = "VCC", Direction = "Power" },
    ];

    private static ComponentImageEntry Baseline(string pin, string name, string file) =>
        new() { BoardLabel = "U27", Pin = pin, Name = name, File = file };

    // The steps of the run's lines over `pins`.
    private static List<BaselineRunStep> Steps(IEnumerable<BaselineRunPin> pins, IReadOnlyList<ComponentImageEntry> baselines) =>
        BaselineRunPins.Lines(pins, baselines).Select(line => line.Step).ToList();

    // The case's every pin, in number order, each with what the pinout says of it - and how many baselines it has.
    [Fact]
    public void The_pins_are_the_cases_in_number_order_with_the_pinouts_names_and_directions()
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of(
            "DIP-14",
            Pinout,
            datasheet: null,
            [Baseline("3", "1Y", "Scope baseline/U27_3.png"), Baseline("3", "1Y", "Scope baseline/U27_3_b.png")]);

        Assert.Equal(Enumerable.Range(1, 14).Select(number => number.ToString()), pins.Select(pin => pin.Pin));
        Assert.Equal(("1Y", "Output", 2), (pins[2].Name, pins[2].Direction, pins[2].Baselines));
        Assert.Equal(("", "", 0), (pins[1].Name, pins[1].Direction, pins[1].Baselines));
        Assert.Equal(3, pins[2].Number);
    }

    // ###########################################################################################
    // With no case CRT can count, the pins are those the pinout and the baselines name - one per NUMBER, so "2 [64K]" is pin 2
    // - and a pin named rather than numbered (a regulator's "12V") comes after them.
    // ###########################################################################################
    [Fact]
    public void Without_a_case_the_pins_are_those_named_one_per_number_and_a_named_pin_after_them()
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of(
            package: null,
            [new ComponentPinEntry { Pin = "5", Name = "OUT" }],
            datasheet: null,
            [Baseline("12V", "Supply", "a.png"), Baseline("2 [64K]", "D0", "b.png"), Baseline("2", "D0", "c.png")]);

        Assert.Equal(["2", "5", "12V"], pins.Select(pin => pin.Pin));
        Assert.Equal((2, "D0"), (pins[0].Baselines, pins[0].Name));
        Assert.Null(pins[2].Number);
    }

    // A pin the pinout does not name takes the datasheet's name and direction, else - its name - its first baseline's.
    [Fact]
    public void A_pin_the_pinout_leaves_out_takes_the_datasheets_name_and_direction_else_a_baselines_name()
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of(
            "DIP-14",
            [],
            [new KiCadPin("2", "1B", "Input", string.Empty)],
            [Baseline("4", "2A", "a.png")]);

        Assert.Equal(("1B", "Input"), (pins[1].Name, pins[1].Direction));
        Assert.Equal(("2A", ""), (pins[3].Name, pins[3].Direction));
    }

    // ###########################################################################################
    // The quick picks: the pins with a baseline, the pins without one, every pin. GROUND, SUPPLY AND UNCONNECTED PINS ARE
    // PICKED AS ANY OTHER (owner report, 2026-11-18: "Pins without a baseline" left out GND and VCC - "This is wrong, as also
    // those pins should be measured").
    // ###########################################################################################
    [Fact]
    public void The_quick_picks_leave_no_pin_out_for_its_direction()
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of(
            "DIP-14",
            [.. Pinout, new ComponentPinEntry { Pin = "13", Direction = "Not connected" }],
            datasheet: null,
            [Baseline("3", "1Y", "a.png"), Baseline("14", "VCC", "b.png")]);

        Assert.Equal(["3", "14"], BaselineRunPins.WithBaseline(pins).Select(pin => pin.Pin));

        IReadOnlyList<BaselineRunPin> without = BaselineRunPins.WithoutBaseline(pins);
        Assert.Equal(12, without.Count);
        Assert.Contains(without, pin => pin.Pin == "7");
        Assert.Contains(without, pin => pin.Pin == "13");
        Assert.DoesNotContain(without, pin => pin.Pin is "3" or "14");

        Assert.Equal(14, BaselineRunPins.EveryPin(pins).Count);
    }

    // What a run starts with picked: the pins with a baseline - updating - else every pin.
    [Fact]
    public void A_run_starts_with_the_pins_with_a_baseline_else_every_pin()
    {
        IReadOnlyList<BaselineRunPin> withBaselines = BaselineRunPins.Of("DIP-14", Pinout, null, [Baseline("1", "1A", "a.png")]);
        Assert.Equal(["1"], BaselineRunPins.FirstChoice(withBaselines).Select(pin => pin.Pin));

        IReadOnlyList<BaselineRunPin> none = BaselineRunPins.Of("DIP-14", Pinout, null, []);
        Assert.Equal(14, BaselineRunPins.FirstChoice(none).Count);
    }

    // ###########################################################################################
    // The steps: each baseline of a pin picked - its own pin text, signal name and file, a pin measured twice twice - and one
    // new baseline, named as the pin, for a pin picked with none - the pins in the order given, a pin's baselines in the order
    // the component lists them.
    // ###########################################################################################
    [Fact]
    public void Each_baseline_of_a_pin_picked_is_a_step_and_a_pin_with_none_gets_a_new_one()
    {
        ComponentImageEntry[] baselines =
        [
            Baseline("3", "1Y", "Scope baseline/U27_3.png"),
            Baseline("1", "1A", "Scope baseline/U27_1.png"),
            Baseline("3 [running]", "1Y", "Scope baseline/U27_3_b.png"),
        ];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, baselines);

        IReadOnlyList<BaselineRunStep> steps = Steps([pins[0], pins[1], pins[2]], baselines);

        Assert.Equal(
            [("1", "1A", "Scope baseline/U27_1.png", true), ("2", "", "", false), ("3", "1Y", "Scope baseline/U27_3.png", true), ("3 [running]", "1Y", "Scope baseline/U27_3_b.png", true)],
            steps.Select(step => (step.PinText, step.Name, step.File, step.IsUpdate)));
    }

    [Fact]
    public void A_new_step_is_named_as_its_pin_and_captured_into_a_file_noted_as_it_goes()
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, []);
        BaselineRunStep step = Assert.Single(Steps([pins[0]], []));

        Assert.Equal(("1", "1A", false, false), (step.PinText, step.Name, step.IsUpdate, step.IsCaptured));

        // Captured, it waits - no file of its own until it is accepted.
        step.Captured();
        Assert.Equal(("", true, true), (step.File, step.IsCaptured, step.IsWaiting));

        Assert.True(step.Accepted(" Scope baseline/U27_1_1A.png "));

        // Accepted, it is still the new baseline it was - not one the component had.
        Assert.Equal(("Scope baseline/U27_1_1A.png", "", false), (step.File, step.OriginalFile, step.IsWaiting));
        Assert.False(step.IsUpdate);
    }

    // ###########################################################################################
    // A capture waits until it is accepted - and only then is it in the draft (owner requests, 2026-11-20: "it must stay on the pin,
    // so I can check that it is good ... "Accept and continue to next pin"", and "this must be accepted before it will be used"): only
    // a capture can be accepted, a capture taken again waits again, and the baseline the run started from is kept for the comparison.
    // ###########################################################################################
    [Fact]
    public void Only_a_capture_is_accepted_and_one_taken_again_waits_again()
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, [Baseline("1", "1A", "Scope baseline/U27_1.png")]);
        BaselineRunStep step = Steps([pins[0]], [Baseline("1", "1A", "Scope baseline/U27_1.png")])[0];

        Assert.False(step.Accepted("Scope baseline/U27_1_1A.png"));
        Assert.Equal(("Scope baseline/U27_1.png", false), (step.File, step.IsAccepted));

        step.Captured();
        Assert.True(step.IsWaiting);
        Assert.Equal("Scope baseline/U27_1.png", step.File);

        Assert.True(step.Accepted("Scope baseline/U27_1_1A.png"));
        Assert.True(step.IsAccepted);

        step.Captured();
        Assert.True(step.IsWaiting);
        Assert.Equal(("Scope baseline/U27_1_1A.png", "Scope baseline/U27_1.png", true), (step.File, step.OriginalFile, step.IsUpdate));
    }

    // Steps of pins 1 to `count` of a DIP, the ones in `picked` checked and the ones in `accepted` captured and accepted.
    private static List<BaselineRunStep> StepsOf(int count, int[] picked, int[] accepted)
    {
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of($"DIP-{count}", [], null, []);
        List<BaselineRunStep> steps = Steps(pins, []).ToList();

        foreach (BaselineRunStep step in steps)
        {
            step.IsPicked = picked.Contains(step.Pin.Number!.Value);

            if (accepted.Contains(step.Pin.Number!.Value))
            {
                Accepted(step);
            }
        }

        return steps;
    }

    private static void Accepted(BaselineRunStep step)
    {
        step.Captured();
        step.Accepted($"Scope baseline/{step.PinText}.png");
    }

    // ###########################################################################################
    // "Accept and continue to next pin" goes on to the next PICKED pin not accepted yet - round past the end, never the pin it is
    // on - and says when none is left (owner requests, 2026-11-20: "then what it the next pin I should do", and the accepting).
    // ###########################################################################################
    [Fact]
    public void The_next_to_accept_is_the_next_picked_pin_not_accepted_round_past_the_end()
    {
        // Pins 2, 4 and 6 picked of eight, 4 accepted.
        List<BaselineRunStep> steps = StepsOf(8, picked: [2, 4, 6], accepted: [4]);

        Assert.Equal(5, BaselineRunPins.NextToAccept(steps, 1));
        Assert.Equal(5, BaselineRunPins.NextToAccept(steps, 3));
        Assert.Equal(1, BaselineRunPins.NextToAccept(steps, 5));
        Assert.Equal(1, BaselineRunPins.NextToAccept(steps, 7));

        // Captured but not accepted is still to do.
        steps[5].Captured();
        Assert.Equal(5, BaselineRunPins.NextToAccept(steps, 1));

        // Never the one it is on: pin 6 accepted too, pin 2 is the only one left - none after it, from it.
        steps[5].Accepted("Scope baseline/6.png");
        Assert.Equal(-1, BaselineRunPins.NextToAccept(steps, 1));
        Assert.Equal(1, BaselineRunPins.NextToAccept(steps, 3));

        // Every picked pin accepted: none - a pin not picked is never next.
        Accepted(steps[1]);
        Assert.Equal(-1, BaselineRunPins.NextToAccept(steps, 0));
        Assert.Equal(-1, BaselineRunPins.NextToAccept([], 0));
    }

    // Left and Right go round the picked pins, past either end - the pin it is on when it is the only one, none with none picked.
    [Fact]
    public void Left_and_right_go_round_the_picked_pins_past_either_end()
    {
        List<BaselineRunStep> steps = StepsOf(8, picked: [2, 4, 6], accepted: []);

        Assert.Equal(3, BaselineRunPins.Neighbour(steps, 1, 1));
        Assert.Equal(5, BaselineRunPins.Neighbour(steps, 1, -1));
        Assert.Equal(1, BaselineRunPins.Neighbour(steps, 5, 1));

        // From a pin not picked, the picked one either side.
        Assert.Equal(1, BaselineRunPins.Neighbour(steps, 0, 1));
        Assert.Equal(5, BaselineRunPins.Neighbour(steps, 0, -1));

        List<BaselineRunStep> one = StepsOf(8, picked: [4], accepted: []);
        Assert.Equal(3, BaselineRunPins.Neighbour(one, 3, 1));

        List<BaselineRunStep> none = StepsOf(8, picked: [], accepted: []);
        Assert.Equal(-1, BaselineRunPins.Neighbour(none, 3, 1));
    }

    // ###########################################################################################
    // The number keys go to the pin typed, as the popup's go to a pin's first image: its first step - "2" finds "2 [64K]" too -
    // and none for a pin the component has no step of.
    // ###########################################################################################
    [Fact]
    public void A_pin_typed_is_found_as_its_first_step_whatever_its_baseline_writes_after_the_number()
    {
        ComponentImageEntry[] baselines =
        [
            Baseline("2 [64K]", "D0", "a.png"),
            Baseline("2", "D0", "b.png"),
            Baseline("12V", "Supply", "c.png"),
        ];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of(null, [new ComponentPinEntry { Pin = "1" }], null, baselines);
        IReadOnlyList<BaselineRunStep> steps = Steps(pins, baselines);

        Assert.Equal(["1", "2 [64K]", "2", "12V"], steps.Select(step => step.PinText));
        Assert.Equal(1, BaselineRunPins.IndexOfPin(steps, "2"));
        Assert.Equal(0, BaselineRunPins.IndexOfPin(steps, "1"));
        Assert.Equal(3, BaselineRunPins.IndexOfPin(steps, "12v"));
        Assert.Equal(-1, BaselineRunPins.IndexOfPin(steps, "3"));
        Assert.Equal(-1, BaselineRunPins.IndexOfPin(steps, ""));
    }

    // ###########################################################################################
    // A BASELINE WITH NO IMAGE YET IS A LINE OF ITS OWN, captured into that very baseline (code review, 2026-11-21: it read "No
    // baseline yet", and its capture went in as a second baseline beside it, which still had none). Each line says the baseline
    // it is of.
    // ###########################################################################################
    [Fact]
    public void A_baseline_with_no_image_is_a_line_of_its_own_and_each_line_says_its_baseline()
    {
        ComponentImageEntry[] baselines = [Baseline("3", "1Y", "Scope baseline/U27_3.png"), Baseline("1", "1A", "")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, baselines);

        IReadOnlyList<BaselineRunLine> lines = BaselineRunPins.Lines([pins[0], pins[1], pins[2]], baselines);

        Assert.Equal([1, -1, 0], lines.Select(line => line.Row));
        Assert.Equal(
            [(true, false), (false, false), (true, true)],
            lines.Select(line => (line.Step.HadBaseline, line.Step.IsUpdate)));
        Assert.Equal(1, BaselineRunPins.Match(lines.Select(line => line.Step).ToList(), baselines)[0]);
    }

    // ###########################################################################################
    // EACH BASELINE ONE LINE'S, FOUND BY ITS FILE FIRST, THEN ITS VERY PIN TEXT (code review, 2026-11-21: two baselines of pin 2
    // with one signal name, "2 [64K]" and "2 [256K]" - a line that lost its file was matched by the pin's NUMBER, the first of
    // them, and its capture replaced the 64K baseline with the 256K screen).
    // ###########################################################################################
    [Fact]
    public void A_line_finds_its_own_baseline_by_file_then_its_very_pin_text_and_never_one_another_line_holds()
    {
        ComponentImageEntry[] before = [Baseline("2 [64K]", "/CAS", "a.png"), Baseline("2 [256K]", "/CAS", "b.png")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-16", [], null, before);
        List<BaselineRunStep> steps = Steps(pins, before).Where(step => step.HadBaseline).ToList();

        // Both images changed in the pane since: neither file is there - each line still finds its own pin's baseline.
        ComponentImageEntry[] changed = [Baseline("2 [64K]", "/CAS", "a2.png"), Baseline("2 [256K]", "/CAS", "b2.png")];
        Assert.Equal([0, 1], BaselineRunPins.Match(steps, changed));

        // Listed the other way round, still each its own.
        Assert.Equal([1, 0], BaselineRunPins.Match(steps, [changed[1], changed[0]]));

        // The 256K baseline's capture accepted, then taken back by an "Undo": the line is no longer accepted, and finds the
        // baseline it started from - not the 64K one.
        steps[1].Captured();
        steps[1].Accepted("b-captured.png");
        IReadOnlyList<BaselineRunLine> lines = BaselineRunPins.Follow(steps, pins, before);

        Assert.False(steps[1].IsAccepted);
        Assert.Equal(("b.png", false), (steps[1].File, steps[1].IsCaptured));
        Assert.Equal([0, 1], lines.Where(line => line.Step.HadBaseline).Select(line => line.Row));
        Assert.Same(steps[1], lines.Single(line => line.Row == 1).Step);
    }

    // ###########################################################################################
    // AN ACCEPTED CAPTURE THE DRAFT NO LONGER HAS IS NO LONGER ACCEPTED (code review, 2026-11-21: an "Undo" of the capture that made
    // pin 5's baseline left its line "Accepted", counted, framed on the case and skipped by "Accept and continue"): the line is the
    // new line it was before it, the very line - and accepted again, the baseline is its own once more.
    // ###########################################################################################
    [Fact]
    public void An_accepted_capture_the_draft_no_longer_has_is_its_new_line_again()
    {
        ComponentImageEntry[] baselines = [Baseline("1", "1A", "Scope baseline/U27_1.png")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, baselines);
        List<BaselineRunStep> steps = Steps(pins, baselines).ToList();
        BaselineRunStep pin5 = steps[4];

        pin5.Captured();
        pin5.Accepted("Scope baseline/U27_5.png");

        // In the draft: held by its baseline.
        ComponentImageEntry[] with = [.. baselines, Baseline("5", "2A", "Scope baseline/U27_5.png")];
        IReadOnlyList<BaselineRunLine> accepted = BaselineRunPins.Follow(steps, BaselineRunPins.Of("DIP-14", Pinout, null, with), with);
        Assert.True(pin5.IsAccepted);
        Assert.Equal(1, accepted.Single(line => ReferenceEquals(line.Step, pin5)).Row);

        // Taken back by an "Undo".
        IReadOnlyList<BaselineRunLine> undone = BaselineRunPins.Follow(steps, pins, baselines);
        Assert.Equal((false, false, ""), (pin5.IsAccepted, pin5.IsCaptured, pin5.File));
        Assert.Equal(-1, undone.Single(line => ReferenceEquals(line.Step, pin5)).Row);
        Assert.Equal(14, undone.Count);
        Assert.False(pin5.LostAcceptance());

        // A "Redo" brings it back: accepted again, held by its baseline (code review, 2026-11-21).
        IReadOnlyList<BaselineRunLine> redone = BaselineRunPins.Follow(steps, BaselineRunPins.Of("DIP-14", Pinout, null, with), with);
        Assert.Equal((true, "Scope baseline/U27_5.png", ""), (pin5.IsAccepted, pin5.File, pin5.LostFile));
        Assert.Equal(1, redone.Single(line => ReferenceEquals(line.Step, pin5)).Row);

        // Captured again after an "Undo", a "Redo" leaves the capture waiting alone.
        BaselineRunPins.Follow(steps, pins, baselines);
        pin5.Captured();
        BaselineRunPins.Follow(steps, BaselineRunPins.Of("DIP-14", Pinout, null, with), with);
        Assert.True(pin5.IsWaiting);
    }

    // ###########################################################################################
    // A BASELINE NO LINE HOLDS GETS ONE (code review, 2026-11-21: one added, or given another pin, in the pane while the window was
    // open had no line, and its "Capture again" did nothing): in its pin's place, in place of that pin's new line with nothing
    // captured - picked as that line was. Every other line is the very line it was, its pick kept.
    // ###########################################################################################
    [Fact]
    public void A_baseline_added_gets_its_line_in_place_of_its_pins_empty_new_line_picked_as_that_was()
    {
        ComponentImageEntry[] baselines = [Baseline("1", "1A", "Scope baseline/U27_1.png")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, baselines);
        List<BaselineRunStep> steps = Steps(pins, baselines).ToList();
        steps[4].IsPicked = true;
        steps[0].IsPicked = true;

        ComponentImageEntry[] added = [.. baselines, Baseline("5", "2A", "Scope baseline/U27_5.png")];
        IReadOnlyList<BaselineRunLine> lines = BaselineRunPins.Follow(steps, BaselineRunPins.Of("DIP-14", Pinout, null, added), added);

        Assert.Equal(14, lines.Count);
        BaselineRunLine pin5 = lines[4];
        Assert.Equal(("5", "2A", "Scope baseline/U27_5.png", 1, true, true), (pin5.Step.PinText, pin5.Step.Name, pin5.Step.File, pin5.Row, pin5.Step.HadBaseline, pin5.Step.IsPicked));
        Assert.DoesNotContain(lines, line => ReferenceEquals(line.Step, steps[4]));

        // The rest are the lines they were.
        Assert.All(lines.Where(line => line.Step.PinText != "5"), line => Assert.Contains(line.Step, steps));
        Assert.True(lines[0].Step.IsPicked);
        Assert.False(lines[1].Step.IsPicked);
    }

    // ###########################################################################################
    // A LINE OF A BASELINE GONE GOES - its pin back to a new line - UNLESS A CAPTURE OF IT WAITS: that is never dropped, and goes in
    // as a new baseline of its pin once accepted. A capture waiting on a pin the run no longer lists stays too, last.
    // ###########################################################################################
    [Fact]
    public void A_line_of_a_baseline_gone_goes_unless_a_capture_of_it_waits()
    {
        ComponentImageEntry[] baselines = [Baseline("1", "1A", "Scope baseline/U27_1.png"), Baseline("3", "1Y", "Scope baseline/U27_3.png")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, baselines);
        List<BaselineRunStep> steps = Steps(pins, baselines).ToList();
        BaselineRunStep pin3 = steps[2];

        ComponentImageEntry[] without = [baselines[0]];
        IReadOnlyList<BaselineRunPin> pinsNow = BaselineRunPins.Of("DIP-14", Pinout, null, without);

        // Nothing captured: pin 3 a new line again.
        IReadOnlyList<BaselineRunLine> gone = BaselineRunPins.Follow(steps, pinsNow, without);
        Assert.Equal(14, gone.Count);
        Assert.DoesNotContain(gone, line => ReferenceEquals(line.Step, pin3));
        Assert.Equal(("3", "1Y", false, -1), (gone[2].Step.PinText, gone[2].Step.Name, gone[2].Step.HadBaseline, gone[2].Row));

        // A capture waiting on it: the very line stays, in its pin's place.
        pin3.Captured();
        IReadOnlyList<BaselineRunLine> waiting = BaselineRunPins.Follow(steps, pinsNow, without);
        Assert.Equal(14, waiting.Count);
        Assert.Same(pin3, waiting[2].Step);
        Assert.Equal(-1, waiting[2].Row);

        // On a pin the run no longer lists - a named one, its baseline gone - it stays, last.
        ComponentImageEntry[] named = [Baseline("12V", "Supply", "a.png")];
        IReadOnlyList<BaselineRunPin> namedPins = BaselineRunPins.Of(null, [], null, named);
        List<BaselineRunStep> namedSteps = Steps(namedPins, named).ToList();
        namedSteps[0].Captured();

        IReadOnlyList<BaselineRunLine> after = BaselineRunPins.Follow(namedSteps, [], []);
        Assert.Same(namedSteps[0], Assert.Single(after).Step);
    }

    // ###########################################################################################
    // A BASELINE GIVEN ANOTHER PIN IN THE PANE IS THAT PIN'S (code review, 2026-11-21: matched by its file, the line kept pin 3 - its
    // label, its place on the case - and a capture of pin 3 went into the baseline that was now pin 4's). Its pin 3 line goes, pin 3
    // a new line again, and pin 4 gets the baseline's line. A baseline renamed keeps its line, saying the name it has now.
    // ###########################################################################################
    [Fact]
    public void A_baseline_given_another_pin_is_that_pins_and_one_renamed_keeps_its_line_under_its_new_name()
    {
        ComponentImageEntry[] baselines = [Baseline("1", "1A", "Scope baseline/U27_1.png"), Baseline("3", "1Y", "Scope baseline/U27_3.png")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, baselines);
        List<BaselineRunStep> steps = Steps(pins, baselines);
        BaselineRunStep pin1 = steps[0];
        BaselineRunStep pin3 = steps[2];

        ComponentImageEntry[] now = [Baseline("1", "A1", "Scope baseline/U27_1.png"), Baseline("4", "1Y", "Scope baseline/U27_3.png")];
        Assert.Equal([0, -1], BaselineRunPins.Match([pin1, pin3], now));

        IReadOnlyList<BaselineRunLine> lines = BaselineRunPins.Follow(steps, BaselineRunPins.Of("DIP-14", Pinout, null, now), now);

        Assert.Equal(14, lines.Count);
        Assert.Same(pin1, lines[0].Step);
        Assert.Equal("A1", pin1.Name);
        Assert.DoesNotContain(lines, line => ReferenceEquals(line.Step, pin3));
        Assert.Equal(("3", -1, false), (lines[2].Step.PinText, lines[2].Row, lines[2].Step.HadBaseline));
        Assert.Equal(("4", 1, "Scope baseline/U27_3.png"), (lines[3].Step.PinText, lines[3].Row, lines[3].Step.File));
    }

    // ###########################################################################################
    // A BASELINE WITH NO IMAGE, RENAMED IN THE PANE, IS STILL ITS LINE'S (code review, 2026-11-21): it has no file to be known by,
    // and a capture waiting on it went in as a second baseline beside it. The only one with no image of its pin left is it - two
    // such, and neither is guessed.
    // ###########################################################################################
    [Fact]
    public void A_baseline_with_no_image_renamed_is_still_its_lines_while_it_is_the_only_one_of_its_pin()
    {
        ComponentImageEntry[] before = [Baseline("3", "PA1", "")];
        IReadOnlyList<BaselineRunPin> pins = BaselineRunPins.Of("DIP-14", Pinout, null, before);
        BaselineRunStep step = Steps(pins, before)[2];
        step.Captured();

        Assert.Equal([0], BaselineRunPins.Match([step], [Baseline("3", "PA1 out", "")]));

        IReadOnlyList<BaselineRunLine> lines = BaselineRunPins.Follow([step], pins, [Baseline("3", "PA1 out", "")]);
        Assert.Same(step, lines[2].Step);
        Assert.Equal((0, "PA1 out"), (lines[2].Row, step.Name));

        // Renamed again, beside another with no image: neither is guessed.
        Assert.Equal([-1], BaselineRunPins.Match([step], [Baseline("3", "PA1 in", ""), Baseline("3", "PA2", "")]));
    }
}
