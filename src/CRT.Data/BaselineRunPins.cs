using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // "CAPTURE FROM OSCILLOSCOPE PIN-BY-PIN" (owner requests, 2026-11-18: "how this better can be integrated with the pinout
    // info and how I easily can take an entire oscilloscope baseline in one go - e.g. for updating it"; of the pins a run goes
    // round, "Maybe the user can easily select the pins in scope?"; and 2026-11-20, a window of its own): which pins of a
    // component a run can go round, the quick picks of them, the run's steps - one per baseline of a pin, or one new baseline
    // for a pin that has none - and going round them. The window shows and drives it (BaselineCaptureWindow, opened by
    // ComponentPane.BaselineRun.cs); every rule is here.
    //
    //   - THE PINS: the case's (1 to its count), every pin the pinout lists, and every pin a baseline names - one per NUMBER
    //     ("2 [64K]" is pin 2, ComponentVocabulary.PinNumberOf), in number order; a pin named rather than numbered (a
    //     regulator's "12V") after them, in the order they come. Each with its name and direction: the pinout's, else the
    //     datasheet's, else (a name) the first baseline's of it.
    //   - EVERY PIN IS WORTH A BASELINE, ground and supply too (owner report, 2026-11-18: "Pins without a baseline" left out GND
    //     and VCC - "This is wrong, as also those pins should be measured"): no quick pick leaves a pin out for its direction.
    //   - THE STEPS: every baseline of each pin chosen, in the order the component lists them - a pin measured in two states has
    //     two - and a new one for a chosen pin that has none.
    // ###########################################################################################
    public static class BaselineRunPins
    {
        public static IReadOnlyList<BaselineRunPin> Of(
            string? package,
            IReadOnlyList<ComponentPinEntry> pins,
            IReadOnlyList<KiCadPin>? datasheet,
            IReadOnlyList<ComponentImageEntry> baselines)
        {
            ArgumentNullException.ThrowIfNull(pins);
            ArgumentNullException.ThrowIfNull(baselines);

            var numbered = new SortedSet<int>();
            var named = new List<string>();

            void Take(string? pin)
            {
                string text = (pin ?? string.Empty).Trim();

                if (text.Length == 0)
                {
                    return;
                }

                if (ComponentVocabulary.PinNumberOf(text) is int number)
                {
                    numbered.Add(number);
                }
                else if (!named.Contains(text, StringComparer.OrdinalIgnoreCase))
                {
                    named.Add(text);
                }
            }

            if (ComponentVocabulary.PinCountOf(package) is int count)
            {
                for (int number = 1; number <= count; number++)
                {
                    numbered.Add(number);
                }
            }

            foreach (ComponentPinEntry pin in pins)
            {
                Take(pin.Pin);
            }

            foreach (ComponentImageEntry baseline in baselines)
            {
                Take(baseline.Pin);
            }

            return numbered
                .Select(number => number.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Concat(named)
                .Select(pin => BaselineRunPins.PinOf(pin, pins, datasheet, baselines))
                .ToList();
        }

        private static BaselineRunPin PinOf(
            string pin,
            IReadOnlyList<ComponentPinEntry> pins,
            IReadOnlyList<KiCadPin>? datasheet,
            IReadOnlyList<ComponentImageEntry> baselines)
        {
            ComponentPinEntry? listed = pins.FirstOrDefault(row => ComponentVocabulary.IsSamePin(row.Pin, pin));
            KiCadPin? fromDatasheet = datasheet?.FirstOrDefault(row => ComponentVocabulary.IsSamePin(row.Pin, pin));
            List<ComponentImageEntry> its = baselines.Where(row => ComponentVocabulary.IsSamePin(row.Pin, pin)).ToList();

            string name = BaselineRunPins.FirstFilled(listed?.Name, fromDatasheet?.Name, its.Select(row => row.Name).FirstOrDefault(text => !string.IsNullOrWhiteSpace(text)));
            string direction = BaselineRunPins.FirstFilled(listed?.Direction, fromDatasheet?.Direction);

            return new BaselineRunPin(pin, ComponentVocabulary.PinNumberOf(pin), name, direction, its.Count);
        }

        private static string FirstFilled(params string?[] texts) =>
            texts.FirstOrDefault(text => !string.IsNullOrWhiteSpace(text))?.Trim() ?? string.Empty;

        // ---- the quick picks ----------------------------------------------------------------------

        // The pins with a baseline - what a run updates.
        public static IReadOnlyList<BaselineRunPin> WithBaseline(IEnumerable<BaselineRunPin> pins) =>
            pins.Where(pin => pin.HasBaseline).ToList();

        // The pins with no baseline - the gaps a run fills, ground and supply pins among them.
        public static IReadOnlyList<BaselineRunPin> WithoutBaseline(IEnumerable<BaselineRunPin> pins) =>
            pins.Where(pin => !pin.HasBaseline).ToList();

        // Every pin.
        public static IReadOnlyList<BaselineRunPin> EveryPin(IEnumerable<BaselineRunPin> pins) => pins.ToList();

        // What a run starts with chosen: the pins with a baseline, else - none yet - every pin.
        public static IReadOnlyList<BaselineRunPin> FirstChoice(IReadOnlyList<BaselineRunPin> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            IReadOnlyList<BaselineRunPin> withBaseline = BaselineRunPins.WithBaseline(pins);
            return withBaseline.Count > 0 ? withBaseline : BaselineRunPins.EveryPin(pins);
        }

        // ---- the steps ----------------------------------------------------------------------------

        // ###########################################################################################
        // The run's lines over `chosen`, in the order given (the pane gives them in the pins' order): each baseline of a pin (its own
        // pin text, signal name and file), else one new baseline for it - its pin the pin's, its signal name the pin's - each line with
        // the baseline it is of: its index in `baselines`, -1 for a new one. A baseline with no image yet is a line of its own too,
        // captured into that very baseline (code review, 2026-11-21: it read "No baseline yet", and its capture was added as a second
        // baseline beside it).
        // ###########################################################################################
        public static IReadOnlyList<BaselineRunLine> Lines(IEnumerable<BaselineRunPin> chosen, IReadOnlyList<ComponentImageEntry> baselines)
        {
            ArgumentNullException.ThrowIfNull(chosen);
            ArgumentNullException.ThrowIfNull(baselines);

            var lines = new List<BaselineRunLine>();

            foreach (BaselineRunPin pin in chosen)
            {
                List<int> its = Enumerable.Range(0, baselines.Count).Where(index => ComponentVocabulary.IsSamePin(baselines[index].Pin, pin.Pin)).ToList();

                if (its.Count == 0)
                {
                    lines.Add(new BaselineRunLine(new BaselineRunStep(pin, pin.Pin, pin.Name, string.Empty), -1));
                    continue;
                }

                lines.AddRange(its.Select(index => new BaselineRunLine(BaselineRunPins.StepOf(pin, baselines[index]), index)));
            }

            return lines;
        }

        private static BaselineRunStep StepOf(BaselineRunPin pin, ComponentImageEntry baseline) =>
            new(pin, (baseline.Pin ?? string.Empty).Trim(), (baseline.Name ?? string.Empty).Trim(), (baseline.File ?? string.Empty).Trim(), hadBaseline: true);

        // ---- the baselines read again under a run (code review, 2026-11-21) ------------------------------------------------------

        // ###########################################################################################
        // Which baseline each step is of now - its index in `baselines`, -1 for none - EACH BASELINE ONE STEP'S: by the step's file,
        // else the file it started from, else - a step of a baseline the component had, or one accepted - its pin and signal name,
        // the very pin text before the pin's number ("2 [64K]" is never taken for "2 [256K]" while that one is there). Steps first
        // in their order at each rule; a baseline taken is never another's. ALWAYS OF THE STEP'S PIN, the file passes too (code
        // review, 2026-11-21): a baseline given another pin in the pane is another pin's, and a capture of the step's pin went into it.
        // ###########################################################################################
        public static IReadOnlyList<int> Match(IReadOnlyList<BaselineRunStep> steps, IReadOnlyList<ComponentImageEntry> baselines)
        {
            ArgumentNullException.ThrowIfNull(steps);
            ArgumentNullException.ThrowIfNull(baselines);

            int[] rowOf = Enumerable.Repeat(-1, steps.Count).ToArray();
            var taken = new bool[baselines.Count];

            void Pass(Func<BaselineRunStep, bool> asks, Func<BaselineRunStep, ComponentImageEntry, bool> isIt)
            {
                for (int step = 0; step < steps.Count; step++)
                {
                    if (rowOf[step] >= 0 || !asks(steps[step]))
                    {
                        continue;
                    }

                    for (int row = 0; row < baselines.Count; row++)
                    {
                        if (!taken[row] && isIt(steps[step], baselines[row]))
                        {
                            rowOf[step] = row;
                            taken[row] = true;
                            break;
                        }
                    }
                }
            }

            static string Trimmed(string? text) => (text ?? string.Empty).Trim();

            static bool SameName(BaselineRunStep step, ComponentImageEntry row) =>
                string.Equals(Trimmed(row.Name), step.Name.Trim(), StringComparison.OrdinalIgnoreCase);

            static bool SamePin(BaselineRunStep step, ComponentImageEntry row) => ComponentVocabulary.IsSamePin(row.Pin, step.PinText);

            Pass(step => step.File.Length > 0, (step, row) => SamePin(step, row) && string.Equals(Trimmed(row.File), step.File, StringComparison.Ordinal));
            Pass(step => step.OriginalFile.Length > 0, (step, row) => SamePin(step, row) && string.Equals(Trimmed(row.File), step.OriginalFile, StringComparison.Ordinal));

            static bool OfABaseline(BaselineRunStep step) => step.HadBaseline || step.IsAccepted;

            Pass(OfABaseline, (step, row) => string.Equals(Trimmed(row.Pin), step.PinText, StringComparison.OrdinalIgnoreCase) && SameName(step, row));
            Pass(OfABaseline, (step, row) => SamePin(step, row) && SameName(step, row));

            // A baseline with no image yet has no file to be known by: renamed in the pane, it is the only baseline with no image of
            // its pin left - that one, when there is exactly one (code review, 2026-11-21: its capture went in as a second baseline).
            for (int step = 0; step < steps.Count; step++)
            {
                if (rowOf[step] >= 0 || !steps[step].HadBaseline || steps[step].OriginalFile.Length > 0)
                {
                    continue;
                }

                List<int> imageless = Enumerable.Range(0, baselines.Count)
                    .Where(row => !taken[row] && Trimmed(baselines[row].File).Length == 0 && SamePin(steps[step], baselines[row]))
                    .ToList();

                if (imageless.Count == 1)
                {
                    rowOf[step] = imageless[0];
                    taken[imageless[0]] = true;
                }
            }

            return rowOf;
        }

        // ###########################################################################################
        // THE RUN'S LINES AS THEY STAND NOW, THE BASELINES READ AGAIN (code review, 2026-11-21: an "Undo" of an accepted capture
        // left its line "Accepted", and a baseline added, deleted or given another pin in the pane while the window was open had no
        // line, or one that captured into the wrong row). `steps` the lines as they were, `pins` the pins a run can go round now.
        //   - A capture accepted that the draft no longer has is no longer accepted: its line is as it was before it.
        //   - Every line finds its baseline (Match), and says its signal name as the baseline now does; a baseline no line holds - one
        //     added, or given another pin - gets a new line in its pin's place, in place of the pin's new line with nothing captured.
        //   - A line of a baseline gone goes, unless a capture of it waits to be accepted - that stays, to go in as a new baseline;
        //     a pin left with no line gets a new one, as a pin with no baseline has.
        //   - A new line is picked when a line of its pin was.
        // The lines kept are the very steps they were.
        // ###########################################################################################
        public static IReadOnlyList<BaselineRunLine> Follow(
            IReadOnlyList<BaselineRunStep> steps,
            IReadOnlyList<BaselineRunPin> pins,
            IReadOnlyList<ComponentImageEntry> baselines)
        {
            ArgumentNullException.ThrowIfNull(steps);
            ArgumentNullException.ThrowIfNull(pins);
            ArgumentNullException.ThrowIfNull(baselines);

            var files = baselines.Select(row => (row.File ?? string.Empty).Trim()).Where(file => file.Length > 0).ToHashSet(StringComparer.Ordinal);

            foreach (BaselineRunStep step in steps.Where(step => step.IsAccepted && !files.Contains(step.File)))
            {
                step.LostAcceptance();
            }

            // ... and one a "Redo" brought back is accepted again (code review, 2026-11-21: it stayed "Has a baseline", uncounted).
            foreach (BaselineRunStep step in steps.Where(step => step.LostFile.Length > 0 && files.Contains(step.LostFile)))
            {
                step.RegainAcceptance();
            }

            IReadOnlyList<int> rowOf = BaselineRunPins.Match(steps, baselines);
            var used = new HashSet<BaselineRunStep>(ReferenceEqualityComparer.Instance);
            var lines = new List<BaselineRunLine>();

            bool OfPin(BaselineRunStep step, BaselineRunPin pin) => ComponentVocabulary.IsSamePin(step.Pin.Pin, pin.Pin);

            BaselineRunLine Kept(int index)
            {
                used.Add(steps[index]);

                // Renamed in the pane: the line says so, and a capture of it is filed under the name it has.
                if (rowOf[index] >= 0 && (baselines[rowOf[index]].Name ?? string.Empty).Trim() is { Length: > 0 } name)
                {
                    steps[index].Name = name;
                }

                return new BaselineRunLine(steps[index], rowOf[index]);
            }

            BaselineRunLine Fresh(BaselineRunLine line, BaselineRunPin pin)
            {
                line.Step.IsPicked = steps.Any(step => step.IsPicked && OfPin(step, pin));
                return line;
            }

            foreach (BaselineRunPin pin in pins)
            {
                foreach (BaselineRunLine line in BaselineRunPins.Lines([pin], baselines))
                {
                    // A pin with no baseline: its new line, else a line of a baseline gone with a capture waiting.
                    int holder = line.Row >= 0
                        ? Enumerable.Range(0, steps.Count).FirstOrDefault(index => rowOf[index] == line.Row, -1)
                        : Enumerable.Range(0, steps.Count).FirstOrDefault(
                            index => rowOf[index] < 0 && (!steps[index].HadBaseline || steps[index].IsWaiting) && !used.Contains(steps[index]) && OfPin(steps[index], pin), -1);

                    lines.Add(holder >= 0 ? Kept(holder) : Fresh(line, pin));
                }

                // A capture waiting on a line of this pin no baseline holds now stays, after the pin's lines.
                for (int index = 0; index < steps.Count; index++)
                {
                    if (rowOf[index] < 0 && steps[index].IsWaiting && !used.Contains(steps[index]) && OfPin(steps[index], pin))
                    {
                        lines.Add(Kept(index));
                    }
                }
            }

            // A capture waiting on a pin the run no longer lists stays too, last.
            for (int index = 0; index < steps.Count; index++)
            {
                if (rowOf[index] < 0 && steps[index].IsWaiting && !used.Contains(steps[index]))
                {
                    lines.Add(Kept(index));
                }
            }

            return lines;
        }

        // ---- going round them ("Capture from oscilloscope pin-by-pin"'s window, owner requests 2026-11-20) ------------------------

        // ###########################################################################################
        // The step to go on to after `from` ("Accept and continue to next pin"): the first PICKED step not accepted yet after it,
        // round to the start - never `from` itself. -1 when every other picked step is accepted.
        // ###########################################################################################
        public static int NextToAccept(IReadOnlyList<BaselineRunStep> steps, int from)
        {
            ArgumentNullException.ThrowIfNull(steps);

            for (int offset = 1; offset <= steps.Count; offset++)
            {
                int index = BaselineRunPins.Wrapped(from + offset, steps.Count);

                if (index != from && steps[index].IsPicked && !steps[index].IsAccepted)
                {
                    return index;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // The picked step before (-1) or after (+1) `from`, round past either end - Left and Right in the window, as the popup goes
        // round its images. `from` itself when no other step is picked, -1 when none is.
        // ###########################################################################################
        public static int Neighbour(IReadOnlyList<BaselineRunStep> steps, int from, int direction)
        {
            ArgumentNullException.ThrowIfNull(steps);

            if (direction == 0 || steps.Count == 0)
            {
                return from;
            }

            int by = direction > 0 ? 1 : -1;

            for (int offset = 1; offset <= steps.Count; offset++)
            {
                int index = BaselineRunPins.Wrapped(from + (offset * by), steps.Count);

                if (steps[index].IsPicked)
                {
                    return index;
                }
            }

            return -1;
        }

        // The first step of the pin `typed` names - the number keys, as the popup's go to a pin's first image ("2" is "2 [64K]"
        // too). -1 for none.
        public static int IndexOfPin(IReadOnlyList<BaselineRunStep> steps, string? typed)
        {
            ArgumentNullException.ThrowIfNull(steps);

            for (int index = 0; index < steps.Count; index++)
            {
                if (ComponentVocabulary.IsSamePin(steps[index].PinText, typed) || ComponentVocabulary.IsSamePin(steps[index].Pin.Pin, typed))
                {
                    return index;
                }
            }

            return -1;
        }

        private static int Wrapped(int index, int count) => ((index % count) + count) % count;
    }

    // ###########################################################################################
    // A pin a run can go round: as written ("5", or "12V" for one named rather than numbered), its number, its name and
    // direction (BaselineRunPins.PinOf), and how many baselines it has.
    // ###########################################################################################
    public sealed record BaselineRunPin(string Pin, int? Number, string Name, string Direction, int Baselines)
    {
        public bool HasBaseline => this.Baselines > 0;
    }

    // ###########################################################################################
    // One step of a run: the pin, the baseline's own pin text and signal name, and the file of the baseline it captures again -
    // "" for a new one. The window notes each capture of it as the run goes (the step is its own, for the run).
    //
    // CAPTURED, THEN ACCEPTED - AND ONLY THEN IN THE DRAFT (owner requests, 2026-11-20: "When I have captured an image ... then it
    // must stay on the pin, so I can check that it is good ... "Accept and continue to next pin". The "Previous pin" should not
    // "unmark" it as accepted", then "New capture - not yet accepted", "so it is clear this must be accepted before it will be
    // used"): a capture waits to be checked, and accepting it puts it into the draft under a file of its own - so a capture taken
    // again waits again, and only a capture can be accepted. Going back and forth changes neither.
    // ###########################################################################################
    public sealed class BaselineRunStep(BaselineRunPin pin, string pinText, string name, string file, bool hadBaseline = false)
    {
        public BaselineRunPin Pin { get; } = pin;

        public string PinText { get; } = pinText;

        public string Name { get; set; } = name;

        // The baseline's file as the draft has it - "" for a new one until a capture of it is accepted.
        public string File { get; private set; } = file;

        // The baseline's file as the run started - what a capture is compared with. "" for a new baseline, or one with no image.
        public string OriginalFile { get; } = file;

        // Whether it captures again a baseline the component had with a picture when the run started - shown beside it.
        public bool IsUpdate => this.OriginalFile.Length > 0;

        // Whether it is of a baseline the component had when the run started - one with no image yet too, whose capture goes into it.
        public bool HadBaseline { get; } = hadBaseline || file.Length > 0;

        // Captured in this run - waiting to be accepted, or accepted.
        public bool IsCaptured { get; private set; }

        // Its capture checked, accepted and in the draft.
        public bool IsAccepted { get; private set; }

        // A capture not yet accepted - not in the draft.
        public bool IsWaiting => this.IsCaptured && !this.IsAccepted;

        // Picked to go round - the window's checkbox, its quick picks and the case's pins set it.
        public bool IsPicked { get; set; }

        // The file of an accepted capture an "Undo" took out of the draft - "" while there is none. A "Redo" that brings it back
        // accepts it again (RegainAcceptance); a capture taken since lets go of it.
        public string LostFile { get; private set; } = string.Empty;

        // A capture taken: waiting to be accepted - again, when one was before.
        public void Captured()
        {
            this.IsCaptured = true;
            this.IsAccepted = false;
            this.LostFile = string.Empty;
        }

        // Its capture accepted, in the draft as `acceptedFile` - false, and nothing changed, when it has none in this run.
        public bool Accepted(string acceptedFile)
        {
            ArgumentNullException.ThrowIfNull(acceptedFile);

            if (!this.IsCaptured)
            {
                return false;
            }

            this.File = acceptedFile.Trim();
            this.IsAccepted = true;
            return true;
        }

        // ###########################################################################################
        // Its accepted capture no longer in the draft - an "Undo" took it (code review, 2026-11-21): as it was before it, its baseline
        // the one the run started from, nothing captured. False, and nothing changed, when it had none accepted.
        // ###########################################################################################
        public bool LostAcceptance()
        {
            if (!this.IsAccepted)
            {
                return false;
            }

            this.LostFile = this.File;
            this.File = this.OriginalFile;
            this.IsAccepted = false;
            this.IsCaptured = false;
            return true;
        }

        // Its accepted capture back in the draft - a "Redo" of the step an "Undo" took: accepted again. False, and nothing changed,
        // when it lost none, or a capture was taken since.
        public bool RegainAcceptance()
        {
            if (this.LostFile.Length == 0 || this.IsCaptured)
            {
                return false;
            }

            this.File = this.LostFile;
            this.LostFile = string.Empty;
            this.IsCaptured = true;
            this.IsAccepted = true;
            return true;
        }
    }

    // A line of a run: its step, and the baseline it is of - its index in the component's baselines, -1 for a new one.
    public sealed record BaselineRunLine(BaselineRunStep Step, int Row);
}
