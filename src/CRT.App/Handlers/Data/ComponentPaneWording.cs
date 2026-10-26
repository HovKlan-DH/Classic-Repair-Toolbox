using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // The board's own lists the component pane shows (slice 4a) - one at a time, in place of a component.
    internal enum PaneBoardPart
    {
        FilesAndLinks,
        Credits,
        ImportantSignals,

        // The pictures and their colours (slice 4b).
        Pictures,
    }

    // ###########################################################################################
    // WHAT THE COMPONENT PANE SAYS (CRT 4.0 step 2, 2026-10-25) - its title, the region line, the
    // offer to add the case's pins, its sections' headings and what a scope baseline still needs. Words only, so
    // they are tested without a display.
    // ###########################################################################################
    internal static class ComponentPaneWording
    {
        // "U8 - VIC-II | 6569" - the label, then what the component list shows beside it.
        public static string Title(ComponentEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            string label = row.BoardLabel?.Trim() ?? string.Empty;
            string[] names = new[] { row.FriendlyName, row.TechnicalNameOrValue }
                .Select(name => name?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .ToArray();

            return names.Length == 0 ? label : $"{label} - {string.Join(" | ", names)}";
        }

        // Said only for a component with a row of its own per region - which one is open.
        public static string RegionLine(string? rowRegion)
        {
            string region = rowRegion?.Trim() ?? string.Empty;

            return region.Length == 0
                ? string.Empty
                : $"The {region} version - each region has a row of its own, and this is the one CRT shows now.";
        }

        public static string NotOnBoard(string label) =>
            $"[{label}] is not on the board in your draft any more. If that was not meant, Undo in the draft's header brings it back.";

        // ---- the component as a whole (slice 3c) -----------------------------------------------------

        public const string NewComponentTitle = "New component";

        // ###########################################################################################
        // The regions a new component can be in, as the drop-down list says them - first the one a
        // component is in most often - and the region each stands for (blank: every region, the
        // component list's rule).
        // ###########################################################################################
        public static IReadOnlyList<string> RegionChoices { get; } = ["PAL and NTSC", "PAL only", "NTSC only"];

        public static string RegionOfChoice(string? choice) => choice switch
        {
            "PAL only" => "PAL",
            "NTSC only" => "NTSC",
            _ => string.Empty,
        };

        // Why a typed label cannot be added - null when it can, or while nothing is typed yet.
        public static string? NewLabelProblem(ContributionPackaging.NewComponentProblem problem, string? label) =>
            problem == ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists
                ? $"This board already has [{label?.Trim()}] - pick it in the list on the left to change it."
                : null;

        // Under a new component's empty category - it is added all the same.
        public const string NoCategoryYet =
            "Without a category it is not in the component list beside the schematics - only here. It can be given one later.";

        public static string ShowOnBoardTip(string label, bool marked) =>
            marked
                ? $"Show [{label}] on the schematics - the Schematics tab, with only this component selected"
                : $"[{label}] is not marked on any picture yet, so there is nothing to show";

        public static string DeleteTip(string label) =>
            $"Take [{label}] off the board in your draft - every row of it in every region, with its pictures, scope baselines, pins and files. Undo brings it back.";

        // Offered while the step that deleted it is the latest one.
        public static string BringBack(string label) => $"Bring [{label}] back";

        // ---- the board's own lists (slice 4a) -------------------------------------------------------

        public static string Title(PaneBoardPart part) => part switch
        {
            PaneBoardPart.FilesAndLinks => "Files and links",
            PaneBoardPart.Credits => "Credits",
            PaneBoardPart.Pictures => "Pictures",
            _ => "Important signals",
        };

        // Under an important signal whose net the Schematics tab's panel will not find.
        public static string NetNotFound(string net) =>
            $"[{net.Trim()}] is no net in the board's KiCad data - the Schematics tab will not light this signal up.";

        public const string NoKiCadNets =
            "The board has no KiCad data yet, so there is no net to pick from. A net can still be typed.";

        public const string ReadingKiCadNets = "Reading the board's KiCad nets...";

        // What the "Changed" mark says of a component, and of a board part.
        public const string ComponentChangedTip =
            "Your draft changes this component - the published board reads otherwise until your draft is submitted and published";

        public const string BoardPartChangedTip =
            "Your draft changes this part of the board - the published board reads otherwise until your draft is submitted and published";

        // A section's heading with how many it holds - "Scope baselines (3)", as the plan draws it.
        public static string Heading(string name, int count) =>
            $"{name} ({count.ToString(CultureInfo.InvariantCulture)})";

        // ###########################################################################################
        // What a scope baseline still needs, or null when it has it all: its pin and one scope setting
        // - the one rule that makes a row a baseline and not a picture (ComponentImageKinds) - and a
        // screenshot (the plan: "a screenshot, which is required").
        // ###########################################################################################
        public static string? BaselineProblem(string? pin, string? timeDiv, string? voltsDiv, string? triggerLevel, string? file)
        {
            var missing = new List<string>(3);

            if (string.IsNullOrWhiteSpace(pin))
                missing.Add("its pin");

            if (string.IsNullOrWhiteSpace(file))
                missing.Add("a screenshot");

            if (string.IsNullOrWhiteSpace(timeDiv) && string.IsNullOrWhiteSpace(voltsDiv) && string.IsNullOrWhiteSpace(triggerLevel))
                missing.Add("one of T/DIV, V/DIV or trigger level");

            if (missing.Count == 0)
                return null;

            string listed = missing.Count == 1
                ? missing[0]
                : $"{string.Join(", ", missing.Take(missing.Count - 1))} and {missing[^1]}";

            return $"A scope baseline needs {listed}.";
        }

        // Under a baseline whose edit is held back: what it needs, and that the draft keeps it as it was.
        public static string BaselineHeldBack(string problem) => $"{problem} Until it has, your draft keeps it as it was.";

        // ###########################################################################################
        // Above a component's scope baselines: how the board's were measured is in its "Scope
        // measurements" files - or not written down yet, and where it goes (owner decision, 2026-10-26).
        // ###########################################################################################
        public static string MeasuringNotesLine(bool hasNotes) =>
            hasNotes
                ? "How this board's baselines were measured:"
                : $"How this board's baselines were measured is not written down yet - add a file to the board under Files and links, in the category \"{ScopeMeasurementFiles.Category}\".";

        // ---- copy from another board (slice 5) --------------------------------------------------------

        public const string CopyNeedsPartNumber =
            "Give the component its part number first - the same part is found on the other boards by it.";

        public static string CopyFoundNothing(string partNumber) =>
            $"No other board on this machine has a component with the part number [{partNumber.Trim()}].";

        public static string CopyIntro(string partNumber) =>
            $"The same part - part number [{partNumber.Trim()}] - on the other boards. Pick one, and what to take from it; what this component has of it is replaced.";

        // ---- also on other boards (slice 5) -----------------------------------------------------------

        public static string AlsoIntro(string partNumber) =>
            $"The same part - part number [{partNumber.Trim()}] - is on these boards. What you changed here goes into each ticked board's own draft, as one step there - each is reviewed by that board's maintainers. What the part has there of what is ticked is replaced.";

        public const string AlsoNothingChanged =
            "Your draft changes nothing of this part that could go to another board yet - its package, notes, pins, replacement parts or links.";

        // ###########################################################################################
        // What came of it: the boards changed, and those that were not, with why - "Changed on 2 boards,
        // each in its own draft: A, B. Not changed on C - its draft could not be saved."
        // ###########################################################################################
        public static string AlsoDone(IReadOnlyList<string> changed, IReadOnlyList<(string Board, string Why)> notChanged)
        {
            var parts = new List<string>();

            if (changed.Count > 0)
            {
                parts.Add(changed.Count == 1
                    ? $"Changed on 1 board, in its own draft: {changed[0]}."
                    : $"Changed on {changed.Count.ToString(CultureInfo.InvariantCulture)} boards, each in its own draft: {string.Join(", ", changed)}.");
            }

            parts.AddRange(notChanged.Select(board => $"Not changed on {board.Board} - {board.Why}."));

            return string.Join(" ", parts);
        }

        public const string AlsoAlreadySo = "it is like that there already";

        public const string AlsoNoDraft = "no draft could be made of it";

        public const string AlsoNotSaved = "its draft could not be saved";

        // A file of the board that is in neither the draft nor the data on this machine.
        public static string FileNotFound(string? name) =>
            $"[{name?.Trim()}] is not on this machine yet - it comes with the next data update.";

        public const string CaptureFailed =
            "Nothing came from the oscilloscope - the Oscilloscope tab says why.";

        public const string NotAnImage =
            "That file is no picture CRT can show - pick a PNG, JPEG, GIF, BMP or WebP.";

        public const string FileNotCopied =
            "The file could not be copied into your draft - see the log for why.";

        // How many of a case's pin numbers the pins do not list yet - 0 for a case whose count CRT
        // cannot tell.
        public static int MissingCasePins(string? package, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            if (ComponentVocabulary.PinCountOf(package) is not int count)
            {
                return 0;
            }

            var listed = new HashSet<string>(pins.Select(pin => pin.Pin?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);

            return Enumerable.Range(1, count).Count(number => !listed.Contains(number.ToString(CultureInfo.InvariantCulture)));
        }

        // The button that adds them: every pin of the case while none is listed, else the ones missing.
        public static string AddCasePins(string? package, int missing, bool noPinsYet)
        {
            string name = package?.Trim() ?? string.Empty;

            return noPinsYet && ComponentVocabulary.PinCountOf(name) is int count
                ? $"Add the {count.ToString(CultureInfo.InvariantCulture)} pins of {name}"
                : missing == 1
                    ? "Add the missing pin"
                    : $"Add the {missing.ToString(CultureInfo.InvariantCulture)} missing pins";
        }
    }
}
