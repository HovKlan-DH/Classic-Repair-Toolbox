using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // The board's own lists the component pane shows (slice 4a) - one at a time, in place of a component.
    // Files and Links were one part ("Files and links") until 2026-10-29 (owner request: one entry each
    // in the Contribute tab's left column).
    internal enum PaneBoardPart
    {
        Files,
        Links,
        Credits,
        ImportantSignals,

        // The schematic images and their colours (slice 4b) - "Pictures" on screen until 2026-10-29.
        Pictures,

        // The regions the board comes in, and its component categories (2026-11-01).
        Regions,
        Categories,
    }

    // ###########################################################################################
    // THE PAGES OF A COMPONENT, its tabs over the pane (owner decision, 2026-11-01: "the functionality in
    // this is too massive, if all data is filled, so I think we need to split this into multiple menu
    // items" - as tabs on the right, so the left column keeps its room for the component list). In the
    // order the tabs stand.
    // ###########################################################################################
    internal enum ComponentPage
    {
        Basics,
        Replacements,
        Pinout,
        Pictures,
        Baselines,
        Files,
        Links,
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
            $"[{label}] is not on the board in your draft any more. If that was not meant, \"Undo\" in the draft's header brings it back.";

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

        // Under a new component's empty category - every component needs one (owner decision, 2026-11-01).
        public const string NoCategoryYet =
            "Every component needs a category - pick one of the board's, or type a new one.";

        // ---- a component's label (2026-11-01) --------------------------------------------------------

        // Why a label typed for the component open cannot be taken - null when it can.
        public static string? LabelProblem(string? typed, IEnumerable<string?> otherLabels)
        {
            string wanted = typed?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return "A component needs a label - the one printed on the board.";
            }

            return otherLabels.Any(other => string.Equals(other?.Trim(), wanted, StringComparison.OrdinalIgnoreCase))
                ? $"This board already has [{wanted}] - a label names one component."
                : null;
        }

        // ---- regions (2026-11-01) ---------------------------------------------------------------------

        // What a row of a board with regions says it applies to - the words of its drop-down list.
        public const string AllRegions = "All regions";

        // The choices, and a word the board already holds that is not one of them.
        public static IReadOnlyList<string> AppliesToChoices(string? region)
        {
            string own = region?.Trim() ?? string.Empty;

            return own.Length == 0 || BoardRegions.All.Contains(own, StringComparer.OrdinalIgnoreCase)
                ? [AllRegions, .. BoardRegions.All]
                : [AllRegions, .. BoardRegions.All, own];
        }

        public static string AppliesToOf(string? region)
        {
            string own = region?.Trim() ?? string.Empty;

            return own.Length == 0 ? AllRegions : BoardRegions.All.FirstOrDefault(each => string.Equals(each, own, StringComparison.OrdinalIgnoreCase)) ?? own;
        }

        public static string RegionOfAppliesTo(string? choice) =>
            string.Equals(choice, AllRegions, StringComparison.Ordinal) ? string.Empty : choice?.Trim() ?? string.Empty;

        // A row changed to the other region leaves what this one shows - said, or it reads as gone.
        public static string MovedToRegion(string region) =>
            $"Moved to {region.Trim()} - it no longer shows here, as this is the other region. Pick {region.Trim()} above to see it.";

        // A component or its pins made the same in every region: what went, and how to get it back.
        public static string MadeTheSameInEveryRegion(string what, string region) =>
            $"{what} now the same in every region, as {region.Trim()} had it. \"Undo\" brings the other region's back.";

        // "Different per region" ticked on a component with no pins: nothing to tell apart yet, so no step.
        public const string NoPinsToTellApart = "It has no pins yet, so there is nothing to tell apart per region - add its pins first.";

        // ... taken off while the region shown has none: its pins would be every region's, so all would go.
        public static string NoPinsInRegion(string region) =>
            $"{region.Trim()} has no pins of its own to make the same in every region. Pick the region whose pins it should be, and take \"Different per region\" off there.";

        public const string RegionsHasOnBoard =
            "This board comes in PAL and NTSC versions";

        // Why the tick cannot be taken away while rows say a region.
        public static string RegionsInUse(int rows) => rows == 1
            ? "1 row of this board says PAL or NTSC, so this stays ticked until it applies to all regions."
            : $"{rows.ToString(CultureInfo.InvariantCulture)} rows of this board say PAL or NTSC, so this stays ticked until they apply to all regions.";

        // ---- the board's categories (2026-11-01) ------------------------------------------------------

        public static string CategoryCount(int components) =>
            components == 1 ? "1 component" : $"{components.ToString(CultureInfo.InvariantCulture)} components";

        public const string CategoryUnused = "No component yet";

        public const string CategoryRowsNeedOne =
            "These components have no category. Type one here - or the name of one the board has - and they all go under it.";

        public const string CategoryNameEmpty = "A category needs a name.";

        public const string CategoryDeleteNeedsAnother =
            "There is no other category to move its components to - add one first.";

        public static string CategoryDeleteAsk(int components) =>
            components == 1
                ? "Its component needs a category - which one should it move to?"
                : $"Its {components.ToString(CultureInfo.InvariantCulture)} components need a category - which one should they move to?";

        public static string CategoryAlreadyThere(string name) => $"This board already has the category [{name.Trim()}].";

        // ---- TEMPORARY: the tool that makes pinouts of scope baselines ---------------------------------
        // *** REMOVED BEFORE CRT 4.0 GOES LIVE - see BaselinePinMigration. ***

        public const string MigrationNote =
            "Temporary, until CRT 4.0 goes live: make the pins from the scope baselines - each pin's number and signal name, and the package when that is certain. What you typed stays, and \"Undo\" takes it back.";

        public const string MigrationNothing = "The pins already have everything the scope baselines say.";

        public static string MigratedBoard(int components) =>
            components == 1
                ? "Made the pins of 1 component from its scope baselines. \"Undo\" takes it back."
                : $"Made the pins of {components.ToString(CultureInfo.InvariantCulture)} components from their scope baselines. \"Undo\" takes them all back.";

        public static string ShowOnBoardTip(string label, bool marked) =>
            marked
                ? $"Show [{label}] on the schematics - the \"Schematics\" tab, with only this component selected"
                : $"[{label}] is not marked on any picture yet, so there is nothing to show";

        public static string DeleteTip(string label) =>
            $"Take [{label}] off the board in your draft - every row of it in every region, with its pictures, scope baselines, pins and files. \"Undo\" brings it back.";

        // Offered while the step that deleted it is the latest one.
        public static string BringBack(string label) => $"Bring [{label}] back";

        // ---- the board's own lists (slice 4a) -------------------------------------------------------

        // A board part's headline names its group, as the left column's tree does (owner request,
        // 2026-11-01: "Board" and "KiCad" over the entries) - a component has its own "Files" and "Links".
        public static string Title(PaneBoardPart part) => part switch
        {
            PaneBoardPart.Files => "Board: Files",
            PaneBoardPart.Links => "Board: Links",
            PaneBoardPart.Credits => "Board: Credits",
            PaneBoardPart.Pictures => SchematicImages,
            PaneBoardPart.Regions => "Board: Regions",
            PaneBoardPart.Categories => "Board: Categories",
            _ => ImportantSignals,
        };

        // A component's page as its tab says it, with how many it holds when it holds any.
        public static string PageTitle(ComponentPage page, int count)
        {
            string name = page switch
            {
                ComponentPage.Basics => "Basics",
                ComponentPage.Replacements => "Replacement parts",
                ComponentPage.Pinout => "Pinout",
                ComponentPage.Pictures => "Pictures",
                ComponentPage.Baselines => "Scope baseline",
                ComponentPage.Files => "Files",
                _ => "Links",
            };

            return count > 0 ? ComponentPaneWording.Heading(name, count) : name;
        }

        // A tab marked because of what is on its page.
        public const string PageHasErrorTip = "Something on this page stops the draft being submitted";

        // What the board's pictures are called wherever the part is named (owner requests, 2026-10-29:
        // "rename 'Pictures' to 'Schematic images'", and 2026-11-01: "'Board: Schematic images'").
        public const string SchematicImages = "Board: Schematic images";

        // And its important signals (owner requests, 2026-11-01: "rename 'Important KiCad signals' to
        // 'KiCad: Important signals'") - each lights up a net of the board's KiCad data. The undo steps
        // keep the plain words (DraftStepWording.ImportantSignals), the sheet its stored name.
        public const string ImportantSignals = "KiCad: Important signals";

        // Under the KiCad views "KiCad: Files" lists - where a picture is matched to one, the part named
        // as the left column names it (code review, 2026-10-30: it still said "Pictures").
        public const string KiCadViewsHint =
            "Each of these can be matched to one picture of the board. Pick it under \"Shows KiCad view\" on the picture's card - \"" +
            SchematicImages + "\" on the left of the \"Contribute\" tab.";

        // Under an important signal whose net the Schematics tab's panel will not find.
        public static string NetNotFound(string net) =>
            $"[{net.Trim()}] is no net in the board's KiCad data - the \"Schematics\" tab will not light this signal up.";

        // A signal's net is picked from the board's KiCad data, never typed (owner request, 2026-10-29).
        public const string NoKiCadNets =
            "The board has no KiCad data yet, so there is no net to pick from - add the board's KiCad project under \"KiCad: Files\" first.";

        public const string ReadingKiCadNets = "Reading the board's KiCad nets...";

        // "Add signal" beside the net picker: what it does, or why it is greyed out.
        public const string AddSignalTip = "Add an important signal on this KiCad net - then give it the name people know it by";

        // Under a link's address the checks refuse (BoardDataChecks.IsRefusedLink).
        public const string LinkNotWeb = "Write the whole address, starting with http:// or https://";

        public const string PickANetTip ="Pick one of the board's KiCad nets that is not an important signal yet";

        public const string AddSignalNeedsKiCadTip = "Needs the board's KiCad data - there is no net to pick from yet";

        // The pane's title over a board part - its name and how many it holds, the one headline of the
        // part (owner request, 2026-10-29: "two headlines for the same").
        public static string PartTitle(PaneBoardPart part, int count) => ComponentPaneWording.Heading(ComponentPaneWording.Title(part), count);

        // "Open" on a board file or link that could not be opened.
        public static string FileNotOpened(string? name) =>
            $"[{name?.Trim()}] could not be opened - CRT opens documents, pictures and text files only.";

        public static string LinkNotOpened(string? name) =>
            $"[{name?.Trim()}] has no web address CRT can open.";

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
                : $"How this board's baselines were measured is not written down yet - add a file to the board under \"Board: Files\", in the category \"{ScopeMeasurementFiles.Category}\".";

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
            "Your draft changes nothing of this part that could go to another board yet - its package, notes, pins, replacement parts, links, files or pictures.";

        // A file replaced with nothing on the other boards to change: no part number to find the part
        // by, or none that has it, and no other board with the file it replaced.
        public static string AlsoFoundNothing(string partNumber) =>
            partNumber.Trim().Length == 0
                ? "No other board on this machine has the files you replaced."
                : $"No other board on this machine has a component with the part number [{partNumber.Trim()}], or the files you replaced.";

        // ###########################################################################################
        // Above the files the draft replaced (2026-10-26 - the plan: "This datasheet is also used by 11
        // other boards - change it there too?"): the same bytes as published here, on the boards listed
        // under each - and what happens to them.
        // ###########################################################################################
        public const string AlsoFilesIntro =
            "The files you replaced are on other boards too - the same file as it is published here. Each ticked board gets your new file in its place, in its own draft: every entry there showing the old file shows the new one, and that board's \"Undo\" takes it back. A board whose draft has a version of its own keeps it.";

        // The board's own files: the same, offered under Files.
        public const string AlsoBoardFilesIntro =
            "The board files you replaced are on other boards too - the same file as it is published here. Each ticked board gets your new file in its place, in its own draft: every entry there showing the old file shows the new one, and that board's \"Undo\" takes it back. A board whose draft has a version of its own keeps it.";

        public const string AlsoBoardFilesNothingChanged =
            "Your draft replaces none of the board's files yet - pick another file for one with Change file first.";

        public const string AlsoBoardFilesFoundNothing =
            "No other board on this machine has the board files you replaced.";

        // ###########################################################################################
        // What came of it: the boards changed, and those that were not, with why - "Changed on 2 boards,
        // each in its own draft: A, B. Not changed on C - its draft could not be saved." - then each
        // component not changed on a board where something else was ("[U7] not changed on B - ..."; code
        // review, 2026-10-26), then each file not put on a board it was ticked for: "[Datasheet
        // (6510.pdf)] not put on D - ...".
        // ###########################################################################################
        public static string AlsoDone(
            IReadOnlyList<string> changed,
            IReadOnlyList<(string Board, string Why)> notChanged,
            IReadOnlyList<(string File, string Board, string Why)>? filesNotPut = null,
            IReadOnlyList<(string Label, string Board, string Why)>? componentsNotChanged = null)
        {
            var parts = new List<string>();

            if (changed.Count > 0)
            {
                parts.Add(changed.Count == 1
                    ? $"Changed on 1 board, in its own draft: {changed[0]}."
                    : $"Changed on {changed.Count.ToString(CultureInfo.InvariantCulture)} boards, each in its own draft: {string.Join(", ", changed)}.");
            }

            parts.AddRange(notChanged.Select(board => $"Not changed on {board.Board} - {board.Why}."));
            parts.AddRange((componentsNotChanged ?? []).Select(component => $"[{component.Label.Trim()}] not changed on {component.Board} - {component.Why}."));
            parts.AddRange((filesNotPut ?? []).Select(file => $"[{file.File}] not put on {file.Board} - {file.Why}."));

            return string.Join(" ", parts);
        }

        public const string AlsoAlreadySo = "it is like that there already";

        public const string AlsoNoDraft = "no draft could be made of it";

        public const string AlsoNotSaved = "its draft could not be saved";

        // Why a file was not put on a board (ChangedFiles.CanPut, and the copy itself).
        public const string AlsoFileNoLongerThere = "that board's draft no longer has it";

        public const string AlsoFileOwnVersion = "that board's draft has a version of its own";

        public const string AlsoFileNotCopied = "it could not be copied into that board's draft";

        public const string AlsoFileNotFound = "your new file is not in your draft";

        // A file of the board that is in neither the draft nor the data on this machine.
        public static string FileNotFound(string? name) =>
            $"[{name?.Trim()}] is not on this machine yet - it comes with the next data update.";

        public const string CaptureFailed =
            "Nothing came from the oscilloscope - the \"Oscilloscope\" tab says why.";

        public const string NotAnImage =
            "That file is no picture CRT can show - pick a PNG, JPEG, GIF, BMP or WebP.";

        public const string FileNotCopied =
            "The file could not be copied into your draft - see the log for why.";

        // ---- the board's schematic images (step 3 - the "Schematic images" window, moved inline) ----

        // `notCopied`: images picked that could not be copied in - said, for they got no card.
        public static string PicturesAdded(int count, int notCopied = 0)
        {
            string added = count == 1 ? "Added 1 schematic image." : $"Added {count.ToString(CultureInfo.InvariantCulture)} schematic images.";

            return notCopied switch
            {
                <= 0 => added,
                1 => $"{added} 1 could not be copied - it is open in another program, or no longer where it was picked.",
                _ => $"{added} {notCopied.ToString(CultureInfo.InvariantCulture)} could not be copied - they are open in another program, or no longer where they were picked.",
            };
        }

        public const string NoPicturesToAdd =
            "Nothing was added - pick PNG, JPEG, GIF, BMP or WebP files.";

        // "Delete" since 2026-10-29 (owner question: "will that actually not delete the schematic
        // image?"): the image leaves the board, and its file the draft - the contributor's own copy,
        // where it was picked from, is never touched.
        public static string PictureRemoved(string name) =>
            $"Deleted [{name.Trim()}] from the board. \"Undo\" at the top of the tab brings it back.";

        // ---- the previews on a schematic image's card (2026-10-29) ----------------------------------

        public const string ThisSideTracesTip =
            "Draw every trace on this side of the board on the preview, in the color of the traces on this side - on the \"Schematics\" tab only the traces of the net lit up are drawn so";

        public const string OtherSideTracesTip =
            "Draw every trace on the board's other side on the preview, in the color of the traces on the other side - on the \"Schematics\" tab only the traces of the net lit up are drawn so";

        // Why the traces cannot be shown - the image names no KiCad view, or is not aligned with it.
        public const string ShowTracesNeedsAlignTip =
            "Needs KiCad data, a KiCad view of a side of the board picked for this image, and KiCad aligned on it (\"Align KiCad data\")";

        // The KiCad options of an image with no KiCad data (owner request, 2026-10-29: "if the
        // schematic image has no KiCad data, then the KiCad (traces) options should be disabled").
        public const string TraceColourTip = "With KiCad data: the color of a lit trace on the board's other side";

        // Each side's traces have a colour and an opacity of their own (owner request, 2026-10-30).
        public const string ThisSideTraceColourTip =
            "With KiCad data: the color of a lit trace on the image's own side - the Schematic image color until one is picked";

        public const string ThisSideTraceOpacityTip = "With KiCad data: how strongly a lit trace on the image's own side shows";

        public const string OtherSideTraceOpacityTip = "With KiCad data: how strongly a lit trace on the board's other side shows";

        public const string TraceColourNeedsKiCadTip =
            "Only for an image with KiCad data - the board's KiCad project, and a KiCad view picked for this image";

        public const string ReadingTraces = "Reading the board's KiCad traces...";

        // ###########################################################################################
        // Under the ticks, once the traces are read: a side ticked that has no copper says so - "" when
        // every side ticked has some (or none is ticked).
        // ###########################################################################################
        public static string TracesNote(bool thisSide, int thisSideCount, bool otherSide, int otherSideCount)
        {
            bool thisEmpty = thisSide && thisSideCount == 0;
            bool otherEmpty = otherSide && otherSideCount == 0;

            return (thisEmpty, otherEmpty) switch
            {
                (true, true) => "The board's KiCad data has no trace on either side of this image.",
                (true, false) => "The board's KiCad data has no trace on this side of the image.",
                (false, true) => "The board's KiCad data has no trace on the other side of this image.",
                _ => string.Empty,
            };
        }

        // ---- important signals suggested ----------------------------------------------------------

        public static string SignalSuggestionsHeading(int count) =>
            $"Suggested from the KiCad nets ({count.ToString(CultureInfo.InvariantCulture)})";

        // The name is a box since 2026-10-30 (owner request) - said here, so it reads as one to change.
        public const string SignalSuggestionsIntro =
            "Nets of this board whose names say what they are - a supply, ground, a clock, reset, a bus line. Add the ones worth checking first - change a name before adding it if you want another; \"Show on board\" lights a net up on the \"Schematics\" tab, so you can see which it is.";

        // The board's other nets, under the suggestions (owner request, 2026-11-01).
        public static string OtherNetsHeading(int count) =>
            $"Other KiCad nets ({count.ToString(CultureInfo.InvariantCulture)})";

        public const string OtherNetsIntro =
            "Every other net of this board, each with a proposed name you can change. \"Add\" makes it an important KiCad signal under that name; \"Show on board\" lights it up on the \"Schematics\" tab.";

        // What a new component needs next (owner request, 2026-11-01).
        public const string NewComponentMarkNote =
            "After adding it, mark it by hand on at least one schematic image: press \"Mark components manually\" on the image under \"Board: Schematic images\". A component marked on no image is flagged as an orphan, as CRT cannot point it out on the board.";

        // Under a picture's thumbnail, once it shows a KiCad view (owner requests, 2026-11-01).
        public static string PictureComponentsAvailable(int count) =>
            $"- {ComponentPaneWording.Components(count)} available from KiCad data";

        public static string PictureComponentsAutomaticallyMarked(int count) =>
            $"- {ComponentPaneWording.Components(count)} automatically marked from KiCad data";

        public static string PictureComponentsManuallyMarked(int count) =>
            $"- {ComponentPaneWording.Components(count)} manually marked";

        // After the number, which is bold.
        public static string PictureComponentsMarkedWords(int count) => count == 1 ? " component marked" : " components marked";

        private static string Components(int count) =>
            count == 1 ? "1 component" : $"{count.ToString(CultureInfo.InvariantCulture)} components";

        // ---- fill from KiCad ----------------------------------------------------------------------

        public static string FillFromKiCadTip(string label) =>
            $"Add every pin the board's KiCad data has for [{label.Trim()}], with the name and direction its symbol gives it, and fill in the names and directions still empty - nothing you typed is changed. \"Undo\" takes it back.";

        public const string NothingToFillFromKiCad =
            "The pins already have everything the board's KiCad data has for this component.";

        // What "Fill from KiCad" did: "Filled from KiCad: 38 pins added, 2 named."
        public static string FilledFromKiCad(int added, int filled)
        {
            var parts = new List<string>();

            if (added > 0)
                parts.Add(added == 1 ? "1 pin added" : $"{added.ToString(CultureInfo.InvariantCulture)} pins added");

            if (filled > 0)
                parts.Add(filled == 1 ? "1 pin filled in" : $"{filled.ToString(CultureInfo.InvariantCulture)} pins filled in");

            return $"Filled from KiCad: {string.Join(", ", parts)}.";
        }

        // Above the nets: what they are - read from the board's KiCad data, never stored.
        public const string NetHeadingTip =
            "The net each pin is on, on this board - read from the board's KiCad data as it is shown, not stored with the pins.";

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
