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
    // THE PAGES OF A COMPONENT, under "Components" in the left column (owner decision, 2026-11-01: "the
    // functionality in this is too massive, if all data is filled, so I think we need to split this into
    // multiple menu items"; tabs over the pane until 2026-11-03). In the order their entries stand.
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
    // offer to add the case's pins, its sections' headings and what an oscilloscope baseline still needs. Words only, so
    // they are tested without a display.
    // ###########################################################################################
    internal static class ComponentPaneWording
    {
        // "Component: U8 | VIC-II | 6569" - what it is, as a board part's headline says "Board", then exactly
        // the component popup's title by its one rule (owner requests, 2026-11-04).
        public static string Title(ComponentEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return ComponentPaneWording.Title(ComponentPopupContent.TitleOf(row.BoardLabel, row));
        }

        public static string Title(string name) => $"Component: {name}";

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

        // Why a typed label cannot be added - null when it can, or while nothing is typed yet.
        public static string? NewLabelProblem(ContributionPackaging.NewComponentProblem problem, string? label) =>
            problem == ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists
                ? $"This board already has [{label?.Trim()}] - pick it in the list on the left to change it."
                : null;

        // Under a new component's empty category - every component needs one (owner decision, 2026-11-01).
        public const string NoCategoryYet =
            "Every component needs a category - pick one of the board's. A new one is added under \"Board: Categories\".";

        // ... and with none to pick from: the one place a category is made (owner request, 2026-11-03).
        public const string NoCategoriesToPick =
            "Every component needs a category, and this board has none yet - add one under \"Board: Categories\" first.";

        // Which region a page shows, now that the buttons switching it are CRT's own (owner request,
        // 2026-11-03: "use the regions from the left-side instead of adding new region buttons").
        public static string RegionShownLine(string? region) =>
            $"Showing {region?.Trim()} - the \"PAL\" and \"NTSC\" buttons under the component list switch it.";

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

        // Under the label of a component told apart per region: what the other boxes are not.
        public const string LabelInEveryRegion =
            "The same in every region - what is printed on the board. The other boxes here are per region.";

        // ---- regions (2026-11-01) ---------------------------------------------------------------------

        // What a row of a board with regions says it applies to - the words of its drop-down list.
        public const string AllRegions = "All regions";

        private static readonly IReadOnlyList<string> UsualAppliesToChoices = [AllRegions, .. BoardRegions.All];

        // The choices, and a word the board already holds that is not one of them. The usual ones are ONE
        // list: a box's list replaced in the middle of a pick undoes the pick (ComponentPaneRegionalRow).
        public static IReadOnlyList<string> AppliesToChoices(string? region)
        {
            string own = region?.Trim() ?? string.Empty;

            return own.Length == 0 || BoardRegions.All.Contains(own, StringComparer.OrdinalIgnoreCase)
                ? UsualAppliesToChoices
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
            $"Moved to {region.Trim()} - it no longer shows here, as this is the other region. Pick \"{region.Trim()}\" to see it.";

        // A component or its pins made the same in every region: what went, and how to get it back.
        public static string MadeTheSameInEveryRegion(string what, string region) =>
            $"{what} now the same in every region, as {region.Trim()} had it. \"Undo\" brings the other region's back.";

        // "This data is different per region" ticked on a component with no pins: nothing to tell apart yet, so no step.
        public const string NoPinsToTellApart = "It has no pins yet, so there is nothing to tell apart per region - add its pins first.";

        // ... taken off while the region shown has none: its pins would be every region's, so all would go.
        public static string NoPinsInRegion(string region) =>
            $"{region.Trim()} has no pins of its own to make the same in every region. Pick the region whose pins it should be, and take \"This data is different per region\" off there.";

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

        // ---- TEMPORARY: the tool that makes pinouts of oscilloscope baselines ---------------------------
        // *** REMOVED BEFORE CRT 4.0 GOES LIVE - see BaselinePinMigration. ***

        public const string MigrationNote =
            "Temporary, until CRT 4.0 goes live: make the pins from the oscilloscope baselines - each pin's number and signal name, and the package when that is certain. \"Create pinouts for the whole board\" also fills in every component the board's KiCad data has, with the package its footprint says. What you typed stays, and \"Undo\" takes it back.";

        public const string MigrationNothing = "The pins already have everything the oscilloscope baselines and the KiCad data say.";

        public static string MigratedBoard(int components) =>
            components == 1
                ? "Made the pins of 1 component from its oscilloscope baselines and the KiCad data. \"Undo\" takes it back."
                : $"Made the pins of {components.ToString(CultureInfo.InvariantCulture)} components from their oscilloscope baselines and the KiCad data. \"Undo\" takes them all back.";

        public static string ShowOnBoardTip(string label, bool marked) =>
            marked
                ? $"Show [{label}] on the schematics - the \"Schematics\" tab, with only this component selected"
                : $"[{label}] is not marked on any picture yet, so there is nothing to show";

        public static string DeleteTip(string label) =>
            $"Take [{label}] off the board in your draft - every row of it in every region, with its pictures, oscilloscope baselines, pins and files. \"Undo\" brings it back.";

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

        // A component's page by the name its entry has.
        public static string PageName(ComponentPage page) => page switch
        {
            ComponentPage.Basics => "Basics",
            ComponentPage.Replacements => "Replacement parts",
            ComponentPage.Pinout => "Pinout",
            ComponentPage.Pictures => "Pictures",
            ComponentPage.Baselines => "Oscilloscope baseline",
            ComponentPage.Files => "Files",
            _ => "Links",
        };

        // A page's entry in the left column marked because of what is on the page.
        public const string PageHasErrorTip = "Something on this page stops the draft being submitted";

        // Under the short description once it is too long for the popup's line (owner report, 2026-11-03:
        // the counter's color alone was not seen). A warning: the draft can still be submitted.
        public static string DescriptionTooLong(int length) =>
            $"This is {length - ComponentVocabulary.ShortDescriptionLimit} too long - the popup has room for {ComponentVocabulary.ShortDescriptionLimit} characters. It can still be submitted, but please shorten it.";

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
            "Each of these can be matched to one picture of the board. Pick it under \"Associate KiCad file\" on the picture's card - \"" +
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

        // What the "Draft" mark says of a component, and of a board part - neither on a new board, where
        // nothing is published to differ from (owner request, 2026-11-02).
        public const string ComponentChangedTip =
            "Your draft changes this component - the published board reads otherwise until your draft is submitted and published";

        public const string BoardPartChangedTip =
            "Your draft changes this part of the board - the published board reads otherwise until your draft is submitted and published";

        // A section's heading with how many it holds - "Pictures (3)", as the plan draws it.
        public static string Heading(string name, int count) =>
            $"{name} ({count.ToString(CultureInfo.InvariantCulture)})";

        // ###########################################################################################
        // What an oscilloscope baseline still needs, or null when it has it all: its pin - what makes a
        // row a baseline and not a picture (ComponentImageKinds) - its signal name and an image (owner
        // request, 2026-11-02: "the Pin + Signal name + image are mandatory"). T/DIV, V/DIV and the
        // trigger level may stay empty; one of them was needed until then.
        // ###########################################################################################
        public static string? BaselineProblem(string? pin, string? signalName, string? file)
        {
            var missing = new List<string>(3);

            if (string.IsNullOrWhiteSpace(pin))
                missing.Add("its pin");

            if (string.IsNullOrWhiteSpace(signalName))
                missing.Add("its signal name");

            if (string.IsNullOrWhiteSpace(file))
                missing.Add("an image");

            if (missing.Count == 0)
                return null;

            string listed = missing.Count == 1
                ? missing[0]
                : $"{string.Join(", ", missing.Take(missing.Count - 1))} and {missing[^1]}";

            return $"An oscilloscope baseline needs {listed}.";
        }

        // Under a baseline whose edit is held back: what it needs, and that the draft keeps it as it was.
        public static string BaselineHeldBack(string problem) => $"{problem} Until it has, your draft keeps it as it was.";

        // ###########################################################################################
        // Above a component's oscilloscope baselines: how the board's were measured is in its "Scope
        // measurements" files - or not written down yet, and where it goes (owner decision, 2026-10-26).
        // `elsewhere`: the region the board's notes are of when none is of the region shown - never "not
        // written down", which asks for a second copy of notes the board has (code review, 2026-11-02).
        // ###########################################################################################
        public static string MeasuringNotesLine(bool hasNotes, string? elsewhere = null)
        {
            if (hasNotes)
                return "How this board's baselines were measured:";

            string where = $"under \"Board: Files\", in the category \"{ScopeMeasurementFiles.Category}\"";

            return string.IsNullOrWhiteSpace(elsewhere)
                ? $"How this board's baselines were measured is not written down yet - add a file to the board {where}."
                : $"How this board's baselines were measured is written down for {elsewhere.Trim()} only - show {elsewhere.Trim()} to read it, or add a file for this region {where}.";
        }

        // ---- clone a component from another board (owner request, 2026-11-04) ---------------------------

        public const string CloneIntro =
            "Pick the component to clone. This component takes over everything it has - every field, its pins, replacement parts, pictures, oscilloscope baselines, files and links - and what it holds of those now is replaced. Its board label stays, and so do the places it is marked on this board's schematic images. \"Undo\" takes it all back.";

        public const string CloneNoOtherBoards =
            "There is no other published board on this computer to clone a component from.";

        public const string CloneBoardNotRead =
            "That board could not be read - its data may still be downloading. Try again in a moment.";

        public const string CloneBoardHasNoComponents = "That board has no components.";

        // What is shown of the component picked, a line each - the names in bold (BoldRuns), its title first.
        public static List<string> CloneFacts(ComponentCloneFacts facts)
        {
            ArgumentNullException.ThrowIfNull(facts);

            // What another board's data says is its words alone: a "<b>" in it says nothing of what is bold here
            // (code review, 2026-11-04).
            static string Its(string? data) => Handlers.Theming.BoldRuns.Plain(data).Trim();

            var lines = new List<string> { $"<b>{Its(ComponentPopupContent.TitleOf(facts.Row.BoardLabel, facts.Row))}</b>" };

            void Line(string name, string? value)
            {
                if (!string.IsNullOrWhiteSpace(Its(value)))
                {
                    lines.Add($"<b>{name}:</b> {Its(value)}");
                }
            }

            Line("Category", facts.Row.Category);
            Line("Part number", facts.Row.PartNumber);
            Line("Package", facts.Row.Package);
            Line("Short description", facts.Row.Description);

            lines.Add(
                $"<b>Holds:</b> {ContributeProblems.Counted(facts.Pins, "pin", "pins")}, {ContributeProblems.Counted(facts.Replacements, "replacement part", "replacement parts")}, " +
                $"{ContributeProblems.Counted(facts.Pictures, "picture", "pictures")}, {ContributeProblems.Counted(facts.Baselines, "oscilloscope baseline", "oscilloscope baselines")}, " +
                $"{ContributeProblems.Counted(facts.Files, "file", "files")}, {ContributeProblems.Counted(facts.Links, "link", "links")}");

            if (facts.EveryRegionComes)
            {
                lines.Add($"It differs per region and every region comes along - this is what {facts.Region} shows.");
            }

            if (facts.OtherRegionsLeftOut)
            {
                lines.Add($"It differs per region and this component has room for one - {facts.Region} is the one taken.");
            }

            return lines;
        }

        // Said once it is done - with the files that did not come, whose rows were left out.
        public static string Cloned(string fromLabel, string fromBoard, IReadOnlyCollection<string> filesLeftOut)
        {
            string done = $"Cloned from [{fromLabel}] of {fromBoard}. \"Undo\" takes it back.";

            return filesLeftOut.Count == 0
                ? done
                : $"{done} These files could not be brought over, so what cites them was left out: {string.Join(", ", filesLeftOut)}.";
        }

        // ---- the same part on other boards (slice 5) ------------------------------------------------------

        public const string CopyNeedsPartNumber =
            "Give the component its part number first - the same part is found on the other boards by it.";

        public static string CopyFoundNothing(string partNumber) =>
            $"No other board on this machine has a component with the part number [{partNumber.Trim()}].";

        // ---- also on other boards (slice 5) -----------------------------------------------------------

        public static string AlsoIntro(string partNumber) =>
            $"The same part - part number [{partNumber.Trim()}] - is on these boards. What you changed here goes into each ticked board's own draft, as one step there - each is reviewed by that board's maintainers. What the part has there of what is ticked is replaced.";

        public const string AlsoNothingChanged =
            "Your draft changes nothing of this part that could go to another board yet - its package, short description, pins, replacement parts, links, files or pictures.";

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

        // That board's own draft renamed or deleted the component (code review, 2026-11-02).
        public const string AlsoComponentNotInDraft = "that board's draft no longer has this component";

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

        // A file of no bytes is never added (owner decision, 2026-11-02: "Refuse empty files").
        public const string FileEmpty =
            "That file is empty - it holds no bytes, so it was not added. Pick a good copy of it.";

        // ---- the board's schematic images (step 3 - the "Schematic images" window, moved inline) ----

        // `notCopied`: images picked that could not be copied in - said, for they got no card. `empty`:
        // those left out as they hold no bytes (owner decision, 2026-11-02), said for what they are.
        public static string PicturesAdded(int count, int notCopied = 0, int empty = 0)
        {
            string added = count == 1 ? "Added 1 schematic image." : $"Added {count.ToString(CultureInfo.InvariantCulture)} schematic images.";

            string copied = notCopied switch
            {
                <= 0 => added,
                1 => $"{added} 1 could not be copied - it is open in another program, or no longer where it was picked.",
                _ => $"{added} {notCopied.ToString(CultureInfo.InvariantCulture)} could not be copied - they are open in another program, or no longer where they were picked.",
            };

            return empty switch
            {
                <= 0 => copied,
                1 => $"{copied} 1 was left out as it is empty.",
                _ => $"{copied} {empty.ToString(CultureInfo.InvariantCulture)} were left out as they are empty.",
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

        // Above the nets ("Net name" - "Net on this board" until 2026-11-02, owner request): what they
        // are - read from the board's KiCad data, never stored.
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
