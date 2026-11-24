using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

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

        // Said only for a component with a row of its own per region - which one is open, the region in bold (owner request,
        // 2026-11-09: "please bold the selected region (for all menus where this is applicable)"; BoldRuns.ShowMarked).
        public static string RegionLine(string? rowRegion)
        {
            // A row's region is typed data (a table cell): its own "<b>" or "</b>" taken out (code review, 2026-11-09).
            string region = Handlers.Theming.BoldRuns.Plain(rowRegion).Trim();

            return region.Length == 0
                ? string.Empty
                : $"The <b>{region}</b> version - each region has a row of its own, and this is the one CRT shows now.";
        }

        public static string NotOnBoard(string label) =>
            $"[{label}] is not on the board in your draft any more. If that was not meant, \"Undo\" in the draft's header brings it back.";

        // ---- a credit's date (2026-11-14) - the server's, set as the credit is published --------------

        public const string CreditDateToCome = "Dated when published";

        public static string CreditPublishedOn(string date) => $"Published {date}";

        // ---- no credit for the contributor's own changes (owner request, 2026-11-24) --------------------

        // Over "Board: Credits" while the draft's work goes without a credit because the contributor took it away
        // (SubmissionCredit.SendsNoCredit) - what the submission then says it changed is the changes, as "Review changes" words them.
        public const string NoCreditForChanges =
            "Your changes are sent without a credit, as you chose. The maintainers see them described by the list of what changed.";

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
        // The board's buttons by name - PAL's and NTSC's, or its own (2026-11-09); null for PAL and NTSC. The region shown in
        // bold (owner request, 2026-11-09; BoldRuns.ShowMarked).
        public static string RegionShownLine(string? region, IReadOnlyList<string>? regions = null)
        {
            IReadOnlyList<string> buttons = regions is { Count: > 0 } ? regions : BoardRegions.PalAndNtsc;
            string named = BoardRegions.Joined(buttons.Select(each => $"\"{each}\"").ToList());

            // The region can be a row's, typed in a table cell: its own "<b>" or "</b>" taken out (second code review, 2026-11-09).
            string shown = Handlers.Theming.BoldRuns.Plain(region).Trim();

            return buttons.Count == 1
                ? $"Showing <b>{shown}</b> - the {named} button under the component list switches it."
                : $"Showing <b>{shown}</b> - the {named} buttons under the component list switch it.";
        }

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
        public const string AllRegions = BoardRegions.EveryRegion;

        // The usual lists, one per set of regions: a box's list replaced in the middle of a pick undoes the pick
        // (ComponentPaneRegionalRow), so the same regions always give the very same list.
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, IReadOnlyList<string>> UsualAppliesToChoices = new(StringComparer.Ordinal);

        // ###########################################################################################
        // The choices for a row of a board coming in `regions` (BoardRegions.Of - 2026-11-09: names of its own since), and a
        // word the row already holds that is not one of them.
        // ###########################################################################################
        public static IReadOnlyList<string> AppliesToChoices(IReadOnlyList<string> regions, string? region)
        {
            ArgumentNullException.ThrowIfNull(regions);

            string own = region?.Trim() ?? string.Empty;
            IReadOnlyList<string> usual = UsualAppliesToChoices.GetOrAdd(string.Join('\u001F', regions), _ => [AllRegions, .. regions]);

            return own.Length == 0 || regions.Contains(own, StringComparer.OrdinalIgnoreCase)
                ? usual
                : [.. usual, own];
        }

        // The choice a row's region is, spelt as the board spells it.
        public static string AppliesToOf(IReadOnlyList<string> regions, string? region)
        {
            ArgumentNullException.ThrowIfNull(regions);

            string own = region?.Trim() ?? string.Empty;

            return own.Length == 0 ? AllRegions : regions.FirstOrDefault(each => string.Equals(each, own, StringComparison.OrdinalIgnoreCase)) ?? own;
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

        // "Board: Regions"' tick, in the owner's words (2026-11-09; "This board comes in PAL and NTSC versions" until then).
        public const string RegionsHasOnBoard =
            "This board comes with multiple regions";

        // What the board comes in, named - "Board status"' line.
        public static string RegionsOfBoard(IReadOnlyList<string> regions) =>
            $"This board comes in {BoardRegions.Joined(regions)} versions.";

        // Why the tick cannot be taken away while rows say a region - worded as a region's own line, the count bold (owner
        // request, 2026-11-09; BoldRuns).
        public static string RegionsInUse(int rows) => rows == 1
            ? "<b>1</b> row of this board has a region applied, so this stays checked until it applies to all regions."
            : $"<b>{rows.ToString(CultureInfo.InvariantCulture)}</b> rows of this board have a region applied, so this stays checked until they apply to all regions.";

        // Why a region cannot be deleted, under it on "Board: Regions" - the count bold (owner request, 2026-11-09: "change this
        // to "<b>2</b> rows of this board have it applied, so it ..""; BoldRuns.MarkedText shows it).
        public static string RegionInUse(int rows) => rows == 1
            ? "<b>1</b> row of this board has it applied, so it stays until that row applies to another region or to all of them."
            : $"<b>{rows.ToString(CultureInfo.InvariantCulture)}</b> rows of this board have it applied, so it stays until they apply to another region or to all of them.";

        // The board's own files and links with a region applied, under its components - nothing for none; the counts bold.
        public static string RegionOnBoardRows(int files, int links)
        {
            string Counted(int count, string one, string many) =>
                $"<b>{count.ToString(CultureInfo.InvariantCulture)}</b> {(count == 1 ? one : many)}";

            List<string> parts = [];
            List<string> pages = [];

            if (files > 0)
            {
                parts.Add(Counted(files, "board file", "board files"));
                pages.Add("\"Board: Files\"");
            }

            if (links > 0)
            {
                parts.Add(Counted(links, "board link", "board links"));
                pages.Add("\"Board: Links\"");
            }

            // Only the pages that hold them are named (code review, 2026-11-11: one link sent the reader to "Board: Files" too).
            return parts.Count == 0 ? string.Empty : $"Also on {string.Join(" and ", parts)} - under {string.Join(" and ", pages)}.";
        }

        // What goes with a region deleted - the rows of components told apart per region it has, and their pins (code review,
        // 2026-11-09: they went without a word); the count bold, as RegionInUse's.
        public static string RegionTakesRows(int rows) => rows == 1
            ? "Deleting it takes <b>1</b> row with it - a component's own data for this region. \"Undo\" brings it back."
            : $"Deleting it takes <b>{rows.ToString(CultureInfo.InvariantCulture)}</b> rows with it - components' own data for this region. \"Undo\" brings them back.";

        // ... and why the last two stay (BoardRegions.FewestRegions).
        public const string RegionsAtFewest =
            "A board with regions comes in two at least - to have none, uncheck \"This board comes with multiple regions\".";

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


        public static string ShowOnBoardTip(string label, bool marked) =>
            marked
                ? $"Show [{label}] on the schematics - the \"Schematics\" tab, with only this component selected"
                : $"[{label}] is not marked on any schematic image yet, so there is nothing to show";

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
            "Each of these can be matched to one schematic image of the board. Pick it under \"Associate KiCad file\" on the schematic image's card - \"" +
            SchematicImages + "\" on the left of the \"Contribute\" tab.";

        // Under an important signal whose net the Schematics tab's panel will not find - "highlighted", never "lit up"
        // (owner request, 2026-11-11; code review, 2026-11-12: this one was missed).
        public static string NetNotFound(string net) =>
            $"[{net.Trim()}] is no net in the board's KiCad data - the \"Schematics\" tab cannot highlight this signal.";

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

        // "Reveal" on a file the file manager could not be asked to show.
        public static string FileNotRevealed(string? name) =>
            $"[{name?.Trim()}] could not be shown in your file manager.";

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
            Line(BoardWorkbookSchema.DescriptionHeading, facts.Row.Description);
            Line(BoardWorkbookSchema.ColPurpose, RichNote.PlainOf(facts.Row.Purpose));

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

        // A file of the board that is in neither the draft nor the data on this machine.
        public static string FileNotFound(string? name) =>
            $"[{name?.Trim()}] is not on this machine yet - it comes with the next data update.";

        public const string CaptureFailed =
            "Nothing came from the oscilloscope - the \"Oscilloscope\" tab says why.";

        // ---- "Capture from oscilloscope pin-by-pin" (owner requests, 2026-11-18 and 2026-11-20; BaselineCaptureWindow) -------------

        public static string CaptureWindowTitle(string label) => $"Capture from oscilloscope pin-by-pin - {label}";

        public const string CaptureWindowIntro =
            "Select the pins to capture, or click them on the case. Probe the pin shown and press Enter to capture the oscilloscope's " +
            "screen and how it is set. Check it - beside the old one, when it had one - then \"Accept and continue to next pin\" makes it " +
            "that pin's baseline in your draft. Each capture accepted is one step \"Undo\" takes back.";

        public const string ScopeConnected = "Oscilloscope connected - Enter captures.";

        public const string ScopeNotConnected =
            "The oscilloscope is not connected - connect it on the \"Oscilloscope\" tab to capture. What was captured is kept.";

        public const string RunNoPins =
            "This component has no pins to go round yet - give it its case on its \"Pinout\" page, or add its pins there.";

        // Said as a capture is accepted - capturing never asks for one (code review, 2026-11-21: "before it is captured" sent
        // contributors back to capturing again).
        public const string RunNeedsName = "This pin has no signal name yet - type one, then accept the capture.";

        // A capture whose baseline changed in the pane, or whose component went, while it was going into the draft (code review,
        // 2026-11-21: it was filed under whatever the pane showed by then).
        public const string RunCaptureNotPut =
            "The component's baselines changed while this capture was going in, so it was not put in - it is still waiting to be accepted.";

        // A capture whose line went while it was being taken - the pin's baselines changed in the pane meanwhile.
        public const string RunCaptureLost = "The pin's baselines changed while it was being captured - capture it again.";

        // "Capture again" on a baseline the window cannot find - it changed meanwhile.
        public const string RunBaselineNotFound =
            "That baseline changed meanwhile - pick its pin in the \"Capture from oscilloscope pin-by-pin\" window.";

        // "Capture again" while the window is taking a capture or putting one into the draft - it moves to no other pin meanwhile
        // (code review, 2026-11-22: it came to the front, and nothing was captured or said).
        public const string RunWindowBusy =
            "The \"Capture from oscilloscope pin-by-pin\" window is still busy with a capture - press \"Capture again\" once it is done.";

        public const string RunNothingPicked = "No pins picked - select the pins to capture, or pick them with the buttons over the list.";

        // How many of the picked baselines are accepted (bold, BoldRuns) - a pin may have several.
        public static string RunProgress(int picked, int accepted) =>
            picked == 0
                ? ComponentPaneWording.RunNothingPicked
                : $"<b>{accepted.ToString(CultureInfo.InvariantCulture)}</b> of <b>{picked.ToString(CultureInfo.InvariantCulture)}</b> picked " +
                  $"{(picked == 1 ? "baseline" : "baselines")} accepted";

        // A step's line in the list: its pin and signal name.
        public static string RunTargetLabel(BaselineRunStep step)
        {
            ArgumentNullException.ThrowIfNull(step);
            return step.Name.Trim().Length == 0 ? step.PinText : $"{step.PinText}  {step.Name.Trim()}";
        }

        // Under it: accepted, captured and to be checked, a baseline to capture again, or none yet.
        public static string RunTargetState(BaselineRunStep step)
        {
            ArgumentNullException.ThrowIfNull(step);

            return step.IsAccepted ? "Accepted"
                : step.IsCaptured ? $"Captured - {NotYetAccepted}"
                : step.IsUpdate ? "Has a baseline"
                : step.HadBaseline ? "Baseline with no image"
                : "No baseline yet";
        }

        // The pin shown, its number and signal name bold (BoldRuns).
        public static string RunCurrent(BaselineRunStep step)
        {
            ArgumentNullException.ThrowIfNull(step);
            return $"Pin <b>{Handlers.Theming.BoldRuns.Plain(step.PinText)}</b>{ComponentPaneWording.NamePart(step)}";
        }

        // Under it: what to do with it now, and what the baseline was measured in, so the machine can be put in that state.
        public static string RunStepDetails(BaselineRunStep step, string machineState)
        {
            ArgumentNullException.ThrowIfNull(step);

            string what = step.IsAccepted
                ? "Accepted. Enter captures it again."
                : step.IsCaptured
                    ? "Captured - not yet accepted, so not in your draft. Check it, then \"Accept and continue to next pin\". Enter captures it again."
                    : step.IsUpdate
                        ? "Enter captures its baseline again."
                        : step.HadBaseline
                            ? "Its baseline has no image yet - Enter captures one, to be accepted."
                            : "No baseline yet - Enter captures one, to be accepted.";

            return string.IsNullOrWhiteSpace(machineState)
                ? what
                : $"{what} Machine state: {machineState.Trim()}";
        }

        // The pin "Accept and continue to next pin" goes on to, or that none is left (bold, BoldRuns).
        public static string RunNext(BaselineRunStep? next) =>
            next is null
                ? "Every picked pin is accepted."
                : $"Next: pin <b>{Handlers.Theming.BoldRuns.Plain(next.PinText)}</b>{ComponentPaneWording.NamePart(next)}";

        // With none after it, the pin shown picked and not accepted yet.
        public const string RunShownIsLast = "This is the last picked pin to accept.";

        private static string NamePart(BaselineRunStep step) =>
            step.Name.Trim().Length == 0 ? string.Empty : $" - <b>{Handlers.Theming.BoldRuns.Plain(step.Name).Trim()}</b>";

        // ###########################################################################################
        // The window's keys - the component info window's with an oscilloscope connected (Controlling-oscilloscope-with-keyboard;
        // owner request, 2026-11-20: "do support the same keyboard options as available in the component info window"), Enter
        // capturing and the number keys going to a pin as there; Right is "Accept and continue to next pin" (owner request, the same
        // day: the button that was "Next pin").
        // ###########################################################################################
        public static IReadOnlyList<(string Keys, string What)> RunKeys { get; } =
        [
            ("Enter", "Capture the pin shown - again to retake it"),
            ("Right", "Accept and continue to the next pin"),
            ("Left", "The picked pin before it"),
            ("0 - 9", "The pin with that number"),
            ("Space", "The first picked pin"),
            ("Up / Down", "Trigger level up or down"),
            ("Numpad + / -", "Time per division, a step either way"),
            ("Numpad 1 / 2", "1 V or 2 V per division"),
            ("Numpad .", "Clear statistics"),
            ("Numpad *", "Single"),
            ("Numpad /", "Run"),
            ("Escape", "Close this window"),
        ];

        // ---- the pictures: the old baseline beside the new capture (owner requests, 2026-11-20: "I want also to see the old baseline
        // image", and, "Differences" taken away the same day, "have it more clear/visible what is the old and new image") ----------

        public const string PictureOld = "Old baseline";

        public const string PictureNew = "New capture";

        // An existing baseline not captured in this run.
        public const string PictureBaseline = "Baseline";

        public const string PictureOldUnreadable = "The old baseline's picture could not be read.";

        // The new capture's title over it - checked or not yet.
        public static string PictureNewTitle(bool accepted) => accepted ? $"{PictureNew} - accepted" : $"{PictureNew} - {NotYetAccepted}";

        // A capture waiting to be checked and accepted - not in the draft until it is (owner request, 2026-11-20: "Rename the "New
        // capture - to check" to "New capture - not yet accepted", so it is clear this must be accepted before it will be used").
        public const string NotYetAccepted = "not yet accepted";

        // ###########################################################################################
        // Closing the window with captures not accepted: they are in no draft, so the window asks before they go (`count` of them).
        // ###########################################################################################
        public static string RunNotAcceptedOnClose(int count) =>
            count == 1
                ? "<b>1</b> capture is not yet accepted, so it is not in your draft. Accept it before closing, or discard it?"
                : $"<b>{count.ToString(CultureInfo.InvariantCulture)}</b> captures are not yet accepted, so they are not in your draft. Accept them before closing, or discard them?";

        // Over a picture: what it is, bold, and how the oscilloscope was set for it (BoldRuns) - the settings are the data's words, a
        // "<b>" in them no mark (code review, 2026-11-21).
        public static string RunPictureHeader(string title, string settings) =>
            string.IsNullOrWhiteSpace(settings) ? $"<b>{title}</b>" : $"<b>{title}</b>   {Handlers.Theming.BoldRuns.Plain(settings).Trim()}";

        // How the oscilloscope was set for a baseline, short - its T/DIV, V/DIV and trigger level, those it has, as the data spells
        // them: "100us, 2V, 1.25V". A line of "Capture from oscilloscope pin-by-pin"'s list, and the strip over each picture.
        public static string RunSettings(ComponentImageEntry baseline)
        {
            ArgumentNullException.ThrowIfNull(baseline);

            return string.Join(", ", new[] { baseline.TimeDiv, baseline.VoltsDiv, baseline.TriggerLevelVolts }
                .Select(value => value?.Trim() ?? string.Empty)
                .Where(value => value.Length > 0));
        }

        // ###########################################################################################
        // How the oscilloscope was set for a baseline, each value with its name - the line under a picture seen large (owner requests,
        // 2026-11-20: "have the headers show also, so it is clear what the value is", then "again do remember to have the values
        // "bold""); one not set says so, not bold. The viewer shows a note with formatting, so this is one (RichNote): written by
        // RichNote.Serialize, a "*" or "__" a value holds stays a character, never a mark.
        // ###########################################################################################
        public static string RunSettingsNamed(ComponentImageEntry baseline)
        {
            ArgumentNullException.ThrowIfNull(baseline);

            var text = new StringBuilder();
            var styles = new List<RichNoteStyle>();

            void Add(string words, RichNoteStyle style)
            {
                text.Append(words);
                styles.AddRange(Enumerable.Repeat(style, words.Length));
            }

            void Field(string name, string? value)
            {
                if (text.Length > 0)
                {
                    Add("   |   ", RichNoteStyle.None);
                }

                Add($"{name}: ", RichNoteStyle.None);

                if (value?.Trim() is { Length: > 0 } words)
                {
                    Add(words, RichNoteStyle.Bold);
                }
                else
                {
                    Add("(none)", RichNoteStyle.None);
                }
            }

            Field(FactTimeDiv, baseline.TimeDiv);
            Field(FactVoltsDiv, baseline.VoltsDiv);
            Field(FactTriggerLevel, baseline.TriggerLevelVolts);

            return RichNote.Serialize(new RichNoteText(text.ToString(), styles));
        }

        // A pin's picture seen large (WorklogPhotoViewerWindow): "Pin 3 - 1Y", and which of its pictures it is.
        public static string RunPictureTitle(BaselineRunStep step, string? which)
        {
            ArgumentNullException.ThrowIfNull(step);

            string pin = step.Name.Trim().Length == 0 ? $"Pin {step.PinText}" : $"Pin {step.PinText} - {step.Name.Trim()}";
            return string.IsNullOrEmpty(which) ? pin : $"{pin} - {which}";
        }

        // Following the pointer over a picture - never a tooltip, whose window under the pointer takes the next click.
        public const string PictureHint = "Click to see it larger";

        // ---- every field of the baseline shown ------------------------------------------------------------------------------------

        public const string FactPin = "Pin";
        public const string FactSignalName = "Signal name";
        public const string FactExpectedReading = "Expected reading";
        public const string FactTimeDiv = "T/DIV";
        public const string FactVoltsDiv = "V/DIV";
        public const string FactTriggerLevel = "Trigger level";
        public const string FactMachineState = "Machine state";
        public const string FactNote = "Note";
        public const string FactAppliesTo = "Applies to";

        // ###########################################################################################
        // Beside a baseline's picture, every field the component info popup shows of it (owner request, 2026-11-18: "Everything that
        // will be shown, should also be shown here, so I can validate it"), in the popup's order - pin, signal name, expected
        // reading, T/DIV, V/DIV, trigger level, then machine state and note - and what it applies to on a board with regions; the
        // window lines them up in a table (owner request, 2026-11-20: "align the values, so it looks nicer and easier to get an
        // overview"). A field left empty says so: the popup leaves it out, and a setting the oscilloscope did not answer is what is
        // to be checked. THE VALUE BOLD, NOT ITS NAME (owner request, the same day: "I want the value to be the important one here");
        // a value that changed since the baseline the run started from (`before`) says what it was. Each its name and its value as
        // BoldRuns shows it; what the data says is its words alone, a "<b>" in it no mark.
        // ###########################################################################################
        public static IReadOnlyList<(string Name, string Value)> RunFacts(ComponentImageEntry baseline, bool boardHasRegions, ComponentImageEntry? before = null)
        {
            ArgumentNullException.ThrowIfNull(baseline);

            static string Its(string? data) => Handlers.Theming.BoldRuns.Plain(data).Trim();

            (string Name, string Value) Fact(string name, Func<ComponentImageEntry, string?> valueOf)
            {
                string now = Its(valueOf(baseline));
                string value = now.Length > 0 ? $"<b>{now}</b>" : "(none)";

                if (before is not null && Its(valueOf(before)) is var earlier && !string.Equals(earlier, now, StringComparison.Ordinal))
                {
                    value += earlier.Length > 0 ? $"  (was {earlier})" : "  (was none)";
                }

                return (name, value);
            }

            var facts = new List<(string Name, string Value)>
            {
                Fact(FactPin, entry => entry.Pin),
                Fact(FactSignalName, entry => entry.Name),
                Fact(FactExpectedReading, entry => entry.ExpectedOscilloscopeReading),
                Fact(FactTimeDiv, entry => entry.TimeDiv),
                Fact(FactVoltsDiv, entry => entry.VoltsDiv),
                Fact(FactTriggerLevel, entry => entry.TriggerLevelVolts),
                Fact(FactMachineState, entry => entry.MachineState),
                Fact(FactNote, entry => RichNote.PlainOf(entry.Note)),
            };

            if (boardHasRegions)
            {
                facts.Add(Fact(FactAppliesTo, entry => entry.Region is { Length: > 0 } region ? region : ComponentPaneWording.AllRegions));
            }

            return facts;
        }

        // ###########################################################################################
        // While no oscilloscope is connected, over the baselines: what connecting one offers here (owner request, 2026-11-18: "it
        // should be shown somewhere, that the user can connect his oscilloscope to get more options here, as I at first thought it
        // was not implemented") - and, with the "Oscilloscope" tab turned off, where it is turned on.
        // ###########################################################################################
        public static string ScopeHint(bool tabTurnedOn)
        {
            const string offers =
                "its screen becomes a baseline once you accept it: \"Capture from oscilloscope pin-by-pin\" goes round the component's pins, " +
                "and \"Capture again\" on a baseline opens it on that pin - each taking how the oscilloscope is set too.";

            return tabTurnedOn
                ? $"Connect an oscilloscope on your network in the \"Oscilloscope\" tab, and {offers}"
                : $"Check \"Enable network connected oscilloscope tab\" in the \"Configuration\" tab, connect an oscilloscope on your network in the \"Oscilloscope\" tab, and {offers}";
        }

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
            "Draw every trace on this side of the board on the preview, in the color of the traces on this side - on the \"Schematics\" tab only the traces of the highlighted net are drawn so";

        public const string OtherSideTracesTip =
            "Draw every trace on the board's other side on the preview, in the color of the traces on the other side - on the \"Schematics\" tab only the traces of the highlighted net are drawn so";

        // Why the traces cannot be shown - the image names no KiCad view, or is not aligned with it.
        public const string ShowTracesNeedsAlignTip =
            "Needs KiCad data, a KiCad view of a side of the board picked for this image, and KiCad aligned on it (\"Align KiCad data\")";

        // The KiCad options of an image with no KiCad data (owner request, 2026-10-29: "if the
        // schematic image has no KiCad data, then the KiCad (traces) options should be disabled").
        public const string TraceColourTip = "With KiCad data: the color of a highlighted trace on the board's other side";

        // Each side's traces have a colour and an opacity of their own (owner request, 2026-10-30).
        public const string ThisSideTraceColourTip =
            "With KiCad data: the color of a highlighted trace on the image's own side - the Schematic image color until one is picked";

        public const string ThisSideTraceOpacityTip = "With KiCad data: how strongly a highlighted trace on the image's own side shows";

        public const string OtherSideTraceOpacityTip = "With KiCad data: how strongly a highlighted trace on the board's other side shows";

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
            "Nets of this board whose names say what they are - a supply, ground, a clock, reset, a bus line. Add the ones worth checking first - change a name before adding it if you want another; \"Show on board\" highlights a net on the \"Schematics\" tab, so you can see which it is.";

        // The board's other nets, under the suggestions (owner request, 2026-11-01).
        public static string OtherNetsHeading(int count) =>
            $"Other KiCad nets ({count.ToString(CultureInfo.InvariantCulture)})";

        public const string OtherNetsIntro =
            "Every other net of this board, each with a proposed name you can change. \"Add\" makes it an important KiCad signal under that name; \"Show on board\" highlights it on the \"Schematics\" tab.";

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

        // What "Fill from KiCad" did: "Filled from KiCad: 38 pins added, 2 pins filled in."
        public static string FilledFromKiCad(int added, int filled) => ComponentPaneWording.Filled("KiCad", added, filled);

        // ---- fill from the datasheet (owner request, 2026-11-17 - DatasheetPins.PinoutOf) ---------------

        public static string FillFromDatasheetTip(string label, string chip) =>
            $"Add every pin of the {chip}'s datasheet that [{label.Trim()}] lacks, with its name and direction, and fill in the names and directions still empty - nothing you typed is changed. \"Undo\" takes it back.";

        // What "Fill pins from datasheet" did: "Filled from the datasheet: 40 pins added."
        public static string FilledFromDatasheet(int added, int filled) => ComponentPaneWording.Filled("the datasheet", added, filled);

        // A press with nothing left to fill in - the last empty box filled as the press committed it (code review, 2026-11-17:
        // the press said nothing at all).
        public const string NothingToFillFromDatasheet =
            "The pins already have everything the datasheet has for this component.";

        private static string Filled(string source, int added, int filled)
        {
            var parts = new List<string>();

            if (added > 0)
                parts.Add(added == 1 ? "1 pin added" : $"{added.ToString(CultureInfo.InvariantCulture)} pins added");

            if (filled > 0)
                parts.Add(filled == 1 ? "1 pin filled in" : $"{filled.ToString(CultureInfo.InvariantCulture)} pins filled in");

            return $"Filled from {source}: {string.Join(", ", parts)}.";
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
