using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Tours
{
    // ###########################################################################################
    // *** THE INTRODUCTION TOURS (owner request, 2026-11-11: "I do not like Wiki documentation (too much to read), nor
    // creating YouTube videos ... can we make the tool itself provide basic introduction per tab, when the user see it the
    // first time?" - and then "Everyone should see it", "this is a view-demo only", Schematics, Contribute and Workbooks
    // first). *** Every step of every tour, and the one sentence that says what each tab is for. Main.Tours.cs decides
    // when one is shown, TourPlan which steps, Main/TourOverlay draws them.
    //
    // HOW A TOUR GROWS WITH CRT (asked by the owner the same day: "How to handle updates, and then new functionalities?
    // ... what should happen if one update and has skipped multiple previous versions?"):
    //
    //   - every step carries the CRT version it came with, `Since`. A step about something that CHANGED gets the version
    //     of the change, so it counts as new again;
    //   - CRT remembers per tab the newest `Since` of the tour last seen there (UserSettings.IntroductionToursSeen);
    //   - somebody who never saw the tab's tour gets its INTRODUCTION - the steps with InIntroduction, of every version;
    //   - somebody who did gets "Tour of what's new": the steps newer than what they saw, of EVERY version since, in the
    //     tab's own order - so skipping 4.1 and 4.2 shows what both brought, in one tour;
    //   - a step worth showing those who update, but no part of a first look (a small change), is InIntroduction: false.
    //
    // A step names the controls it frames by x:Name - "HardwareComboBox" in Main, "TabDrafts/ViewSwitch" inside a tab -
    // and MainToursTests fails on a name its window no longer has. One that is not on screen when the tour starts is left
    // out of that showing (the thumbnails in a window of their own, "My devices" with no device). A step naming nothing
    // shows its card in the middle of the window.
    //
    // Words: hobbyists, not professionals (CLAUDE.md), "color", plain ASCII, the board's pictures "schematic images", and
    // "highlight", never "light up" or "lit" (owner requests, 2026-11-11). EVERY TAB, BUTTON, PANEL AND KEY A TOUR NAMES
    // IS BOLD, NEVER IN QUOTES (owner request, 2026-11-11: "in all introductions, then try and use bold instead of
    // encapsulating in """) - "<b>...</b>" in a Body (BoldRuns); a Headline is bold as a whole and marks nothing.
    // ###########################################################################################
    // The tabs with a tour (owner request, 2026-11-11: "OK, now do the rest") - all but "Resources", "About" and "Maintainer"
    // (owner request, 2026-11-12: "no need to have an introduction for the "About" and "Resources" tabs", and the
    // Maintainer tab "is also much more limited and specialized"). MainToursTests fails on any other tab without one.
    internal enum TourTab
    {
        Schematics,
        Contribute,
        Workbooks,
        Overview,
        Oscilloscope,
        Configuration,
        Feedback,
    }

    // What the window is made to show for a step, and put back as it is left (Main.Tours.cs does it).
    internal enum TourStage
    {
        AsItIs,

        // The "Schematics" tab's three panels expanded (owner request, 2026-11-11: "Please show them expanded").
        SchematicsPanelsExpanded,

        // The page scrolled so the step's first control is in view - it lies further down a page that scrolls - and
        // scrolled back as it was once the step is left.
        ScrolledIntoView,
    }

    // One paragraph of a step's words - said always, or only while the control it is about (Shows, one of the step's
    // Targets) is on screen as the tour starts (owner question, 2026-11-11: "Will this step always show the current board
    // selected, as that may not include the regions buttons? Then what?"): a board without regions is told nothing of
    // "PAL" and "NTSC", a schematic image without KiCad data nothing of "Global settings".
    internal sealed record TourPart(string Text, string? Shows = null);

    internal sealed class TourBody
    {
        public TourBody(params TourPart[] parts) => this.Parts = parts;

        public IReadOnlyList<TourPart> Parts { get; }

        // A step whose words are all said always.
        public static implicit operator TourBody(string text) => new(new TourPart(text));

        // What the card says: the paragraphs whose control is on screen, a blank line between them.
        public string TextFor(Func<string, bool> isShown) =>
            string.Join("\n\n", this.Parts.Where(part => part.Shows is null || isShown(part.Shows)).Select(part => part.Text));

        // Everything it can say.
        public string Whole => this.TextFor(_ => true);

        public override string ToString() => this.Whole;
    }

    internal sealed record TourStep(
        string Since,
        IReadOnlyList<string> Targets,
        string Headline,
        TourBody Body,
        bool InIntroduction = true,
        TourStage Stage = TourStage.AsItIs);

    internal sealed record TourOfTab(TourTab Tab, string TabName, string Purpose, IReadOnlyList<TourStep> Steps);

    internal static class TourCatalogue
    {
        // The version every step of the first tours came with.
        private const string Crt4 = "4.0.0";

        public static readonly IReadOnlyList<TourOfTab> All =
        [
            new TourOfTab(
                TourTab.Schematics,
                "Schematics",
                "Here you see your board's schematic images, and find every component on them.",
                [
                    // The owner's steps and words (2026-11-11 - "if you can adapt the wordings so it sounds better, then do
                    // so"), the filter's possibilities filled in from ComponentListBuilder.MatchesSearch: every word typed,
                    // in the label, the name or the value.
                    new TourStep(
                        Crt4,
                        ["ComponentListsGrid", "ComponentSearchTextBox"],
                        "Find a component",
                        "Click a component in the list to highlight it on the schematic image to the right.\n\n" +
                        "To select a range, click the first component, then hold down <b>Shift</b> and click the last one.\n\n" +
                        "To narrow the list, type a board label, a name or a value in the filter box below it - such as <b>U1</b>, <b>CIA</b> or <b>6526</b>. Every component that matches is listed and highlighted, and several words find only those matching all of them.\n\n" +
                        "The categories selected above the list decide which components it shows."),
                    new TourStep(
                        Crt4,
                        ["ClearComponentsButton", "MarkAllComponentsButton", "RegionButtonsGrid", "FullscreenSchematicsButton", "BlinkSelectedCheckBox"],
                        "Below the component list",
                        new TourBody(
                            new TourPart("<b>Clear</b> empties the filter box and unselects every component."),
                            new TourPart("<b>Mark all</b> selects every component in the list."),
                            new TourPart("The region buttons, such as <b>PAL</b> and <b>NTSC</b>, switch between the board's versions for each region.", Shows: "RegionButtonsGrid"),
                            new TourPart("<b>Fullscreen</b> shows the schematic image on the whole screen - so does <b>F11</b>."),
                            new TourPart("<b>Blink selected</b> makes the selected components blink, so they are easy to spot."))),
                    new TourStep(
                        Crt4,
                        ["TabSchematicsControl/SchematicsContainer"],
                        "Move around the schematic image",
                        "Scroll to zoom in and out where the pointer is, and drag with the right or middle mouse button to move around.\n\n" +
                        "Click a component on the schematic image to see all about it - its pinout, pictures and oscilloscope readings."),
                    // Shown only for a schematic image with KiCad data - the badge beside "Draw trace" says it has (owner
                    // request, 2026-11-12: "show an example of how to select a trace, if it has KiCad data").
                    new TourStep(
                        Crt4,
                        ["TabSchematicsControl/KiCadDataBadge"],
                        "Highlight a trace",
                        "The small <b>Has KiCad data</b> badge beside <b>Draw trace</b> says the board's KiCad data is shown on this schematic image, so its traces can be followed.\n\n" +
                        "Point at a trace or a pin, and its whole net is highlighted. Click it to keep it highlighted - on every schematic image with KiCad data - and click it again to let go. Several can be kept at once."),
                    new TourStep(
                        Crt4,
                        ["TabSchematicsControl/GlobalSettingsPanel", "TabSchematicsControl/LabelsPanel", "TabSchematicsControl/ImportantSignalsPanel"],
                        "Panels on the schematic image",
                        new TourBody(
                            new TourPart("<b>Global settings</b> decides how the board's KiCad traces are highlighted as you point at or select a trace or a component, and whether the other side's traces and a component's first pin are shown.", Shows: "TabSchematicsControl/GlobalSettingsPanel"),
                            new TourPart("<b>Labels visible</b> writes the board label, technical name or friendly name on each component - or on the selected ones only."),
                            new TourPart("<b>Important KiCad signals</b> lists the board's key signals, such as power, ground and the clock - click one to highlight where it runs.", Shows: "TabSchematicsControl/ImportantSignalsPanel"),
                            new TourPart("Click a panel's headline to fold it away or open it again.")),
                        Stage: TourStage.SchematicsPanelsExpanded),
                    new TourStep(
                        Crt4,
                        ["TabSchematicsControl/InlineThumbnailGallery"],
                        "Switch between schematic images",
                        "Each thumbnail is one of the board's schematic images - click it to show it large.\n\n" +
                        "With components selected, the thumbnails holding them get a blue name bar, and the others are dimmed.\n\n" +
                        "Drag the edge of the thumbnail panel to make it wider or narrower - the thumbnails resize to fit."),
                    new TourStep(
                        Crt4,
                        ["DataSyncStatusIconBorder"],
                        "Get the newest data",
                        "CRT looks for new or updated board data each time it starts. Click this icon to look again now - it turns while it does, and pointing at it says how it went."),
                ]),

            new TourOfTab(
                TourTab.Contribute,
                "Contribute",
                "Here you improve the data of the board shown in CRT - fix a mistake, or add a picture, a datasheet or a whole component.",
                [
                    new TourStep(
                        Crt4,
                        ["TabDrafts/TopLineText", "TabDrafts/MyDraftsComboBox", "TabDrafts/AddNewBoardButton"],
                        "The board you work on",
                        "Everything here is about the data of the board shown in the left-side panel. Select another board there if you want to work on other data.\n\n" +
                        "<b>My drafts</b> takes you to a board you have made changes to, and <b>Add a new board</b> starts one CRT does not have yet."),
                    new TourStep(
                        Crt4,
                        ["ComponentListsGrid", "ContributeSearchTextBox"],
                        "Pick a component",
                        "Click a component to change its data. Hold down <b>Ctrl</b> (<b>Cmd</b> on a Mac) or <b>Shift</b> to pick several - a change then reaches all of them at once."),
                    new TourStep(
                        Crt4,
                        ["TabDrafts/BoardNavScroll", "TabDrafts/AddComponentButton"],
                        "Pick what to work on",
                        "The board's own parts are listed here - its name, schematic images, files, links and credits - and under <b>Components</b> the pages of the component picked. A red badge marks something to fix before you can submit, and a yellow one something worth a look.\n\n" +
                        "<b>Add component</b> adds one the board does not have yet."),
                    new TourStep(
                        Crt4,
                        ["TabDrafts/ViewSwitch"],
                        "Two ways to edit",
                        "<b>UI</b> shows a page with fields to fill in. <b>Table</b> shows the same data as rows, where copy, paste and <b>Replace</b> make many changes at once."),
                    new TourStep(
                        Crt4,
                        [],
                        "Your changes stay on this computer",
                        "Your first change makes a draft, which saves itself as you go - there is nothing to save by hand.\n\n" +
                        "Nobody else sees it until you press <b>Submit</b> and a maintainer has looked at it. <b>My submissions</b> shows the status of each, with the review comments from a board maintainer."),
                ]),

            new TourOfTab(
                TourTab.Workbooks,
                "Workbooks",
                "Here you keep notes of your own repairs: what you found, what you did and what it cost.",
                [
                    new TourStep(
                        Crt4,
                        ["WorklogBar"],
                        "Start a repair",
                        "A workbook is one repair of one board. Start one with <b>Add workbook</b>, then <b>Add worklog</b> marks where on the board you worked, and what you did.\n\n" +
                        "<b>Add device</b> adds a machine of your own, to keep its repairs together."),
                    // How a worklog works (owner request, 2026-11-12: "I would like to also intro the worklogs and how this
                    // works") - in the middle, as "Add worklog" shows only once a workbook is there.
                    new TourStep(
                        Crt4,
                        [],
                        "Worklogs: what you found and did",
                        "A worklog is one finding or one piece of work in a repair. Press <b>Add worklog</b>, then drag a rectangle on the schematic image where it is - the worklog opens.\n\n" +
                        "Give it a title, and say whether it is a <b>Note</b>, something <b>Cosmetic</b> or an <b>Issue</b>. Describe it, and add the work done, comments, photos, files and links - and tick the components in its area as they are completed.\n\n" +
                        "A worklog stays <b>Open</b> until you set it <b>Closed</b>. On the <b>Schematics</b> tab each stands on the schematic image as a badge - click one to open it, and untick <b>Show worklogs</b> to hide them."),
                    new TourStep(
                        Crt4,
                        ["TabWorkbooks/WorkbookListArea", "TabWorkbooks/DeviceListArea"],
                        "Your repairs",
                        new TourBody(
                            new TourPart("Every workbook is listed here - click one to open it."),
                            new TourPart("Your devices are listed too - click one to see it and every repair made on it.", Shows: "TabWorkbooks/DeviceListArea"))),
                    new TourStep(
                        Crt4,
                        ["TabWorkbooks/FindRepairTextBox"],
                        "Find a previous repair",
                        "Type a word to search every workbook, worklog and device - a component, a symptom or a note."),
                    new TourStep(
                        Crt4,
                        ["TabWorkbooks/WorkbookViewGrid", "TabWorkbooks/ExportWorkbookButton"],
                        "The repair, step by step",
                        new TourBody(
                            new TourPart("The workbook picked is shown here: its worklogs on the board's schematic images, and what each one says."),
                            new TourPart("<b>Export to PDF</b> or <b>Export to ZIP</b> makes a write-up to keep or share.", Shows: "TabWorkbooks/ExportWorkbookButton"))),
                ]),

            new TourOfTab(
                TourTab.Overview,
                "Overview",
                "Here you see every component of the board as one list - and print it, or a bill of materials.",
                [
                    new TourStep(
                        Crt4,
                        ["TabOverview/OverviewItemsControl"],
                        "Every component at a glance",
                        "One row per component: its label, names, part number, package and notes, and its files and links.\n\n" +
                        "Click a component's label to see all about it, or a file or a link to open it. The categories and the filter box on the left narrow this list too."),
                    new TourStep(
                        Crt4,
                        ["TabOverview/PrintButtonsPanel"],
                        "Print a list or a bill of materials",
                        "Tick the components to print at the left of each row - the tick at the top ticks or unticks them all.\n\n" +
                        "<b>Print component list</b> lists them with their details, and <b>Print BOM</b> counts how many of each part the board needs. Both open in your web browser, ready to print."),
                ]),

            new TourOfTab(
                TourTab.Oscilloscope,
                "Oscilloscope",
                "Here you connect CRT to your own oscilloscope over the network, so it can be set up as each oscilloscope baseline was taken.",
                [
                    new TourStep(
                        Crt4,
                        ["TabOscilloscope/VendorComboBox", "TabOscilloscope/SeriesOrModelComboBox", "TabOscilloscope/ScopeNotListedButton"],
                        "Your oscilloscope",
                        "Pick its vendor, then its series or model.\n\n" +
                        "Not in the list? <b>My scope is not listed</b> makes a model of your own from a similar one, where each command can be tried out on your oscilloscope."),
                    new TourStep(
                        Crt4,
                        ["TabOscilloscope/HostTextBox", "TabOscilloscope/PortTextBox", "TabOscilloscope/AutoConnectOscilloscopeCheckBox", "TabOscilloscope/ConnectToOscilloscopeButton"],
                        "Connect over the network",
                        "Type your oscilloscope's network address, then press <b>Connect to oscilloscope</b> - or tick <b>Auto-connect oscilloscope</b>, and CRT connects by itself."),
                    new TourStep(
                        Crt4,
                        ["TabOscilloscope/OscilloscopeImageFolderTextBox", "TabOscilloscope/SelectOscilloscopeImageFolderButton"],
                        "Use it while you repair",
                        "Once connected, picking an oscilloscope baseline in a component's information window sets your oscilloscope up as that baseline was taken, so you can compare what you measure with it.\n\n" +
                        "Pictures you capture from your oscilloscope's screen are saved in the folder chosen here."),
                ]),

            new TourOfTab(
                TourTab.Configuration,
                "Configuration",
                "Here you decide how CRT looks and behaves, and which hardware and boards it lists.",
                [
                    new TourStep(
                        Crt4,
                        ["TabConfiguration/SettingsScrollViewer"],
                        "Settings",
                        "The theme, the component information window, updates of CRT and its data - and which tabs show: <b>Workbooks</b>, <b>Oscilloscope</b> and <b>Maintainer</b> can each be switched on or off here."),
                    new TourStep(
                        Crt4,
                        ["TabConfiguration/CataloguePanelBorder"],
                        "Hide what you do not need",
                        "Untick a hardware, a board or a schematic image you never use, and it is hidden from the lists."),
                    new TourStep(
                        Crt4,
                        ["TabConfiguration/TakeToursAgainButton", "TabConfiguration/ShowProblemsButton", "TabConfiguration/OpenDataFolderButton", "TabConfiguration/OpenWorkbooksFolderButton", "TabConfiguration/OpenLogsFolderButton"],
                        "Help when you need it",
                        "<b>Take the introduction tours again</b> shows every tab's introduction once more.\n\n" +
                        "<b>Show errors and warnings</b> lists everything that went wrong since CRT started, with <b>Copy all</b> to paste it into the <b>Feedback</b> tab.\n\n" +
                        "The buttons under it open the folders CRT keeps its data, your workbooks and its logs in.",
                        Stage: TourStage.ScrolledIntoView),
                ]),

            new TourOfTab(
                TourTab.Feedback,
                "Feedback",
                "Here you send a problem, an idea or a question straight to the developer.",
                [
                    new TourStep(
                        Crt4,
                        ["TabFeedback/EmailTextBox", "TabFeedback/FeedbackText"],
                        "Write your feedback",
                        "Describe what you found or would like - a screenshot can be pasted or dropped right into the text.\n\n" +
                        "Your email address is optional, but needed if you would like a reply."),
                    new TourStep(
                        Crt4,
                        ["TabFeedback/AttachmentsPanel", "TabFeedback/SubmitButton"],
                        "Attach and send",
                        "CRT's logfile and crash reports help find what went wrong - <b>Attach file(s)</b> and <b>Attach folder</b> add anything else.\n\n" +
                        "<b>Send feedback</b> sends it all to the developer."),
                ]),

        ];

        public static TourOfTab Of(TourTab tab) => TourCatalogue.All.First(tour => tour.Tab == tab);
    }
}
