using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPaneWording - what the component pane says (CRT 4.0 step 2, 2026-10-25): its title,
// the region line, the offer to add the case's pins, what an oscilloscope baseline still needs, the new
// component's form and the whole component's buttons (slice 3c), and the board's own lists (4a).
// ###########################################################################################
public sealed class ComponentPaneWordingTests
{
    [Fact]
    public void The_title_is_the_label_and_what_the_component_list_shows_beside_it()
    {
        Assert.Equal("Component: U8 | CPU | 6510", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510" }));
        Assert.Equal("Component: C12 | 10uF", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "C12", TechnicalNameOrValue = " 10uF " }));
        Assert.Equal("Component: R1", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "R1" }));
    }

    // ###########################################################################################
    // Why a region stays on "Board: Regions": how many rows have it applied, the count bold and said once (owner request,
    // 2026-11-09: "remove the first "2 rows say it" as it is both bad English and double-wording ... change this to "<b>2</b>
    // rows of this board have it applied, so it .."").
    // ###########################################################################################
    [Fact]
    public void A_region_in_use_says_how_many_rows_have_it_applied_the_count_bold()
    {
        Assert.Equal(
            "<b>2</b> rows of this board have it applied, so it stays until they apply to another region or to all of them.",
            ComponentPaneWording.RegionInUse(2));
        Assert.Equal(
            "<b>1</b> row of this board has it applied, so it stays until that row applies to another region or to all of them.",
            ComponentPaneWording.RegionInUse(1));
        Assert.StartsWith("<b>10</b> rows of this board have a region applied, so this stays checked", ComponentPaneWording.RegionsInUse(10), StringComparison.Ordinal);

        // What goes with a region deleted is said too (code review, 2026-11-09: it went without a word).
        Assert.Equal(
            "Deleting it takes <b>4</b> rows with it - components' own data for this region. \"Undo\" brings them back.",
            ComponentPaneWording.RegionTakesRows(4));
        Assert.StartsWith("Deleting it takes <b>1</b> row with it", ComponentPaneWording.RegionTakesRows(1), StringComparison.Ordinal);
    }

    // Said only for a component with a row of its own per region, the region in bold (owner request, 2026-11-09).
    [Fact]
    public void The_region_line_names_the_region_of_a_regional_row_only()
    {
        Assert.Equal(string.Empty, ComponentPaneWording.RegionLine(""));
        Assert.Equal(string.Empty, ComponentPaneWording.RegionLine(null));
        Assert.StartsWith("The <b>PAL</b> version", ComponentPaneWording.RegionLine(" PAL "), StringComparison.Ordinal);

        // A region typed in a table cell with a mark of its own bolds nothing wrong (code review, 2026-11-09) - nor in the line
        // saying which region a page shows.
        Assert.StartsWith("The <b>EUx</b> version", ComponentPaneWording.RegionLine("</b>EU<b>x"), StringComparison.Ordinal);
        Assert.StartsWith("Showing <b>NTSCx</b> - ", ComponentPaneWording.RegionShownLine("NTSC</b>x"), StringComparison.Ordinal);
    }

    // The line saying which region a page shows names the board's buttons - word for word as before on a PAL and NTSC board,
    // its own on a board of its own regions (2026-11-09) - and the region shown in bold, nothing else (owner request,
    // 2026-11-09: "please bold the selected region").
    [Fact]
    public void The_region_shown_line_names_the_boards_own_buttons()
    {
        Assert.Equal(
            "Showing <b>PAL</b> - the \"PAL\" and \"NTSC\" buttons under the component list switch it.",
            ComponentPaneWording.RegionShownLine("PAL"));
        Assert.Equal(
            "Showing <b>US</b> - the \"EU\", \"US\" and \"Japan\" buttons under the component list switch it.",
            ComponentPaneWording.RegionShownLine("US", ["EU", "US", "Japan"]));
        Assert.Equal(
            "Showing <b>EU</b> - the \"EU\" button under the component list switches it.",
            ComponentPaneWording.RegionShownLine("EU", ["EU"]));
    }

    // A section's heading says how many it holds, as the plan draws the pane.
    [Fact]
    public void A_sections_heading_counts_what_it_holds()
    {
        Assert.Equal("Oscilloscope baseline (3)", ComponentPaneWording.Heading("Oscilloscope baseline", 3));
        Assert.Equal("Pictures (0)", ComponentPaneWording.Heading("Pictures", 0));
    }

    // ###########################################################################################
    // WHAT AN OSCILLOSCOPE BASELINE STILL NEEDS (owner request, 2026-11-02: "the Pin + Signal name +
    // image are mandatory"): T/DIV, V/DIV and the trigger level may stay empty. Nothing when it has all three.
    // ###########################################################################################
    [Fact]
    public void What_an_oscilloscope_baseline_still_needs_is_said_in_plain_words()
    {
        Assert.Null(ComponentPaneWording.BaselineProblem("1", "PA0", "Scope baseline/U8_1.png"));

        Assert.Equal("An oscilloscope baseline needs its pin.", ComponentPaneWording.BaselineProblem(" ", "PA0", "x.png"));
        Assert.Equal("An oscilloscope baseline needs its signal name.", ComponentPaneWording.BaselineProblem("1", " ", "x.png"));
        Assert.Equal("An oscilloscope baseline needs an image.", ComponentPaneWording.BaselineProblem("1", "PA0", null));
        Assert.Equal("An oscilloscope baseline needs its pin, its signal name and an image.", ComponentPaneWording.BaselineProblem("", "", ""));
        Assert.Equal("An oscilloscope baseline needs its pin and an image.", ComponentPaneWording.BaselineProblem("", "PA0", null));

        Assert.Equal(
            "An oscilloscope baseline needs its pin. Until it has, your draft keeps it as it was.",
            ComponentPaneWording.BaselineHeldBack("An oscilloscope baseline needs its pin."));
    }

    // ###########################################################################################
    // The offer to add the case's pins: every pin while none is listed, then the ones still missing
    // - and nothing for a case CRT cannot count.
    // ###########################################################################################
    [Fact]
    public void The_case_pins_offer_counts_what_is_missing()
    {
        Assert.Equal(8, ComponentPaneWording.MissingCasePins("DIP-8", []));
        Assert.Equal(6, ComponentPaneWording.MissingCasePins("DIP-8", [new ComponentPinEntry { Pin = "1" }, new ComponentPinEntry { Pin = " 8 " }]));
        Assert.Equal(0, ComponentPaneWording.MissingCasePins("SIP", []));
        Assert.Equal(0, ComponentPaneWording.MissingCasePins("TO-92", [new ComponentPinEntry { Pin = "1" }, new ComponentPinEntry { Pin = "2" }, new ComponentPinEntry { Pin = "3" }]));

        Assert.Equal("Add the 40 pins of DIP-40", ComponentPaneWording.AddCasePins("DIP-40", 40, noPinsYet: true));
        Assert.Equal("Add the 3 missing pins", ComponentPaneWording.AddCasePins("DIP-40", 3, noPinsYet: false));
        Assert.Equal("Add the missing pin", ComponentPaneWording.AddCasePins("DIP-40", 1, noPinsYet: false));
    }

    [Fact]
    public void A_component_no_longer_on_the_board_says_how_to_bring_it_back()
    {
        Assert.Contains("[U8]", ComponentPaneWording.NotOnBoard("U8"), StringComparison.Ordinal);
        Assert.Contains("Undo", ComponentPaneWording.NotOnBoard("U8"), StringComparison.Ordinal);
    }

    // ---- the component as a whole (slice 3c) ------------------------------------------------------

    // What "Clone component from another board" says of the component picked: its title as the popup's,
    // what is filled, and every count - "0" as much as any.
    [Fact]
    public void What_is_shown_of_a_component_to_clone_is_its_title_what_is_filled_and_its_counts()
    {
        var facts = new ComponentCloneFacts(
            new ComponentEntry { BoardLabel = "U7", FriendlyName = "PLA", TechnicalNameOrValue = "906114", PartNumber = "906114-01", Category = "IC" },
            "PAL",
            Pins: 28,
            Replacements: 0,
            Pictures: 1,
            Baselines: 2,
            Files: 0,
            Links: 3,
            EveryRegionComes: false,
            OtherRegionsLeftOut: false);

        Assert.Equal(
            [
                "<b>U7 | PLA | 906114</b>",
                "<b>Category:</b> IC",
                "<b>Part number:</b> 906114-01",
                "<b>Holds:</b> 28 pins, 0 replacement parts, 1 picture, 2 oscilloscope baselines, 0 files, 3 links",
            ],
            ComponentPaneWording.CloneFacts(facts));
    }

    // What another board's data says is its words: a "<b>" in it says nothing of what is bold in the panel, and
    // only the panel's own headings are (code review, 2026-11-04).
    [Fact]
    public void A_tag_in_another_boards_data_is_neither_shown_nor_read_as_bold()
    {
        var facts = new ComponentCloneFacts(
            new ComponentEntry { BoardLabel = "U7", FriendlyName = "<b>PLA", Category = "<b>IC</b>", Description = "Glue <b>logic", PartNumber = "<b></b>" },
            "PAL",
            0, 0, 0, 0, 0, 0,
            EveryRegionComes: false,
            OtherRegionsLeftOut: false);

        List<string> lines = ComponentPaneWording.CloneFacts(facts);

        Assert.Equal(
            [
                "<b>U7 | PLA</b>",
                "<b>Category:</b> IC",
                "<b>Other technical details:</b> Glue logic",
            ],
            lines.Take(3));

        Assert.Equal(
            [("Other technical details:", true), (" Glue logic", false)],
            Handlers.Theming.BoldRuns.Marked(lines[2]));
    }

    [Fact]
    public void A_component_differing_per_region_says_what_comes_of_that()
    {
        var row = new ComponentEntry { BoardLabel = "U7" };

        Assert.Contains("every region comes along", ComponentPaneWording.CloneFacts(new ComponentCloneFacts(row, "PAL", 0, 0, 0, 0, 0, 0, true, false))[^1], StringComparison.Ordinal);
        Assert.Contains("NTSC is the one taken", ComponentPaneWording.CloneFacts(new ComponentCloneFacts(row, "NTSC", 0, 0, 0, 0, 0, 0, false, true))[^1], StringComparison.Ordinal);
    }

    [Fact]
    public void What_a_clone_says_names_where_it_came_from_and_any_file_left_out()
    {
        Assert.Equal("Cloned from [U7] of C64 / 250425. \"Undo\" takes it back.", ComponentPaneWording.Cloned("U7", "C64 / 250425", []));
        Assert.Contains("left out: Images/a.png", ComponentPaneWording.Cloned("U7", "C64 / 250425", ["Images/a.png"]), StringComparison.Ordinal);
    }

    // Only a label taken is said - nothing typed yet says nothing, and no category is said elsewhere.
    [Fact]
    public void Only_a_label_taken_is_said_under_the_new_label()
    {
        Assert.Equal(
            "This board already has [U8] - pick it in the list on the left to change it.",
            ComponentPaneWording.NewLabelProblem(ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists, " U8 "));
        Assert.Null(ComponentPaneWording.NewLabelProblem(ContributionPackaging.NewComponentProblem.BoardLabelMissing, ""));
        Assert.Null(ComponentPaneWording.NewLabelProblem(ContributionPackaging.NewComponentProblem.CategoryMissing, "U9"));
        Assert.Null(ComponentPaneWording.NewLabelProblem(ContributionPackaging.NewComponentProblem.None, "U9"));
    }

    [Fact]
    public void The_whole_components_buttons_say_what_they_do_and_why_not()
    {
        Assert.StartsWith("Show [U8] on the schematics", ComponentPaneWording.ShowOnBoardTip("U8", marked: true), StringComparison.Ordinal);
        Assert.Equal("[U8] is not marked on any schematic image yet, so there is nothing to show", ComponentPaneWording.ShowOnBoardTip("U8", marked: false));
        Assert.Contains("every region", ComponentPaneWording.DeleteTip("U8"), StringComparison.Ordinal);
        Assert.Contains("\"Undo\" brings it back", ComponentPaneWording.DeleteTip("U8"), StringComparison.Ordinal);
        Assert.Equal("Bring [U8] back", ComponentPaneWording.BringBack("U8"));
    }

    // ---- the board's own lists (slice 4a) -----------------------------------------------------------

    [Fact]
    public void A_board_part_is_named_and_a_net_not_found_says_it_will_not_light_up()
    {
        // Files and links apart, the pictures named schematic images (owner request, 2026-10-29), and each
        // headline naming its group of the left column's tree (owner request, 2026-11-01).
        Assert.Equal("Board: Files", ComponentPaneWording.Title(PaneBoardPart.Files));
        Assert.Equal("Board: Links", ComponentPaneWording.Title(PaneBoardPart.Links));
        Assert.Equal("Board: Schematic images", ComponentPaneWording.Title(PaneBoardPart.Pictures));
        Assert.Equal("Board: Credits", ComponentPaneWording.Title(PaneBoardPart.Credits));
        Assert.Equal("Board: Regions", ComponentPaneWording.Title(PaneBoardPart.Regions));
        Assert.Equal("Board: Categories", ComponentPaneWording.Title(PaneBoardPart.Categories));
        Assert.Equal("KiCad: Important signals", ComponentPaneWording.Title(PaneBoardPart.ImportantSignals));
        Assert.Equal(
            "[NRESET] is no net in the board's KiCad data - the \"Schematics\" tab cannot highlight this signal.",
            ComponentPaneWording.NetNotFound(" NRESET "));

        // The part's one headline, with its count (owner request, 2026-10-29).
        Assert.Equal("Board: Files (6)", ComponentPaneWording.PartTitle(PaneBoardPart.Files, 6));
        Assert.Equal("Board: Schematic images (0)", ComponentPaneWording.PartTitle(PaneBoardPart.Pictures, 0));
    }

    // What "Fill from KiCad" did, counted - one pin, or several.
    [Fact]
    public void Fill_from_KiCad_says_what_it_added_and_filled_in()
    {
        Assert.Equal("Filled from KiCad: 38 pins added, 2 pins filled in.", ComponentPaneWording.FilledFromKiCad(38, 2));
        Assert.Equal("Filled from KiCad: 1 pin added.", ComponentPaneWording.FilledFromKiCad(1, 0));
        Assert.Equal("Filled from KiCad: 1 pin filled in.", ComponentPaneWording.FilledFromKiCad(0, 1));
        Assert.Contains("[U8]", ComponentPaneWording.FillFromKiCadTip(" U8 "), StringComparison.Ordinal);
    }

    // Every sentence the pane says is plain ASCII and names neither Excel nor the old editor window.
    // The Pictures part's own lines (step 3 - the "Schematic images" window's work moved there): one
    // picture or several, and the name trimmed in what was removed.
    [Fact]
    public void Pictures_added_and_removed_are_said_in_plain_words()
    {
        Assert.Equal("Added 1 schematic image.", ComponentPaneWording.PicturesAdded(1));
        Assert.Equal("Added 3 schematic images.", ComponentPaneWording.PicturesAdded(3));
        Assert.StartsWith("Added 2 schematic images. 1 could not be copied - ", ComponentPaneWording.PicturesAdded(2, notCopied: 1), StringComparison.Ordinal);
        Assert.StartsWith("Added 0 schematic images. 2 could not be copied - they", ComponentPaneWording.PicturesAdded(0, notCopied: 2), StringComparison.Ordinal);
        Assert.StartsWith("Deleted [Top] from the board. ", ComponentPaneWording.PictureRemoved(" Top "), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // A SENTENCE SENDING THE CONTRIBUTOR TO THE BOARD'S PICTURES NAMES THE PART AS THE LEFT COLUMN DOES
    // (code review, 2026-10-30): the part became "Schematic images" on 2026-10-29, and the refusal of
    // Align KiCad and the "KiCad data" window's note still sent people to "Pictures", which the column
    // no longer has.
    // ###########################################################################################
    [Fact]
    public void Every_sentence_sending_the_contributor_to_the_pictures_names_the_part_as_the_column_does()
    {
        Assert.Equal("Board: Schematic images", ComponentPaneWording.Title(PaneBoardPart.Pictures));

        foreach (string sentence in new[] { PictureToolWording.NoKiCadView("Top"), ComponentPaneWording.KiCadViewsHint })
        {
            Assert.Contains(ComponentPaneWording.SchematicImages, sentence, StringComparison.Ordinal);
            Assert.DoesNotContain("Pictures", sentence, StringComparison.Ordinal);
        }
    }

    // Under the traces' ticks (2026-10-29): a side ticked with no copper says so; a side not ticked,
    // or one with copper, says nothing.
    [Fact]
    public void A_side_ticked_with_no_traces_says_so()
    {
        Assert.Equal(string.Empty, ComponentPaneWording.TracesNote(thisSide: true, 12, otherSide: true, 8));
        Assert.Equal(string.Empty, ComponentPaneWording.TracesNote(thisSide: false, 0, otherSide: true, 8));
        Assert.Equal("The board's KiCad data has no trace on this side of the image.", ComponentPaneWording.TracesNote(true, 0, true, 8));
        Assert.Equal("The board's KiCad data has no trace on the other side of this image.", ComponentPaneWording.TracesNote(false, 0, true, 0));
        Assert.Equal("The board's KiCad data has no trace on either side of this image.", ComponentPaneWording.TracesNote(true, 0, true, 0));
    }

    [Fact]
    public void The_new_sentences_are_plain_ASCII()
    {
        string[] sentences =
        [
            ComponentPaneWording.NewComponentTitle,
            ComponentPaneWording.NoCategoryYet,
            ComponentPaneWording.NoKiCadNets,
            ComponentPaneWording.ReadingKiCadNets,
            ComponentPaneWording.AddSignalTip,
            ComponentPaneWording.PickANetTip,
            ComponentPaneWording.AddSignalNeedsKiCadTip,
            ComponentPaneWording.FileNotOpened("Manual"),
            ComponentPaneWording.LinkNotOpened("Lemon64"),
            ComponentPaneWording.ComponentChangedTip,
            ComponentPaneWording.BoardPartChangedTip,
            ComponentPaneWording.DeleteTip("U8"),
            ComponentPaneWording.NetNotFound("X"),
            ComponentPaneWording.FillFromKiCadTip("U8"),
            ComponentPaneWording.NothingToFillFromKiCad,
            ComponentPaneWording.FilledFromKiCad(2, 3),
            ComponentPaneWording.NetHeadingTip,
            ComponentPaneWording.SignalSuggestionsHeading(3),
            ComponentPaneWording.SignalSuggestionsIntro,
            ComponentPaneWording.PicturesAdded(2),
            ComponentPaneWording.PicturesAdded(2, notCopied: 1),
            ComponentPaneWording.PicturesAdded(1, notCopied: 3),
            ComponentPaneWording.NoPicturesToAdd,
            ComponentPaneWording.PictureRemoved("Top"),
            ComponentPaneWording.KiCadViewsHint,
        ];

        Assert.All(sentences, sentence =>
        {
            Assert.All(sentence, character => Assert.True(character < 128, $"[{sentence}] is not plain ASCII"));
            Assert.DoesNotContain("Excel", sentence, StringComparison.Ordinal);
            Assert.DoesNotContain("component editor", sentence, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // In "Capture from oscilloscope pin-by-pin"'s window, every field the popup shows of a baseline (owner request, 2026-11-18: "Everything
    // that will be shown, should also be shown here, so I can validate it"), in the popup's order, an empty one said as such -
    // a setting the oscilloscope did not answer is what is to be checked. THE VALUE BOLD, NOT ITS NAME (owner request, 2026-11-20:
    // "I want the value to be the important one here"). The note as its words, what it applies to only on a board with regions,
    // and what the data says never read as a bold mark.
    // ###########################################################################################
    [Fact]
    public void A_capture_is_shown_with_every_field_the_popup_shows_its_value_bold_an_empty_one_said()
    {
        var baseline = new ComponentImageEntry
        {
            Pin = "3",
            Name = "<b>1Y</b>",
            ExpectedOscilloscopeReading = "Pulsing",
            TimeDiv = "100us",
            VoltsDiv = "2V",
            MachineState = "At the READY prompt",
            Note = "Measured **twice**",
            Region = "PAL",
        };

        IReadOnlyList<(string Name, string Value)> facts = ComponentPaneWording.RunFacts(baseline, boardHasRegions: false);

        Assert.Equal(
            ["Pin: 3", "Signal name: 1Y", "Expected reading: Pulsing", "T/DIV: 100us", "V/DIV: 2V", "Trigger level: (none)", "Machine state: At the READY prompt", "Note: Measured twice"],
            facts.Select(fact => $"{fact.Name}: {Handlers.Theming.BoldRuns.Plain(fact.Value)}"));
        Assert.Equal(("Pin", "<b>3</b>"), facts[0]);
        Assert.Equal(("Signal name", "<b>1Y</b>"), facts[1]);
        Assert.Equal(("Trigger level", "(none)"), facts[5]);

        Assert.Equal((ComponentPaneWording.FactAppliesTo, "<b>PAL</b>"), ComponentPaneWording.RunFacts(baseline, boardHasRegions: true)[^1]);
        Assert.Equal(
            $"<b>{ComponentPaneWording.AllRegions}</b>",
            ComponentPaneWording.RunFacts(new ComponentImageEntry { Pin = "1" }, boardHasRegions: true)[^1].Value);
    }

    // ###########################################################################################
    // A baseline captured again says what each changed value was in the baseline it replaced (owner request, 2026-11-20: "maybe
    // there can be a comparison, so I can easily see the diff?") - a value that stayed says nothing more, one emptied or newly
    // given says so.
    // ###########################################################################################
    [Fact]
    public void A_capture_says_what_each_changed_value_was_before()
    {
        var before = new ComponentImageEntry { Pin = "3", Name = "1Y", TimeDiv = "50us", VoltsDiv = "1V", MachineState = "At the READY prompt" };
        var now = new ComponentImageEntry { Pin = "3", Name = "1Y", TimeDiv = "100us", VoltsDiv = "1V", TriggerLevelVolts = "1.25V" };

        IReadOnlyList<(string Name, string Value)> facts = ComponentPaneWording.RunFacts(now, boardHasRegions: false, before);

        Assert.Equal(
            ["Pin: 3", "Signal name: 1Y", "Expected reading: (none)", "T/DIV: 100us  (was 50us)", "V/DIV: 1V", "Trigger level: 1.25V  (was none)", "Machine state: (none)  (was At the READY prompt)", "Note: (none)"],
            facts.Select(fact => $"{fact.Name}: {Handlers.Theming.BoldRuns.Plain(fact.Value)}"));
        Assert.Equal("<b>100us</b>  (was 50us)", facts[3].Value);
    }

    // ###########################################################################################
    // While no oscilloscope is connected, the baselines' page says what one offers there (owner request, 2026-11-18: "I at first
    // thought it was not implemented") - and, the "Oscilloscope" tab turned off, the setting that turns it on.
    // ###########################################################################################
    [Fact]
    public void The_hint_names_what_an_oscilloscope_offers_and_where_its_tab_is_turned_on()
    {
        string on = ComponentPaneWording.ScopeHint(tabTurnedOn: true);
        string off = ComponentPaneWording.ScopeHint(tabTurnedOn: false);

        Assert.All(new[] { on, off }, hint =>
        {
            Assert.Contains("\"Capture from oscilloscope pin-by-pin\"", hint, StringComparison.Ordinal);
            Assert.Contains("\"Capture again\"", hint, StringComparison.Ordinal);
            Assert.Contains("\"Oscilloscope\" tab", hint, StringComparison.Ordinal);
            Assert.DoesNotContain("Capture from my oscilloscope", hint, StringComparison.Ordinal);
            Assert.All(hint, character => Assert.True(character < 128, $"[{hint}] is not plain ASCII"));

            // A capture is a baseline only once accepted - never "straight into" one (code review, 2026-11-21).
            Assert.Contains("once you accept it", hint, StringComparison.Ordinal);
        });

        Assert.DoesNotContain("Configuration", on, StringComparison.Ordinal);
        Assert.Contains("\"Enable network connected oscilloscope tab\" in the \"Configuration\" tab", off, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // The capture window says how far it is - the picked baselines accepted, singular or plural - what is next, or that every
    // picked pin is accepted, and what to do with the pin shown: capture it, check and accept its capture, or retake it (owner
    // requests, 2026-11-20: "and then what it the next pin I should do", "it must stay on the pin, so I can check that it is good").
    // ###########################################################################################
    [Fact]
    public void The_capture_window_says_how_far_it_is_what_is_next_and_what_to_do_with_the_pin_shown()
    {
        var pin = new BaselineRunPin("3", 3, "1Y", "Output", 1);
        var update = new BaselineRunStep(pin, "3", "1Y", "Scope baseline/U27_3.png");
        var fresh = new BaselineRunStep(new BaselineRunPin("7", 7, "", "Ground", 0), "7", "", "");

        Assert.Equal("1 of 2 picked baselines accepted", Handlers.Theming.BoldRuns.Plain(ComponentPaneWording.RunProgress(2, 1)));
        Assert.Equal("0 of 1 picked baseline accepted", Handlers.Theming.BoldRuns.Plain(ComponentPaneWording.RunProgress(1, 0)));
        Assert.Equal(ComponentPaneWording.RunNothingPicked, ComponentPaneWording.RunProgress(0, 0));

        Assert.Equal("Next: pin <b>3</b> - <b>1Y</b>", ComponentPaneWording.RunNext(update));
        Assert.Equal("Next: pin 7", Handlers.Theming.BoldRuns.Plain(ComponentPaneWording.RunNext(fresh)));
        Assert.Equal("Every picked pin is accepted.", ComponentPaneWording.RunNext(null));

        Assert.Equal(("3  1Y", "Has a baseline"), (ComponentPaneWording.RunTargetLabel(update), ComponentPaneWording.RunTargetState(update)));
        Assert.Equal(("7", "No baseline yet"), (ComponentPaneWording.RunTargetLabel(fresh), ComponentPaneWording.RunTargetState(fresh)));
        Assert.Equal("Enter captures its baseline again.", ComponentPaneWording.RunStepDetails(update, " "));
        Assert.Equal("No baseline yet - Enter captures one, to be accepted. Machine state: Running", ComponentPaneWording.RunStepDetails(fresh, " Running "));

        // Captured, it is not yet accepted - so not in the draft (owner request, 2026-11-20: "so it is clear this must be accepted
        // before it will be used").
        update.Captured();
        Assert.Equal("Captured - not yet accepted", ComponentPaneWording.RunTargetState(update));
        Assert.Equal(
            "Captured - not yet accepted, so not in your draft. Check it, then \"Accept and continue to next pin\". Enter captures it again.",
            ComponentPaneWording.RunStepDetails(update, ""));

        update.Accepted("Scope baseline/U27_3_1Y.png");
        Assert.Equal("Accepted", ComponentPaneWording.RunTargetState(update));
        Assert.Equal("Accepted. Enter captures it again.", ComponentPaneWording.RunStepDetails(update, ""));

        // The viewer's title: the pin, and which of its pictures.
        Assert.Equal("Pin 3 - 1Y - Old baseline", ComponentPaneWording.RunPictureTitle(update, ComponentPaneWording.PictureOld));
        Assert.Equal("Pin 7", ComponentPaneWording.RunPictureTitle(fresh, null));

        // A signal name typed with a mark in it is its words, never a bold run of its own.
        Assert.Equal("Pin 3 - 1Y", Handlers.Theming.BoldRuns.Plain(ComponentPaneWording.RunCurrent(new BaselineRunStep(pin, "3", "<b>1Y</b>", ""))));

        // Right is the button that was "Next pin".
        Assert.Contains(ComponentPaneWording.RunKeys, key => key == ("Right", "Accept and continue to the next pin"));
        Assert.All(ComponentPaneWording.RunKeys, key => Assert.All(key.Keys + key.What, character => Assert.True(character < 128, $"[{key}] is not plain ASCII")));
    }

    // ###########################################################################################
    // Over each picture a strip says which it is - the old baseline, a baseline not captured again, or the new capture and whether
    // it is accepted (owner requests, 2026-11-20: "have it more clear/visible what is the old and new image", "Rename the "New capture
    // - to check" to "New capture - not yet accepted"") - its name bold, the oscilloscope's settings after it.
    // ###########################################################################################
    [Fact]
    public void A_pictures_strip_names_it_bold_with_its_settings()
    {
        Assert.Equal("<b>Old baseline</b>   50us, 1V, 1.5V", ComponentPaneWording.RunPictureHeader(ComponentPaneWording.PictureOld, "50us, 1V, 1.5V"));
        Assert.Equal("<b>New capture - not yet accepted</b>", ComponentPaneWording.RunPictureHeader(ComponentPaneWording.PictureNewTitle(accepted: false), " "));
        Assert.Equal("New capture - accepted", ComponentPaneWording.PictureNewTitle(accepted: true));
        Assert.Equal("Baseline", ComponentPaneWording.PictureBaseline);

        // The settings are the data's words: a "<b>" typed in a T/DIV cell is no mark (code review, 2026-11-21).
        string header = ComponentPaneWording.RunPictureHeader(ComponentPaneWording.PictureBaseline, "<b>50us, 1V");
        Assert.Equal([("Baseline", true), ("   50us, 1V", false)], Handlers.Theming.BoldRuns.Marked(header));
    }

    // ###########################################################################################
    // A baseline with no image yet is a line of its own - captured into it, not added beside it (code review, 2026-11-21: it read
    // "No baseline yet", and its capture went in as a second baseline).
    // ###########################################################################################
    [Fact]
    public void A_baseline_with_no_image_says_so_and_that_Enter_captures_one()
    {
        var step = new BaselineRunStep(new BaselineRunPin("2", 2, "1B", "Input", 1), "2", "1B", "", hadBaseline: true);

        Assert.Equal("Baseline with no image", ComponentPaneWording.RunTargetState(step));
        Assert.Equal("Its baseline has no image yet - Enter captures one, to be accepted.", ComponentPaneWording.RunStepDetails(step, ""));
    }

    // A line of "Capture from oscilloscope pin-by-pin"'s list, and the strip over each picture: the settings the baseline has, short,
    // as the data spells them - one not set left out, spaces round a value taken away, none at all nothing.
    [Fact]
    public void The_short_settings_are_those_the_baseline_has_in_order()
    {
        Assert.Equal("100us, 2V, 1.25V", ComponentPaneWording.RunSettings(new ComponentImageEntry { TimeDiv = "100us", VoltsDiv = "2V", TriggerLevelVolts = "1.25V" }));
        Assert.Equal("50us, 0.75V", ComponentPaneWording.RunSettings(new ComponentImageEntry { TimeDiv = " 50us ", VoltsDiv = " ", TriggerLevelVolts = "0.75V" }));
        Assert.Equal(string.Empty, ComponentPaneWording.RunSettings(new ComponentImageEntry()));
    }

    // ###########################################################################################
    // Under a picture seen large, its settings named, the VALUES bold (owner requests, 2026-11-20: "have the headers show also", then
    // "again do remember to have the values "bold"") - a note with formatting, so a "*" or "__" in a value stays a character, never
    // a mark; a value not set says so, not bold.
    // ###########################################################################################
    [Fact]
    public void The_settings_under_a_picture_are_named_their_values_bold_and_a_mark_in_one_stays_a_character()
    {
        string line = ComponentPaneWording.RunSettingsNamed(new ComponentImageEntry { TimeDiv = "100us", VoltsDiv = "2**V", TriggerLevelVolts = " " });
        RichNoteText note = RichNote.Parse(line);

        Assert.Equal("T/DIV: 100us   |   V/DIV: 2**V   |   Trigger level: (none)", note.Text);

        string Styled(RichNoteStyle style) =>
            new(note.Text.Where((_, index) => note.Styles[index].HasFlag(style)).ToArray());

        Assert.Equal("100us2**V", Styled(RichNoteStyle.Bold));
    }

    // Closing with captures not yet accepted asks first - they are in no draft - the count bold, one or several.
    [Fact]
    public void Closing_with_captures_not_yet_accepted_says_they_are_in_no_draft()
    {
        Assert.Equal(
            "1 capture is not yet accepted, so it is not in your draft. Accept it before closing, or discard it?",
            Handlers.Theming.BoldRuns.Plain(ComponentPaneWording.RunNotAcceptedOnClose(1)));
        Assert.StartsWith("<b>3</b> captures are not yet accepted, so they are not in your draft.", ComponentPaneWording.RunNotAcceptedOnClose(3), StringComparison.Ordinal);
    }
}
