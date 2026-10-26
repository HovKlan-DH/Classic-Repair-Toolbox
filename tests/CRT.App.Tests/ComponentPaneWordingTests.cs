using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPaneWording - what the component pane says (CRT 4.0 step 2, 2026-10-25): its title,
// the region line, the offer to add the case's pins, what a scope baseline still needs, the new
// component's form and the whole component's buttons (slice 3c), and the board's own lists (4a).
// ###########################################################################################
public sealed class ComponentPaneWordingTests
{
    [Fact]
    public void The_title_is_the_label_and_what_the_component_list_shows_beside_it()
    {
        Assert.Equal("U8 - CPU | 6510", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510" }));
        Assert.Equal("C12 - 10uF", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "C12", TechnicalNameOrValue = " 10uF " }));
        Assert.Equal("R1", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "R1" }));
    }

    // Said only for a component with a row of its own per region.
    [Fact]
    public void The_region_line_names_the_region_of_a_regional_row_only()
    {
        Assert.Equal(string.Empty, ComponentPaneWording.RegionLine(""));
        Assert.Equal(string.Empty, ComponentPaneWording.RegionLine(null));
        Assert.StartsWith("The PAL version", ComponentPaneWording.RegionLine(" PAL "), StringComparison.Ordinal);
    }

    // A section's heading says how many it holds, as the plan draws the pane.
    [Fact]
    public void A_sections_heading_counts_what_it_holds()
    {
        Assert.Equal("Scope baselines (3)", ComponentPaneWording.Heading("Scope baselines", 3));
        Assert.Equal("Pictures (0)", ComponentPaneWording.Heading("Pictures", 0));
    }

    // ###########################################################################################
    // WHAT A SCOPE BASELINE STILL NEEDS (slice 3b): its pin and one scope setting - the one rule that
    // makes a row a baseline and not a picture (ComponentImageKinds) - and a screenshot (the plan: "a
    // screenshot, which is required"). Nothing when it has all three.
    // ###########################################################################################
    [Fact]
    public void What_a_scope_baseline_still_needs_is_said_in_plain_words()
    {
        Assert.Null(ComponentPaneWording.BaselineProblem("1", "50us", "", "", "Scope baseline/U8_1.png"));
        Assert.Null(ComponentPaneWording.BaselineProblem("1", "", "", "1.5V", "Scope baseline/U8_1.png"));

        Assert.Equal("A scope baseline needs its pin.", ComponentPaneWording.BaselineProblem(" ", "50us", "1V", "1.5V", "x.png"));
        Assert.Equal("A scope baseline needs one of T/DIV, V/DIV or trigger level.", ComponentPaneWording.BaselineProblem("1", "", " ", "", "x.png"));
        Assert.Equal(
            "A scope baseline needs its pin, a screenshot and one of T/DIV, V/DIV or trigger level.",
            ComponentPaneWording.BaselineProblem("", "", "", "", ""));
        Assert.Equal("A scope baseline needs its pin and a screenshot.", ComponentPaneWording.BaselineProblem("", "1us", "", "", null));

        Assert.Equal(
            "A scope baseline needs its pin. Until it has, your draft keeps it as it was.",
            ComponentPaneWording.BaselineHeldBack("A scope baseline needs its pin."));
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

    // The region a new component is in: first the one most components are in - every region, blank.
    [Fact]
    public void A_new_components_region_choices_stand_for_the_regions_the_list_knows()
    {
        Assert.Equal(["PAL and NTSC", "PAL only", "NTSC only"], ComponentPaneWording.RegionChoices);
        Assert.Equal(string.Empty, ComponentPaneWording.RegionOfChoice("PAL and NTSC"));
        Assert.Equal("PAL", ComponentPaneWording.RegionOfChoice("PAL only"));
        Assert.Equal("NTSC", ComponentPaneWording.RegionOfChoice("NTSC only"));
        Assert.Equal(string.Empty, ComponentPaneWording.RegionOfChoice(null));
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
        Assert.Equal("[U8] is not marked on any picture yet, so there is nothing to show", ComponentPaneWording.ShowOnBoardTip("U8", marked: false));
        Assert.Contains("every region", ComponentPaneWording.DeleteTip("U8"), StringComparison.Ordinal);
        Assert.Contains("Undo brings it back", ComponentPaneWording.DeleteTip("U8"), StringComparison.Ordinal);
        Assert.Equal("Bring [U8] back", ComponentPaneWording.BringBack("U8"));
    }

    // ---- the board's own lists (slice 4a) -----------------------------------------------------------

    [Fact]
    public void A_board_part_is_named_and_a_net_not_found_says_it_will_not_light_up()
    {
        Assert.Equal("Files and links", ComponentPaneWording.Title(PaneBoardPart.FilesAndLinks));
        Assert.Equal("Credits", ComponentPaneWording.Title(PaneBoardPart.Credits));
        Assert.Equal("Important signals", ComponentPaneWording.Title(PaneBoardPart.ImportantSignals));
        Assert.Equal(
            "[NRESET] is no net in the board's KiCad data - the Schematics tab will not light this signal up.",
            ComponentPaneWording.NetNotFound(" NRESET "));
    }

    // What "Also on other boards" did: the boards changed, then each that was not, with why.
    [Fact]
    public void Also_on_other_boards_says_what_it_changed_and_what_not()
    {
        Assert.Equal(
            "Changed on 2 boards, each in its own draft: A, B. Not changed on C - its draft could not be saved.",
            ComponentPaneWording.AlsoDone(["A", "B"], [("C", ComponentPaneWording.AlsoNotSaved)]));
        Assert.Equal("Changed on 1 board, in its own draft: A.", ComponentPaneWording.AlsoDone(["A"], []));
        Assert.Equal("Not changed on C - it is like that there already.", ComponentPaneWording.AlsoDone([], [("C", ComponentPaneWording.AlsoAlreadySo)]));
    }

    // Every sentence the pane says is plain ASCII and names neither Excel nor the old editor window.
    [Fact]
    public void The_new_sentences_are_plain_ASCII()
    {
        string[] sentences =
        [
            ComponentPaneWording.NewComponentTitle,
            ComponentPaneWording.NoCategoryYet,
            ComponentPaneWording.NoKiCadNets,
            ComponentPaneWording.ReadingKiCadNets,
            ComponentPaneWording.ComponentChangedTip,
            ComponentPaneWording.BoardPartChangedTip,
            ComponentPaneWording.DeleteTip("U8"),
            ComponentPaneWording.NetNotFound("X"),
        ];

        Assert.All(sentences, sentence =>
        {
            Assert.All(sentence, character => Assert.True(character < 128, $"[{sentence}] is not plain ASCII"));
            Assert.DoesNotContain("Excel", sentence, StringComparison.Ordinal);
            Assert.DoesNotContain("component editor", sentence, StringComparison.Ordinal);
        });
    }
}
