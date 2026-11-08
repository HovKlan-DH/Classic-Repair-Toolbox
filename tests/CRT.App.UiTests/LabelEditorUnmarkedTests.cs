using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.LogicalTree;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE COMPONENTS NOT MARKED ON THE PICTURE, BESIDE THE LABEL EDITOR (TabSchematics.LabelEditor
// .Unmarked.cs - CRT 4.0 step 3, slice 3c: "Photos and scanned schematics use the label editor as
// today, with the unmarked components listed beside it").
//
// While the editor is on, the board's components with no rectangle on the picture are listed; one
// picked names the next rectangle drawn - no prompt - and leaves the list; an undo puts it back. The
// list is changed in place, never rebuilt under a press. Driven through the seams in bitmap pixels,
// as LabelEditorInteractionTests are.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class LabelEditorUnmarkedTests
{
    private const string Schematic = "Board top";

    private static TabSchematics CreateTab()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "u2", Category = "IC", Region = "NTSC" },
                new ComponentEntry { BoardLabel = "C3", Category = "Capacitor" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC" },
            ],
            ComponentHighlights =
            [
                new ComponentHighlightEntry { SchematicName = Schematic, BoardLabel = "U1", X = "100", Y = "100", Width = "50", Height = "40" },
                new ComponentHighlightEntry { SchematicName = "Board bottom", BoardLabel = "U9", X = "10", Y = "10", Width = "10", Height = "10" },
            ],
        };

        return new TabSchematics
        {
            CurrentBoardDataOverrideForTests = board,
            SelectedThumbnail = new SchematicThumbnail { Name = Schematic },
        };
    }

    [Fact]
    public void The_editor_lists_the_components_not_marked_on_the_picture()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            Assert.False(tab.GetControl<Border>("UnmarkedPanelBorder").IsVisible);

            tab.BeginLabelEditorModeForTests();

            Assert.True(tab.GetControl<Border>("UnmarkedPanelBorder").IsVisible);
            Assert.Equal(["U2", "C3", "U9"], tab.UnmarkedLabelsForTests);
            Assert.Equal(PictureToolWording.UnmarkedHeading(3), tab.GetControl<TextBlock>("UnmarkedHeadingText").Text);

            tab.CancelLabelEditorChangesForTests();
            Assert.False(tab.GetControl<Border>("UnmarkedPanelBorder").IsVisible);
        });
    }

    // ###########################################################################################
    // The list is in the picture tool's bar - the label editor only ever starts from "Mark
    // components" - never over the picture's corners: a render showed it laid over the Global
    // settings panel at the bottom-left (2026-10-26), and the bottom-right holds the traces.
    // ###########################################################################################
    [Fact]
    public void The_list_is_in_the_tool_bar_clear_of_the_pictures_corners()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            Border panel = tab.GetControl<Border>("UnmarkedPanelBorder");

            Assert.Contains(tab.GetControl<Border>("PictureToolBarBorder"), panel.GetLogicalAncestors());
        });
    }

    // ###########################################################################################
    // A component picked names the next rectangle drawn - its label and category, no prompt - and
    // leaves the list; the next rectangle, nothing picked, asks again. An undo puts it back.
    // ###########################################################################################
    [Fact]
    public void A_picked_component_names_the_next_rectangle_and_leaves_the_list()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            ToggleButton c3Button = tab.GetControl<WrapPanel>("UnmarkedButtonsPanel").Children.OfType<ToggleButton>().Single(button => Equals(button.Tag, "C3"));

            tab.PickUnmarkedComponentForTests("U2");
            Assert.Equal("U2", tab.PickedUnmarkedLabelForTests);

            tab.AddDrawnLabelEditorRectangleForTests(new Rect(300, 300, 40, 20));

            Assert.Contains(("U2", new Rect(300, 300, 40, 20)), tab.LabelEditorWorkingRowsForTests);
            Assert.False(tab.GetControl<Border>("SchematicsNewLabelPromptBorder").IsVisible);
            Assert.Null(tab.PickedUnmarkedLabelForTests);
            Assert.Equal(["C3", "U9"], tab.UnmarkedLabelsForTests);

            // The list changed in place - C3's button is the very one.
            Assert.Same(c3Button, tab.GetControl<WrapPanel>("UnmarkedButtonsPanel").Children.OfType<ToggleButton>().Single(button => Equals(button.Tag, "C3")));

            // Nothing picked: the prompt asks as it always did.
            tab.AddDrawnLabelEditorRectangleForTests(new Rect(500, 500, 40, 20));
            Assert.True(tab.GetControl<Border>("SchematicsNewLabelPromptBorder").IsVisible);
        });
    }

    // ###########################################################################################
    // A component the board HAS needs no category from the editor - the writer leaves an existing
    // component alone (code review, 2026-10-26: one with no category, picked and drawn - or marked
    // automatically, then opened - had Done refused for "missing a category", with no way to give
    // its rectangle one). A label new to the board still needs its category.
    // ###########################################################################################
    [Fact]
    public void A_component_the_board_has_needs_no_category_to_be_marked()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.CurrentBoardDataOverrideForTests!.Components.Add(new ComponentEntry { BoardLabel = "J7" });
            tab.CurrentBoardDataOverrideForTests.ComponentHighlights.Add(
                new ComponentHighlightEntry { SchematicName = Schematic, BoardLabel = "J7", X = "200", Y = "200", Width = "20", Height = "20" });
            tab.CurrentBoardDataOverrideForTests.Components.Add(new ComponentEntry { BoardLabel = "J8" });
            tab.BeginLabelEditorModeForTests();

            tab.PickUnmarkedComponentForTests("J8");
            tab.AddDrawnLabelEditorRectangleForTests(new Rect(300, 300, 40, 20));

            Assert.True(tab.TryValidateLabelEditorSaveForTests(out string error), error);
        });
    }

    [Fact]
    public void An_undo_puts_a_component_back_and_picking_it_again_lets_it_go()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.PickUnmarkedComponentForTests("C3");
            tab.AddDrawnLabelEditorRectangleForTests(new Rect(300, 300, 40, 20));
            Assert.Equal(["U2", "U9"], tab.UnmarkedLabelsForTests);

            Assert.True(tab.TryUndoLabelEditorChangeForTests());
            Assert.Equal(["U2", "C3", "U9"], tab.UnmarkedLabelsForTests);

            tab.PickUnmarkedComponentForTests("U9");
            Assert.True(tab.GetControl<WrapPanel>("UnmarkedButtonsPanel").Children.OfType<ToggleButton>().Single(button => Equals(button.Tag, "U9")).IsChecked);

            tab.PickUnmarkedComponentForTests("U9");
            Assert.Null(tab.PickedUnmarkedLabelForTests);
            Assert.False(tab.GetControl<WrapPanel>("UnmarkedButtonsPanel").Children.OfType<ToggleButton>().Single(button => Equals(button.Tag, "U9")).IsChecked);
        });
    }
}
