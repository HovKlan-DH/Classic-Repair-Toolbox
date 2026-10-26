using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT;

// ###########################################################################################
// THE COMPONENTS NOT MARKED ON THIS PICTURE YET, BESIDE THE LABEL EDITOR (CRT 4.0 step 3, slice
// 3c - the plan: "Photos and scanned schematics use the label editor as today, with the unmarked
// components listed beside it").
//
// While the label editor is on, the picture tool's bar lists the board's components with no
// rectangle on the picture in the working copy (UnmarkedComponents) - each a button; in the bar, as
// the editor only starts from "Mark components", never over the picture's corners, which hold the
// settings and the traces (a render showed it laid over Global settings, 2026-10-26). One
// picked stays picked until a rectangle is drawn: that rectangle is named after it with its
// category, and no prompt asks (AddDrawnLabelEditorRectangle); picked again, it is let go. The list
// follows every change of the working copy - an undo puts a component back in it - and is changed
// IN PLACE (InPlaceList.Sync), never rebuilt under a press.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // The component picked in the list - the next rectangle drawn is its.
    private UnmarkedComponent? thisPickedUnmarkedComponent;

    internal IReadOnlyList<string> UnmarkedLabelsForTests =>
        this.UnmarkedButtonsPanel.Children.OfType<ToggleButton>().Select(button => (string)button.Tag!).ToList();

    internal string? PickedUnmarkedLabelForTests => this.thisPickedUnmarkedComponent?.Label;

    // A component picked in the list, as its button does.
    internal void PickUnmarkedComponentForTests(string label) =>
        this.PickUnmarkedComponent(this.UnmarkedButtonsPanel.Children.OfType<ToggleButton>().Single(button => Equals(button.Tag, label)));

    // A rectangle drawn, in bitmap pixels - what the pointer handlers hand over once converted.
    internal void AddDrawnLabelEditorRectangleForTests(Rect rect) => this.AddDrawnLabelEditorRectangle(rect, default);

    // ###########################################################################################
    // The list shown, worked out from the working copy - hidden with the label editor off.
    // ###########################################################################################
    private void ShowUnmarkedComponents()
    {
        if (!this.thisIsLabelEditorMode || this.CurrentBoardDataForLabelEditor is not { } board)
        {
            this.thisPickedUnmarkedComponent = null;
            this.UnmarkedPanelBorder.IsVisible = false;
            return;
        }

        string schematicName = this.GetCurrentSchematicName();

        List<UnmarkedComponent> unmarked = UnmarkedComponents.Of(
            board.Components,
            this.thisLabelEditorWorkingHighlights
                .Where(row => string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                .Select(row => row.BoardLabel));

        // A picked component marked meanwhile (typed into the prompt, say) is let go.
        if (this.thisPickedUnmarkedComponent is { } picked && !unmarked.Any(row => string.Equals(row.Label, picked.Label, StringComparison.OrdinalIgnoreCase)))
        {
            this.thisPickedUnmarkedComponent = null;
        }

        InPlaceList.Sync(
            this.UnmarkedButtonsPanel.Children,
            unmarked,
            (control, component) => control is ToggleButton { Tag: string label } && string.Equals(label, component.Label, StringComparison.Ordinal),
            this.UnmarkedComponentButton);

        foreach (ToggleButton button in this.UnmarkedButtonsPanel.Children.OfType<ToggleButton>())
        {
            button.IsChecked = this.thisPickedUnmarkedComponent is { } chosen && Equals(button.Tag, chosen.Label);
        }

        this.UnmarkedHeadingText.Text = PictureToolWording.UnmarkedHeading(unmarked.Count);
        this.UnmarkedHelpText.Text = unmarked.Count == 0 ? PictureToolWording.AllMarked : PictureToolWording.UnmarkedHelp;
        this.UnmarkedPanelBorder.IsVisible = true;
    }

    private Control UnmarkedComponentButton(UnmarkedComponent component)
    {
        var button = new ToggleButton
        {
            Content = component.Label,
            Tag = component.Label,
            Margin = new Thickness(0, 0, 6, 4),
            Focusable = false,
        };

        ToolTip.SetTip(button, $"Draw the rectangle of [{component.Label}] next");
        button.Click += (_, _) => this.PickUnmarkedComponent(button);
        return button;
    }

    // A component's button pressed: picked - or let go, when it was.
    private void PickUnmarkedComponent(ToggleButton button)
    {
        string label = (string)button.Tag!;

        this.thisPickedUnmarkedComponent = this.thisPickedUnmarkedComponent is { } picked && string.Equals(picked.Label, label, StringComparison.Ordinal)
            ? null
            : UnmarkedComponents.Of(this.CurrentBoardDataForLabelEditor?.Components ?? [], [])
                .FirstOrDefault(row => string.Equals(row.Label, label, StringComparison.Ordinal));

        this.ShowUnmarkedComponents();
        this.SchematicsContainer.Focus();
    }

    // The picked component, taken for the rectangle being drawn - none is picked after.
    private UnmarkedComponent? TakePickedUnmarkedComponent()
    {
        UnmarkedComponent? picked = this.thisPickedUnmarkedComponent;
        this.thisPickedUnmarkedComponent = null;
        return picked;
    }
}
