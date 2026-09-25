using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Handlers.DataHandling;
using Handlers.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace CRT;

// ###########################################################################################
// Test seams for the label editor.
//
// The editor is driven entirely by private methods reached through pointer and key handlers, so
// a headless test had no way in and TabSchematics.LabelEditor.cs sat at ~1% coverage. Rather
// than widen those methods, this part exposes a narrow set of `...ForTests` members that call
// straight through to them - the same convention TabWorkbooks already uses
// (SelectWorkbookForTests, ActivateWorkbookOverrideForTests, CurrentBoardDataOverrideForTests).
//
// Two things make this a seam rather than a back door:
//
//  - Every member here delegates to the SHIPPED method. Nothing reimplements editor behaviour,
//    so a test cannot pass against logic the application does not run.
//  - CurrentBoardDataOverrideForTests exists because BeginLabelEditorMode reads its working copy
//    from MainWindow.CurrentBoardData, and no test constructs Main. It mirrors the override of
//    the same name on TabWorkbooks exactly.
//
// The label editor's own pointer plumbing (hit-testing a handle, converting container points to
// bitmap pixels) needs a laid-out control and a loaded bitmap, neither of which exists headlessly.
// So these seams enter at the level BELOW that - StartDrag/UpdateDrag/CompleteDrag in bitmap-pixel
// space - which is exactly what the real handlers call once they have converted the pointer
// position. What is NOT covered here is that conversion itself; that needs the app.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Stands in for MainWindow.CurrentBoardData, which no test can supply because no test builds
    // Main. Same shape and same reasoning as TabWorkbooks.CurrentBoardDataOverrideForTests.
    internal BoardData? CurrentBoardDataOverrideForTests { get; set; }

    private BoardData? CurrentBoardDataForLabelEditor =>
        this.CurrentBoardDataOverrideForTests ?? this.MainWindow?.CurrentBoardData;

    internal bool IsLabelEditorModeForTests => this.thisIsLabelEditorMode;

    internal string LabelEditorSchematicNameForTests => this.thisLabelEditorSchematicName;

    // The editor's working copy: the rows it is editing, before anything is written back.
    // Exposed as rectangles plus labels so a test can assert on geometry without reaching into
    // EditableComponentHighlight itself.
    internal IReadOnlyList<(string BoardLabel, Rect Rect)> LabelEditorWorkingRowsForTests =>
        this.thisLabelEditorWorkingHighlights
            .Select(row => (row.BoardLabel, new Rect(row.X, row.Y, row.Width, row.Height)))
            .ToList();

    internal int SelectedLabelEditorCountForTests => this.thisSelectedLabelEditorHighlights.Count;

    internal void BeginLabelEditorModeForTests() => this.BeginLabelEditorMode();

    internal void CancelLabelEditorChangesForTests() => this.CancelLabelEditorChanges();

    // Selection, by index into the working rows.
    internal void SelectLabelEditorRowForTests(int workingIndex)
    {
        this.SetSingleSelectedLabelEditorHighlight(this.thisLabelEditorWorkingHighlights[workingIndex]);
    }

    internal void ToggleLabelEditorRowForTests(int workingIndex)
    {
        this.ToggleSelectedLabelEditorHighlight(this.thisLabelEditorWorkingHighlights[workingIndex]);
    }

    internal void ClearLabelEditorSelectionForTests() => this.ClearSelectedLabelEditorHighlights();

    internal bool IsLabelEditorRowSelectedForTests(int workingIndex) =>
        this.IsSelectedLabelEditorHighlight(this.thisLabelEditorWorkingHighlights[workingIndex]);

    // A drag, in BITMAP PIXELS - the space the real pointer handlers convert to before calling
    // these same three methods. dragMode picks move vs. which edge/corner is being resized.
    internal void StartLabelEditorDragForTests(int workingIndex, Point startPixelPoint, LabelEditorDragMode dragMode)
    {
        this.StartLabelEditorDrag(workingIndex, startPixelPoint, dragMode);
    }

    internal void UpdateLabelEditorDragForTests(Point currentPixelPoint, KeyModifiers modifiers = KeyModifiers.None)
    {
        this.UpdateLabelEditorDrag(currentPixelPoint, modifiers);
    }

    internal void CompleteLabelEditorDragForTests() => this.CompleteLabelEditorDrag();

    internal void DeleteLabelEditorRowForTests(int workingIndex) => this.DeleteLabelEditorHighlight(workingIndex);

    internal bool TryUndoLabelEditorChangeForTests() => this.TryUndoLabelEditorChange();

    internal bool TryRedoLabelEditorChangeForTests() => this.TryRedoLabelEditorChange();

    // The keyboard nudge/resize path, which is a different entry point from a pointer drag.
    internal bool ApplySelectedLabelEditorKeyboardStepForTests(Key key, KeyModifiers modifiers) =>
        this.ApplySelectedLabelEditorKeyboardStep(key, modifiers);

    // The save/validation error dialog, BUILT but not shown.
    //
    // The two shipped entry points end in ShowDialog, which blocks on a real window with nothing
    // to dismiss it headlessly - so a test driving them would hang instead of failing. This seam
    // returns the constructed Window instead, which is the same object the app shows: the shipped
    // method does nothing to it between building and showing.
    //
    // Worth a seam because these two dialogs were 48 identical lines until they were merged behind
    // one builder, and the four things that differ between them (title, headline, fallback text,
    // and whether a hint paragraph exists at all) are now arguments that a caller can get wrong
    // silently - a swapped pair still compiles and still shows a plausible-looking dialog.
    // Each delegates to the SAME private builder call the shipped dialog uses - deliberately not a
    // copy of its arguments here, which would be the very duplication this refactor removed and
    // would let the test pass against strings the application does not use.
    internal Window BuildLabelEditorSaveFailedDialogForTests(string errorMessage) =>
        this.BuildLabelEditorSaveFailedDialog(errorMessage);

    internal Window BuildLabelEditorValidationFailedDialogForTests(string errorMessage) =>
        this.BuildLabelEditorValidationFailedDialog(errorMessage);
}
