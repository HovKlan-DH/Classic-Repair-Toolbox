using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT;

// ###########################################################################################
// TabSchematics, part: THE VIEW'S SMALL TOOLS (CRT 4.0 step 5 - the plan's "Smaller items"):
//
//   - PANNING with the right OR the middle button, one helper for every mode (TryBeginPan) - and a
//     ONE-TIME TIP saying so in the view's corner ("a middle-button drag pans too, with a one-time
//     hint in the corner"), put away by "Got it" or by the first pan, and never shown again
//     (UserSettings.SchematicsPanHintSeen). A middle click that does not move does nothing more:
//     only a right click opens the component menu.
//   - "DRAW TRACE" ("Traces are drawn only with Draw trace"): traces are drawn and changed ONLY
//     while it is on. Off, a press on the picture never starts, bends or deletes a trace - a left
//     press selects what is under it, a right drag over a trace pans - and the traces are only
//     drawn. Escape, another tool, or another board turns it off; while another tool is on it is
//     greyed out, its tooltip saying why (FollowToolsInDrawTrace). "Done drawing" is a plain button
//     - it turns a mode off and saves nothing, so the Accept colour is not its (code review,
//     2026-10-29).
//   - "Global settings" behind a gear button: the markup's flyout, nothing here.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the file map.
// ###########################################################################################
public partial class TabSchematics
{
    // The one-time tip, and what the mode hint says while traces are drawn.
    internal const string PanHint = "Drag with the right or middle mouse button to move the picture, and scroll to zoom.";

    internal const string DrawTraceHint =
        "Drag on the picture to draw a trace, or drag a trace's points. Right-click a trace to remove it. Escape stops.";

    private const string DrawTraceLabel = "Draw trace";

    private const string StopDrawingLabel = "Done drawing";

    internal const string DrawTraceNotNow = "Finish what you are doing on the picture first - then you can draw traces.";

    // "Draw trace" on.
    private bool thisIsDrawingTraces;

    // The pan under way was started by the middle button - its release is no click.
    private bool thisPanIsMiddleButton;

    // The picture whose traces the trace manager holds - null while none were read for the picture
    // chosen (it is still being read, or has no file).
    private string? thisTracesReadFor;

    internal bool IsDrawingTraces => this.thisIsDrawingTraces;

    // ###########################################################################################
    // *** A PICTURE'S TRACES ARE ITS OWN (code review, 2026-10-29). *** While the next picture was
    // read, the last one's traces stayed loaded and editable under its name: a point dragged then -
    // or a colour ticked on a picture with no file - saved them over the new picture's own. So a
    // switch lets them go at once (LetGoOfTraces), they are read with the picture (ReadTracesFor),
    // and they are changed and saved only while they are the chosen picture's (TracesAreThePictures).
    // ###########################################################################################
    private void LetGoOfTraces()
    {
        this.polylineManager?.StopEditing(saveDrag: false);
        this.polylineManager?.Reset();
        this.thisTracesReadFor = null;
    }

    private void ReadTracesFor(string? boardKey, string pictureName)
    {
        if (this.polylineManager == null || string.IsNullOrEmpty(boardKey))
        {
            return;
        }

        this.polylineManager.ImportTraces(TraceStorage.GetTraces(boardKey, pictureName));
        this.thisTracesReadFor = pictureName;
    }

    // Whether the traces held are the chosen picture's - with none chosen, or none read, they are not
    // (code review, 2026-10-29: "none chosen" and "none read" were taken for the same picture).
    private bool TracesAreThePictures =>
        this.thisTracesReadFor is string read &&
        string.Equals(read, this.SelectedThumbnail?.Name, StringComparison.Ordinal);

    // ###########################################################################################
    // Whether a press on the picture landed on the GREYED-OUT "Draw trace" - disabled, a button takes
    // no pointer, so the press went through to the picture under it: in the label editor it let go of
    // the rectangles selected (code review, 2026-10-29). OnSchematicsPointerPressed leaves such a press
    // alone. An enabled button is left to itself: it takes its own click, and a right or middle press
    // on it still pans the picture, as it always did.
    // ###########################################################################################
    private bool IsPointerOnDrawTraceButton(Point containerPoint)
    {
        if (!this.DrawTraceButton.IsVisible || this.DrawTraceButton.IsEnabled)
        {
            return false;
        }

        Point? topLeft = this.DrawTraceButton.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
        return topLeft.HasValue && new Rect(topLeft.Value, this.DrawTraceButton.Bounds.Size).Contains(containerPoint);
    }

    internal void ResetTracesForTests() => this.LetGoOfTraces();

    // Traces drawn and changed: "Draw trace" on, over the chosen picture's own traces.
    private bool CanEditTraces => this.thisIsDrawingTraces && this.TracesAreThePictures;

    // The import a picture's load does when its picture is read - headless tests decode none.
    internal void ImportTracesForTests(string pictureName) => this.ReadTracesFor(this.MainWindow?.GetCurrentBoardKey(), pictureName);

    // ###########################################################################################
    // A pan for a press of the right or the middle button - true when one started. Every mode pans
    // the same way; the tip is put away by the first pan.
    // ###########################################################################################
    private bool TryBeginPan(PointerPressedEventArgs e, PointerPoint pointer, Point point)
    {
        if (!pointer.Properties.IsRightButtonPressed && !pointer.Properties.IsMiddleButtonPressed)
        {
            return false;
        }

        this.isPanning = true;
        this.thisPanIsMiddleButton = !pointer.Properties.IsRightButtonPressed;
        this.panStartPoint = point;
        this.panStartMatrix = this.schematicsMatrix;

        this.HideSchematicsHoverUi();
        this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);

        e.Pointer.Capture(this.SchematicsContainer);
        e.Handled = true;

        this.PutPanHintAway();
        return true;
    }

    // ###########################################################################################
    // The picture shown, or none: "Draw trace" offered with one, and the tip shown with the first
    // one until it has been put away.
    // ###########################################################################################
    internal void ShowViewToolsForPicture(bool hasPicture)
    {
        if (!hasPicture)
        {
            this.HideViewToolsForNoPicture();
            return;
        }

        this.DrawTraceButton.IsVisible = true;

        if (!UserSettings.SchematicsPanHintSeen)
        {
            this.PanHintText.Text = PanHint;
            this.PanHintBorder.IsVisible = true;
        }
    }

    // No picture: no tools - and no drawing (code review, 2026-10-29: the button went, drawing stayed
    // on, and a drag changed the previous picture's traces).
    private void HideViewToolsForNoPicture()
    {
        this.SetDrawingTraces(false, navigating: true);
        this.DrawTraceButton.IsVisible = false;
        this.PanHintBorder.IsVisible = false;
    }

    // The tip put away for good - by "Got it", or by the first pan while it shows.
    private void PutPanHintAway()
    {
        if (!this.PanHintBorder.IsVisible)
        {
            return;
        }

        this.PanHintBorder.IsVisible = false;
        UserSettings.SchematicsPanHintSeen = true;
    }

    private void OnPanHintGotItClick(object? sender, RoutedEventArgs e) => this.PutPanHintAway();

    // A press on the tip is the tip's - never a click on the picture under it.
    private void OnPanHintPointerPressed(object? sender, PointerPressedEventArgs e) => e.Handled = true;

    private void OnDrawTraceClick(object? sender, RoutedEventArgs e) => this.SetDrawingTraces(!this.thisIsDrawingTraces);

    // ###########################################################################################
    // "Draw trace" on or off. On only over a picture and with no other tool on (the label editor,
    // a worklog area, aligning KiCad) - each of those turns it off as it starts. Off, whatever was
    // being drawn or dragged is let go (PolylineManagement.StopEditing).
    // ###########################################################################################
    // `navigating` - another board or picture: a point being dragged is dropped, not saved (it would
    // be saved under the new one's name).
    internal void SetDrawingTraces(bool on, bool navigating = false)
    {
        if (on && (this.currentFullResBitmap == null || this.thisIsLabelEditorMode || this.thisIsWorklogEntryMode || this.thisIsKiCadTraceCalibrationMode))
        {
            on = false;
        }

        if (this.thisIsDrawingTraces == on)
        {
            return;
        }

        this.thisIsDrawingTraces = on;
        this.DrawTraceButton.Content = on ? StopDrawingLabel : DrawTraceLabel;

        if (on)
        {
            this.MainWindow?.ShowModeHint(DrawTraceHint);
            this.SchematicsContainer.Focus();
            return;
        }

        this.polylineManager?.StopEditing(saveDrag: !navigating);
        this.MainWindow?.HideModeHint();
    }

    // ###########################################################################################
    // Another tool on (the label editor, a worklog area, aligning KiCad): "Draw trace" greyed out,
    // its tooltip saying why (code review, 2026-10-29: clickable, it did nothing and said nothing).
    // Called wherever one of them starts or ends - FollowPictureTool for the label editor and the
    // calibration, the worklog area's own begin and cancel.
    // ###########################################################################################
    private void FollowToolsInDrawTrace()
    {
        bool anotherTool = this.thisIsLabelEditorMode || this.thisIsWorklogEntryMode || this.thisIsKiCadTraceCalibrationMode;

        this.DrawTraceButton.IsEnabled = !anotherTool;
        ToolTip.SetTip(this.DrawTraceButton, anotherTool ? DrawTraceNotNow : null);
        ToolTip.SetShowOnDisabled(this.DrawTraceButton, true);
    }
}
