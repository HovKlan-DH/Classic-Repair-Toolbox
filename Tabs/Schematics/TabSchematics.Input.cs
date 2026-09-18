using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.GestureRecognizers;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// Raw pointer, wheel, gesture and keyboard event handlers for the schematics surface.
// These only dispatch - the actual behaviour lives in the Viewport, LabelEditor,
// KiCad and Highlights parts.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // ###########################################################################################
    // Handles mouse wheel zoom on the Schematics image, centered on the cursor position.
    // Wheel input over interactive overlay panels is consumed there and must not zoom the
    // schematic viewer underneath, even when a nested scroll viewer reaches its top or bottom.
    // ###########################################################################################
    private void OnSchematicsZoom(object? sender, PointerWheelEventArgs e)
    {
        var zoomCenterInContainer = e.GetPosition(this.SchematicsContainer);

        if (this.IsPointerInsideInteractiveOverlayPanel(zoomCenterInContainer))
        {
            e.Handled = true;
            return;
        }

        double zoomFactor = ViewportMath.ComputeWheelZoomFactor(e.Delta.Y, AppConfig.SchematicsZoomFactor);

        this.ApplySchematicsZoom(zoomFactor, zoomCenterInContainer);

        e.Handled = true;
    }

    // ###########################################################################################
    // Handles trackpad pinch zoom. macOS trackpad pinch does not reliably come through as a mouse
    // wheel event, so this explicit gesture path is needed.
    // ###########################################################################################
    private void OnSchematicsPinch(object? sender, PinchEventArgs e)
    {
        if (this.currentFullResBitmap == null)
        {
            return;
        }

        Point zoomCenterInContainer = new(
            this.SchematicsContainer.Bounds.Width / 2.0,
            this.SchematicsContainer.Bounds.Height / 2.0);

        this.ApplySchematicsZoom(e.Scale, zoomCenterInContainer);

        e.Handled = true;
    }

    // ###########################################################################################
    // Handles two-finger trackpad pan gestures independently of right-mouse panning.
    // Uses strict edge clamping so manual pan cannot drag the image below or beyond the viewport.
    // If the direction feels reversed on a specific platform, flip the signs below.
    // ###########################################################################################
    private void OnSchematicsScrollGesture(object? sender, ScrollGestureEventArgs e)
    {
        if (this.currentFullResBitmap == null)
        {
            return;
        }

        if (this.isPanning)
        {
            return;
        }

        Vector delta = e.Delta;

        if (Math.Abs(delta.X) < 0.001 && Math.Abs(delta.Y) < 0.001)
        {
            return;
        }

        this.schematicsMatrix = this.schematicsMatrix * Matrix.CreateTranslation(delta.X, delta.Y);
//hest        this.schematicsMatrix = this.schematicsMatrix * Matrix.CreateTranslation(-delta.X, -delta.Y); // replace above line if two-finger pan feels inverted on macOS
        this.ClampSchematicsMatrix(useStrictEdgeClamp: true);

        e.Handled = true;
    }

    // ###########################################################################################
    // Handles right-click for panning on the schematic view and selection toggling on release.
    // Left-click selects hovered component, single-click opens component info popup, and while the
    // new KiCad trace calibration mode is active the same pointer pipeline is reused for moving and
    // resizing the temporary calibration box.
    // ###########################################################################################
    private void OnSchematicsPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        var point = e.GetPosition(this.SchematicsContainer);
        var pointer = e.GetCurrentPoint(this.SchematicsContainer);

        this.UpdateInteractiveCadTraceHoverShiftState(e.KeyModifiers);

        if (this.IsPointerInsideKiCadNetConnectionsPanel(point))
        {
            this.ClearTransientHoverForKiCadNetConnectionsPanel();
            return;
        }

        if (this.IsPointerInsideLabelEditorMenu(point) || this.IsPointerInsideNewLabelPrompt(point))
        {
            e.Handled = true;
            return;
        }

        if (this.thisIsLabelEditorMode && this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            e.Handled = true;
            return;
        }

        if (pointer.Properties.IsLeftButtonPressed && this.thisIsShowingLabelEditorMenu)
        {
            this.HideLabelEditorMenu();
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            if (pointer.Properties.IsRightButtonPressed)
            {
                this.isPanning = true;
                this.panStartPoint = point;
                this.panStartMatrix = this.schematicsMatrix;

                this.HideSchematicsHoverUi();
                this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);

                e.Pointer.Capture(this.SchematicsContainer);
                e.Handled = true;
                return;
            }

            if (pointer.Properties.IsLeftButtonPressed)
            {
                if (!this.TryGetSchematicsImagePixelPoint(point, out var pixelPoint))
                {
                    e.Handled = true;
                    return;
                }

                if (this.TryGetKiCadTraceCalibrationHandleAtContainerPoint(point, out var resizeMode))
                {
                    this.StartKiCadTraceCalibrationDrag(pixelPoint, resizeMode);
                    this.UpdateKiCadTraceCalibrationCursor(point);

                    // Captured for the same reason every other drag on this surface captures: without
                    // it the drag only lasts as long as the pointer happens to stay over the handle's
                    // own hit rect. A handle is ~10 px, and the pointer routinely outruns the overlay
                    // redraw, so the box kept letting go mid-drag - reported as the box being hard to
                    // resize, and worse the slower the redraw is.
                    //
                    // Remembered on the tab (rather than only released from the matching release
                    // branch) so ESC or Apply mid-drag can release it too - see
                    // thisKiCadTraceCalibrationCapturedPointer.
                    e.Pointer.Capture(this.SchematicsContainer);
                    this.thisKiCadTraceCalibrationCapturedPointer = e.Pointer;
                    e.Handled = true;
                    return;
                }

                if (this.IsPointerInsideCurrentKiCadCalibrationBounds(point))
                {
                    this.StartKiCadTraceCalibrationDrag(pixelPoint, LabelEditorDragMode.Move);
                    this.UpdateKiCadTraceCalibrationCursor(point);

                    e.Pointer.Capture(this.SchematicsContainer);
                    this.thisKiCadTraceCalibrationCapturedPointer = e.Pointer;
                    e.Handled = true;
                    return;
                }

                e.Handled = true;
                return;
            }
        }

        if (this.thisIsLabelEditorMode)
        {
            if (pointer.Properties.IsRightButtonPressed)
            {
                this.isPanning = true;
                this.panStartPoint = point;
                this.panStartMatrix = this.schematicsMatrix;

                this.HideSchematicsHoverUi();
                this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);

                e.Pointer.Capture(this.SchematicsContainer);
                e.Handled = true;
                return;
            }

            if (pointer.Properties.IsLeftButtonPressed)
            {
                bool isCtrlDown = e.KeyModifiers.HasFlag(KeyModifiers.Control);

                if (!this.TryGetLabelEditorPixelPoint(point, out var pixelPoint))
                {
                    if (!isCtrlDown)
                    {
                        this.ClearSelectedLabelEditorHighlight();
                    }

                    e.Handled = true;
                    return;
                }

                if (isCtrlDown)
                {
                    if (this.TryGetLabelEditorHighlightAtContainerPoint(point, out var toggleIndex))
                    {
                        this.ToggleSelectedLabelEditorHighlight(this.thisLabelEditorWorkingHighlights[toggleIndex]);
                    }

                    e.Handled = true;
                    return;
                }

                if (this.TryGetSelectedLabelEditorHandleAtContainerPoint(point, out var handleIndex, out var resizeMode))
                {
                    this.StartLabelEditorDrag(handleIndex, pixelPoint, resizeMode);
                    e.Handled = true;
                    return;
                }

                if (this.TryGetSelectedLabelEditorHighlightAtContainerPoint(point, out var selectedWorkingIndex))
                {
                    this.StartLabelEditorDrag(selectedWorkingIndex, pixelPoint, LabelEditorDragMode.Move);
                    e.Handled = true;
                    return;
                }

                if (this.TryGetLabelEditorHighlightAtContainerPoint(point, out var workingIndex))
                {
                    this.SetSingleSelectedLabelEditorHighlight(this.thisLabelEditorWorkingHighlights[workingIndex], refresh: false);
                    this.StartLabelEditorDrag(workingIndex, pixelPoint, LabelEditorDragMode.Move);
                    e.Handled = true;
                    return;
                }

                this.ClearSelectedLabelEditorHighlights(refresh: false);
                this.StartDrawingLabelEditorRectangle(pixelPoint);

                e.Handled = true;
                return;
            }
        }

        if (this.thisIsWorklogEntryMode)
        {
            if (pointer.Properties.IsRightButtonPressed)
            {
                this.isPanning = true;
                this.panStartPoint = point;
                this.panStartMatrix = this.schematicsMatrix;

                this.HideSchematicsHoverUi();
                this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);

                e.Pointer.Capture(this.SchematicsContainer);
                e.Handled = true;
                return;
            }

            if (pointer.Properties.IsLeftButtonPressed)
            {
                if (this.TryGetSchematicsImagePixelPoint(point, out var worklogPixelPoint))
                {
                    this.StartDrawingWorklogEntryRectangle(worklogPixelPoint);
                }
            }

            e.Handled = true;
            return;
        }

        // Resizing a marked worklog area takes precedence over panning and component selection:
        // the pointer is on a handle or inside an area the user can see, so a drag there means the
        // area, not the view. Only claims the press when a handle or area was actually hit, so
        // clicking anywhere else behaves exactly as before.
        if (pointer.Properties.IsLeftButtonPressed && this.TryBeginWorklogEntryResize(point))
        {
            this.HideSchematicsHoverUi();
            e.Pointer.Capture(this.SchematicsContainer);
            e.Handled = true;
            return;
        }

        bool hoveringComponent = this.TryGetHoveredBoardLabel(point, out var boardLabel, out var displayText);
        string? activeHoveredKiCadNetName = this.GetActiveHoveredKiCadNetName();
        bool hoveringKiCadNet = !string.IsNullOrWhiteSpace(activeHoveredKiCadNetName);

        if (RectGeometry.TryInvert(this.schematicsMatrix, out var inv) && !hoveringKiCadNet)
        {
            var localPoint = new Point(
                (point.X * inv.M11) + (point.Y * inv.M21) + inv.M31,
                (point.X * inv.M12) + (point.Y * inv.M22) + inv.M32);

            if (this.polylineManager != null && this.polylineManager.OnPointerPressed(point, localPoint, pointer, hoveringComponent))
            {
                e.Handled = true;
                return;
            }
        }

        if (pointer.Properties.IsRightButtonPressed)
        {
            this.isPanning = true;
            this.panStartPoint = point;
            this.panStartMatrix = this.schematicsMatrix;

            this.HideSchematicsHoverUi();
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);

            e.Pointer.Capture(this.SchematicsContainer);
            e.Handled = true;
            return;
        }

        if (pointer.Properties.IsLeftButtonPressed && !this.thisIsLabelEditorMode)
        {
            bool lockedChanged = false;

            if (hoveringKiCadNet)
            {
                if (this.thisLockedKiCadNetNames.Contains(activeHoveredKiCadNetName!))
                {
                    this.thisLockedKiCadNetNames.Remove(activeHoveredKiCadNetName!);
                    this.thisHoveredKiCadNetName = null;
                }
                else
                {
                    this.thisLockedKiCadNetNames.Add(activeHoveredKiCadNetName!);
                }

                lockedChanged = true;
                this.RefreshKiCadOverlay();
                this.RefreshBlinkStateFromCurrentSelection();
            }

            if (lockedChanged)
            {
                e.Handled = true;
                return;
            }

            if (hoveringComponent)
            {
                this.SelectComponentByBoardLabel(boardLabel);

                if (e.ClickCount == 1 && this.MainWindow != null)
                {
                    this.MainWindow.OpenComponentInfoPopup(boardLabel, displayText);
                }

                e.Handled = true;
                return;
            }
        }
    }

    // ###########################################################################################
    // Translates the schematics image while the right mouse button is held down.
    // Routes movement and shift key state to Polyline Manager, label editor, and the new KiCad
    // trace calibration box interaction mode.
    // ###########################################################################################
    private void OnSchematicsPointerMoved(object? sender, PointerEventArgs e)
    {
        var point = e.GetPosition(this.SchematicsContainer);
        bool isShiftDown = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        this.UpdateInteractiveCadTraceHoverShiftState(e.KeyModifiers);

        // An in-progress worklog area drag owns the pointer until release - checked before the
        // net-connections panel guard below, so dragging across that panel keeps updating the area
        // instead of silently stalling the gesture half-way.
        if (this.thisIsResizingWorklogEntry)
        {
            this.UpdateWorklogEntryResize(point);
            e.Handled = true;
            return;
        }

        if (!this.isPanning && this.IsPointerInsideKiCadNetConnectionsPanel(point))
        {
            this.ClearTransientHoverForKiCadNetConnectionsPanel();
            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode && this.thisKiCadTraceCalibrationDragMode != LabelEditorDragMode.None)
        {
            if (this.TryGetSchematicsImagePixelPoint(point, out var pixelPoint))
            {
                this.UpdateKiCadTraceCalibrationDrag(pixelPoint);
            }

            this.UpdateKiCadTraceCalibrationCursor(point);
            e.Handled = true;
            return;
        }

        if (this.thisIsLabelEditorMode && this.thisLabelEditorDragMode != LabelEditorDragMode.None)
        {
            if (this.SchematicsLabelEditorOverlay.HoveredIndex != -1)
            {
                this.SetLabelEditorOverlayTransientState(hoveredIndex: -1);
            }

            this.UpdateLabelEditorCursor(point);

            if (this.TryGetLabelEditorPixelPoint(point, out var pixelPoint))
            {
                this.UpdateLabelEditorDrag(pixelPoint, e.KeyModifiers);
            }

            e.Handled = true;
            return;
        }

        if (this.thisIsLabelEditorMode && this.thisIsDrawingLabelEditorRectangle)
        {
            if (this.SchematicsLabelEditorOverlay.HoveredIndex != -1)
            {
                this.SetLabelEditorOverlayTransientState(hoveredIndex: -1);
            }

            this.UpdateLabelEditorCursor(point);

            if (this.TryGetLabelEditorPixelPoint(point, out var pixelPoint))
            {
                this.UpdateDrawingLabelEditorRectangle(pixelPoint, e.KeyModifiers);
            }

            e.Handled = true;
            return;
        }

        if (this.thisIsWorklogEntryMode && this.thisIsDrawingWorklogEntryRectangle)
        {
            if (this.TryGetSchematicsImagePixelPoint(point, out var worklogPixelPoint))
            {
                this.UpdateDrawingWorklogEntryRectangle(worklogPixelPoint);
            }

            e.Handled = true;
            return;
        }

        if (!this.thisIsLabelEditorMode && !this.thisIsKiCadTraceCalibrationMode && !this.thisIsWorklogEntryMode && RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            var localPoint = new Point(
                (point.X * inv.M11) + (point.Y * inv.M21) + inv.M31,
                (point.X * inv.M12) + (point.Y * inv.M22) + inv.M32);

            if (this.polylineManager != null && this.polylineManager.OnPointerMoved(localPoint, isShiftDown))
            {
                e.Handled = true;
                return;
            }
        }

        if (this.isPanning)
        {
            var delta = point - this.panStartPoint;
            this.schematicsMatrix = this.panStartMatrix * Matrix.CreateTranslation(delta.X, delta.Y);
            this.ClampSchematicsMatrix();
            e.Handled = true;
            return;
        }

        if (this.thisIsWorklogEntryMode)
        {
            e.Handled = true;
            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            this.UpdateKiCadTraceCalibrationCursor(point);
            this.SchematicsHoverLabelBorder.IsVisible = false;
            this.SchematicsHoverLabelText.Text = string.Empty;
            this.SchematicsHoverPadBorder.IsVisible = false;
            this.SchematicsHoverPadText.Text = string.Empty;

            if (this.MainWindow != null)
            {
                this.MainWindow.isHoveringComponent = false;
            }

            e.Handled = true;
            return;
        }

        if (!this.thisIsLabelEditorMode)
        {
            if (this.ShouldProcessKiCadHoverHitTest(point))
            {
                this.HitTestKiCadOverlayForHover(point);
            }
        }

        if (this.thisIsLabelEditorMode)
        {
            int hoveredIndex = -1;

            if (this.TryGetSelectedLabelEditorHighlightAtContainerPoint(point, out var hoveredSelectedIndex))
            {
                hoveredIndex = hoveredSelectedIndex;
            }

            if (this.SchematicsLabelEditorOverlay.HoveredIndex != hoveredIndex)
            {
                this.SetLabelEditorOverlayTransientState(hoveredIndex: hoveredIndex);
            }
        }

        this.UpdateSchematicsHoverUi(point);
    }

    // ###########################################################################################
    // Exits pan mode when the right mouse button is released, finalizes label-editor operations,
    // and handles the new KiCad trace calibration move/resize workflow including empty-space
    // right-click access to Apply or Discard actions.
    // Keeps keyboard focus on the schematics control while KiCad trace calibration mode is active.
    // ###########################################################################################
    private void OnSchematicsPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        var point = e.GetPosition(this.SchematicsContainer);

        this.UpdateInteractiveCadTraceHoverShiftState(e.KeyModifiers);

        // BEFORE the net-connections panel guard, not after.
        //
        // That guard returns early, so a resize released with the pointer over the panel never
        // reached CompleteWorklogEntryResize: the drag was never saved AND thisIsResizingWorklogEntry
        // stayed true forever, after which every pointer-move was swallowed by the resize branch -
        // panning, hover and selection all dead until the board was switched. An in-flight gesture
        // has to be finished wherever the pointer happens to be when the button comes up.
        if (this.thisIsResizingWorklogEntry)
        {
            this.CompleteWorklogEntryResize();

            // Releases the capture taken in TryBeginWorklogEntryResize. Without this the pointer
            // stays captured to SchematicsContainer after the drag, routing later events there even
            // once the pointer has left - the same fault the worklog-drawing branch below carries a
            // comment about having already had to fix once.
            e.Pointer.Capture(null);

            this.UpdateWorklogEntryResizeHover(point);
            e.Handled = true;
            return;
        }

        if (!this.isPanning && this.IsPointerInsideKiCadNetConnectionsPanel(point))
        {
            this.ClearTransientHoverForKiCadNetConnectionsPanel();
            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode && this.thisKiCadTraceCalibrationDragMode != LabelEditorDragMode.None)
        {
            this.CompleteKiCadTraceCalibrationDrag();

            // Released here rather than left hanging: the press captured the pointer to keep the drag
            // alive, and a capture that outlives the drag routes every later event to
            // SchematicsContainer even once the pointer has left it - the fault the worklog branch
            // above carries its own comment about.
            this.ReleaseKiCadTraceCalibrationPointerCapture();

            this.UpdateKiCadTraceCalibrationCursor(point);
            this.SchematicsContainer.Focus();
            this.Focus();
            e.Handled = true;
            return;
        }

        if (this.thisIsLabelEditorMode && this.thisLabelEditorDragMode != LabelEditorDragMode.None)
        {
            this.CompleteLabelEditorDrag();
            this.UpdateLabelEditorCursor(point);
            e.Handled = true;
            return;
        }

        if (this.thisIsLabelEditorMode && this.thisIsDrawingLabelEditorRectangle)
        {
            if (this.TryGetLabelEditorPixelPoint(point, out var pixelPoint))
            {
                this.CompleteDrawingLabelEditorRectangle(point, pixelPoint, e.KeyModifiers);
            }
            else
            {
                this.thisIsDrawingLabelEditorRectangle = false;
                this.thisLabelEditorDraftRectangle = null;
                this.RefreshLabelEditorOverlay();
            }

            this.UpdateLabelEditorCursor(point);
            e.Handled = true;
            return;
        }

        if (this.thisIsWorklogEntryMode && this.thisIsDrawingWorklogEntryRectangle)
        {
            if (this.TryGetSchematicsImagePixelPoint(point, out var worklogPixelPoint))
            {
                this.CompleteDrawingWorklogEntryRectangle(worklogPixelPoint);
            }
            else
            {
                this.thisIsDrawingWorklogEntryRectangle = false;
                this.thisWorklogEntryDraftRectangle = null;
                this.RefreshWorklogEntryOverlay();
            }

            // This branch returns before the isPanning teardown below, so it must do that teardown
            // itself. Pressing the right button while already drag-drawing an area starts a pan;
            // leaving it armed here left the pointer captured and every later move panning the
            // board with no button held, recoverable only by switching board.
            if (this.isPanning)
            {
                this.isPanning = false;
                e.Pointer.Capture(null);
                this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.Cross);
            }

            e.Handled = true;
            return;
        }

        if (!this.thisIsLabelEditorMode && !this.thisIsKiCadTraceCalibrationMode && !this.thisIsWorklogEntryMode && RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            var localPoint = new Point(
                (point.X * inv.M11) + (point.Y * inv.M21) + inv.M31,
                (point.X * inv.M12) + (point.Y * inv.M22) + inv.M32);

            if (this.polylineManager != null && this.polylineManager.OnPointerReleased(point, localPoint))
            {
                e.Handled = true;
                return;
            }
        }

        if (!this.isPanning)
        {
            if (this.thisIsKiCadTraceCalibrationMode)
            {
                this.UpdateKiCadTraceCalibrationCursor(point);
                this.SchematicsContainer.Focus();
                this.Focus();
            }

            return;
        }

        this.isPanning = false;
        e.Pointer.Capture(null);

        var delta = point - this.panStartPoint;
        bool isStationaryRightClick = Math.Abs(delta.X) < 4 && Math.Abs(delta.Y) < 4;

        if (isStationaryRightClick)
        {
            if (this.thisIsWorklogEntryMode)
            {
                // Right-click while marking a worklog entry area only pans - no context menu.
            }
            else if (this.thisIsKiCadTraceCalibrationMode)
            {
                this.ShowLabelEditorMenu(point);
            }
            else if (this.thisIsLabelEditorMode)
            {
                if (this.TryGetLabelEditorHighlightAtContainerPoint(point, out var workingIndex))
                {
                    this.DeleteLabelEditorHighlight(workingIndex);
                    this.HideLabelEditorMenu();
                }
                else
                {
                    this.ShowLabelEditorMenu(point);
                }
            }
            else
            {
                string? activeHoveredKiCadNetName = this.GetActiveHoveredKiCadNetName();

                if (this.TryGetHoveredBoardLabel(point, out var boardLabel, out _))
                {
                    this.ToggleComponentSelectionByBoardLabel(boardLabel);
                }
                else if (!string.IsNullOrWhiteSpace(activeHoveredKiCadNetName) && this.thisLockedKiCadNetNames.Contains(activeHoveredKiCadNetName))
                {
                    this.thisLockedKiCadNetNames.Remove(activeHoveredKiCadNetName);
                    this.thisHoveredKiCadNetName = null;
                    this.RefreshKiCadOverlay();
                    this.RefreshBlinkStateFromCurrentSelection();
                }
                else if (this.thisLockedKiCadNetNames.Count > 0)
                {
                    this.thisLockedKiCadNetNames.Clear();
                    this.RefreshKiCadOverlay();
                    this.RefreshBlinkStateFromCurrentSelection();
                }
                else
                {
                    this.ShowLabelEditorMenu(point);
                }
            }
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            this.SchematicsContainer.Focus();
            this.Focus();
        }

        if (this.thisIsWorklogEntryMode)
        {
            // UpdateSchematicsHoverUi has no worklog branch, so letting it run here undid the mode:
            // it hit-tests components, shows the hover label, leaves MainWindow.isHoveringComponent
            // set, and replaces the Cross cursor with Hand/Default. Restore the mode's own cursor
            // and leave the hover UI alone instead.
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.Cross);
            e.Handled = true;
            return;
        }

        this.UpdateSchematicsHoverUi(e.GetPosition(this.SchematicsContainer));
        e.Handled = true;
    }

    // ###########################################################################################
    // Clears hover UI when pointer exits schematic area.
    // Uses a batched overlay update so exiting the editor does not trigger extra redraw churn.
    // ###########################################################################################
    private void OnSchematicsPointerExited(object? sender, PointerEventArgs e)
    {
        if (this.isPanning)
        {
            return;
        }

        if (this.thisIsLabelEditorMode && this.SchematicsLabelEditorOverlay.HoveredIndex != -1)
        {
            this.SetLabelEditorOverlayTransientState(hoveredIndex: -1);
        }

        this.HideSchematicsHoverUi();
    }

    // ###########################################################################################
    // Handles keyboard interaction for label-editor and KiCad calibration workflows.
    // Ctrl+Z undoes label-editor changes and Ctrl+Y redoes them within the current editor session.
    // Pressing D duplicates the currently selected editor rectangle and opens the new-label prompt.
    // ###########################################################################################
    private void OnSchematicsKeyDown(object? sender, KeyEventArgs e)
    {
        this.UpdateInteractiveCadTraceHoverShiftState(e.KeyModifiers);

        if (this.thisIsWorklogEntryMode)
        {
            // Escape cancels the mode outright. Nothing is at risk: the mode is only ever waiting
            // for an area to be drawn now, and the editor that opens once one IS drawn is a modal
            // window with its own Escape handling, so this handler cannot fire underneath it.
            if (e.Key == Key.Escape)
            {
                this.CancelWorklogEntryMode();
                e.Handled = true;
            }

            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            if (e.Key == Key.Escape)
            {
                this.CancelKiCadTraceCalibrationMode();
                e.Handled = true;
                return;
            }

            if (this.ApplyKiCadTraceCalibrationKeyboardStep(e.Key, e.KeyModifiers))
            {
                e.Handled = true;
            }

            return;
        }

        if (!this.thisIsLabelEditorMode)
        {
            return;
        }

        if (this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            return;
        }

        bool thisIsCtrlDown = e.KeyModifiers.HasFlag(KeyModifiers.Control);

        if (thisIsCtrlDown && e.Key == Key.Z)
        {
            if (this.TryUndoLabelEditorChange())
            {
                e.Handled = true;
            }

            return;
        }

        if (thisIsCtrlDown && e.Key == Key.Y)
        {
            if (this.TryRedoLabelEditorChange())
            {
                e.Handled = true;
            }

            return;
        }

        if (!thisIsCtrlDown &&
            !e.KeyModifiers.HasFlag(KeyModifiers.Alt) &&
            e.Key == Key.D)
        {
            if (this.TryDuplicateSelectedLabelEditorHighlight())
            {
                e.Handled = true;
            }

            return;
        }

        if (this.ApplySelectedLabelEditorKeyboardStep(e.Key, e.KeyModifiers))
        {
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Tracks key releases so SHIFT-based KiCad hover highlighting updates immediately.
    // ###########################################################################################
    private void OnSchematicsKeyUp(object? sender, KeyEventArgs e)
    {
        this.UpdateInteractiveCadTraceHoverShiftState(e.KeyModifiers);
    }
}