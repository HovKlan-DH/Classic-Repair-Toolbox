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
// Interactive KiCad trace calibration mode: aligning the KiCad geometry to the schematic
// image by dragging or nudging the calibration box, and persisting the result.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private bool thisIsKiCadTraceCalibrationMode;

    private double thisKiCadCalibrationImageLeft;

    private double thisKiCadCalibrationImageTop;

    private double thisKiCadCalibrationImageRight;

    private double thisKiCadCalibrationImageBottom;

    private double thisKiCadCalibrationStartImageLeft;

    private double thisKiCadCalibrationStartImageTop;

    private double thisKiCadCalibrationStartImageRight;

    private double thisKiCadCalibrationStartImageBottom;

    private LabelEditorDragMode thisKiCadTraceCalibrationDragMode;

    private Point thisKiCadTraceCalibrationDragStartPixelPoint;

    // ###########################################################################################
    // The pointer the calibration press captured to SchematicsContainer, held so the capture can
    // be released by whatever ends the drag - not only by OnSchematicsPointerReleased.
    //
    // The release handler only releases inside its "drag mode is not None" branch, and ESC
    // (CancelKiCadTraceCalibrationMode) and Apply (ApplyKiCadTraceCalibration) both clear that
    // mode while the button is still down. The release that follows then falls through to the
    // "not panning, so nothing to tear down" early return and the capture outlives the gesture,
    // routing every later pointer event to SchematicsContainer even once the pointer has left it -
    // the same fault the worklog-resize and worklog-drawing branches each carry a comment about
    // having already had to fix, recoverable only by switching board.
    // ###########################################################################################
    private IPointer? thisKiCadTraceCalibrationCapturedPointer;

    // ###########################################################################################
    // Releases the calibration drag's pointer capture, if one is still held. Safe to call when
    // nothing is captured, so every path that ends a calibration drag can call it unconditionally.
    // ###########################################################################################
    private void ReleaseKiCadTraceCalibrationPointerCapture()
    {
        this.thisKiCadTraceCalibrationCapturedPointer?.Capture(null);
        this.thisKiCadTraceCalibrationCapturedPointer = null;
    }

    // ###########################################################################################
    // What makes a calibration drag cheap: the overlay geometry is built ONCE, at the calibration
    // recorded here, and every pointer move after that only updates a matrix on the render control.
    //
    // The problem this solves, measured on a real board (451 nets, ~5,200 primitives): rebuilding
    // the overlay costs ~390-540 ms, essentially all of it in the per-net geometry loop
    // (SetGeometry is only ~15-20 ms). The per-net primitive cache cannot absorb any of it during
    // a drag, because the calibration values sit in its generation key (see KiCadOverlayCacheKeys)
    // - it recorded 0 hits and 451 MISSES on every single frame. So each frame rebuilt identical
    // shapes at slightly shifted coordinates.
    //
    // A calibration change is a pure affine transform though - scale, optional mirror, translate,
    // no rotation or shear - so the shift can be applied at draw time instead. That is exactly how
    // zoom and pan already behave here (KiCadOverlayRenderControl.ArrangeOverride deliberately does
    // not re-record primitives when only the transform changes). The matrix is derived by
    // Handlers/Geometry/KiCadCalibrationTransform, which is unit tested against the very same
    // world-to-local mapping the geometry was built with, so "transformed" and "rebuilt" provably
    // land in the same place.
    //
    // MEASURED before and after, on the same board, so the next person does not have to re-derive
    // any of it. Before: ~390-540 ms per drag frame, essentially all of it the per-net geometry
    // loop. After: ~0.7-1.0 ms in this handler, ~0.7 ms in SetGeometry and ~6 ms in the control's
    // Render - roughly 7 ms of measurable work per frame, with no rebuild at all.
    //
    // What is left is Avalonia rasterising the primitives themselves: Render reports drawing 5,853
    // of 5,853, because at calibration zoom the whole board is on screen and the viewport cull has
    // nothing to discard. That is the floor for this approach, and it is NOT worth chasing with
    // another "cheaper geometry while dragging" attempt - one was tried (dropping zone fills) and
    // measured as buying nothing, because the primitive count barely moved. Reducing it further
    // means drawing genuinely fewer traces, which costs the user the alignment feedback that is the
    // entire point of calibration mode.
    //
    // These three are only meaningful while a drag is in progress; HasValue on the reference is
    // what says a drag-time transform is currently possible at all.
    // ###########################################################################################
    private KiCadCalibrationValues? thisKiCadTraceCalibrationDragReferenceCalibration;

    private Rect thisKiCadTraceCalibrationDragReferenceWorldBounds;

    private Rect thisKiCadTraceCalibrationDragReferenceContentRect;

    // ###########################################################################################
    // The four calibration edges as the one value KiCadCalibrationGeometry works on, and back.
    //
    // The fields stay the storage - this is only the adapter that lets the maths live outside
    // the tab. Note the edges are passed through EXACTLY as stored, mirror-inverted ordering
    // included: KiCadCalibrationBox is built for that and derives its mirror flags from it.
    // ###########################################################################################
    private KiCadCalibrationBox GetKiCadCalibrationBox() =>
        new KiCadCalibrationBox(
            this.thisKiCadCalibrationImageLeft,
            this.thisKiCadCalibrationImageTop,
            this.thisKiCadCalibrationImageRight,
            this.thisKiCadCalibrationImageBottom);

    private void SetKiCadCalibrationBox(KiCadCalibrationBox box)
    {
        this.thisKiCadCalibrationImageLeft = box.Left;
        this.thisKiCadCalibrationImageTop = box.Top;
        this.thisKiCadCalibrationImageRight = box.Right;
        this.thisKiCadCalibrationImageBottom = box.Bottom;
    }

    private KiCadCalibrationBox GetKiCadCalibrationStartBox() =>
        new KiCadCalibrationBox(
            this.thisKiCadCalibrationStartImageLeft,
            this.thisKiCadCalibrationStartImageTop,
            this.thisKiCadCalibrationStartImageRight,
            this.thisKiCadCalibrationStartImageBottom);

    // ###########################################################################################
    // Applies saved KiCad mirror flags onto the calibration-box coordinates by swapping edges.
    // Calibration mode encodes mirroring by having Left>Right and/or Top>Bottom.
    // ###########################################################################################
    private void ApplyKiCadCalibrationMirrorFlagsToBox(bool mirrorX, bool mirrorY)
    {
        this.SetKiCadCalibrationBox(
            KiCadCalibrationGeometry.ApplyMirrorFlags(this.GetKiCadCalibrationBox(), mirrorX, mirrorY));
    }

    // ###########################################################################################
    // Applies keyboard move, expand, or shrink operations to the KiCad trace calibration box.
    // Arrow keys move by 1 px, Shift expands in the pressed direction, and Alt shrinks from
    // the opposite side of the pressed direction, matching the component label editor behavior.
    // ###########################################################################################
    private bool ApplyKiCadTraceCalibrationKeyboardStep(Key key, KeyModifiers modifiers)
    {
        if (!this.thisIsKiCadTraceCalibrationMode ||
            this.currentFullResBitmap == null ||
            this.thisKiCadTraceCalibrationDragMode != LabelEditorDragMode.None ||
            this.SchematicsLabelEditorMenuBorder.IsVisible)
        {
            return false;
        }

        // The maths - which modifier does what, and the refusal to shrink the box to nothing -
        // is KiCadCalibrationGeometry's. This end only supplies the box and stores the result.
        if (!KiCadCalibrationGeometry.TryApplyKeyboardStep(
                this.GetKiCadCalibrationBox(),
                key,
                modifiers,
                out var updatedBox))
        {
            return false;
        }

        this.SetKiCadCalibrationBox(updatedBox);

        this.RefreshKiCadOverlay(forceImmediate: true);
        return true;
    }

    // ###########################################################################################
    // Enters interactive KiCad trace calibration mode and seeds the resize box from the currently
    // active calibration if one exists, otherwise from the default full-image KiCad bounds.
    // The temporary traces-and-pads visibility toggle always defaults to checked on entry.
    // ###########################################################################################
    private void BeginKiCadTraceCalibrationMode()
    {
        if (this.currentFullResBitmap == null || this.thisKiCadProject == null)
        {
            return;
        }

        var view = this.ResolveKiCadViewForCurrentSchematic();
        if (view == null)
        {
            return;
        }

        Rect imageBounds = this.BuildKiCadCalibrationImageBounds(view);

        this.thisKiCadCalibrationImageLeft = imageBounds.Left;
        this.thisKiCadCalibrationImageTop = imageBounds.Top;
        this.thisKiCadCalibrationImageRight = imageBounds.Right;
        this.thisKiCadCalibrationImageBottom = imageBounds.Bottom;

        // Load persisted mirror flags and re-apply them onto the calibration box.
        string excelPath = this.MainWindow?.GetCurrentBoardExcelPath() ?? string.Empty;
        string schematicName = this.GetCurrentSchematicName();
        bool hasSavedCalibration = BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                excelPath,
                schematicName,
                out _,
                out _,
                out _,
                out _,
                out _,
                out bool mirrorX,
                out bool mirrorY);

        if (hasSavedCalibration)
        {
            this.ApplyKiCadCalibrationMirrorFlagsToBox(mirrorX, mirrorY);
        }
        else
        {
            // First-time calibration for this schematic: imageBounds is the raw full-image
            // bounds, whose corners sit exactly on the image edges - reported directly as the
            // resize handles being clamped onto the viewport frame and hard to grab. Pull them
            // in so every handle starts comfortably inside the visible schematic. A schematic
            // that already has a saved calibration skips this entirely, so re-entering
            // calibration mode on an already-aligned board never moves the box the user set.
            this.SetKiCadCalibrationBox(
                KiCadCalibrationGeometry.InsetForInitialVisibility(
                    this.GetKiCadCalibrationBox(),
                    this.currentFullResBitmap.PixelSize.Width,
                    this.currentFullResBitmap.PixelSize.Height));
        }

        this.thisKiCadCalibrationStartImageLeft = this.thisKiCadCalibrationImageLeft;
        this.thisKiCadCalibrationStartImageTop = this.thisKiCadCalibrationImageTop;
        this.thisKiCadCalibrationStartImageRight = this.thisKiCadCalibrationImageRight;
        this.thisKiCadCalibrationStartImageBottom = this.thisKiCadCalibrationImageBottom;

        // Same one-sided exclusion as the label editor: worklog entry mode grabs the pointer
        // handlers first, so it has to be ended rather than left running underneath this.
        this.CancelWorklogEntryMode();

        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;
        this.thisIsKiCadTraceCalibrationMode = true;

        // Entering the mode switches the active calibration source, so anything recorded from the
        // previous geometry is stale; the refresh at the end of this method records it afresh.
        this.ClearKiCadTraceCalibrationDragTransform();

        this.CheckGlobalShowCalibrationTracesAndPads.IsChecked = true;

        // Hover hit-test caches bake in whichever calibration was active when they were built,
        // and entering calibration mode switches GetKiCadViewCalibration from the persisted box
        // to the live drag box - see InvalidateKiCadHoverHitTestCachesForCalibrationChange.
        this.InvalidateKiCadHoverHitTestCachesForCalibrationChange();

        this.HideLabelEditorMenu();
        this.UpdateInteractiveCadTraceHoverModeUi();
        this.SchematicsContainer.Focus();
        this.Focus();
        this.RefreshKiCadOverlay(forceImmediate: true);
        this.UpdateSchematicsHoverUi(new Point(0, 0));

        Logger.Info($"KiCad trace calibration mode enabled for schematic [{this.GetCurrentSchematicName()}]");
    }

    // ###########################################################################################
    // Cancels the current interactive KiCad trace calibration session and restores the persisted
    // calibration without writing anything to disk.
    // ###########################################################################################
    private void CancelKiCadTraceCalibrationMode()
    {
        this.ClearKiCadTraceCalibrationDragTransform();

        // ESC can arrive with the button still down, mid-drag. Clearing the drag mode below is what
        // makes the eventual PointerReleased skip its own release, so the capture has to go here.
        this.ReleaseKiCadTraceCalibrationPointerCapture();

        this.thisIsKiCadTraceCalibrationMode = false;
        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;
        this.thisKiCadCalibrationImageLeft = 0.0;
        this.thisKiCadCalibrationImageTop = 0.0;
        this.thisKiCadCalibrationImageRight = 0.0;
        this.thisKiCadCalibrationImageBottom = 0.0;
        this.thisKiCadCalibrationStartImageLeft = 0.0;
        this.thisKiCadCalibrationStartImageTop = 0.0;
        this.thisKiCadCalibrationStartImageRight = 0.0;
        this.thisKiCadCalibrationStartImageBottom = 0.0;

        this.CheckGlobalShowCalibrationTracesAndPads.IsChecked = true;

        // Same reasoning as on entry: leaving calibration mode switches the active calibration
        // back to the persisted box, and any cache built against the live drag box must go with it.
        this.InvalidateKiCadHoverHitTestCachesForCalibrationChange();

        this.HideLabelEditorMenu();
        this.UpdateInteractiveCadTraceHoverModeUi();
        this.RefreshKiCadOverlay(forceImmediate: true);
        this.SchematicsContainer.Focus();

        Logger.Info("KiCad trace calibration mode canceled");
    }

    // ###########################################################################################
    // Saves the current interactive KiCad trace calibration box into the board JSON file and then
    // exits calibration mode so the persisted transform becomes the active transform immediately.
    // ###########################################################################################
    private void ApplyKiCadTraceCalibration()
    {
        if (!this.thisIsKiCadTraceCalibrationMode || this.currentFullResBitmap == null)
        {
            return;
        }

        string schematicName = this.GetCurrentSchematicName();
        string excelPath = this.MainWindow?.GetCurrentBoardExcelPath() ?? string.Empty;
        string cadName = this.schematicByName.TryGetValue(schematicName, out var entry)
            ? entry.CadName?.Trim() ?? string.Empty
            : string.Empty;

        if (string.IsNullOrWhiteSpace(excelPath) || string.IsNullOrWhiteSpace(schematicName))
        {
            return;
        }

        double left = Math.Min(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
        double right = Math.Max(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
        double top = Math.Min(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);
        double bottom = Math.Max(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);

        bool mirrorX = this.thisKiCadCalibrationImageLeft > this.thisKiCadCalibrationImageRight;
        bool mirrorY = this.thisKiCadCalibrationImageTop > this.thisKiCadCalibrationImageBottom;

        double scaleX = (right - left) / this.currentFullResBitmap.PixelSize.Width;
        double scaleY = (bottom - top) / this.currentFullResBitmap.PixelSize.Height;
        double offsetX = left;
        double offsetY = top;

        BoardComponentHighlightStorage.SaveKiCadCalibration(
            excelPath,
            schematicName,
            cadName,
            offsetX,
            offsetY,
            scaleX,
            scaleY,
            mirrorX,
            mirrorY);

        this.ClearKiCadTraceCalibrationDragTransform();

        // Same reason as the cancel path: Apply can be reached from the keyboard while a drag is
        // still live, and clearing the drag mode below stops PointerReleased releasing it.
        this.ReleaseKiCadTraceCalibrationPointerCapture();

        this.thisIsKiCadTraceCalibrationMode = false;
        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;
        this.CheckGlobalShowCalibrationTracesAndPads.IsChecked = true;

        // The bug this fixes: the newly-saved calibration takes effect for RENDERING immediately
        // (RefreshKiCadOverlay recomputes every primitive on every call), but the hover hit-test
        // caches do not recompute - they are built once and served from a dictionary keyed by
        // schematic/pcb identity only, never by calibration. Without this, hover kept testing
        // against whichever calibration (live drag box, or the pre-calibration persisted one) was
        // active the first time hover ran, so a freshly-calibrated schematic highlighted nothing
        // where the visible traces now are, until the app restarted and rebuilt the caches fresh.
        this.InvalidateKiCadHoverHitTestCachesForCalibrationChange();

        this.HideLabelEditorMenu();
        this.UpdateInteractiveCadTraceHoverModeUi();
        this.RefreshKiCadOverlay(forceImmediate: true);
        this.SchematicsContainer.Focus();

        Logger.Info(
            $"KiCad trace calibration saved for schematic [{schematicName}] " +
            $"OffsetX=[{offsetX.ToString("0.######", CultureInfo.InvariantCulture)}] " +
            $"OffsetY=[{offsetY.ToString("0.######", CultureInfo.InvariantCulture)}] " +
            $"ScaleX=[{scaleX.ToString("0.######", CultureInfo.InvariantCulture)}] " +
            $"ScaleY=[{scaleY.ToString("0.######", CultureInfo.InvariantCulture)}] " +
            $"MirrorX=[{mirrorX}] MirrorY=[{mirrorY}]");
    }

    // ###########################################################################################
    // Builds the current KiCad calibration box in image-pixel coordinates by mapping the active
    // KiCad view bounds through the currently active calibration.
    // ###########################################################################################
    private Rect BuildKiCadCalibrationImageBounds(KiCadProjectView view)
    {
        if (this.currentFullResBitmap == null)
        {
            return default;
        }

        Rect worldBounds;
        string currentSchematicName = this.GetCurrentSchematicName();

        if (string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase))
        {
            if (view.SourceIndex < 0 || view.SourceIndex >= this.thisKiCadProject!.Root.Pcb.Count)
            {
                return new Rect(0, 0, this.currentFullResBitmap.PixelSize.Width, this.currentFullResBitmap.PixelSize.Height);
            }

            worldBounds = this.GetKiCadPcbWorldBounds(this.thisKiCadProject.Root.Pcb[view.SourceIndex]);
        }
        else
        {
            if (view.SourceIndex < 0 || view.SourceIndex >= this.thisKiCadProject!.Root.Schematics.Count)
            {
                return new Rect(0, 0, this.currentFullResBitmap.PixelSize.Width, this.currentFullResBitmap.PixelSize.Height);
            }

            worldBounds = this.GetKiCadSchematicWorldBounds(this.thisKiCadProject.Root.Schematics[view.SourceIndex]);
        }

        if (worldBounds.Width <= 0 || worldBounds.Height <= 0)
        {
            return new Rect(0, 0, this.currentFullResBitmap.PixelSize.Width, this.currentFullResBitmap.PixelSize.Height);
        }

        var calibration = KiCadViewCalibration.Identity;

        string excelPath = this.MainWindow?.GetCurrentBoardExcelPath() ?? string.Empty;
        if (BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                excelPath,
                currentSchematicName,
                out _,
                out double offsetX,
                out double offsetY,
                out double scaleX,
                out double scaleY,
                out bool mirrorX,
                out bool mirrorY))
        {
            calibration = new KiCadViewCalibration
            {
                OffsetX = offsetX,
                OffsetY = offsetY,
                ScaleX = scaleX,
                ScaleY = scaleY,
                MirrorX = mirrorX,
                MirrorY = mirrorY
            };
        }

        Point topLeft = this.MapKiCadWorldToImagePixel(worldBounds.Left, worldBounds.Top, worldBounds, calibration);
        Point topRight = this.MapKiCadWorldToImagePixel(worldBounds.Right, worldBounds.Top, worldBounds, calibration);
        Point bottomLeft = this.MapKiCadWorldToImagePixel(worldBounds.Left, worldBounds.Bottom, worldBounds, calibration);
        Point bottomRight = this.MapKiCadWorldToImagePixel(worldBounds.Right, worldBounds.Bottom, worldBounds, calibration);

        double left = new[] { topLeft.X, topRight.X, bottomLeft.X, bottomRight.X }.Min();
        double right = new[] { topLeft.X, topRight.X, bottomLeft.X, bottomRight.X }.Max();
        double top = new[] { topLeft.Y, topRight.Y, bottomLeft.Y, bottomRight.Y }.Min();
        double bottom = new[] { topLeft.Y, topRight.Y, bottomLeft.Y, bottomRight.Y }.Max();

        return new Rect(left, top, Math.Max(1.0, right - left), Math.Max(1.0, bottom - top));
    }

    // ###########################################################################################
    // Maps one KiCad world coordinate directly into image-pixel coordinates using the current
    // non-affine box calibration model.
    // ###########################################################################################
    private Point MapKiCadWorldToImagePixel(
        double worldX,
        double worldY,
        Rect worldBounds,
        KiCadViewCalibration calibration)
    {
        if (this.currentFullResBitmap == null || worldBounds.Width <= 0 || worldBounds.Height <= 0)
        {
            return default;
        }

        double nx = (worldX - worldBounds.X) / worldBounds.Width;
        double ny = (worldY - worldBounds.Y) / worldBounds.Height;

        if (calibration.MirrorX)
        {
            nx = 1.0 - nx;
        }

        if (calibration.MirrorY)
        {
            ny = 1.0 - ny;
        }

        double imageX = calibration.OffsetX + (nx * calibration.ScaleX * this.currentFullResBitmap.PixelSize.Width);
        double imageY = calibration.OffsetY + (ny * calibration.ScaleY * this.currentFullResBitmap.PixelSize.Height);

        return new Point(imageX, imageY);
    }

    // ###########################################################################################
    // Converts an image-pixel rectangle into schematic-local coordinates so the calibration border
    // can be drawn on top of the current image using the same mapping as other overlays.
    // ###########################################################################################
    private Rect ConvertImagePixelRectToLocalRect(Rect imagePixelRect)
    {
        if (this.currentFullResBitmap == null ||
            this.currentFullResBitmap.PixelSize.Width <= 0 ||
            this.currentFullResBitmap.PixelSize.Height <= 0)
        {
            return default;
        }

        var contentRect = this.GetImageContentRect();

        double x = contentRect.X + ((imagePixelRect.X / this.currentFullResBitmap.PixelSize.Width) * contentRect.Width);
        double y = contentRect.Y + ((imagePixelRect.Y / this.currentFullResBitmap.PixelSize.Height) * contentRect.Height);
        double width = (imagePixelRect.Width / this.currentFullResBitmap.PixelSize.Width) * contentRect.Width;
        double height = (imagePixelRect.Height / this.currentFullResBitmap.PixelSize.Height) * contentRect.Height;

        return new Rect(x, y, width, height);
    }

    // ###########################################################################################
    // Builds the visible KiCad calibration border box and explicit corner/side handle markers so the
    // user can see where resize interaction is available while aligning the temporary KiCad overlay.
    // The border is drawn slightly outside the actual KiCad data bounds to avoid covering details.
    // ###########################################################################################
    private KiCadOverlayPrimitive BuildKiCadCalibrationBoxPrimitive()
    {
        Rect thisBorderImageRect = this.GetKiCadCalibrationBorderImageRect();
        Rect thisLocalRect = this.ConvertImagePixelRectToLocalRect(thisBorderImageRect);

        double thisScale = Math.Max(0.0001, this.schematicsMatrix.M11);
        double thisHandleSize = Math.Clamp(10.0 / thisScale, 5.0, 12.0);
        double thisHalfHandleSize = thisHandleSize / 2.0;

        // IndianRed, matching the worklog/selection accent used elsewhere in this tab - asked for
        // directly in place of the original solid LimeGreen, which read as just another trace
        // colour rather than as the calibration tool's own overlay. The border is dashed so it is
        // visually distinct from both KiCad copper and the solid worklog-area rectangles; the
        // handles stay solid squares since dashing a mark this small would leave it barely visible.
        var thisHandleBrush = new SolidColorBrush(Colors.IndianRed, 1.0);
        var thisHandlePen = new Pen(thisHandleBrush, 1.0);
        var thisBorderPen = new Pen(thisHandleBrush, 1.0) { DashStyle = DashStyle.Dash };

        var thisHandleCenters = new[]
        {
        new Point(thisLocalRect.Left, thisLocalRect.Top),
        new Point(thisLocalRect.Center.X, thisLocalRect.Top),
        new Point(thisLocalRect.Right, thisLocalRect.Top),
        new Point(thisLocalRect.Right, thisLocalRect.Center.Y),
        new Point(thisLocalRect.Right, thisLocalRect.Bottom),
        new Point(thisLocalRect.Center.X, thisLocalRect.Bottom),
        new Point(thisLocalRect.Left, thisLocalRect.Bottom),
        new Point(thisLocalRect.Left, thisLocalRect.Center.Y)
    };

        // One StreamGeometry holding both the border figure and the eight handle figures, exactly
        // as before - RefreshKiCadCalibrationBoxPrimitiveOnly and the full rebuild path both rely
        // on the calibration box being exactly one primitive, always last (see that method's own
        // header), so the border and handles cannot become two separate primitives.
        //
        // Each figure's own isFilled controls whether IT is filled when the geometry is drawn with
        // a non-null Fill brush - it is a per-figure flag, not a per-geometry one, so the border can
        // stay unfilled (isFilled: false) while the handles are filled (isFilled: true) within this
        // one geometry/one DrawGeometry call. The border is stroked with the dashed pen; the shared
        // Pen also strokes the handle squares' 1px edge, but at handle size that dash is not visible
        // against the filled square underneath it. The original code built each handle primitive
        // with Fill set but folded every figure's isFilled through "Fill != null" against the WRONG
        // (per-sub-primitive) Fill, then discarded all of that and drew the combined geometry with
        // Fill = null - so the handles were never actually filled despite the Fill assignment
        // implying they should be. This restores what that assignment always intended.
        var thisGeometry = new StreamGeometry();

        using (var thisGeometryContext = thisGeometry.Open())
        {
            thisGeometryContext.BeginFigure(thisLocalRect.TopLeft, isFilled: false);
            thisGeometryContext.LineTo(thisLocalRect.TopRight);
            thisGeometryContext.LineTo(thisLocalRect.BottomRight);
            thisGeometryContext.LineTo(thisLocalRect.BottomLeft);
            thisGeometryContext.EndFigure(isClosed: true);

            foreach (var thisHandleCenter in thisHandleCenters)
            {
                var thisHandleRect = new Rect(
                    thisHandleCenter.X - thisHalfHandleSize,
                    thisHandleCenter.Y - thisHalfHandleSize,
                    thisHandleSize,
                    thisHandleSize);

                thisGeometryContext.BeginFigure(thisHandleRect.TopLeft, isFilled: true);
                thisGeometryContext.LineTo(thisHandleRect.TopRight);
                thisGeometryContext.LineTo(thisHandleRect.BottomRight);
                thisGeometryContext.LineTo(thisHandleRect.BottomLeft);
                thisGeometryContext.EndFigure(isClosed: true);
            }
        }

        return new KiCadOverlayPrimitive
        {
            Kind = KiCadOverlayPrimitiveKind.Geometry,
            Geometry = thisGeometry,
            Pen = thisBorderPen,
            Fill = thisHandleBrush
        };
    }

    // ###########################################################################################
    // Returns true when the pointer is inside the currently visible KiCad calibration rectangle.
    // This is used for move-drag behavior while calibration mode is active.
    // ###########################################################################################
    private bool IsPointerInsideCurrentKiCadCalibrationBounds(Point pointerInContainer)
    {
        if (!this.thisIsKiCadTraceCalibrationMode)
        {
            return false;
        }

        if (!this.TryGetSchematicsImagePixelPoint(pointerInContainer, out var pixelPoint))
        {
            return false;
        }

        double left = Math.Min(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
        double right = Math.Max(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
        double top = Math.Min(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);
        double bottom = Math.Max(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);

        return pixelPoint.X >= left &&
               pixelPoint.X <= right &&
               pixelPoint.Y >= top &&
               pixelPoint.Y <= bottom;
    }

    // ###########################################################################################
    // Builds the visual calibration-border rectangle in image-pixel space.
    // The border is intentionally expanded slightly outside the actual KiCad data bounds so the
    // visible box and handles do not sit directly on top of traces and pads.
    // ###########################################################################################
    private Rect GetKiCadCalibrationBorderImageRect()
    {
        const double thisBorderPaddingPixels = 10.0;

        double thisLeft = Math.Min(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
        double thisRight = Math.Max(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
        double thisTop = Math.Min(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);
        double thisBottom = Math.Max(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);

        double thisExpandedLeft = thisLeft - thisBorderPaddingPixels;
        double thisExpandedTop = thisTop - thisBorderPaddingPixels;
        double thisExpandedRight = thisRight + thisBorderPaddingPixels;
        double thisExpandedBottom = thisBottom + thisBorderPaddingPixels;

        if (this.currentFullResBitmap != null)
        {
            thisExpandedLeft = Math.Clamp(thisExpandedLeft, 0.0, this.currentFullResBitmap.PixelSize.Width);
            thisExpandedTop = Math.Clamp(thisExpandedTop, 0.0, this.currentFullResBitmap.PixelSize.Height);
            thisExpandedRight = Math.Clamp(thisExpandedRight, 0.0, this.currentFullResBitmap.PixelSize.Width);
            thisExpandedBottom = Math.Clamp(thisExpandedBottom, 0.0, this.currentFullResBitmap.PixelSize.Height);
        }

        return new Rect(
            thisExpandedLeft,
            thisExpandedTop,
            Math.Max(1.0, thisExpandedRight - thisExpandedLeft),
            Math.Max(1.0, thisExpandedBottom - thisExpandedTop));
    }

    // ###########################################################################################
    // Tries to resolve which KiCad calibration resize handle is under the pointer so the box can
    // be resized from edges or corners and flipped naturally by dragging across opposite sides.
    // Hit-testing uses the expanded visual border rectangle so the handles match what is drawn.
    // ###########################################################################################
    private bool TryGetKiCadTraceCalibrationHandleAtContainerPoint(
        Point pointerInContainer,
        out LabelEditorDragMode dragMode)
    {
        dragMode = LabelEditorDragMode.None;

        if (!this.thisIsKiCadTraceCalibrationMode ||
            this.currentFullResBitmap == null)
        {
            return false;
        }

        if (!RectGeometry.TryInvert(this.schematicsMatrix, out var thisInverseMatrix))
        {
            return false;
        }

        var thisLocalPoint = new Point(
            (pointerInContainer.X * thisInverseMatrix.M11) + (pointerInContainer.Y * thisInverseMatrix.M21) + thisInverseMatrix.M31,
            (pointerInContainer.X * thisInverseMatrix.M12) + (pointerInContainer.Y * thisInverseMatrix.M22) + thisInverseMatrix.M32);

        var thisContentRect = this.GetImageContentRect();
        if (thisContentRect.Width <= 0 || thisContentRect.Height <= 0 || !thisContentRect.Contains(thisLocalPoint))
        {
            return false;
        }

        Rect thisBorderImageRect = this.GetKiCadCalibrationBorderImageRect();
        Rect thisLocalRect = this.ConvertImagePixelRectToLocalRect(thisBorderImageRect);
        double thisScale = Math.Max(0.0001, this.schematicsMatrix.M11);

        foreach (var thisHitTarget in LabelEditorGeometry.BuildLabelEditorHandleHitRects(thisLocalRect, thisScale))
        {
            if (!thisHitTarget.HitRect.Contains(thisLocalPoint))
            {
                continue;
            }

            dragMode = thisHitTarget.DragMode;
            return true;
        }

        return false;
    }

    // ###########################################################################################
    // Remaps a visually hit KiCad calibration handle to the underlying stored edge/corner definition.
    // This keeps resize behavior correct after horizontal and/or vertical flips, because the visible
    // top-left corner may no longer correspond to the stored left/top values.
    //
    // The mapping itself is KiCadCalibrationGeometry.RemapDragModeForFlip - eight handles across
    // four flip states, which is exactly the kind of table that needs tests rather than eyes.
    // ###########################################################################################
    private LabelEditorDragMode RemapKiCadTraceCalibrationDragModeForCurrentFlip(LabelEditorDragMode dragMode) =>
        KiCadCalibrationGeometry.RemapDragModeForFlip(this.GetKiCadCalibrationBox(), dragMode);

    // ###########################################################################################
    // Starts a KiCad calibration move or resize drag by capturing both the pointer start pixel and
    // the current box edges so drag updates remain stable and do not accumulate rounding drift.
    // Visual resize handles are remapped to the stored flipped edge/corner definition first.
    // ###########################################################################################
    private void StartKiCadTraceCalibrationDrag(Point startPixelPoint, LabelEditorDragMode dragMode)
    {
        this.thisKiCadTraceCalibrationDragMode =
            dragMode == LabelEditorDragMode.Move
                ? LabelEditorDragMode.Move
                : this.RemapKiCadTraceCalibrationDragModeForCurrentFlip(dragMode);

        this.thisKiCadTraceCalibrationDragStartPixelPoint = startPixelPoint;
        this.thisKiCadCalibrationStartImageLeft = this.thisKiCadCalibrationImageLeft;
        this.thisKiCadCalibrationStartImageTop = this.thisKiCadCalibrationImageTop;
        this.thisKiCadCalibrationStartImageRight = this.thisKiCadCalibrationImageRight;
        this.thisKiCadCalibrationStartImageBottom = this.thisKiCadCalibrationImageBottom;

        // One full rebuild at the box's starting position, so the geometry the rest of this drag
        // transforms is known to exist and is known to match the reference recorded by
        // NoteKiCadOverlayBuiltCalibration. Every move after this is a matrix update.
        this.RefreshKiCadOverlay(forceImmediate: true);
    }

    // ###########################################################################################
    // Updates the temporary KiCad calibration box during drag. Moving preserves the box size while
    // resize modes allow edge crossing so horizontal and vertical flipping happen automatically.
    // ###########################################################################################
    private void UpdateKiCadTraceCalibrationDrag(Point currentPixelPoint)
    {
        if (!this.thisIsKiCadTraceCalibrationMode ||
            this.thisKiCadTraceCalibrationDragMode == LabelEditorDragMode.None)
        {
            return;
        }

        // Offset from where the drag STARTED, applied to the box as it stood then - not an
        // increment on the current box. See KiCadCalibrationGeometry.ApplyDrag for why.
        double dx = currentPixelPoint.X - this.thisKiCadTraceCalibrationDragStartPixelPoint.X;
        double dy = currentPixelPoint.Y - this.thisKiCadTraceCalibrationDragStartPixelPoint.Y;

        this.SetKiCadCalibrationBox(
            KiCadCalibrationGeometry.ApplyDrag(
                this.GetKiCadCalibrationStartBox(),
                this.thisKiCadTraceCalibrationDragMode,
                dx,
                dy));

        this.RefreshKiCadTraceCalibrationDragOverlay();
    }

    // ###########################################################################################
    // Redraws the calibration overlay during a drag WITHOUT rebuilding the trace/pad geometry.
    //
    // The traces were built once at the drag's starting calibration; this only works out how that
    // geometry has to be shifted and scaled to match the box's current position, hands the result
    // to the render control as a matrix, and rebuilds the box/handles (which are a handful of
    // rectangles, not thousands of primitives).
    //
    // Falls back to a full rebuild whenever the delta cannot be built - no reference recorded yet,
    // or a degenerate calibration. Correct output always wins over a fast one.
    // ###########################################################################################
    private void RefreshKiCadTraceCalibrationDragOverlay()
    {
        if (!this.TryApplyKiCadTraceCalibrationDragTransform())
        {
            this.RefreshKiCadOverlay(forceImmediate: true);
            return;
        }

        this.RefreshKiCadCalibrationBoxPrimitiveOnly();
    }

    // ###########################################################################################
    // Works out the drag-time overlay transform and applies it, returning false when there is no
    // usable reference to transform from.
    // ###########################################################################################
    private bool TryApplyKiCadTraceCalibrationDragTransform()
    {
        if (this.thisKiCadTraceCalibrationDragReferenceCalibration == null ||
            this.currentFullResBitmap == null)
        {
            return false;
        }

        var current = this.GetKiCadViewCalibration(this.GetCurrentSchematicName());

        if (!KiCadCalibrationTransform.TryBuildDeltaTransform(
                this.thisKiCadTraceCalibrationDragReferenceCalibration.Value,
                new KiCadCalibrationValues(
                    current.ScaleX,
                    current.ScaleY,
                    current.OffsetX,
                    current.OffsetY,
                    current.MirrorX,
                    current.MirrorY),
                this.thisKiCadTraceCalibrationDragReferenceWorldBounds,
                this.thisKiCadTraceCalibrationDragReferenceContentRect,
                this.currentFullResBitmap.PixelSize.Width,
                this.currentFullResBitmap.PixelSize.Height,
                out var transform))
        {
            return false;
        }

        this.SchematicsKiCadOverlayCanvas.OverlayTransform = transform;
        return true;
    }

    // ###########################################################################################
    // Records the calibration (and the world/content rects) that the overlay geometry currently on
    // screen was built against, so a drag can transform that geometry instead of rebuilding it.
    //
    // Called from both render paths rather than from the drag, because only the render knows what
    // it actually baked in - and a rebuild triggered by anything else mid-drag (a hover, a
    // background net cache finishing) must re-point the reference at the new geometry, or the next
    // move would transform from a calibration that is no longer on screen.
    //
    // It also clears any transform already applied: the geometry has just been rebuilt at the
    // current calibration, so it needs no shifting, and leaving a stale matrix on the control would
    // double-apply the drag so far.
    // ###########################################################################################
    private void NoteKiCadOverlayBuiltCalibration(
        KiCadViewCalibration calibration,
        Rect worldBounds,
        Rect contentRect)
    {
        this.thisKiCadTraceCalibrationDragReferenceCalibration = new KiCadCalibrationValues(
            calibration.ScaleX,
            calibration.ScaleY,
            calibration.OffsetX,
            calibration.OffsetY,
            calibration.MirrorX,
            calibration.MirrorY);

        this.thisKiCadTraceCalibrationDragReferenceWorldBounds = worldBounds;
        this.thisKiCadTraceCalibrationDragReferenceContentRect = contentRect;

        this.SchematicsKiCadOverlayCanvas.OverlayTransform = Matrix.Identity;
    }

    // ###########################################################################################
    // Drops the drag-time reference and any transform applied from it. Anything that changes which
    // geometry is on screen, or that ends calibration, goes through here so a later drag can never
    // transform from a reference belonging to different geometry.
    // ###########################################################################################
    private void ClearKiCadTraceCalibrationDragTransform()
    {
        this.thisKiCadTraceCalibrationDragReferenceCalibration = null;
        this.thisKiCadTraceCalibrationDragReferenceWorldBounds = default;
        this.thisKiCadTraceCalibrationDragReferenceContentRect = default;

        this.SchematicsKiCadOverlayCanvas.OverlayTransform = Matrix.Identity;
    }

    // ###########################################################################################
    // Replaces just the calibration box/handles primitive in the already-drawn overlay, leaving
    // whichever trace/pad primitives were last rebuilt untouched. RefreshKiCadOverlayNow always
    // appends the box as the LAST primitive (calibration mode's branch), which is what makes this
    // safe - there is exactly one to replace and it is always at the end.
    // ###########################################################################################
    private void RefreshKiCadCalibrationBoxPrimitiveOnly()
    {
        var primitives = this.SchematicsKiCadOverlayCanvas.Primitives;
        var updatedPrimitives = primitives.Count > 0
            ? primitives.Take(primitives.Count - 1).ToList()
            : new List<KiCadOverlayPrimitive>();

        updatedPrimitives.Add(this.BuildKiCadCalibrationBoxPrimitiveForCurrentOverlayTransform());
        this.SchematicsKiCadOverlayCanvas.SetGeometry(updatedPrimitives);
    }

    // ###########################################################################################
    // The calibration box, pre-compensated for whatever overlay transform is currently applied.
    //
    // The box shares a canvas with the traces, so it is drawn through the same transform - but the
    // traces NEED that transform (they were built at the drag's starting calibration) while the box
    // is built fresh at the box's actual position and needs none. Without this the box would be
    // shifted twice and would drift away from the pointer as the drag went on.
    //
    // Pre-multiplying by the inverse cancels the transform for this one primitive. A non-invertible
    // transform cannot arise from a usable calibration - TryBuildDeltaTransform refuses to produce
    // one - but it falls back to the plain box rather than dropping the box entirely.
    // ###########################################################################################
    private KiCadOverlayPrimitive BuildKiCadCalibrationBoxPrimitiveForCurrentOverlayTransform()
    {
        var primitive = this.BuildKiCadCalibrationBoxPrimitive();
        var overlayTransform = this.SchematicsKiCadOverlayCanvas.OverlayTransform;

        if (overlayTransform == Matrix.Identity ||
            primitive.Geometry == null ||
            !overlayTransform.TryInvert(out var inverseTransform))
        {
            return primitive;
        }

        var compensatedGeometry = primitive.Geometry.Clone();
        compensatedGeometry.Transform = new MatrixTransform(inverseTransform);

        return new KiCadOverlayPrimitive
        {
            Kind = primitive.Kind,
            Geometry = compensatedGeometry,
            Pen = primitive.Pen,
            Fill = primitive.Fill
        };
    }

    // ###########################################################################################
    // Completes the active KiCad calibration drag, clears the transient drag mode, and forces one
    // final full-quality refresh - the throttle above can otherwise leave the traces/pads one step
    // behind the box's true resting position after the last pointer move of the drag.
    // ###########################################################################################
    private void CompleteKiCadTraceCalibrationDrag()
    {
        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;

        // One real rebuild at the box's final position, which also re-points the drag reference at
        // the freshly built geometry and clears the transform (NoteKiCadOverlayBuiltCalibration).
        // The transformed overlay is geometrically identical to this, so nothing visibly jumps -
        // but from here on the primitives are genuinely built at the calibration they represent,
        // which is what the hover hit-test caches and the next drag both need.
        this.RefreshKiCadOverlay(forceImmediate: true);
    }

    // ###########################################################################################
    // Updates the cursor while KiCad trace calibration mode is active so resize handles and move
    // areas feel consistent with the component label editor interactions.
    // ###########################################################################################
    private void UpdateKiCadTraceCalibrationCursor(Point pointerInContainer)
    {
        if (!this.thisIsKiCadTraceCalibrationMode)
        {
            return;
        }

        if (this.thisKiCadTraceCalibrationDragMode != LabelEditorDragMode.None)
        {
            this.SchematicsContainer.Cursor = this.thisKiCadTraceCalibrationDragMode == LabelEditorDragMode.Move
                ? new Cursor(StandardCursorType.SizeAll)
                : new Cursor(StandardCursorType.Hand);
            return;
        }

        if (this.TryGetKiCadTraceCalibrationHandleAtContainerPoint(pointerInContainer, out _))
        {
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.Hand);
            return;
        }

        if (this.IsPointerInsideCurrentKiCadCalibrationBounds(pointerInContainer))
        {
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);
            return;
        }

        this.SchematicsContainer.Cursor = Cursor.Default;
    }
}