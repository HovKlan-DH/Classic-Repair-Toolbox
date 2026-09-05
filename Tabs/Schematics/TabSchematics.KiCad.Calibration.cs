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
        if (BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                excelPath,
                schematicName,
                out _,
                out _,
                out _,
                out _,
                out _,
                out bool mirrorX,
                out bool mirrorY))
        {
            this.ApplyKiCadCalibrationMirrorFlagsToBox(mirrorX, mirrorY);
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

        this.CheckGlobalShowCalibrationTracesAndPads.IsChecked = true;

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

        this.thisIsKiCadTraceCalibrationMode = false;
        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;
        this.CheckGlobalShowCalibrationTracesAndPads.IsChecked = true;

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

        var thisHandleBrush = new SolidColorBrush(Colors.LimeGreen, 1.0);
        var thisHandlePen = new Pen(thisHandleBrush, 1.0);
        var thisBorderPen = new Pen(thisHandleBrush, 1.0);

        var thisPrimitives = new List<KiCadOverlayPrimitive>
    {
        new KiCadOverlayPrimitive
        {
            Kind = KiCadOverlayPrimitiveKind.Rectangle,
            Rect = thisLocalRect,
            Pen = thisBorderPen,
            Fill = null
        }
    };

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

        foreach (var thisHandleCenter in thisHandleCenters)
        {
            thisPrimitives.Add(new KiCadOverlayPrimitive
            {
                Kind = KiCadOverlayPrimitiveKind.Rectangle,
                Rect = new Rect(
                    thisHandleCenter.X - thisHalfHandleSize,
                    thisHandleCenter.Y - thisHalfHandleSize,
                    thisHandleSize,
                    thisHandleSize),
                Pen = thisHandlePen,
                Fill = thisHandleBrush
            });
        }

        var thisGeometry = new StreamGeometry();

        using (var thisGeometryContext = thisGeometry.Open())
        {
            foreach (var thisPrimitive in thisPrimitives)
            {
                if (thisPrimitive.Kind != KiCadOverlayPrimitiveKind.Rectangle)
                {
                    continue;
                }

                thisGeometryContext.BeginFigure(thisPrimitive.Rect.TopLeft, isFilled: thisPrimitive.Fill != null);
                thisGeometryContext.LineTo(thisPrimitive.Rect.TopRight);
                thisGeometryContext.LineTo(thisPrimitive.Rect.BottomRight);
                thisGeometryContext.LineTo(thisPrimitive.Rect.BottomLeft);
                thisGeometryContext.EndFigure(isClosed: true);
            }
        }

        return new KiCadOverlayPrimitive
        {
            Kind = KiCadOverlayPrimitiveKind.Geometry,
            Geometry = thisGeometry,
            Pen = thisBorderPen,
            Fill = null
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

        this.RefreshKiCadOverlay(forceImmediate: true);
    }

    // ###########################################################################################
    // Completes the active KiCad calibration drag and clears the transient drag mode so the overlay
    // returns to idle calibration interaction state.
    // ###########################################################################################
    private void CompleteKiCadTraceCalibrationDrag()
    {
        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;
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