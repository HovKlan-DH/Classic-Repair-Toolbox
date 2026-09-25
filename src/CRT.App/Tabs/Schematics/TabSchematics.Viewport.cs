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
// Zoom, pan and the schematic transform matrix: applying zoom, clamping the matrix to the
// image bounds, and resolving image/content/viewport rectangles and image pixel points.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Zoom
    internal Matrix schematicsMatrix = Matrix.Identity;

    // Panning
    private bool isPanning;

    private Point panStartPoint;

    private Matrix panStartMatrix;

    // ###########################################################################################
    // Applies zoom around a container-space anchor point and reuses the same clamping logic for
    // mouse wheel zoom and pinch zoom gestures.
    // ###########################################################################################
    private void ApplySchematicsZoom(double zoomFactor, Point zoomCenterInContainer)
    {
        if (this.currentFullResBitmap == null)
        {
            return;
        }

        if (double.IsNaN(zoomFactor) || double.IsInfinity(zoomFactor) || zoomFactor <= 0)
        {
            return;
        }


        double currentScale = this.schematicsMatrix.M11;
        double newScale = currentScale * zoomFactor;

        if (newScale > AppConfig.SchematicsMaxZoom)
        {
            zoomFactor = AppConfig.SchematicsMaxZoom / currentScale;
            newScale = currentScale * zoomFactor;
        }

        // The image is already fully fitted by Stretch="Uniform", so do not allow zooming out
        // below the baseline matrix scale of 1.0.
        if (newScale < 1.0)
        {
            this.schematicsMatrix = Matrix.Identity;
            this.ClampSchematicsMatrix();
            return;
        }

        var zoomMatrix =
            Matrix.CreateTranslation(-zoomCenterInContainer.X, -zoomCenterInContainer.Y) *
            Matrix.CreateScale(zoomFactor, zoomFactor) *
            Matrix.CreateTranslation(zoomCenterInContainer.X, zoomCenterInContainer.Y);

        // Apply zoom in container space, matching the same row-vector composition used by panning.
        this.schematicsMatrix = this.schematicsMatrix * zoomMatrix;
        this.ClampSchematicsMatrix();
    }

    // ###########################################################################################
    // Returns the rectangle (in local overlay coordinates) that the actual bitmap content occupies.
    // Must match the overlay renderer mapping exactly, with no centering offset applied.
    // ###########################################################################################
    internal Rect GetImageContentRect()
    {
        return this.GetSchematicsContentRect();
    }

    // ###########################################################################################
    // Computes the schematic image content rect using the same top-left anchored logic as all
    // overlay renderers so labels, hit testing, and editor rectangles always share one mapping.
    // ###########################################################################################
    private Rect GetSchematicsContentRect()
    {
        var bitmap = this.currentFullResBitmap;

        Size controlSize = this.SchematicsHighlightsOverlay.Bounds.Size;
        if (controlSize.Width <= 0 || controlSize.Height <= 0)
        {
            controlSize = this.SchematicsContainer.Bounds.Size;
        }

        if (bitmap == null || controlSize.Width <= 0 || controlSize.Height <= 0)
        {
            return new Rect(controlSize);
        }

        double containerAspect = controlSize.Width / controlSize.Height;
        double bitmapAspect = (double)bitmap.PixelSize.Width / bitmap.PixelSize.Height;

        if (bitmapAspect > containerAspect)
        {
            return new Rect(0, 0, controlSize.Width, controlSize.Width / bitmapAspect);
        }
        else
        {
            return new Rect(0, 0, controlSize.Height * bitmapAspect, controlSize.Height);
        }
    }

    // ###########################################################################################
    // Computes the effective visible viewport inside the schematics container after subtracting
    // only the panel edges that should actually constrain panning. Bottom-docked utility panels
    // reserve bottom space, while the net connections panel reserves right-side space. This avoids
    // false top-edge shrinkage when a corner panel appears, which otherwise causes jumpy panning.
    // ###########################################################################################
    private Rect GetSchematicsVisibleViewportRect()
    {
        Size containerSize = this.SchematicsContainer.Bounds.Size;
        if (containerSize.Width <= 0 || containerSize.Height <= 0)
        {
            return new Rect(containerSize);
        }

        double leftInset = 0.0;
        double topInset = 0.0;
        double rightInset = 0.0;
        double bottomInset = 0.0;

        void IncludeOverlay(
            Control? overlay,
            bool reserveLeft = false,
            bool reserveTop = false,
            bool reserveRight = false,
            bool reserveBottom = false)
        {
            if (overlay == null ||
                !overlay.IsVisible ||
                overlay.Bounds.Width <= 0 ||
                overlay.Bounds.Height <= 0)
            {
                return;
            }

            Point? translatedTopLeft = overlay.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
            if (!translatedTopLeft.HasValue)
            {
                return;
            }

            double left = Math.Max(0.0, translatedTopLeft.Value.X);
            double top = Math.Max(0.0, translatedTopLeft.Value.Y);
            double right = Math.Min(containerSize.Width, translatedTopLeft.Value.X + overlay.Bounds.Width);
            double bottom = Math.Min(containerSize.Height, translatedTopLeft.Value.Y + overlay.Bounds.Height);

            if (right <= left || bottom <= top)
            {
                return;
            }

            if (reserveLeft)
            {
                leftInset = Math.Max(leftInset, right);
            }

            if (reserveTop)
            {
                topInset = Math.Max(topInset, bottom);
            }

            if (reserveRight)
            {
                rightInset = Math.Max(rightInset, containerSize.Width - left);
            }

            if (reserveBottom)
            {
                bottomInset = Math.Max(bottomInset, containerSize.Height - top);
            }
        }

        IncludeOverlay(this.GlobalSettingsPanel, reserveBottom: true);
        IncludeOverlay(this.LabelsPanel, reserveBottom: true);
        IncludeOverlay(this.ImportantSignalsPanel, reserveBottom: true);
        IncludeOverlay(this.TracesPanel, reserveBottom: true);
        IncludeOverlay(this.KiCadNetConnectionsPanel, reserveRight: true);

        double viewportLeft = Math.Clamp(leftInset, 0.0, containerSize.Width);
        double viewportTop = Math.Clamp(topInset, 0.0, containerSize.Height);
        double viewportRight = Math.Clamp(containerSize.Width - rightInset, viewportLeft, containerSize.Width);
        double viewportBottom = Math.Clamp(containerSize.Height - bottomInset, viewportTop, containerSize.Height);

        return new Rect(
            viewportLeft,
            viewportTop,
            Math.Max(1.0, viewportRight - viewportLeft),
            Math.Max(1.0, viewportBottom - viewportTop));
    }

    // ###########################################################################################
    // Clamps the current schematics matrix so the image cannot be moved beyond the currently
    // visible viewport, once edge-docked overlay panels have been accounted for, while still
    // leaving room for whatever position a cursor-anchored zoom asks for.
    // Also allows baseline-scale panning when overlay panels reduce the effectively visible area.
    //
    // The two demands are reconciled per axis in ViewportMath.ComputeAxisTranslationRange - read
    // that before changing anything here, because tightening this clamp is what breaks the
    // guarantee that the point under the mouse pointer stays under the mouse pointer.
    // ###########################################################################################
    private void ClampSchematicsMatrix(bool useStrictEdgeClamp = false)
    {
        Size containerSize = this.SchematicsContainer.Bounds.Size;
        if (containerSize.Width <= 0 || containerSize.Height <= 0)
        {
            return;
        }

        Rect viewportRect = this.GetSchematicsVisibleViewportRect();
        if (viewportRect.Width <= 0 || viewportRect.Height <= 0)
        {
            viewportRect = new Rect(containerSize);
        }

        Rect contentRect = this.GetImageContentRect();

        double scale = this.schematicsMatrix.M11;
        if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0)
        {
            scale = 1.0;
        }

        double tx = this.schematicsMatrix.M31;
        double ty = this.schematicsMatrix.M32;

        const double minimumScaleEpsilon = 0.000001;

        // Both the pannable range and the room a cursor-anchored zoom needs - see
        // ViewportMath.ComputeAxisTranslationRange for why the two have to be taken together.
        (double minTx, double maxTx) = ViewportMath.ComputeAxisTranslationRange(
            viewportRect.Left,
            viewportRect.Right,
            contentRect.Left,
            contentRect.Right,
            scale);

        (double minTy, double maxTy) = ViewportMath.ComputeAxisTranslationRange(
            viewportRect.Top,
            viewportRect.Bottom,
            contentRect.Top,
            contentRect.Bottom,
            scale);

        bool shouldUseStrictEdgeClamp = useStrictEdgeClamp || this.isPanning;

        bool shouldAllowBaselinePanX =
            contentRect.Width > viewportRect.Width + 0.01 ||
            viewportRect.Left > 0.01 ||
            viewportRect.Right < containerSize.Width - 0.01;

        bool shouldAllowBaselinePanY =
            contentRect.Height > viewportRect.Height + 0.01 ||
            viewportRect.Top > 0.01 ||
            viewportRect.Bottom < containerSize.Height - 0.01;

        if (scale <= 1.0 + minimumScaleEpsilon)
        {
            tx = shouldAllowBaselinePanX
                ? Math.Clamp(tx, minTx, maxTx)
                : 0.0;

            ty = shouldAllowBaselinePanY
                ? Math.Clamp(ty, minTy, maxTy)
                : 0.0;
        }
        else if (shouldUseStrictEdgeClamp)
        {
            tx = Math.Clamp(tx, minTx, maxTx);
            ty = Math.Clamp(ty, minTy, maxTy);
        }
        else
        {
            if (tx < minTx)
            {
                tx = minTx;
            }
            else if (tx > maxTx)
            {
                tx = maxTx;
            }

            if (ty < minTy)
            {
                ty = minTy;
            }
            else if (ty > maxTy)
            {
                ty = maxTy;
            }
        }

        this.schematicsMatrix = new Matrix(scale, 0, 0, scale, tx, ty);


        ((MatrixTransform)this.SchematicsImage.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHoverHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsLabelEditorOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsPolylineCanvas.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsLabelsCanvas.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsKiCadOverlayCanvas.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsWorklogEntryOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsWorklogEntryBadgeCanvas.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsWorklogEntriesOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsWorklogEntriesBadgeCanvas.RenderTransform!).Matrix = this.schematicsMatrix;

        // The KiCad overlay culls to the visible area, so it needs the same view matrix the
        // highlights overlay uses. Assigning it does not invalidate anything on its own.
        this.SchematicsKiCadOverlayCanvas.ViewMatrix = this.schematicsMatrix;

        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsHighlightsOverlay.InvalidateVisual();

        this.SchematicsHoverHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsHoverHighlightsOverlay.InvalidateVisual();

        this.SchematicsWorklogEntriesOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsWorklogEntriesOverlay.InvalidateVisual();

        this.SchematicsLabelEditorOverlay.ApplyState(
            rectangles: this.SchematicsLabelEditorOverlay.Rectangles,
            selectedIndex: this.SchematicsLabelEditorOverlay.SelectedIndex,
            selectedIndices: this.SchematicsLabelEditorOverlay.SelectedIndices,
            selectionBounds: this.SchematicsLabelEditorOverlay.SelectionBounds,
            hoveredIndex: this.SchematicsLabelEditorOverlay.HoveredIndex,
            draftRectangle: this.SchematicsLabelEditorOverlay.DraftRectangle,
            snapGuides: this.SchematicsLabelEditorOverlay.SnapGuides,
            bitmapPixelSize: this.SchematicsLabelEditorOverlay.BitmapPixelSize,
            viewMatrix: this.schematicsMatrix,
            highlightColor: this.SchematicsLabelEditorOverlay.HighlightColor,
            highlightOpacity: this.SchematicsLabelEditorOverlay.HighlightOpacity,
            isVisible: this.SchematicsLabelEditorOverlay.IsVisible);

        this.RefreshWorklogEntryOverlay();
        this.RescaleWorklogEntriesListBadges();

        this.polylineManager?.UpdateScaleFactor(scale);

        this.UpdateComponentLabelsScale(scale);

    }

    // ###########################################################################################
    // Converts a schematic container pointer position into bitmap pixel coordinates for the
    // currently displayed schematic image.
    // ###########################################################################################
    private bool TryGetSchematicsImagePixelPoint(Point pointerInContainer, out Point pixelPoint)
    {
        pixelPoint = default;

        if (this.currentFullResBitmap == null)
        {
            return false;
        }

        if (!RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            return false;
        }

        var localPoint = new Point(
            (pointerInContainer.X * inv.M11) + (pointerInContainer.Y * inv.M21) + inv.M31,
            (pointerInContainer.X * inv.M12) + (pointerInContainer.Y * inv.M22) + inv.M32);

        var contentRect = this.GetImageContentRect();
        if (contentRect.Width <= 0 || contentRect.Height <= 0 || !contentRect.Contains(localPoint))
        {
            return false;
        }

        double px = ((localPoint.X - contentRect.X) / contentRect.Width) * this.currentFullResBitmap.PixelSize.Width;
        double py = ((localPoint.Y - contentRect.Y) / contentRect.Height) * this.currentFullResBitmap.PixelSize.Height;

        pixelPoint = new Point(px, py);
        return true;
    }

    // ###########################################################################################
    // The same mapping as TryGetSchematicsImagePixelPoint, but WITHOUT requiring the pointer to be
    // over the image - the returned point may sit outside the bitmap's bounds.
    //
    // Needed by drags that continue past the edge of the board. The bounded version returns false
    // there, so a caller using it simply stops receiving updates: a worklog area being widened
    // froze at the last in-bounds sample instead of continuing to follow the pointer, and the
    // clamp that was written to handle exactly that case never ran. Callers are expected to clamp
    // the RESULT (see LabelEditorGeometry.ClampRectToBounds) rather than the input.
    // ###########################################################################################
    private bool TryGetSchematicsImagePixelPointUnbounded(Point pointerInContainer, out Point pixelPoint)
    {
        pixelPoint = default;

        if (this.currentFullResBitmap == null)
        {
            return false;
        }

        if (!RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            return false;
        }

        var localPoint = new Point(
            (pointerInContainer.X * inv.M11) + (pointerInContainer.Y * inv.M21) + inv.M31,
            (pointerInContainer.X * inv.M12) + (pointerInContainer.Y * inv.M22) + inv.M32);

        var contentRect = this.GetImageContentRect();
        if (contentRect.Width <= 0 || contentRect.Height <= 0)
        {
            return false;
        }

        pixelPoint = new Point(
            ((localPoint.X - contentRect.X) / contentRect.Width) * this.currentFullResBitmap.PixelSize.Width,
            ((localPoint.Y - contentRect.Y) / contentRect.Height) * this.currentFullResBitmap.PixelSize.Height);

        return true;
    }
}