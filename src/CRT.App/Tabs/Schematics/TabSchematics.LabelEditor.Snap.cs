using Avalonia;
using System;
using System.Collections.Generic;
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// The label editor's rim onto the snapping maths.
//
// The maths itself - ~950 lines of it - lives in Handlers/Geometry/LabelEditorSnapGeometry.cs.
// It was private members of this class until it was extracted, where nothing could test it; it
// is pure arithmetic over EditableComponentHighlight rows in bitmap-pixel space and never
// touches a control, so the only thing that has to stay here is reading this tab's own state.
//
// What this file owns is exactly that: building a LabelEditorSnapContext. The one piece with
// real work in it is BuildVisiblePixelRect, which turns the container bounds, the full-res
// bitmap and the view matrix into the visible region in bitmap pixels - the value that decides
// which neighbours are allowed to participate in a snap.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // The selection predicate the context carries, cached.
    //
    // Passing "this.IsSelectedLabelEditorHighlight" as a method group allocates a fresh delegate
    // closing over this on EVERY call, and a context is built per pointer-move event while a drag
    // is running. The delegate never varies - it is always this instance's method - so it is made
    // once. Lazily, because the field initialiser would otherwise run before the rest of
    // construction.
    private Func<EditableComponentHighlight, bool>? thisLabelEditorIsSelectedPredicate;

    // ###########################################################################################
    // Resolves everything the snapping maths needs from this tab into plain values.
    //
    // dragModeOverride lets a caller snap for an edge other than the one currently being dragged.
    // ###########################################################################################
    private LabelEditorSnapContext BuildLabelEditorSnapContext(LabelEditorDragMode? dragModeOverride = null)
    {
        this.thisLabelEditorIsSelectedPredicate ??= this.IsSelectedLabelEditorHighlight;

        return new LabelEditorSnapContext(
            this.thisLabelEditorWorkingHighlights,
            dragModeOverride ?? this.thisLabelEditorDragMode,
            this.GetCurrentSchematicName(),
            this.BuildVisiblePixelRect(),
            this.thisLabelEditorIsSelectedPredicate);
    }

    // ###########################################################################################
    // The part of the schematic bitmap currently visible in the viewport, in BITMAP PIXELS.
    //
    // Only components inside this rect may take part in snapping, so an edge cannot be dragged by
    // a neighbour that is scrolled off screen. Returns null when there is no usable bitmap, the
    // container has not been laid out, or the view matrix cannot be inverted - callers treat that
    // as "no viewport information" and let every row participate, which is what this code did
    // inline before it was pulled out of the two snap methods that each carried a copy of it.
    // ###########################################################################################
    private Rect? BuildVisiblePixelRect()
    {
        if (this.currentFullResBitmap == null ||
            this.SchematicsContainer.Bounds.Width <= 0 ||
            this.SchematicsContainer.Bounds.Height <= 0 ||
            !RectGeometry.TryInvert(this.schematicsMatrix, out var inverseMatrix))
        {
            return null;
        }

        var contentRect = this.GetLabelEditorImageContentRect();

        if (contentRect.Width <= 0 || contentRect.Height <= 0)
        {
            return null;
        }

        var containerRect = new Rect(this.SchematicsContainer.Bounds.Size);
        var visibleLocalRect = containerRect.TransformToAABB(inverseMatrix);

        double clippedLeft = Math.Max(contentRect.Left, visibleLocalRect.Left);
        double clippedTop = Math.Max(contentRect.Top, visibleLocalRect.Top);
        double clippedRight = Math.Min(contentRect.Right, visibleLocalRect.Right);
        double clippedBottom = Math.Min(contentRect.Bottom, visibleLocalRect.Bottom);

        if (clippedRight <= clippedLeft || clippedBottom <= clippedTop)
        {
            return null;
        }

        double bitmapWidth = this.currentFullResBitmap.PixelSize.Width;
        double bitmapHeight = this.currentFullResBitmap.PixelSize.Height;

        double pixelLeft = Math.Clamp(
            ((clippedLeft - contentRect.X) / contentRect.Width) * bitmapWidth,
            0.0,
            bitmapWidth);

        double pixelTop = Math.Clamp(
            ((clippedTop - contentRect.Y) / contentRect.Height) * bitmapHeight,
            0.0,
            bitmapHeight);

        double pixelRight = Math.Clamp(
            ((clippedRight - contentRect.X) / contentRect.Width) * bitmapWidth,
            0.0,
            bitmapWidth);

        double pixelBottom = Math.Clamp(
            ((clippedBottom - contentRect.Y) / contentRect.Height) * bitmapHeight,
            0.0,
            bitmapHeight);

        if (pixelRight <= pixelLeft || pixelBottom <= pixelTop)
        {
            return null;
        }

        return new Rect(
            pixelLeft,
            pixelTop,
            pixelRight - pixelLeft,
            pixelBottom - pixelTop);
    }

    // ###########################################################################################
    // Snaps active resize edges to nearby neighbour edges. See LabelEditorSnapGeometry for the
    // rules; this only supplies the tab state.
    // ###########################################################################################
    private void ApplyLabelEditorResizeSnap(
        EditableComponentHighlight currentHighlight,
        ref double left,
        ref double top,
        ref double right,
        ref double bottom,
        List<(Point Start, Point End)> snapGuides,
        bool suppressSnap,
        bool snapOnMatch = true,
        LabelEditorDragMode? dragModeOverride = null)
    {
        // The mode goes in via the context and nowhere else - ApplyResizeSnap no longer takes a
        // separate override, so the two cannot disagree.
        LabelEditorSnapGeometry.ApplyResizeSnap(
            this.BuildLabelEditorSnapContext(dragModeOverride),
            currentHighlight,
            ref left,
            ref top,
            ref right,
            ref bottom,
            snapGuides,
            suppressSnap,
            snapOnMatch);
    }

    // ###########################################################################################
    // Snaps the moved selection bounds to nearby neighbour edges while preserving the selection
    // layout. See LabelEditorSnapGeometry for the rules.
    // ###########################################################################################
    private void ApplyLabelEditorMoveSnap(
        IReadOnlyList<EditableComponentHighlight> selectedHighlights,
        IReadOnlyDictionary<EditableComponentHighlight, Rect> sourceRects,
        ref Rect movedSelectionBounds,
        List<(Point Start, Point End)> snapGuides,
        bool suppressSnap,
        bool snapOnMatch = true)
    {
        LabelEditorSnapGeometry.ApplyMoveSnap(
            this.BuildLabelEditorSnapContext(),
            selectedHighlights,
            sourceRects,
            ref movedSelectionBounds,
            snapGuides,
            suppressSnap,
            snapOnMatch);
    }

    // ###########################################################################################
    // Snaps a newly drawn rectangle on all four edges. See LabelEditorSnapGeometry for the rules.
    // ###########################################################################################
    private void ApplyNewLabelEditorRectangleSnap(
        ref Rect rect,
        List<(Point Start, Point End)> snapGuides,
        bool suppressSnap)
    {
        LabelEditorSnapGeometry.ApplyNewRectangleSnap(
            this.BuildLabelEditorSnapContext(),
            ref rect,
            snapGuides,
            suppressSnap);
    }
}
