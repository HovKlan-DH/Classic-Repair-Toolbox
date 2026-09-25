using Avalonia;
using Avalonia.Input;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    internal enum LabelEditorDragMode
    {
        None,
        Move,
        ResizeTopLeft,
        ResizeTop,
        ResizeTopRight,
        ResizeRight,
        ResizeBottomRight,
        ResizeBottom,
        ResizeBottomLeft,
        ResizeLeft
    }
    // ###########################################################################################
    // Geometry for the component label editor: which resize handle is under the pointer, which
    // keyboard chord maps to which resize direction, and whether a drawn rectangle is usable.
    //
    // Extracted from TabSchematics so the handle layout and keyboard mapping can be tested.
    // ###########################################################################################
    internal static class LabelEditorGeometry
    {

    // ###########################################################################################
    // Maps one keyboard resize gesture to the equivalent editor drag mode so exact-match guides
    // can be shown without applying any mouse-style snap movement.
    // ###########################################################################################
    public static bool TryGetKeyboardLabelEditorResizeDragMode(Key key, KeyModifiers modifiers, out LabelEditorDragMode dragMode)
    {
        dragMode = LabelEditorDragMode.None;

        bool isShift = modifiers.HasFlag(KeyModifiers.Shift);
        bool isAlt = modifiers.HasFlag(KeyModifiers.Alt);

        if (isShift == isAlt)
        {
            return false;
        }

        switch (key)
        {
            case Key.Left:
                dragMode = isShift
                    ? LabelEditorDragMode.ResizeLeft
                    : LabelEditorDragMode.ResizeRight;
                return true;

            case Key.Right:
                dragMode = isShift
                    ? LabelEditorDragMode.ResizeRight
                    : LabelEditorDragMode.ResizeLeft;
                return true;

            case Key.Up:
                dragMode = isShift
                    ? LabelEditorDragMode.ResizeTop
                    : LabelEditorDragMode.ResizeBottom;
                return true;

            case Key.Down:
                dragMode = isShift
                    ? LabelEditorDragMode.ResizeBottom
                    : LabelEditorDragMode.ResizeTop;
                return true;

            default:
                return false;
        }
    }

    // ###########################################################################################
    // Returns true when a newly drawn editor rectangle is too small to be considered intentional.
    // This prevents accidental tiny drags from opening the new-label prompt.
    // ###########################################################################################
    public static bool IsLabelEditorRectangleTooSmall(Rect rect)
    {
        const double minimumWidth = 15.0;
        const double minimumHeight = 15.0;
        const double minimumArea = minimumWidth * minimumHeight;

        return rect.Width < minimumWidth ||
               rect.Height < minimumHeight ||
               (rect.Width * rect.Height) < minimumArea;
    }

    // ###########################################################################################
    // Applies a resize drag to a rectangle and returns the result, clamped so it can never invert
    // or collapse below minimumSize.
    //
    // The edge being dragged is the one that moves; the opposite edge stays put. When a drag would
    // push the moving edge past the fixed one, the moving edge is pinned at the minimum instead of
    // being allowed through - dragging the left edge rightwards past the right edge would otherwise
    // produce a negative width, which Rect represents as a rectangle mirrored about its origin.
    //
    // Pure, so the worklog overlay and the label editor can share one definition of what a resize
    // means rather than each growing their own subtly different arithmetic.
    // ###########################################################################################
    public static Rect ResizeRect(Rect original, LabelEditorDragMode dragMode, double dx, double dy, double minimumSize)
    {
        double left = original.Left;
        double top = original.Top;
        double right = original.Right;
        double bottom = original.Bottom;

        switch (dragMode)
        {
            case LabelEditorDragMode.ResizeTopLeft:
                left += dx;
                top += dy;
                break;

            case LabelEditorDragMode.ResizeTop:
                top += dy;
                break;

            case LabelEditorDragMode.ResizeTopRight:
                right += dx;
                top += dy;
                break;

            case LabelEditorDragMode.ResizeRight:
                right += dx;
                break;

            case LabelEditorDragMode.ResizeBottomRight:
                right += dx;
                bottom += dy;
                break;

            case LabelEditorDragMode.ResizeBottom:
                bottom += dy;
                break;

            case LabelEditorDragMode.ResizeBottomLeft:
                left += dx;
                bottom += dy;
                break;

            case LabelEditorDragMode.ResizeLeft:
                left += dx;
                break;

            case LabelEditorDragMode.Move:
                return new Rect(original.X + dx, original.Y + dy, original.Width, original.Height);

            default:
                return original;
        }

        double minimum = Math.Max(0.0, minimumSize);

        // Pin whichever edge the drag was moving, so the rectangle shrinks to the minimum and stops
        // rather than turning inside out.
        if (right < left + minimum)
        {
            if (dragMode == LabelEditorDragMode.ResizeLeft ||
                dragMode == LabelEditorDragMode.ResizeTopLeft ||
                dragMode == LabelEditorDragMode.ResizeBottomLeft)
            {
                left = right - minimum;
            }
            else
            {
                right = left + minimum;
            }
        }

        if (bottom < top + minimum)
        {
            if (dragMode == LabelEditorDragMode.ResizeTop ||
                dragMode == LabelEditorDragMode.ResizeTopLeft ||
                dragMode == LabelEditorDragMode.ResizeTopRight)
            {
                top = bottom - minimum;
            }
            else
            {
                bottom = top + minimum;
            }
        }

        return new Rect(left, top, right - left, bottom - top);
    }

    // ###########################################################################################
    // Clamps a rectangle so it stays inside the source image. A worklog area dragged past the edge
    // of the board would otherwise be saved with coordinates outside the bitmap, where it cannot be
    // seen or grabbed again.
    // ###########################################################################################
    public static Rect ClampRectToBounds(Rect rect, Size bounds)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            return rect;
        }

        double width = Math.Min(rect.Width, bounds.Width);
        double height = Math.Min(rect.Height, bounds.Height);

        double x = Math.Clamp(rect.X, 0, Math.Max(0, bounds.Width - width));
        double y = Math.Clamp(rect.Y, 0, Math.Max(0, bounds.Height - height));

        return new Rect(x, y, width, height);
    }

    // ###########################################################################################
    // Clamps a RESIZED rectangle to the image by trimming the edges that stray outside, leaving
    // every other edge exactly where the resize put it.
    //
    // ClampRectToBounds is a translate-clamp: it preserves width and height and slides the whole
    // rectangle back inside. That is right for a Move, and wrong for a resize - dragging the left
    // edge out past x=0 produced Rect(-40, .., 240, ..), which slid to x=0 and pushed the RIGHT
    // edge from 200 to 240. The user dragged one handle and the opposite edge moved, breaking the
    // same anchored-edge property ResizeRect is careful to guarantee.
    //
    // Trimming instead keeps the anchored edge fixed and simply refuses to let the dragged one
    // leave the board. minimumSize is honoured, so an edge dragged far outside collapses to the
    // minimum against the boundary rather than inverting.
    // ###########################################################################################
    public static Rect ClampResizedRectToBounds(Rect rect, Size bounds, double minimumSize)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            return rect;
        }

        double minimum = Math.Max(0.0, minimumSize);

        double left = Math.Clamp(rect.Left, 0, Math.Max(0, bounds.Width - minimum));
        double right = Math.Clamp(rect.Right, minimum, bounds.Width);
        double top = Math.Clamp(rect.Top, 0, Math.Max(0, bounds.Height - minimum));
        double bottom = Math.Clamp(rect.Bottom, minimum, bounds.Height);

        if (right < left + minimum)
        {
            right = Math.Min(bounds.Width, left + minimum);
        }

        if (bottom < top + minimum)
        {
            bottom = Math.Min(bounds.Height, top + minimum);
        }

        return new Rect(left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));
    }

    // ###########################################################################################
    // Builds non-overlapping resize-handle hit rectangles so corner drags keep two-axis behavior
    // even when the selected component is too small for all handle zones to coexist.
    // ###########################################################################################
    public static List<(Rect HitRect, LabelEditorDragMode DragMode)> BuildLabelEditorHandleHitRects(Rect localRect, double scale)
    {
        double cornerHandleSize = Math.Clamp(9.0 / scale, 4.0, 12.0);
        double sideHandleThickness = Math.Clamp(6.0 / scale, 2.5, 7.0);
        double cornerHalf = cornerHandleSize / 2.0;
        double sideHalf = sideHandleThickness / 2.0;
        double minimumGap = Math.Clamp(2.0 / scale, 1.0, 3.0);

        var hitRects = new List<(Rect HitRect, LabelEditorDragMode DragMode)>(8)
        {
            (new Rect(localRect.Left - cornerHalf, localRect.Top - cornerHalf, cornerHandleSize, cornerHandleSize), LabelEditorDragMode.ResizeTopLeft),
            (new Rect(localRect.Right - cornerHalf, localRect.Top - cornerHalf, cornerHandleSize, cornerHandleSize), LabelEditorDragMode.ResizeTopRight),
            (new Rect(localRect.Right - cornerHalf, localRect.Bottom - cornerHalf, cornerHandleSize, cornerHandleSize), LabelEditorDragMode.ResizeBottomRight),
            (new Rect(localRect.Left - cornerHalf, localRect.Bottom - cornerHalf, cornerHandleSize, cornerHandleSize), LabelEditorDragMode.ResizeBottomLeft)
        };

        double horizontalSideHitLength = Math.Max(0.0, localRect.Width - (cornerHandleSize * 2.0) - minimumGap);
        if (horizontalSideHitLength > 0.0)
        {
            double horizontalSideLeft = localRect.Center.X - (horizontalSideHitLength / 2.0);

            hitRects.Add((new Rect(horizontalSideLeft, localRect.Top - sideHalf, horizontalSideHitLength, sideHandleThickness), LabelEditorDragMode.ResizeTop));
            hitRects.Add((new Rect(horizontalSideLeft, localRect.Bottom - sideHalf, horizontalSideHitLength, sideHandleThickness), LabelEditorDragMode.ResizeBottom));
        }

        double verticalSideHitLength = Math.Max(0.0, localRect.Height - (cornerHandleSize * 2.0) - minimumGap);
        if (verticalSideHitLength > 0.0)
        {
            double verticalSideTop = localRect.Center.Y - (verticalSideHitLength / 2.0);

            hitRects.Add((new Rect(localRect.Right - sideHalf, verticalSideTop, sideHandleThickness, verticalSideHitLength), LabelEditorDragMode.ResizeRight));
            hitRects.Add((new Rect(localRect.Left - sideHalf, verticalSideTop, sideHandleThickness, verticalSideHitLength), LabelEditorDragMode.ResizeLeft));
        }

        return hitRects;
    }
    }
}