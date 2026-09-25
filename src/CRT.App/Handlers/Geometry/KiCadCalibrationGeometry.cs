using Avalonia.Input;
using System;

namespace Handlers.Geometry
{
    // #######################################################################################
    // The KiCad calibration box, in schematic-image pixels.
    //
    // The four edges are stored exactly as the tab stores them, INCLUDING the deliberate
    // inversion that encodes mirroring: calibration mode represents a horizontally flipped
    // board by holding Left > Right, and a vertically flipped one by holding Top > Bottom.
    // That is why this is a plain four-edge record rather than an Avalonia Rect - a Rect
    // normalises its edges and would silently discard the flip.
    //
    // IsMirroredX/IsMirroredY are DERIVED from that ordering rather than stored alongside it,
    // so the two can never disagree. Normalised* give the same box with the edges put back in
    // ascending order, which is what the arithmetic wants to work in.
    // #######################################################################################
    internal readonly struct KiCadCalibrationBox
    {
        public KiCadCalibrationBox(double left, double top, double right, double bottom)
        {
            this.Left = left;
            this.Top = top;
            this.Right = right;
            this.Bottom = bottom;
        }

        public double Left { get; }

        public double Top { get; }

        public double Right { get; }

        public double Bottom { get; }

        // Mirroring is encoded BY the edge ordering - see the type comment.
        public bool IsMirroredX => this.Left > this.Right;

        public bool IsMirroredY => this.Top > this.Bottom;

        public double NormalisedLeft => Math.Min(this.Left, this.Right);

        public double NormalisedRight => Math.Max(this.Left, this.Right);

        public double NormalisedTop => Math.Min(this.Top, this.Bottom);

        public double NormalisedBottom => Math.Max(this.Top, this.Bottom);

        // ###################################################################################
        // Rebuilds a box from ascending edges, re-applying this box's own mirror flags.
        //
        // This is the counterpart to the Normalised* properties: work in ascending edges, then
        // come back through here so a flipped box stays flipped. Doing it any other way - for
        // instance assigning the ascending values straight back - silently un-mirrors the board.
        // ###################################################################################
        public KiCadCalibrationBox WithNormalisedEdges(double left, double top, double right, double bottom) =>
            new KiCadCalibrationBox(
                this.IsMirroredX ? right : left,
                this.IsMirroredY ? bottom : top,
                this.IsMirroredX ? left : right,
                this.IsMirroredY ? top : bottom);
    }

    // #######################################################################################
    // Pure maths for interactive KiCad trace calibration: remapping a grabbed resize handle
    // once the box has been flipped, nudging the box from the keyboard, and applying a pointer
    // drag to it.
    //
    // This was private members of TabSchematics, where nothing could test it and the file sat
    // at 0% coverage. All three operations are arithmetic over four doubles plus an enum, so
    // they moved here whole; TabSchematics.KiCad.Calibration.cs keeps the rim that reads the
    // tab's fields, guards on mode, and refreshes the overlay afterwards. Same split as
    // LabelEditorSnapGeometry and its TabSchematics.LabelEditor.Snap.cs rim.
    //
    // Everything here works in schematic-image pixels and knows nothing about the viewport.
    // #######################################################################################
    internal static class KiCadCalibrationGeometry
    {
        // How far one arrow-key press moves an edge, in image pixels.
        public const double KeyboardStep = 1.0;

        // ###################################################################################
        // Applies saved mirror flags to a box by swapping the edges they affect.
        //
        // Used when entering calibration mode on a board whose saved calibration is mirrored:
        // the box is seeded from unmirrored bounds, then the flags are folded in here. Since
        // mirroring IS the inverted edge order, applying a flag is simply a swap - and applying
        // the same flag twice returns the original box.
        // ###################################################################################
        public static KiCadCalibrationBox ApplyMirrorFlags(
            KiCadCalibrationBox box,
            bool mirrorX,
            bool mirrorY) =>
            new KiCadCalibrationBox(
                mirrorX ? box.Right : box.Left,
                mirrorY ? box.Bottom : box.Top,
                mirrorX ? box.Left : box.Right,
                mirrorY ? box.Top : box.Bottom);

        // How far in from the image edge the box's own edge must sit before its corner handles
        // are comfortably inside the visible schematic rather than clamped onto the viewport
        // frame - see InsetForInitialVisibility.
        private const double MinimumEdgeMarginFraction = 0.03;

        // ###################################################################################
        // Insets a freshly seeded calibration box so its corner handles do not land exactly on
        // the schematic image's own edges - reported directly: on first-time calibration (no
        // saved calibration yet, so the box starts at the full image bounds) all four handles
        // sat right at the image border, some of them clipped by the viewport frame and nearly
        // impossible to grab.
        //
        // Only pulls an edge IN, and only when it is within the margin of the image's own edge -
        // a box that already has room on a given side (for instance because it was seeded from a
        // saved calibration that only covers part of the board) is left exactly as computed, so
        // this only ever affects the "starts at the full image" case it exists for. The margin is
        // a FRACTION of the image dimension rather than a fixed pixel count, so it scales with
        // both a small preview schematic and a large hi-res board scan alike.
        //
        // Deliberately unaware of mirroring: it works in NORMALISED (ascending) edges and comes
        // back out through WithNormalisedEdges, the same pattern every other edge-arithmetic
        // method here uses, so a mirrored box's flip survives the inset untouched.
        // ###################################################################################
        public static KiCadCalibrationBox InsetForInitialVisibility(
            KiCadCalibrationBox box,
            double imageWidth,
            double imageHeight)
        {
            if (imageWidth <= 0.0 || imageHeight <= 0.0)
            {
                return box;
            }

            double marginX = imageWidth * MinimumEdgeMarginFraction;
            double marginY = imageHeight * MinimumEdgeMarginFraction;

            double left = box.NormalisedLeft;
            double top = box.NormalisedTop;
            double right = box.NormalisedRight;
            double bottom = box.NormalisedBottom;

            if (left <= marginX)
            {
                left = marginX;
            }

            if (top <= marginY)
            {
                top = marginY;
            }

            if (right >= imageWidth - marginX)
            {
                right = imageWidth - marginX;
            }

            if (bottom >= imageHeight - marginY)
            {
                bottom = imageHeight - marginY;
            }

            // A degenerate result (an image too small for both margins to fit) falls back to the
            // box exactly as computed rather than producing an inverted or zero-sized rectangle.
            if (right <= left || bottom <= top)
            {
                return box;
            }

            return box.WithNormalisedEdges(left, top, right, bottom);
        }

        // ###################################################################################
        // Remaps a VISUALLY grabbed resize handle onto the stored edge it actually controls.
        //
        // Once the box is mirrored, the handle the user sees in the top-left corner is no longer
        // the one backed by the stored Left/Top values - on a horizontally flipped box it is
        // backed by Right/Top instead. Without this remap, dragging a corner of a mirrored board
        // resizes the opposite edge, which is the kind of fault that is invisible until someone
        // calibrates a flipped board by hand.
        //
        // An unmirrored box returns the mode untouched. Move is not a resize and is returned
        // untouched too, as is None.
        // ###################################################################################
        public static LabelEditorDragMode RemapDragModeForFlip(
            KiCadCalibrationBox box,
            LabelEditorDragMode dragMode)
        {
            bool mirroredX = box.IsMirroredX;
            bool mirroredY = box.IsMirroredY;

            if (!mirroredX && !mirroredY)
            {
                return dragMode;
            }

            return dragMode switch
            {
                LabelEditorDragMode.ResizeTopLeft => mirroredX
                    ? mirroredY
                        ? LabelEditorDragMode.ResizeBottomRight
                        : LabelEditorDragMode.ResizeTopRight
                    : mirroredY
                        ? LabelEditorDragMode.ResizeBottomLeft
                        : LabelEditorDragMode.ResizeTopLeft,

                LabelEditorDragMode.ResizeTop => mirroredY
                    ? LabelEditorDragMode.ResizeBottom
                    : LabelEditorDragMode.ResizeTop,

                LabelEditorDragMode.ResizeTopRight => mirroredX
                    ? mirroredY
                        ? LabelEditorDragMode.ResizeBottomLeft
                        : LabelEditorDragMode.ResizeTopLeft
                    : mirroredY
                        ? LabelEditorDragMode.ResizeBottomRight
                        : LabelEditorDragMode.ResizeTopRight,

                LabelEditorDragMode.ResizeRight => mirroredX
                    ? LabelEditorDragMode.ResizeLeft
                    : LabelEditorDragMode.ResizeRight,

                LabelEditorDragMode.ResizeBottomRight => mirroredX
                    ? mirroredY
                        ? LabelEditorDragMode.ResizeTopLeft
                        : LabelEditorDragMode.ResizeBottomLeft
                    : mirroredY
                        ? LabelEditorDragMode.ResizeTopRight
                        : LabelEditorDragMode.ResizeBottomRight,

                LabelEditorDragMode.ResizeBottom => mirroredY
                    ? LabelEditorDragMode.ResizeTop
                    : LabelEditorDragMode.ResizeBottom,

                LabelEditorDragMode.ResizeBottomLeft => mirroredX
                    ? mirroredY
                        ? LabelEditorDragMode.ResizeTopRight
                        : LabelEditorDragMode.ResizeBottomRight
                    : mirroredY
                        ? LabelEditorDragMode.ResizeTopLeft
                        : LabelEditorDragMode.ResizeBottomLeft,

                LabelEditorDragMode.ResizeLeft => mirroredX
                    ? LabelEditorDragMode.ResizeRight
                    : LabelEditorDragMode.ResizeLeft,

                _ => dragMode
            };
        }

        // ###################################################################################
        // Applies one arrow-key press to the box, matching the component label editor's own
        // keyboard behaviour: a bare arrow MOVES by one pixel, Shift EXPANDS in the pressed
        // direction, and Alt SHRINKS from the side opposite the pressed direction.
        //
        // Returns false and leaves updatedBox at the input when the press changes nothing - a
        // non-arrow key, Shift and Alt held together (ambiguous, so deliberately refused), or an
        // Alt shrink that would collapse the box to nothing. The caller uses that to decide
        // whether to mark the key handled and repaint.
        //
        // The maths runs on ASCENDING edges and comes back through WithNormalisedEdges, so a
        // mirrored box is nudged in the direction the user sees on screen and stays mirrored.
        // ###################################################################################
        public static bool TryApplyKeyboardStep(
            KiCadCalibrationBox box,
            Key key,
            KeyModifiers modifiers,
            out KiCadCalibrationBox updatedBox)
        {
            updatedBox = box;

            bool isShift = modifiers.HasFlag(KeyModifiers.Shift);
            bool isAlt = modifiers.HasFlag(KeyModifiers.Alt);

            // Expand and shrink at once has no meaning; refuse rather than pick one.
            if (isShift && isAlt)
            {
                return false;
            }

            double left = box.NormalisedLeft;
            double right = box.NormalisedRight;
            double top = box.NormalisedTop;
            double bottom = box.NormalisedBottom;

            bool changed = false;

            if (!isShift && !isAlt)
            {
                switch (key)
                {
                    case Key.Left:
                        left -= KeyboardStep;
                        right -= KeyboardStep;
                        changed = true;
                        break;

                    case Key.Right:
                        left += KeyboardStep;
                        right += KeyboardStep;
                        changed = true;
                        break;

                    case Key.Up:
                        top -= KeyboardStep;
                        bottom -= KeyboardStep;
                        changed = true;
                        break;

                    case Key.Down:
                        top += KeyboardStep;
                        bottom += KeyboardStep;
                        changed = true;
                        break;
                }
            }
            else if (isShift)
            {
                switch (key)
                {
                    case Key.Left:
                        left -= KeyboardStep;
                        changed = true;
                        break;

                    case Key.Right:
                        right += KeyboardStep;
                        changed = true;
                        break;

                    case Key.Up:
                        top -= KeyboardStep;
                        changed = true;
                        break;

                    case Key.Down:
                        bottom += KeyboardStep;
                        changed = true;
                        break;
                }
            }
            else
            {
                // Alt shrinks, and every arm refuses to shrink a box that is already one step
                // wide or tall - otherwise the box collapses to zero and cannot be grabbed again.
                switch (key)
                {
                    case Key.Left:
                        if ((right - left) > KeyboardStep)
                        {
                            right -= KeyboardStep;
                            changed = true;
                        }

                        break;

                    case Key.Right:
                        if ((right - left) > KeyboardStep)
                        {
                            left += KeyboardStep;
                            changed = true;
                        }

                        break;

                    case Key.Up:
                        if ((bottom - top) > KeyboardStep)
                        {
                            bottom -= KeyboardStep;
                            changed = true;
                        }

                        break;

                    case Key.Down:
                        if ((bottom - top) > KeyboardStep)
                        {
                            top += KeyboardStep;
                            changed = true;
                        }

                        break;
                }
            }

            if (!changed)
            {
                return false;
            }

            updatedBox = box.WithNormalisedEdges(left, top, right, bottom);
            return true;
        }

        // ###################################################################################
        // Applies a pointer drag to the box.
        //
        // startBox is the box as it stood when the drag began and (dx, dy) the total offset from
        // the drag's start point - not the offset since the last move event. Accumulating deltas
        // per event would let rounding drift build up over a long drag; re-deriving from the
        // start each time cannot.
        //
        // Resize modes deliberately allow an edge to cross the one opposite it. That is how
        // flipping happens: drag the left edge past the right one and the box comes back with
        // Left > Right, which IS the mirrored-X encoding. So no clamping here, on purpose.
        // ###################################################################################
        public static KiCadCalibrationBox ApplyDrag(
            KiCadCalibrationBox startBox,
            LabelEditorDragMode dragMode,
            double dx,
            double dy)
        {
            double left = startBox.Left;
            double top = startBox.Top;
            double right = startBox.Right;
            double bottom = startBox.Bottom;

            switch (dragMode)
            {
                case LabelEditorDragMode.Move:
                    left += dx;
                    right += dx;
                    top += dy;
                    bottom += dy;
                    break;

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
            }

            return new KiCadCalibrationBox(left, top, right, bottom);
        }
    }
}
