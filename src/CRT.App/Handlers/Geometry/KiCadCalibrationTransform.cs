using Avalonia;
using System;

namespace Handlers.Geometry
{
    // #######################################################################################
    // The calibration half of the KiCad world-to-local mapping, expressed as a MATRIX.
    //
    // Why this exists
    // ---------------
    // The overlay bakes the active calibration into every coordinate it builds: each pad centre,
    // trace point, arc sample and zone polygon vertex is mapped world-to-local up front, and the
    // resulting primitives are cached per net. That is fine while the calibration holds still,
    // but dragging the calibration box changes it on every pointer move, and the calibration
    // values sit in the per-net cache's generation key (see KiCadOverlayCacheKeys) - so every
    // frame dropped all 451 nets and rebuilt identical shapes at slightly shifted coordinates.
    // Measured on a real board that was ~390-540 ms per frame, with the per-net cache recording
    // 0 hits and 451 misses every single time.
    //
    // The mapping is a plain affine transform though - a scale, an optional mirror, and a
    // translation - with NO rotation and no shear. So the whole calibration can be lifted out of
    // the baked coordinates and applied at RENDER time instead, exactly as zoom and pan already
    // are. Dragging the box then costs a matrix update rather than a rebuild.
    //
    // What the caller does with it
    // ----------------------------
    // Build the primitives ONCE at some reference calibration, then for any other calibration ask
    // for BuildDeltaTransform(reference, current, ...). The result maps the already-built local
    // coordinates onto where that same geometry would have landed had it been built at `current`.
    // Handing back Matrix.Identity when the two calibrations match is not a special case, it falls
    // out of the algebra, and it means a drag that returns to its starting box draws exactly what
    // it started with.
    //
    // The algebra, so it can be checked rather than trusted
    // -----------------------------------------------------
    // TabSchematics.MapKiCadWorldToLocal computes, for X (Y is the same with its own terms):
    //
    //     nx      = (worldX - worldBounds.X) / worldBounds.Width     // 0..1 across the board
    //     nx      = 1 - nx                                           // only when MirrorX
    //     nx      = nx * ScaleX
    //     localX  = contentRect.X + nx * contentRect.Width + OffsetX * (contentRect.Width / pixelWidth)
    //
    // Substituting and collecting terms in worldX gives localX = A * worldX + B, with
    //
    //     A = (MirrorX ? -1 : 1) * ScaleX * contentRect.Width / worldBounds.Width
    //     B = contentRect.X
    //       + (MirrorX ? ScaleX * contentRect.Width : 0)
    //       - (MirrorX ? -1 : 1) * ScaleX * contentRect.Width * worldBounds.X / worldBounds.Width
    //       + OffsetX * contentRect.Width / pixelWidth
    //
    // That is the forward map for ONE calibration. To turn coordinates built at calibration P into
    // coordinates for calibration Q we need Q(P⁻¹(local)), which for two affine maps
    // localP = Ap*world + Bp and localQ = Aq*world + Bq is itself affine:
    //
    //     localQ = (Aq/Ap) * localP + (Bq - Bp * Aq/Ap)
    //
    // which is a scale of Aq/Ap about the origin followed by that translation. Ap is never zero
    // for a usable calibration (a zero ScaleX would collapse the board to a line), but it is
    // guarded anyway rather than producing an infinity that silently poisons every coordinate.
    // #######################################################################################
    public static class KiCadCalibrationTransform
    {
        // Below this the forward map is treated as degenerate: the board would be collapsed to a
        // line or a point, so there is no meaningful delta to build from it.
        private const double MinimumUsableScale = 1e-9;

        // ###################################################################################
        // The forward world-to-local map for one axis, as the pair (coefficient, constant) from
        // the derivation above. Exposed because the delta below is built entirely out of it, and
        // because it is far easier to test one axis of a known calibration than a whole matrix.
        // ###################################################################################
        public static (double Coefficient, double Constant) BuildAxisMap(
            double calibrationScale,
            double calibrationOffset,
            bool isMirrored,
            double worldBoundsOrigin,
            double worldBoundsSize,
            double contentOrigin,
            double contentSize,
            double imagePixelSize)
        {
            if (worldBoundsSize <= 0.0 || contentSize <= 0.0)
            {
                return (0.0, contentOrigin);
            }

            double direction = isMirrored ? -1.0 : 1.0;
            double coefficient = direction * calibrationScale * contentSize / worldBoundsSize;

            double constant = contentOrigin;

            if (isMirrored)
            {
                constant += calibrationScale * contentSize;
            }

            constant -= coefficient * worldBoundsOrigin;

            // MapKiCadWorldToLocal scales the offset into local space by the image's pixel size,
            // and falls back to using it unscaled when there is no bitmap to measure against.
            constant += imagePixelSize > 0.0
                ? calibrationOffset * contentSize / imagePixelSize
                : calibrationOffset;

            return (coefficient, constant);
        }

        // ###################################################################################
        // Builds the matrix that carries local coordinates built at the REFERENCE calibration to
        // where they belong under the CURRENT one.
        //
        // worldBounds, contentRect and the image pixel size are the same values the caller passed
        // to MapKiCadWorldToLocal when it built the geometry - they are shared by both
        // calibrations, and only the six calibration numbers differ between them.
        //
        // Returns false when the reference calibration is degenerate, which is the one case where
        // no delta exists; the caller must then rebuild rather than transform.
        // ###################################################################################
        public static bool TryBuildDeltaTransform(
            KiCadCalibrationValues reference,
            KiCadCalibrationValues current,
            Rect worldBounds,
            Rect contentRect,
            double imagePixelWidth,
            double imagePixelHeight,
            out Matrix transform)
        {
            transform = Matrix.Identity;

            if (worldBounds.Width <= 0.0 ||
                worldBounds.Height <= 0.0 ||
                contentRect.Width <= 0.0 ||
                contentRect.Height <= 0.0)
            {
                return false;
            }

            var referenceX = KiCadCalibrationTransform.BuildAxisMap(
                reference.ScaleX,
                reference.OffsetX,
                reference.MirrorX,
                worldBounds.X,
                worldBounds.Width,
                contentRect.X,
                contentRect.Width,
                imagePixelWidth);

            var referenceY = KiCadCalibrationTransform.BuildAxisMap(
                reference.ScaleY,
                reference.OffsetY,
                reference.MirrorY,
                worldBounds.Y,
                worldBounds.Height,
                contentRect.Y,
                contentRect.Height,
                imagePixelHeight);

            if (Math.Abs(referenceX.Coefficient) < MinimumUsableScale ||
                Math.Abs(referenceY.Coefficient) < MinimumUsableScale)
            {
                return false;
            }

            var currentX = KiCadCalibrationTransform.BuildAxisMap(
                current.ScaleX,
                current.OffsetX,
                current.MirrorX,
                worldBounds.X,
                worldBounds.Width,
                contentRect.X,
                contentRect.Width,
                imagePixelWidth);

            var currentY = KiCadCalibrationTransform.BuildAxisMap(
                current.ScaleY,
                current.OffsetY,
                current.MirrorY,
                worldBounds.Y,
                worldBounds.Height,
                contentRect.Y,
                contentRect.Height,
                imagePixelHeight);

            double scaleX = currentX.Coefficient / referenceX.Coefficient;
            double scaleY = currentY.Coefficient / referenceY.Coefficient;

            double translateX = currentX.Constant - (referenceX.Constant * scaleX);
            double translateY = currentY.Constant - (referenceY.Constant * scaleY);

            // Scale first, then translate - the order the derivation above produces. Avalonia's
            // Matrix multiplication applies the left operand first.
            transform = Matrix.CreateScale(scaleX, scaleY) * Matrix.CreateTranslation(translateX, translateY);
            return true;
        }
    }

    // #######################################################################################
    // The six calibration numbers, as a plain value the Handlers side can take.
    //
    // TabSchematics has its own KiCadViewCalibration, but it is a PRIVATE nested class, so it
    // cannot appear in a signature here. This is deliberately a separate, minimal carrier rather
    // than a reason to widen that type's visibility.
    // #######################################################################################
    public readonly struct KiCadCalibrationValues
    {
        public KiCadCalibrationValues(
            double scaleX,
            double scaleY,
            double offsetX,
            double offsetY,
            bool mirrorX,
            bool mirrorY)
        {
            this.ScaleX = scaleX;
            this.ScaleY = scaleY;
            this.OffsetX = offsetX;
            this.OffsetY = offsetY;
            this.MirrorX = mirrorX;
            this.MirrorY = mirrorY;
        }

        public double ScaleX { get; }

        public double ScaleY { get; }

        public double OffsetX { get; }

        public double OffsetY { get; }

        public bool MirrorX { get; }

        public bool MirrorY { get; }
    }
}
