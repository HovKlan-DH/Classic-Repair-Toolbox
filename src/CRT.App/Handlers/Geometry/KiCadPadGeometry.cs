using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Pad shape and orientation for the KiCad overlay, extracted from TabSchematics so it can be
    // unit tested.
    //
    // "Which way round is this pad?" is the question these answer. Getting it wrong draws a
    // vertical rectangular pad as a horizontal one, which looks entirely plausible next to a
    // correct pad and is therefore very hard to spot by eye.
    // ###########################################################################################
    public static class KiCadPadGeometry
    {
        // ###########################################################################################
        // Returns true when a KiCad pad shape should be drawn as a rectangle rather than an ellipse.
        // roundrect and trapezoid are approximated by a plain rectangle; circle, oval and custom
        // pads fall through to the ellipse path.
        // ###########################################################################################
        public static bool IsRectangularShape(string? shape)
        {
            string trimmed = shape?.Trim() ?? string.Empty;

            return string.Equals(trimmed, "rect", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(trimmed, "roundrect", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(trimmed, "trapezoid", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Converts a KiCad pad rotation into the rotation the overlay must apply on screen.
        //
        // Two conventions collide here. KiCad measures a positive pad angle counter-clockwise as
        // drawn, while Avalonia's Matrix.CreateRotation is clockwise in its Y-down device space, so
        // the angle is negated. On top of that the view calibration can mirror the board (the user
        // draws the calibration box right-to-left or bottom-to-top when matching a mirrored photo),
        // and a single mirror reverses the sense of rotation again. Two mirrors compose into a
        // 180-degree turn, which restores it - hence the exclusive-or.
        // ###########################################################################################
        public static double ResolveScreenRotationDegrees(
            double padRotationDegrees,
            bool mirrorX,
            bool mirrorY)
        {
            if (double.IsNaN(padRotationDegrees) || double.IsInfinity(padRotationDegrees))
            {
                return 0.0;
            }

            double screenDegrees = mirrorX ^ mirrorY
                ? padRotationDegrees
                : -padRotationDegrees;

            return NormalizeDegrees(screenDegrees);
        }

        // ###########################################################################################
        // Wraps an angle into [0, 360). A pad rectangle is symmetric about its centre, so 90 and 270
        // draw identically - the normalisation exists so callers can compare angles meaningfully.
        // ###########################################################################################
        public static double NormalizeDegrees(double degrees)
        {
            if (double.IsNaN(degrees) || double.IsInfinity(degrees))
            {
                return 0.0;
            }

            double wrapped = degrees % 360.0;

            return wrapped < 0.0 ? wrapped + 360.0 : wrapped;
        }

        // ###########################################################################################
        // Returns true when the rotation draws identically to no rotation at all, so the renderer can
        // skip the transform push for the overwhelming majority of pads.
        //
        // The test is modulo 180, not 360: both shapes the overlay draws for a pad - a rectangle and
        // an ellipse - are symmetric about their centre, so a half turn changes nothing. That matters
        // in practice because 180-degree footprints are common on a real board.
        // ###########################################################################################
        public static bool IsAxisAligned(double screenRotationDegrees)
        {
            double normalized = KiCadPadGeometry.NormalizeDegrees(screenRotationDegrees) % 180.0;

            return normalized < 0.0001 || normalized > 179.9999;
        }
    }
}
