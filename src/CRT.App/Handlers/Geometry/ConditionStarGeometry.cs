using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE SHAPE OF ONE STAR OF A DEVICE'S CONDITION (owner request, 2026-11-06: 'have this show prettier - not sure how,
    // but maybe more "depth" and coloring') - drawn as a shape of its own rather than a font's glyph, so it can take
    // a gradient, an edge and a shadow (ConditionStars). Five points, the first straight up, the inner corners at
    // InnerRatio of the outer radius - a little fuller than a geometric star, as a rating star reads - all inside a
    // `size` square, centred in it, the shape's lower points on its bottom edge's level less the stroke's room.
    // ###########################################################################################
    public static class ConditionStarGeometry
    {
        // The inner corners' distance from the middle, of the outer points' - 0.5 a full, friendly star.
        public const double InnerRatio = 0.5;

        // ###########################################################################################
        // The ten corners, clockwise from the top point - outer, inner, outer, ... `size` is the square the star
        // stands in; `inset` the room kept free inside it on every side (the edge's half-thickness).
        // ###########################################################################################
        public static IReadOnlyList<Point> Corners(double size, double inset = 0)
        {
            var corners = new List<Point>(10);

            if (!(size > 0) || inset < 0 || inset * 2 >= size)
            {
                return corners;
            }

            double outer = (size / 2.0) - inset;
            double inner = outer * InnerRatio;

            // A star's points reach 1 + cos(36) of its radius from top to bottom - less than its width - so it is moved
            // down by half the difference to stand in the middle of the square.
            double height = outer * (1 + Math.Cos(Math.PI / 5));
            var middle = new Point(size / 2.0, inset + outer + (((size - (2 * inset)) - height) / 2.0));

            for (int corner = 0; corner < 10; corner++)
            {
                double radius = corner % 2 == 0 ? outer : inner;
                double angle = (-Math.PI / 2) + (corner * Math.PI / 5);
                corners.Add(new Point(middle.X + (radius * Math.Cos(angle)), middle.Y + (radius * Math.Sin(angle))));
            }

            return corners;
        }
    }
}
