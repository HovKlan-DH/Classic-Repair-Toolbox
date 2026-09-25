using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Pure 2D polygon and segment maths, extracted from TabSchematics so it can be unit tested.
    //
    // These functions decide which copper the KiCad overlay lights up: whether a point falls in a
    // poured zone, how close a track runs to a zone boundary, whether an arc grazes one. They use
    // Avalonia's Point/Rect value types but touch no control, no rendering and no instance state.
    //
    // Everything here is world-space: coordinates are KiCad millimetres, not screen pixels.
    // ###########################################################################################
    public static class PolygonGeometry
    {
        // ###########################################################################################
        // Returns the shortest distance from a point to the segment v-w.
        // ###########################################################################################
        public static double DistanceToSegment(Point p, double vX, double vY, double wX, double wY)
        {
            double l2 = Math.Pow(wX - vX, 2) + Math.Pow(wY - vY, 2);
            if (l2 == 0.0) return Math.Sqrt(Math.Pow(p.X - vX, 2) + Math.Pow(p.Y - vY, 2));

            double t = Math.Max(0, Math.Min(1, ((p.X - vX) * (wX - vX) + (p.Y - vY) * (wY - vY)) / l2));
            double projX = vX + t * (wX - vX);
            double projY = vY + t * (wY - vY);

            return Math.Sqrt(Math.Pow(p.X - projX, 2) + Math.Pow(p.Y - projY, 2));
        }

        // ###########################################################################################
        // Returns true when the point lies inside the polygon using ray-casting.
        // ###########################################################################################
        public static bool IsPointInPolygon(IReadOnlyList<Point> polygon, Point point)
        {
            if (polygon.Count < 3)
            {
                return false;
            }

            bool inside = false;
            int previousIndex = polygon.Count - 1;

            for (int currentIndex = 0; currentIndex < polygon.Count; currentIndex++)
            {
                Point current = polygon[currentIndex];
                Point previous = polygon[previousIndex];

                bool intersects = ((current.Y > point.Y) != (previous.Y > point.Y)) &&
                                  (point.X < ((previous.X - current.X) * (point.Y - current.Y) / ((previous.Y - current.Y) + 1e-12)) + current.X);

                if (intersects)
                {
                    inside = !inside;
                }

                previousIndex = currentIndex;
            }

            return inside;
        }

        // ###########################################################################################
        // Returns the shortest distance from the point to the polygon boundary.
        // ###########################################################################################
        public static double GetDistanceToPolygonBoundary(Point point, IReadOnlyList<Point> polygon)
        {
            if (polygon.Count < 2)
            {
                return double.MaxValue;
            }

            double minimumDistance = double.MaxValue;

            for (int i = 0; i < polygon.Count; i++)
            {
                Point start = polygon[i];
                Point end = polygon[(i + 1) % polygon.Count];

                double distance = PolygonGeometry.DistanceToSegment(
                    point,
                    start.X,
                    start.Y,
                    end.X,
                    end.Y);

                if (distance < minimumDistance)
                {
                    minimumDistance = distance;
                }
            }

            return minimumDistance;
        }

        // ###########################################################################################
        // Returns true when the point is inside the zone or near its boundary within the supplied
        // tolerance. The closest distance is returned so overlapping candidates can be ranked.
        // ###########################################################################################
        public static bool IsPointInOrNearZone(
            Point point,
            IReadOnlyList<IReadOnlyList<Point>> polygonsWorld,
            double toleranceWorld,
            out double distanceWorld)
        {
            distanceWorld = double.MaxValue;

            foreach (var polygon in polygonsWorld)
            {
                if (PolygonGeometry.IsPointInPolygon(polygon, point))
                {
                    distanceWorld = 0.0;
                    return true;
                }

                double boundaryDistance = PolygonGeometry.GetDistanceToPolygonBoundary(point, polygon);
                if (boundaryDistance < distanceWorld)
                {
                    distanceWorld = boundaryDistance;
                }
            }

            return distanceWorld <= toleranceWorld;
        }

        // ###########################################################################################
        // Returns true when a circular copper feature touches the zone.
        // ###########################################################################################
        public static bool DoesCircleTouchZone(
            Point centerWorld,
            double radiusWorld,
            IReadOnlyList<IReadOnlyList<Point>> polygonsWorld)
        {
            return PolygonGeometry.IsPointInOrNearZone(centerWorld, polygonsWorld, radiusWorld, out _);
        }

        // ###########################################################################################
        // Returns true when a segment touches the zone.
        // Uses fast endpoint checks and an adaptive sample count instead of a fixed heavy sample loop.
        // ###########################################################################################
        public static bool DoesSegmentTouchZone(
            Point startWorld,
            Point endWorld,
            double radiusWorld,
            IReadOnlyList<IReadOnlyList<Point>> polygonsWorld)
        {
            if (polygonsWorld.Count == 0)
            {
                return false;
            }

            if (PolygonGeometry.IsPointInOrNearZone(startWorld, polygonsWorld, radiusWorld, out _) ||
                PolygonGeometry.IsPointInOrNearZone(endWorld, polygonsWorld, radiusWorld, out _))
            {
                return true;
            }

            double dx = endWorld.X - startWorld.X;
            double dy = endWorld.Y - startWorld.Y;
            double segmentLength = Math.Sqrt((dx * dx) + (dy * dy));

            int sampleCount = Math.Clamp(
                (int)Math.Ceiling(segmentLength / Math.Max(0.75, radiusWorld * 3.0)),
                4,
                12);

            for (int i = 1; i < sampleCount; i++)
            {
                double t = (double)i / sampleCount;

                Point samplePoint = new(
                    startWorld.X + (dx * t),
                    startWorld.Y + (dy * t));

                if (PolygonGeometry.IsPointInOrNearZone(samplePoint, polygonsWorld, radiusWorld, out _))
                {
                    return true;
                }
            }

            return false;
        }

        // ###########################################################################################
        // Returns true when a quadratic arc (start / mid / end control points) touches the zone.
        // Uses fast control-point checks and a lightweight adaptive world-space sampler.
        // ###########################################################################################
        public static bool DoesArcTouchZone(
            Point startWorld,
            Point midWorld,
            Point endWorld,
            double radiusWorld,
            IReadOnlyList<IReadOnlyList<Point>> polygonsWorld)
        {
            if (polygonsWorld.Count == 0)
            {
                return false;
            }

            if (PolygonGeometry.IsPointInOrNearZone(startWorld, polygonsWorld, radiusWorld, out _) ||
                PolygonGeometry.IsPointInOrNearZone(midWorld, polygonsWorld, radiusWorld, out _) ||
                PolygonGeometry.IsPointInOrNearZone(endWorld, polygonsWorld, radiusWorld, out _))
            {
                return true;
            }

            double firstLegLength = Math.Sqrt(
                Math.Pow(midWorld.X - startWorld.X, 2.0) +
                Math.Pow(midWorld.Y - startWorld.Y, 2.0));

            double secondLegLength = Math.Sqrt(
                Math.Pow(endWorld.X - midWorld.X, 2.0) +
                Math.Pow(endWorld.Y - midWorld.Y, 2.0));

            double approximateArcLength = firstLegLength + secondLegLength;

            int sampleCount = Math.Clamp(
                (int)Math.Ceiling(approximateArcLength / Math.Max(1.0, radiusWorld * 3.5)),
                6,
                16);

            for (int i = 1; i < sampleCount; i++)
            {
                double t = (double)i / sampleCount;
                double mt = 1.0 - t;

                Point samplePoint = new(
                    (mt * mt * startWorld.X) + (2.0 * mt * t * midWorld.X) + (t * t * endWorld.X),
                    (mt * mt * startWorld.Y) + (2.0 * mt * t * midWorld.Y) + (t * t * endWorld.Y));

                if (PolygonGeometry.IsPointInOrNearZone(samplePoint, polygonsWorld, radiusWorld, out _))
                {
                    return true;
                }
            }

            return false;
        }

        // ###########################################################################################
        // Computes a world-space bounding box for a polygon set. Width and height are floored at a
        // tiny positive value so a degenerate polygon still produces a usable rectangle.
        // ###########################################################################################
        public static Rect GetPolygonSetBounds(IReadOnlyList<IReadOnlyList<Point>> polygonsWorld)
        {
            bool hasValue = false;
            double minX = 0;
            double minY = 0;
            double maxX = 0;
            double maxY = 0;

            foreach (var polygon in polygonsWorld)
            {
                foreach (var point in polygon)
                {
                    if (!hasValue)
                    {
                        minX = maxX = point.X;
                        minY = maxY = point.Y;
                        hasValue = true;
                        continue;
                    }

                    if (point.X < minX) minX = point.X;
                    if (point.X > maxX) maxX = point.X;
                    if (point.Y < minY) minY = point.Y;
                    if (point.Y > maxY) maxY = point.Y;
                }
            }

            return hasValue
                ? new Rect(minX, minY, Math.Max(0.0001, maxX - minX), Math.Max(0.0001, maxY - minY))
                : default;
        }
    }
}
