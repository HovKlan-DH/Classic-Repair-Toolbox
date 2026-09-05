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

namespace CRT;

// ###########################################################################################
// Coordinate mapping between KiCad world units and on-screen local/image space, world
// bounds, curve sampling, and the zone polygon geometry helpers.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // ###########################################################################################
    // Maps one KiCad world-space point into the local image coordinate system currently used by
    // the schematics image and overlays using the active box-based calibration model.
    // ###########################################################################################
    private Point MapKiCadWorldToLocal(
        double worldX,
        double worldY,
        Rect worldBounds,
        Rect contentRect,
        KiCadViewCalibration calibration)
    {
        if (worldBounds.Width <= 0 || worldBounds.Height <= 0)
        {
            return new Point(contentRect.X, contentRect.Y);
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

        nx *= calibration.ScaleX;
        ny *= calibration.ScaleY;

        double localX = contentRect.X + (nx * contentRect.Width);
        double localY = contentRect.Y + (ny * contentRect.Height);

        if (this.currentFullResBitmap != null)
        {
            if (this.currentFullResBitmap.PixelSize.Width > 0)
            {
                localX += calibration.OffsetX * (contentRect.Width / this.currentFullResBitmap.PixelSize.Width);
            }

            if (this.currentFullResBitmap.PixelSize.Height > 0)
            {
                localY += calibration.OffsetY * (contentRect.Height / this.currentFullResBitmap.PixelSize.Height);
            }
        }
        else
        {
            localX += calibration.OffsetX;
            localY += calibration.OffsetY;
        }

        return new Point(localX, localY);
    }

    // ###########################################################################################
    // Converts one KiCad world-space length into the current local overlay coordinate space using
    // the active box-based calibration model.
    // ###########################################################################################
    private static double MapKiCadWorldLengthToLocal(
        double worldLength,
        Rect worldBounds,
        Rect contentRect,
        KiCadViewCalibration calibration)
    {
        double thisScaleX = contentRect.Width / Math.Max(0.0001, worldBounds.Width);
        double thisScaleY = contentRect.Height / Math.Max(0.0001, worldBounds.Height);

        thisScaleX *= Math.Abs(calibration.ScaleX);
        thisScaleY *= Math.Abs(calibration.ScaleY);

        return worldLength * ((thisScaleX + thisScaleY) / 2.0);
    }

    // ###########################################################################################
    // Computes a world bounding box for all PCB geometry used by the MVP overlay.
    // Copper zones are included so rendering and hit testing use the full occupied board area.
    // ###########################################################################################
    private Rect GetKiCadPcbWorldBounds(KiCadPcb pcb)
    {
        bool hasValue = false;
        double minX = 0;
        double minY = 0;
        double maxX = 0;
        double maxY = 0;

        void Include(double x, double y)
        {
            if (!hasValue)
            {
                minX = maxX = x;
                minY = maxY = y;
                hasValue = true;
                return;
            }

            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
        }

        foreach (var segment in pcb.Routing.Segments)
        {
            if (segment.Start != null) Include(segment.Start.X, segment.Start.Y);
            if (segment.End != null) Include(segment.End.X, segment.End.Y);
        }

        foreach (var via in pcb.Routing.Vias)
        {
            if (via.At != null) Include(via.At.X, via.At.Y);
        }

        foreach (var arc in pcb.Routing.Arcs)
        {
            if (arc.Start != null) Include(arc.Start.X, arc.Start.Y);
            if (arc.Mid != null) Include(arc.Mid.X, arc.Mid.Y);
            if (arc.End != null) Include(arc.End.X, arc.End.Y);
        }

        foreach (var footprint in pcb.Footprints)
        {
            foreach (var pad in footprint.Pads)
            {
                if (pad.AbsoluteCenter == null)
                {
                    continue;
                }

                double halfWidth = (pad.Size?.X ?? 0.0) / 2.0;
                double halfHeight = (pad.Size?.Y ?? 0.0) / 2.0;

                Include(pad.AbsoluteCenter.X - halfWidth, pad.AbsoluteCenter.Y - halfHeight);
                Include(pad.AbsoluteCenter.X + halfWidth, pad.AbsoluteCenter.Y + halfHeight);
            }
        }

        foreach (var zone in pcb.Routing.Zones)
        {
            foreach (var polygon in zone.FilledPolygons.Count > 0 ? zone.FilledPolygons : zone.OutlinePolygons)
            {
                foreach (var point in polygon.Points)
                {
                    Include(point.X, point.Y);
                }
            }
        }

        return hasValue
            ? new Rect(minX, minY, Math.Max(0.0001, maxX - minX), Math.Max(0.0001, maxY - minY))
            : default;
    }

    // ###########################################################################################
    // Computes a world bounding box for schematic wires, polylines, and net labels.
    // ###########################################################################################
    private Rect GetKiCadSchematicWorldBounds(KiCadSchematic schematic)
    {
        bool hasValue = false;
        double minX = 0;
        double minY = 0;
        double maxX = 0;
        double maxY = 0;

        void Include(double x, double y)
        {
            if (!hasValue)
            {
                minX = maxX = x;
                minY = maxY = y;
                hasValue = true;
                return;
            }

            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
        }

        foreach (var wire in schematic.Wires)
        {
            foreach (var point in wire.Points)
            {
                Include(point.X, point.Y);
            }
        }

        foreach (var polyline in schematic.Polylines)
        {
            foreach (var point in polyline.Points)
            {
                Include(point.X, point.Y);
            }
        }

        foreach (var label in schematic.Labels.Local)
        {
            if (label.At != null) Include(label.At.X, label.At.Y);
        }

        foreach (var label in schematic.Labels.Global)
        {
            if (label.At != null) Include(label.At.X, label.At.Y);
        }

        foreach (var label in schematic.Labels.Hierarchical)
        {
            if (label.At != null) Include(label.At.X, label.At.Y);
        }

        return hasValue
            ? new Rect(minX, minY, Math.Max(0.0001, maxX - minX), Math.Max(0.0001, maxY - minY))
            : default;
    }

    // ###########################################################################################
    // Samples one quadratic Bézier curve for PCB arc rendering.
    // Uses adaptive subdivision based on on-screen curve length so long arcs stay smooth without
    // exploding the point count for short arcs.
    // ###########################################################################################
    private static List<Point> SampleQuadraticBezier(Point start, Point control, Point end, int steps)
    {
        double firstLegLength = Math.Sqrt(
            Math.Pow(control.X - start.X, 2.0) +
            Math.Pow(control.Y - start.Y, 2.0));

        double secondLegLength = Math.Sqrt(
            Math.Pow(end.X - control.X, 2.0) +
            Math.Pow(end.Y - control.Y, 2.0));

        double approximateScreenLength = firstLegLength + secondLegLength;

        int adaptiveSteps = Math.Clamp(
            (int)Math.Ceiling(approximateScreenLength / 6.0),
            12,
            96);

        int effectiveSteps = Math.Max(2, Math.Max(steps, adaptiveSteps));

        var points = new List<Point>(effectiveSteps + 1);

        for (int i = 0; i <= effectiveSteps; i++)
        {
            double t = (double)i / effectiveSteps;
            double mt = 1.0 - t;

            double x = (mt * mt * start.X) + (2.0 * mt * t * control.X) + (t * t * end.X);
            double y = (mt * mt * start.Y) + (2.0 * mt * t * control.Y) + (t * t * end.Y);

            points.Add(new Point(x, y));
        }

        return points;
    }

    // ###########################################################################################
    // Projects one schematic-local point back into KiCad world coordinates using the active
    // box-based calibration model.
    // ###########################################################################################
    private bool TryMapLocalToKiCadWorld(
        Point localPoint,
        Rect worldBounds,
        Rect contentRect,
        KiCadViewCalibration calibration,
        out Point worldPoint)
    {
        worldPoint = default;

        if (worldBounds.Width <= 0 || worldBounds.Height <= 0)
        {
            return false;
        }

        double thisLocalX = localPoint.X;
        double thisLocalY = localPoint.Y;

        if (this.currentFullResBitmap != null)
        {
            if (this.currentFullResBitmap.PixelSize.Width > 0)
            {
                thisLocalX -= calibration.OffsetX * (contentRect.Width / this.currentFullResBitmap.PixelSize.Width);
            }

            if (this.currentFullResBitmap.PixelSize.Height > 0)
            {
                thisLocalY -= calibration.OffsetY * (contentRect.Height / this.currentFullResBitmap.PixelSize.Height);
            }
        }
        else
        {
            thisLocalX -= calibration.OffsetX;
            thisLocalY -= calibration.OffsetY;
        }

        double thisNormalizedX = (thisLocalX - contentRect.X) / contentRect.Width;
        double thisNormalizedY = (thisLocalY - contentRect.Y) / contentRect.Height;

        if (Math.Abs(calibration.ScaleX) > 1e-10)
        {
            thisNormalizedX /= calibration.ScaleX;
        }

        if (Math.Abs(calibration.ScaleY) > 1e-10)
        {
            thisNormalizedY /= calibration.ScaleY;
        }

        if (calibration.MirrorX)
        {
            thisNormalizedX = 1.0 - thisNormalizedX;
        }

        if (calibration.MirrorY)
        {
            thisNormalizedY = 1.0 - thisNormalizedY;
        }

        worldPoint = new Point(
            (thisNormalizedX * worldBounds.Width) + worldBounds.X,
            (thisNormalizedY * worldBounds.Height) + worldBounds.Y);

        return true;
    }

    // ###########################################################################################
    // Builds one filled geometry for a KiCad copper zone from world-space polygons.
    // ###########################################################################################
    private Geometry? BuildKiCadZoneGeometry(
        IReadOnlyList<IReadOnlyList<Point>> polygonsWorld,
        Rect worldBounds,
        Rect contentRect,
        KiCadViewCalibration calibration)
    {
        if (polygonsWorld.Count == 0)
        {
            return null;
        }

        var geometry = new StreamGeometry();
        bool hasFigure = false;

        using (var geometryContext = geometry.Open())
        {
            foreach (var polygon in polygonsWorld)
            {
                if (polygon.Count < 3)
                {
                    continue;
                }

                var localPoints = polygon
                    .Select(point => this.MapKiCadWorldToLocal(
                        point.X,
                        point.Y,
                        worldBounds,
                        contentRect,
                        calibration))
                    .ToList();

                if (localPoints.Count < 3)
                {
                    continue;
                }

                geometryContext.BeginFigure(localPoints[0], isFilled: true);

                for (int i = 1; i < localPoints.Count; i++)
                {
                    geometryContext.LineTo(localPoints[i]);
                }

                geometryContext.EndFigure(isClosed: true);
                hasFigure = true;
            }
        }

        return hasFigure ? geometry : null;
    }
}