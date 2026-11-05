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
using Handlers.Geometry;
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
    // Maps one KiCad world-space point into the local image coordinate board currently used by
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
    // the active box-based calibration model. The rule is KiCadFootprintMarks' - a card's preview on
    // the Contribute tab draws its traces as wide by the same one.
    // ###########################################################################################
    private static double MapKiCadWorldLengthToLocal(
        double worldLength,
        Rect worldBounds,
        Rect contentRect,
        KiCadViewCalibration calibration) =>
        KiCadFootprintMarks.WorldLengthOnPicture(worldLength, worldBounds, calibration.ScaleX, calibration.ScaleY, contentRect.Width, contentRect.Height);

    // ###########################################################################################
    // Computes a world bounding box for all PCB geometry used by the MVP overlay.
    // Copper zones are included so rendering and hit testing use the full occupied board area.
    // The rule is KiCadFootprintMarks' - "Mark all automatically" maps footprints with the same
    // bounds, so its rectangles land where this tab draws the copper.
    // ###########################################################################################
    private Rect GetKiCadPcbWorldBounds(KiCadPcb pcb) => KiCadFootprintMarks.PcbWorldBounds(pcb);

    // ###########################################################################################
    // The world bounding box of a schematic sheet - KiCadFootprintMarks' rule, so a component's
    // rectangle lands where this tab draws the wires, as on a side of the board.
    // ###########################################################################################
    private Rect GetKiCadSchematicWorldBounds(KiCadSchematic schematic) => KiCadFootprintMarks.SheetWorldBounds(schematic);

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