using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.Geometry;
using System;
using System.Collections.Generic;

namespace Tabs.TabSchematics
{
    public enum KiCadOverlayPrimitiveKind
    {
        Line,
        Rectangle,
        Ellipse,
        Polyline,
        Geometry
    }

    public sealed class KiCadOverlayPrimitive
    {
        public KiCadOverlayPrimitiveKind Kind { get; init; }
        public Point Start { get; init; }
        public Point End { get; init; }
        public Rect Rect { get; init; }
        public IReadOnlyList<Point> Points { get; init; } = Array.Empty<Point>();
        public Pen? Pen { get; init; }
        public IBrush? Fill { get; init; }
        public Geometry? Geometry { get; init; }

        // Rotation about Rect.Center, in Avalonia's clockwise Y-down degrees. Only Rectangle and
        // Ellipse honour it; it carries a rotated KiCad pad's true orientation.
        public double RotationDegrees { get; init; }

        // The stroke geometry, pen and bounds worked out from this primitive, cached on the primitive
        // itself. Everything above is init-only, so anything derived from it can never go stale.
        //
        // This exists because the per-net primitive cache hands the same instances back on every
        // rebuild, but SetGeometry was still reconstructing every StreamGeometry from scratch -
        // profiled at ~50 ms per rebuild for roughly 5,000 geometries that had not changed.
        internal Geometry? PreparedGeometry { get; set; }
        internal Pen? PreparedPen { get; set; }
        internal Rect? PreparedBounds { get; set; }
    }

    public sealed class KiCadOverlayRenderControl : Control
    {
        // Everything Render needs for one primitive, worked out once when the geometry is set rather
        // than on every frame. Profiling showed Render running on every zoom step, so anything
        // rebuilt inside it was being rebuilt several times a second for no reason.
        private sealed class PreparedPrimitive
        {
            public KiCadOverlayPrimitive Primitive { get; init; } = null!;
            public Geometry? Geometry { get; init; }
            public Pen? Pen { get; init; }
            public Rect Bounds { get; init; }
        }

        private IReadOnlyList<KiCadOverlayPrimitive> thisPrimitives = Array.Empty<KiCadOverlayPrimitive>();
        private IReadOnlyList<PreparedPrimitive> thisPrepared = Array.Empty<PreparedPrimitive>();
        private Size thisLastArrangedSize = new(-1, -1);

        public IReadOnlyList<KiCadOverlayPrimitive> Primitives => this.thisPrimitives;

        // The zoom/pan matrix the overlay is displayed under. Render uses it to work out which part
        // of the overlay is actually on screen, exactly as SchematicHighlightsOverlay does. Setting
        // it does not invalidate anything by itself - the render transform change already does.
        public Matrix ViewMatrix { get; set; } = Matrix.Identity;

        // ###########################################################################################
        // An extra transform applied to every primitive at DRAW time, on top of the coordinates they
        // already carry.
        //
        // This is what lets a KiCad calibration drag move the whole overlay without rebuilding it.
        // The primitives bake the calibration that was active when they were built, and rebuilding
        // them costs ~390-540 ms on a real board (451 nets, every one of them a per-net cache miss,
        // because the calibration values are part of that cache's key). A calibration change is a
        // pure affine transform though, so during a drag the geometry is built ONCE and only this
        // matrix is updated - see Handlers/Geometry/KiCadCalibrationTransform, which derives it and
        // is tested against the real world-to-local mapping.
        //
        // Identity is the normal state, and it means the primitives are drawn exactly as built.
        // Setting this invalidates on its own, unlike ViewMatrix: nothing else changes when only
        // this moves, so there is no other invalidation to ride along with.
        // ###########################################################################################
        public Matrix OverlayTransform
        {
            get => this.thisOverlayTransform;
            set
            {
                if (this.thisOverlayTransform == value)
                {
                    return;
                }

                this.thisOverlayTransform = value;
                this.InvalidateVisual();
            }
        }

        private Matrix thisOverlayTransform = Matrix.Identity;

        // ###########################################################################################
        // Replaces the current render geometry, prepares each primitive's cached path, pen and
        // bounds, and triggers a redraw.
        // ###########################################################################################
        public void SetGeometry(IReadOnlyList<KiCadOverlayPrimitive>? primitives)
        {

            this.thisPrimitives = primitives ?? Array.Empty<KiCadOverlayPrimitive>();
            this.thisPrepared = KiCadOverlayRenderControl.PreparePrimitives(this.thisPrimitives);
            this.InvalidateVisual();

        }

        // ###########################################################################################
        // Works out, once per geometry change, everything Render would otherwise recompute per frame:
        // the drawable path, the pen to stroke it with, and the bounding box used to skip it when it
        // is off screen.
        //
        // The pen matters more than it looks. Polyline primitives used to build a fresh round-capped
        // Pen inside Render, so a board with ~5,000 copper runs allocated ~5,000 pens on every frame
        // - about 175,000 allocations across ten seconds of zooming, all of them identical to the
        // frame before.
        // ###########################################################################################
        private static IReadOnlyList<PreparedPrimitive> PreparePrimitives(
            IReadOnlyList<KiCadOverlayPrimitive> primitives)
        {
            var prepared = new PreparedPrimitive[primitives.Count];

            for (int i = 0; i < primitives.Count; i++)
            {
                var primitive = primitives[i];

                // A primitive never changes after it is created, so anything already worked out from
                // it stays valid however many rebuilds it survives.
                if (primitive.PreparedBounds.HasValue)
                {
                    prepared[i] = new PreparedPrimitive
                    {
                        Primitive = primitive,
                        Geometry = primitive.PreparedGeometry,
                        Pen = primitive.PreparedPen,
                        Bounds = primitive.PreparedBounds.Value
                    };

                    continue;
                }

                double thickness = primitive.Pen?.Thickness ?? 1.0;

                Geometry? geometry = null;
                Pen? pen = primitive.Pen;
                Rect bounds;

                switch (primitive.Kind)
                {
                    case KiCadOverlayPrimitiveKind.Polyline:
                        if (primitive.Points.Count >= 2)
                        {
                            geometry = KiCadOverlayRenderControl.BuildPolylineGeometry(primitive.Points);
                        }

                        if (primitive.Pen != null)
                        {
                            pen = KiCadOverlayRenderControl.BuildSmoothedPolylinePen(primitive.Pen);
                        }

                        bounds = OverlayCullGeometry.InflateForStroke(
                            OverlayCullGeometry.BoundsOfPoints(primitive.Points),
                            thickness);
                        break;

                    case KiCadOverlayPrimitiveKind.Geometry:
                        geometry = primitive.Geometry;
                        bounds = geometry == null
                            ? default
                            : OverlayCullGeometry.InflateForStroke(geometry.Bounds, thickness);
                        break;

                    case KiCadOverlayPrimitiveKind.Line:
                        bounds = OverlayCullGeometry.InflateForStroke(
                            OverlayCullGeometry.BoundsOfPoints(new[] { primitive.Start, primitive.End }),
                            thickness);
                        break;

                    default:
                        bounds = OverlayCullGeometry.InflateForStroke(
                            OverlayCullGeometry.BoundsOfRotatedRect(primitive.Rect, primitive.RotationDegrees),
                            thickness);
                        break;
                }

                primitive.PreparedGeometry = geometry;
                primitive.PreparedPen = pen;
                primitive.PreparedBounds = bounds;

                prepared[i] = new PreparedPrimitive
                {
                    Primitive = primitive,
                    Geometry = geometry,
                    Pen = pen,
                    Bounds = bounds
                };
            }


            return prepared;
        }

        // ###########################################################################################
        // Clears all render geometry, resets the prepared list, and triggers a redraw.
        // ###########################################################################################
        public void ClearGeometry()
        {
            this.thisPrimitives = Array.Empty<KiCadOverlayPrimitive>();
            this.thisPrepared = Array.Empty<PreparedPrimitive>();
            this.InvalidateVisual();
        }

        // ###########################################################################################
        // Redraws after layout only when the arranged size actually changed. Zooming alters the render
        // transform, not the layout, so the prepared primitives stay valid and there is nothing to
        // re-record. A genuine resize is caught by the size check, and a Bounds change on the image or
        // container separately triggers RefreshKiCadOverlay, which rebuilds and calls SetGeometry.
        // ###########################################################################################
        protected override Size ArrangeOverride(Size finalSize)
        {
            var result = base.ArrangeOverride(finalSize);


            if (result != this.thisLastArrangedSize)
            {

                this.thisLastArrangedSize = result;
                this.InvalidateVisual();
            }

            return result;
        }

        // ###########################################################################################
        // Builds one continuous geometry for a polyline so Avalonia can render joins and caps
        // smoothly instead of drawing each segment as an isolated line.
        // ###########################################################################################
        private static Geometry BuildPolylineGeometry(IReadOnlyList<Point> points)
        {
            var geometry = new StreamGeometry();

            using (var geometryContext = geometry.Open())
            {
                geometryContext.BeginFigure(points[0], isFilled: false);

                for (int i = 1; i < points.Count; i++)
                {
                    geometryContext.LineTo(points[i]);
                }

                geometryContext.EndFigure(isClosed: false);
            }

            return geometry;
        }

        // ###########################################################################################
        // Returns a pen configured for smooth KiCad trace rendering while preserving the original
        // brush, thickness, dash style, and miter limit.
        // Uses round caps and round joins to match the original trace appearance without recreating
        // thousands of UI elements.
        // ###########################################################################################
        private static Pen BuildSmoothedPolylinePen(Pen sourcePen)
        {
            return new Pen(
                sourcePen.Brush,
                sourcePen.Thickness,
                sourcePen.DashStyle,
                PenLineCap.Round,
                PenLineJoin.Round,
                sourcePen.MiterLimit);
        }

        // ###########################################################################################
        // Pushes the rotation a primitive asks for, turning about the centre of its own rect so a
        // rotated KiCad pad keeps its position and only changes orientation. Axis-aligned primitives
        // - which is nearly all of them - get an identity transform rather than a special case, so
        // the caller can always wrap the draw in a using block.
        // ###########################################################################################
        private static DrawingContext.PushedState PushPrimitiveRotation(
            DrawingContext context,
            KiCadOverlayPrimitive primitive)
        {
            if (KiCadPadGeometry.IsAxisAligned(primitive.RotationDegrees))
            {
                return context.PushTransform(Matrix.Identity);
            }

            Point centre = primitive.Rect.Center;
            double radians = primitive.RotationDegrees * Math.PI / 180.0;

            return context.PushTransform(
                Matrix.CreateTranslation(-centre.X, -centre.Y) *
                Matrix.CreateRotation(radians) *
                Matrix.CreateTranslation(centre.X, centre.Y));
        }

        // ###########################################################################################
        // Draws the KiCad overlay primitives that are currently on screen, in one control rather than
        // thousands of child controls on a Canvas.
        // ###########################################################################################
        public override void Render(DrawingContext context)
        {
            base.Render(context);

            if (this.thisPrepared.Count == 0)
            {
                return;
            }


            var visibleRect = OverlayCullGeometry.GetVisibleLocalRect(
                new Rect(0, 0, this.Bounds.Width, this.Bounds.Height),
                this.ViewMatrix);

            bool hasOverlayTransform = this.thisOverlayTransform != Matrix.Identity;

            // The prepared bounds are in the primitives' OWN coordinates, so when a transform is
            // being applied the visible rect has to be pulled back into that same space before
            // anything is compared against it. Culling against the untransformed rect instead makes
            // traces vanish exactly when the calibration box is dragged far enough to matter.
            if (hasOverlayTransform && this.thisOverlayTransform.TryInvert(out var inverseOverlay))
            {
                visibleRect = visibleRect.TransformToAABB(inverseOverlay);
            }

            using var overlayTransformState = hasOverlayTransform
                ? context.PushTransform(this.thisOverlayTransform)
                : default;

            int drawn = 0;

            for (int i = 0; i < this.thisPrepared.Count; i++)
            {
                var prepared = this.thisPrepared[i];

                if (!OverlayCullGeometry.IsVisible(prepared.Bounds, visibleRect))
                {
                    continue;
                }

                drawn++;
                var primitive = prepared.Primitive;

                switch (primitive.Kind)
                {
                    case KiCadOverlayPrimitiveKind.Line:
                        if (prepared.Pen != null)
                        {
                            context.DrawLine(prepared.Pen, primitive.Start, primitive.End);
                        }
                        break;

                    case KiCadOverlayPrimitiveKind.Rectangle:
                        using (KiCadOverlayRenderControl.PushPrimitiveRotation(context, primitive))
                        {
                            context.DrawRectangle(primitive.Fill, prepared.Pen, primitive.Rect);
                        }
                        break;

                    case KiCadOverlayPrimitiveKind.Ellipse:
                        using (KiCadOverlayRenderControl.PushPrimitiveRotation(context, primitive))
                        {
                            context.DrawEllipse(
                                primitive.Fill,
                                prepared.Pen,
                                primitive.Rect.Center,
                                primitive.Rect.Width / 2.0,
                                primitive.Rect.Height / 2.0);
                        }
                        break;

                    case KiCadOverlayPrimitiveKind.Polyline:
                        if (prepared.Pen != null && prepared.Geometry != null)
                        {
                            context.DrawGeometry(null, prepared.Pen, prepared.Geometry);
                        }
                        break;

                    case KiCadOverlayPrimitiveKind.Geometry:
                        if (prepared.Geometry != null)
                        {
                            context.DrawGeometry(primitive.Fill, prepared.Pen, prepared.Geometry);
                        }
                        break;
                }
            }

        }
    }
}
