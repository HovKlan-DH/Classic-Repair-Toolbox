using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

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

    // ###########################################################################################
    // Which blink phase a primitive belongs to (code review, 2026-10-13). A blinking net's copper
    // is drawn only while the blink shows (OnPhase); in the hidden phase its zones still show a
    // faint fill (OffPhase) - exactly what its primitives looked like when the blink factor was
    // baked into them, 0 for the hidden phase. Everything else is Always drawn.
    // ###########################################################################################
    public enum KiCadBlinkPhase
    {
        Always,
        OnPhase,
        OffPhase
    }

    public sealed class KiCadOverlayPrimitive
    {
        public KiCadOverlayPrimitiveKind Kind { get; init; }
        public KiCadBlinkPhase BlinkPhase { get; init; } = KiCadBlinkPhase.Always;
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

        // ###########################################################################################
        // Drawn where it was built, whatever the control's OverlayTransform says, and after everything
        // that is transformed (2026-11-06). The calibration box and its handles: built at the box's
        // true place on every move, while the traces under it are the same primitives moved by the
        // transform. Before this flag the box was drawn through the transform with its inverse
        // pre-applied - a clone of its geometry per move, and the box counted as part of what the
        // transform moves, so nothing could tell "the traces are unchanged" while the box moved.
        // ###########################################################################################
        public bool DrawnUntransformed { get; init; }

        // The stroke geometry, pen and bounds worked out from this primitive, cached on the primitive
        // itself. Everything above is init-only, so anything derived from it can never go stale.
        //
        // This exists because the per-net primitive cache hands the same instances back on every
        // rebuild, but SetGeometry was still reconstructing every StreamGeometry from scratch -
        // profiled at ~50 ms per rebuild for roughly 5,000 geometries that had not changed.
        internal Geometry? PreparedGeometry { get; set; }
        internal Pen? PreparedPen { get; set; }
        internal Rect? PreparedBounds { get; set; }

        // This primitive in another blink phase - everything else, prepared parts included, kept.
        internal KiCadOverlayPrimitive WithBlinkPhase(KiCadBlinkPhase phase) => new()
        {
            Kind = this.Kind,
            BlinkPhase = phase,
            Start = this.Start,
            End = this.End,
            Rect = this.Rect,
            Points = this.Points,
            Pen = this.Pen,
            Fill = this.Fill,
            Geometry = this.Geometry,
            RotationDegrees = this.RotationDegrees,
            DrawnUntransformed = this.DrawnUntransformed,
            PreparedGeometry = this.PreparedGeometry,
            PreparedPen = this.PreparedPen,
            PreparedBounds = this.PreparedBounds,
        };
    }

    // ###########################################################################################
    // A blinking KiCad net's primitives, split into blink phases (code review, 2026-10-13).
    //
    // They are built at full strength, and this says what each phase draws - exactly what the old
    // primitives drew with the blink factor baked in: at 1 everything, at 0 a zone's fill clamped
    // to its 0.10 floor and nothing else, except a pin-1 pad, whose brush never carried the blink.
    //   - a pin-1 pad (its fill is `pinOneBrush`) stays Always;
    //   - a zone (Geometry with a fill) is OnPhase, plus an OffPhase copy at ZoneFillFloor opacity
    //     with no outline;
    //   - everything else is OnPhase.
    // ###########################################################################################
    internal static class KiCadBlinkPhases
    {
        // The lowest a blinking net's zone fill went at blink factor 0 (Math.Clamp(0, 0.10, 0.38)).
        internal const double ZoneFillFloor = 0.10;

        internal static List<KiCadOverlayPrimitive> Split(IReadOnlyList<KiCadOverlayPrimitive> netPrimitives, IBrush? pinOneBrush)
        {
            var result = new List<KiCadOverlayPrimitive>(netPrimitives.Count + 4);

            foreach (var primitive in netPrimitives)
            {
                if (pinOneBrush != null && ReferenceEquals(primitive.Fill, pinOneBrush))
                {
                    result.Add(primitive);
                    continue;
                }

                result.Add(primitive.WithBlinkPhase(KiCadBlinkPhase.OnPhase));

                if (primitive.Kind == KiCadOverlayPrimitiveKind.Geometry && primitive.Fill is ISolidColorBrush zoneFill)
                {
                    result.Add(new KiCadOverlayPrimitive
                    {
                        Kind = KiCadOverlayPrimitiveKind.Geometry,
                        BlinkPhase = KiCadBlinkPhase.OffPhase,
                        Geometry = primitive.Geometry,
                        // Never stronger than the zone is when the blink shows it (owner report,
                        // 2026-11-03: a picture's traces at 5% blinked brighter in the hidden phase).
                        Fill = new SolidColorBrush(zoneFill.Color, Math.Min(ZoneFillFloor, zoneFill.Opacity)),
                    });
                }
            }

            return result;
        }
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
        // The blink: true draws the OnPhase primitives, false the OffPhase ones; Always ones are
        // drawn in both (KiCadBlinkPhase). Setting it redraws what is already prepared - nothing is
        // built, and the Schematics tab's per-net primitive cache is not touched (code review,
        // 2026-10-13: the blink rebuilt every blinking net's primitives twice a second).
        //
        // One control, not a layer per phase with its own opacity: the primitives are drawn in the
        // order they were built - a pin-1 pad over its own net's copper, a net over the one before
        // it - and splitting them into layers would change that wherever copper overlaps.
        // ###########################################################################################
        public bool BlinkPhaseVisible
        {
            get => this.thisBlinkPhaseVisible;
            set
            {
                if (this.thisBlinkPhaseVisible == value)
                {
                    return;
                }

                this.thisBlinkPhaseVisible = value;

                if (this.thisHasBlinkingPrimitives)
                {
                    this.InvalidateVisual();
                }
            }
        }

        private bool thisBlinkPhaseVisible = true;
        private bool thisHasBlinkingPrimitives;

        // Whether a primitive of `phase` is drawn when the blink is `phaseVisible`.
        internal static bool IsDrawnInPhase(KiCadBlinkPhase phase, bool phaseVisible) => phase switch
        {
            KiCadBlinkPhase.OnPhase => phaseVisible,
            KiCadBlinkPhase.OffPhase => !phaseVisible,
            _ => true,
        };

        // ###########################################################################################
        // Replaces the current render geometry, prepares each primitive's cached path, pen and
        // bounds, and triggers a redraw.
        // ###########################################################################################
        public void SetGeometry(IReadOnlyList<KiCadOverlayPrimitive>? primitives)
        {
            this.thisPrimitives = primitives ?? Array.Empty<KiCadOverlayPrimitive>();
            this.thisPrepared = KiCadOverlayRenderControl.PreparePrimitives(this.thisPrimitives);
            this.thisHasBlinkingPrimitives = this.thisPrimitives.Any(primitive => primitive.BlinkPhase != KiCadBlinkPhase.Always);

            // A picture of the transformed primitives stays only while they are the very same ones -
            // the calibration box alone replaced on every move of a drag leaves it as it is.
            if (this.thisSnapshot is { } snapshot && !snapshot.ShowsTheSamePrimitivesAs(this.thisPrepared))
            {
                this.ForgetSnapshot();
            }

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

            // The round-capped pen for each pen the polylines carry, made once per pen rather than once
            // per polyline (2026-11-06): a net's runs of one width share one pen, so this was some 2,000
            // identical pens per rebuild on 250469, and as many for the compositor to take in.
            var smoothedPenBySource = new Dictionary<Pen, Pen>(ReferenceEqualityComparer.Instance);

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

                        if (primitive.Pen is Pen sourcePen)
                        {
                            if (!smoothedPenBySource.TryGetValue(sourcePen, out Pen? smoothed))
                            {
                                smoothed = KiCadOverlayRenderControl.BuildSmoothedPolylinePen(sourcePen);
                                smoothedPenBySource[sourcePen] = smoothed;
                            }

                            pen = smoothed;
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
            this.thisHasBlinkingPrimitives = false;
            this.ForgetSnapshot();
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

            // While the calibration box is moved, the picture of the transformed primitives in place
            // of the primitives themselves - one bitmap a frame, not thousands of paths. A picture that
            // no longer fits (another zoom, size, blink phase) is not drawn: the primitives are, as
            // ever, until the next move takes a new one.
            if (!this.TryDrawSnapshot(context))
            {
                this.DrawPrimitives(context, transformed: true, visibleRect);
            }

            this.DrawPrimitives(context, transformed: false, visibleRect);
        }

        // ###########################################################################################
        // Draws the primitives that are on screen: those the OverlayTransform moves (`transformed`), or
        // those drawn where they were built (DrawnUntransformed) - the second pass after the first, so
        // the calibration box stays on top of the traces as it always was, being the last primitive.
        // ###########################################################################################
        private void DrawPrimitives(DrawingContext context, bool transformed, Rect visibleRect)
        {
            bool hasOverlayTransform = transformed && this.thisOverlayTransform != Matrix.Identity;

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

            for (int i = 0; i < this.thisPrepared.Count; i++)
            {
                var prepared = this.thisPrepared[i];

                if (prepared.Primitive.DrawnUntransformed == transformed ||
                    !KiCadOverlayRenderControl.IsDrawnInPhase(prepared.Primitive.BlinkPhase, this.thisBlinkPhaseVisible) ||
                    !OverlayCullGeometry.IsVisible(prepared.Bounds, visibleRect))
                {
                    continue;
                }

                KiCadOverlayRenderControl.DrawPrepared(context, prepared);
            }
        }

        // ###########################################################################################
        // One primitive, as worked out by PreparePrimitives - on the screen, or into a snapshot.
        // ###########################################################################################
        private static void DrawPrepared(DrawingContext context, PreparedPrimitive prepared)
        {
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

        // ---- the snapshot: a picture of the overlay while the calibration box moves (2026-11-06) ----

        // ###########################################################################################
        // *** WHY THE OVERLAY IS DRAWN FROM A PICTURE OF ITSELF WHILE THE BOX MOVES (owner report,
        // 2026-11-06: "incredible slow ... when dragging one of the markers"). ***
        //
        // A drag of the calibration box changes only OverlayTransform - the primitives were already
        // built once (KiCadCalibrationTransform). But every frame still drew all of them: measured on
        // 250469's top side, 5,351 primitives, 35-65 ms of Skia raster per frame on a fast PC, so a
        // drag ran at 10-20 frames a second there and far fewer on a slower one. The picture is taken
        // ONCE, at the zoom on screen and the screen's own resolution, and every frame after draws that
        // one bitmap under the matrix that carries it to where the primitives would be
        // (OverlaySnapshotGeometry) - the same traces in the same places, resampled while they move.
        // The tab stops it when the box is let go of, and the primitives are drawn sharp again.
        //
        // The picture shows what the transformed primitives draw in the blink phase it was taken in;
        // the calibration box (DrawnUntransformed) is never in it and is drawn as ever on top.
        // ###########################################################################################
        private sealed class Snapshot : IDisposable
        {
            // Null: nothing of the overlay near the view, so there is nothing to draw.
            public RenderTargetBitmap? Bitmap { get; init; }

            // Where the picture lies on screen (the control's parent), in DIPs.
            public Rect Area { get; init; }

            // What the transformed primitives cover, in built space - what the view may need of the picture as it moves.
            public Rect BuiltContent { get; init; }

            public Matrix OverlayTransform { get; init; }

            public Matrix VisualTransform { get; init; }

            public Size Size { get; init; }

            public double Scaling { get; init; }

            public bool BlinkPhaseVisible { get; init; }

            // The transformed primitives it shows, in order - it shows them only while they are the
            // very same instances.
            public KiCadOverlayPrimitive[] Primitives { get; init; } = [];

            public bool ShowsTheSamePrimitivesAs(IReadOnlyList<PreparedPrimitive> prepared)
            {
                int index = 0;

                for (int i = 0; i < prepared.Count; i++)
                {
                    var primitive = prepared[i].Primitive;

                    if (primitive.DrawnUntransformed)
                    {
                        continue;
                    }

                    if (index >= this.Primitives.Length || !ReferenceEquals(this.Primitives[index], primitive))
                    {
                        return false;
                    }

                    index++;
                }

                return index == this.Primitives.Length;
            }

            public void Dispose() => this.Bitmap?.Dispose();
        }

        private Snapshot? thisSnapshot;

        private bool thisDrawsFromSnapshot;

        // How many transformed primitives make a picture worth taking (OverlaySnapshotGeometry): fewer
        // are drawn as vectors whatever DrawFromSnapshot asks. A test of a small board sets it lower.
        internal int PrimitivesWorthAPicture { get; set; } = OverlaySnapshotGeometry.PrimitivesWorthAPicture;

        // Whether the frames are drawn from a picture right now - the tab's tests ask.
        internal bool IsDrawingFromSnapshot =>
            this.thisDrawsFromSnapshot && this.thisSnapshot is { } snapshot && this.IsUsable(snapshot);

        // ###########################################################################################
        // From now on the transformed primitives are drawn from a picture - taken now unless the one
        // held still fits what is on screen. Called on every move of the box; a picture that fits
        // makes it cost nothing. When no picture can be taken the primitives are drawn as ever.
        // ###########################################################################################
        internal void DrawFromSnapshot()
        {
            this.thisDrawsFromSnapshot = true;

            if (this.thisSnapshot is not { } snapshot || !this.IsUsable(snapshot))
            {
                this.ForgetSnapshot();
                this.thisSnapshot = this.TakeSnapshot();
            }

            this.InvalidateVisual();
        }

        // ###########################################################################################
        // The primitives themselves again, sharp at the transform they are under. The picture is kept
        // for the next move as long as it fits - a box moved, let go of and moved again takes none.
        // ###########################################################################################
        internal void StopDrawingFromSnapshot()
        {
            if (!this.thisDrawsFromSnapshot)
            {
                return;
            }

            this.thisDrawsFromSnapshot = false;
            this.InvalidateVisual();
        }

        // Lets go of the picture held, if any. Safe while a frame drawn from it is still on its way:
        // what the compositor recorded holds its own reference to the bitmap.
        internal void ForgetSnapshot()
        {
            this.thisSnapshot?.Dispose();
            this.thisSnapshot = null;
        }

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnDetachedFromVisualTree(e);
            this.ForgetSnapshot();
        }

        // ###########################################################################################
        // Whether a picture still shows what is on screen: the same size, resolution and zoom, the
        // blink phase it was taken in, and - the box moved since - still all of the overlay the view
        // shows (OverlaySnapshotGeometry.StillCovers; code review, 2026-11-07: kept across moves, a
        // band of traces went missing once the box had moved past the picture's margin). Its
        // primitives are checked as they are set (SetGeometry).
        // ###########################################################################################
        private bool IsUsable(Snapshot snapshot) =>
            snapshot.Size == this.Bounds.Size &&
            Math.Abs(snapshot.Scaling - this.CurrentScaling()) < 1e-9 &&
            snapshot.VisualTransform == this.VisualTransform() &&
            (!this.thisHasBlinkingPrimitives || snapshot.BlinkPhaseVisible == this.thisBlinkPhaseVisible) &&
            OverlaySnapshotGeometry.StillCovers(
                snapshot.Size,
                snapshot.BuiltContent,
                snapshot.Area,
                snapshot.OverlayTransform,
                this.thisOverlayTransform,
                snapshot.VisualTransform);

        private bool TryDrawSnapshot(DrawingContext context)
        {
            if (!this.thisDrawsFromSnapshot || this.thisSnapshot is not { } snapshot || !this.IsUsable(snapshot))
            {
                return false;
            }

            if (snapshot.Bitmap is null)
            {
                return true;
            }

            if (!OverlaySnapshotGeometry.TryPictureToLocal(
                    snapshot.Area,
                    snapshot.OverlayTransform,
                    snapshot.VisualTransform,
                    this.thisOverlayTransform,
                    out Matrix pictureToLocal))
            {
                return false;
            }

            // The source in the picture's PIXELS (OverlaySnapshotGeometry.SourceRectOf says why: in DIPs
            // a display at 125% drew 80% of it, a 25% zoom as the picture took over).
            using (context.PushTransform(pictureToLocal))
            {
                context.DrawImage(snapshot.Bitmap, OverlaySnapshotGeometry.SourceRectOf(snapshot.Bitmap.PixelSize), new Rect(snapshot.Area.Size));
            }

            return true;
        }

        // ###########################################################################################
        // Draws the transformed primitives, as Render would at this moment, into a bitmap of the part
        // of the screen worth a picture (OverlaySnapshotGeometry.AreaToCapture). Null when no picture
        // can be taken.
        // ###########################################################################################
        private Snapshot? TakeSnapshot()
        {
            Size size = this.Bounds.Size;
            double scaling = this.CurrentScaling();
            Matrix visual = this.VisualTransform();
            Matrix overlay = this.thisOverlayTransform;

            var transformed = this.thisPrepared.Where(prepared => !prepared.Primitive.DrawnUntransformed).ToArray();

            if (!OverlaySnapshotGeometry.IsWorthAPicture(transformed.Length, this.PrimitivesWorthAPicture))
            {
                return null;
            }

            Rect builtContent = transformed.Length == 0
                ? default
                : transformed.Skip(1).Aggregate(transformed[0].Bounds, (union, prepared) => union.Union(prepared.Bounds));

            Rect? area = OverlaySnapshotGeometry.AreaToCapture(size, builtContent.TransformToAABB(overlay * visual), scaling);

            if (area is not { } captured)
            {
                return null;
            }

            RenderTargetBitmap? bitmap = null;

            if (OverlaySnapshotGeometry.PixelSizeOf(captured, scaling) is { } pixels)
            {
                bitmap = new RenderTargetBitmap(pixels, new Vector(96.0 * scaling, 96.0 * scaling));
                Matrix builtToPicture = OverlaySnapshotGeometry.BuiltToPicture(captured, overlay, visual);

                Rect visibleBuilt = builtToPicture.TryInvert(out Matrix pictureToBuilt)
                    ? new Rect(captured.Size).TransformToAABB(pictureToBuilt)
                    : builtContent;

                using DrawingContext context = bitmap.CreateDrawingContext();
                using (context.PushTransform(builtToPicture))
                {
                    foreach (var prepared in transformed.Where(prepared =>
                                 KiCadOverlayRenderControl.IsDrawnInPhase(prepared.Primitive.BlinkPhase, this.thisBlinkPhaseVisible) &&
                                 OverlayCullGeometry.IsVisible(prepared.Bounds, visibleBuilt)))
                    {
                        KiCadOverlayRenderControl.DrawPrepared(context, prepared);
                    }
                }
            }

            return new Snapshot
            {
                Bitmap = bitmap,
                Area = captured,
                BuiltContent = builtContent,
                OverlayTransform = overlay,
                VisualTransform = visual,
                Size = size,
                Scaling = scaling,
                BlinkPhaseVisible = this.thisBlinkPhaseVisible,
                Primitives = transformed.Select(prepared => prepared.Primitive).ToArray(),
            };
        }

        private double CurrentScaling() => TopLevel.GetTopLevel(this)?.RenderScaling ?? 1.0;

        // The control's own visual transform, as the compositor applies it: its RenderTransform about
        // its RenderTransformOrigin (the zoom and pan - ViewMatrix is set to the same).
        private Matrix VisualTransform()
        {
            if (this.RenderTransform is not { } transform)
            {
                return Matrix.Identity;
            }

            Matrix origin = Matrix.CreateTranslation(this.RenderTransformOrigin.ToPixels(this.Bounds.Size));
            return origin.Invert() * transform.Value * origin;
        }
    }
}
