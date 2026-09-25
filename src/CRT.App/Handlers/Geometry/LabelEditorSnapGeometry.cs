using Avalonia;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // #######################################################################################
    // Everything the label editor needs from the tab in order to snap, resolved to plain values.
    //
    // The point of the struct is that the snapping maths never touches a control. The tab reads
    // its own state once - the working highlight rows, the current drag mode, which schematic is
    // showing, which rows are selected, and how much of the bitmap is actually on screen - and
    // hands the answers over. Everything downstream is arithmetic on those values.
    //
    // VisiblePixelRect is the one that used to force this logic to live in the tab. It is the
    // part of the schematic bitmap currently visible in the viewport, in BITMAP PIXELS, and
    // computing it needs the container bounds, the full-res bitmap size and the inverted view
    // matrix. Only components inside it may take part in snapping, so a neighbour scrolled off
    // screen cannot silently drag an edge. Null means "no viewport information", and every rect
    // is then treated as visible - which is what the tab did when the bitmap or the matrix was
    // not usable yet.
    //
    // IsSelected exists because a row being dragged must not snap to itself or to any of its
    // companions in a multi-selection. The tab answers by consulting its own selection state.
    // #######################################################################################
    internal readonly struct LabelEditorSnapContext
    {
        public LabelEditorSnapContext(
            IReadOnlyList<EditableComponentHighlight> workingHighlights,
            LabelEditorDragMode dragMode,
            string schematicName,
            Rect? visiblePixelRect,
            Func<EditableComponentHighlight, bool> isSelected)
        {
            this.WorkingHighlights = workingHighlights;
            this.DragMode = dragMode;
            this.SchematicName = schematicName;
            this.VisiblePixelRect = visiblePixelRect;
            this.IsSelected = isSelected;
        }

        // Every highlight row the editor is working on, across all schematics - the snap methods
        // filter by SchematicName themselves, exactly as the tab version did.
        public IReadOnlyList<EditableComponentHighlight> WorkingHighlights { get; }

        // Which edges the current drag is moving. This is the ONLY place the mode is carried:
        // ApplyResizeSnap used to accept a separate dragModeOverride argument as well, which meant
        // two sources for one value that a caller could set to disagree. A caller that needs a
        // different mode copies the context with WithDragMode instead.
        public LabelEditorDragMode DragMode { get; }

        // The schematic whose rows may participate. Rows belonging to any other schematic are
        // skipped, case-insensitively.
        public string SchematicName { get; }

        // The on-screen part of the bitmap, in bitmap pixels; null when unknown (treat all as
        // visible).
        public Rect? VisiblePixelRect { get; }

        // True for a row that is part of the current selection.
        public Func<EditableComponentHighlight, bool> IsSelected { get; }

        // ###################################################################################
        // The same context with a different drag mode.
        //
        // This is how ApplyNewRectangleSnap drives all four edges in turn: it copies the context
        // once per edge rather than passing an override alongside it. The struct is readonly and
        // holds five references, so the copy is trivial - and it keeps DragMode a single source
        // of truth that cannot be contradicted by a second argument.
        // ###################################################################################
        public LabelEditorSnapContext WithDragMode(LabelEditorDragMode dragMode) =>
            new LabelEditorSnapContext(
                this.WorkingHighlights,
                dragMode,
                this.SchematicName,
                this.VisiblePixelRect,
                this.IsSelected);
    }

    // #######################################################################################
    // Snapping maths for the label editor: aligning a moved, resized or newly drawn highlight to
    // its neighbours, and producing the guide lines drawn while snapping.
    //
    // This was ~970 lines of private members of TabSchematics, where no test could reach it. It
    // is pure arithmetic over EditableComponentHighlight rows in bitmap-pixel space, so it moved
    // here whole; the tab keeps only the thin rim that reads its controls and builds a
    // LabelEditorSnapContext (see TabSchematics.LabelEditor.Snap.cs).
    //
    // The rules the maths encodes, all of which the tests pin down:
    //  - An edge snaps to a neighbouring edge within snapThreshold (2 px, in bitmap pixels).
    //  - A candidate is REJECTED when another highlight sits between the moving edge and the
    //    neighbour it would snap to - you cannot snap through an intervening component.
    //  - Only rows inside VisiblePixelRect may participate.
    //  - snapOnMatch: false means "do not move anything, only draw guides for exact matches",
    //    which is what keyboard nudging uses to show alignment without fighting the keypress.
    //  - When several neighbours align to the same snapped edge, a guide is emitted for each.
    // #######################################################################################
    internal static class LabelEditorSnapGeometry
    {
        // The four edges ApplyNewRectangleSnap snaps in turn, in the order it snaps them. Hoisted
        // out of the loop that reads it because that loop runs per pointer-move event; never
        // mutated.
        private static readonly LabelEditorDragMode[] EdgeSnapOrder =
        {
            LabelEditorDragMode.ResizeTop,
            LabelEditorDragMode.ResizeBottom,
            LabelEditorDragMode.ResizeLeft,
            LabelEditorDragMode.ResizeRight,
        };

        public static void ApplyResizeSnap(
            LabelEditorSnapContext context,
            EditableComponentHighlight currentHighlight,
            ref double left,
            ref double top,
            ref double right,
            ref double bottom,
            List<(Point Start, Point End)> snapGuides,
            bool suppressSnap,
            bool snapOnMatch = true)
        {
            const double snapThreshold = 2.0;
            const double epsilon = 0.001;
            const double guideMatchThreshold = 0.5;

            // The context is the only source of the mode - a caller wanting a different one passes
            // context.WithDragMode(...).
            LabelEditorDragMode dragMode = context.DragMode;

            if (suppressSnap ||
                dragMode == LabelEditorDragMode.None ||
                dragMode == LabelEditorDragMode.Move)
            {
                return;
            }

            string schematicName = context.SchematicName;

            bool resizesTop =
                dragMode == LabelEditorDragMode.ResizeTop ||
                dragMode == LabelEditorDragMode.ResizeTopLeft ||
                dragMode == LabelEditorDragMode.ResizeTopRight;

            bool resizesBottom =
                dragMode == LabelEditorDragMode.ResizeBottom ||
                dragMode == LabelEditorDragMode.ResizeBottomLeft ||
                dragMode == LabelEditorDragMode.ResizeBottomRight;

            bool resizesLeft =
                dragMode == LabelEditorDragMode.ResizeLeft ||
                dragMode == LabelEditorDragMode.ResizeTopLeft ||
                dragMode == LabelEditorDragMode.ResizeBottomLeft;

            bool resizesRight =
                dragMode == LabelEditorDragMode.ResizeRight ||
                dragMode == LabelEditorDragMode.ResizeTopRight ||
                dragMode == LabelEditorDragMode.ResizeBottomRight;

            static bool RangesOverlap(double a1, double a2, double b1, double b2)
            {
                return Math.Min(a2, b2) > Math.Max(a1, b1);
            }

            static Rect BuildCurrentRect(double leftValue, double topValue, double rightValue, double bottomValue)
            {
                return new Rect(
                    leftValue,
                    topValue,
                    Math.Max(1.0, rightValue - leftValue),
                    Math.Max(1.0, bottomValue - topValue));
            }

            Rect? visiblePixelRect = context.VisiblePixelRect;

            bool IsRectVisibleInCurrentView(Rect rect)
            {
                if (!visiblePixelRect.HasValue)
                {
                    return true;
                }

                var visibleRect = visiblePixelRect.Value;

                return rect.Right > visibleRect.Left &&
                       rect.Left < visibleRect.Right &&
                       rect.Bottom > visibleRect.Top &&
                       rect.Top < visibleRect.Bottom;
            }

            bool IsVerticalPathBlocked(double sourceY, double targetY, Rect currentRect, EditableComponentHighlight targetHighlight)
            {
                double minY = Math.Min(sourceY, targetY);
                double maxY = Math.Max(sourceY, targetY);

                foreach (var other in context.WorkingHighlights)
                {
                    if (ReferenceEquals(other, currentHighlight) ||
                        ReferenceEquals(other, targetHighlight) ||
                        !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var otherRect = new Rect(other.X, other.Y, other.Width, other.Height);

                    if (!RangesOverlap(currentRect.Left, currentRect.Right, otherRect.Left, otherRect.Right))
                    {
                        continue;
                    }

                    if (otherRect.Bottom > minY && otherRect.Top < maxY)
                    {
                        return true;
                    }
                }

                return false;
            }

            bool IsHorizontalPathBlocked(double sourceX, double targetX, Rect currentRect, EditableComponentHighlight targetHighlight)
            {
                double minX = Math.Min(sourceX, targetX);
                double maxX = Math.Max(sourceX, targetX);

                foreach (var other in context.WorkingHighlights)
                {
                    if (ReferenceEquals(other, currentHighlight) ||
                        ReferenceEquals(other, targetHighlight) ||
                        !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var otherRect = new Rect(other.X, other.Y, other.Width, other.Height);

                    if (!RangesOverlap(currentRect.Top, currentRect.Bottom, otherRect.Top, otherRect.Bottom))
                    {
                        continue;
                    }

                    if (otherRect.Right > minX && otherRect.Left < maxX)
                    {
                        return true;
                    }
                }

                return false;
            }

            static bool TryBuildHorizontalGuide(Rect currentRect, Rect targetRect, double y, out (Point Start, Point End) guide)
            {
                guide = default;

                double startX = Math.Min(currentRect.Left, targetRect.Left);
                double endX = Math.Max(currentRect.Right, targetRect.Right);

                if (endX - startX <= 0.01)
                {
                    return false;
                }

                guide = (new Point(startX, y), new Point(endX, y));
                return true;
            }

            static bool TryBuildVerticalGuide(Rect currentRect, Rect targetRect, double x, out (Point Start, Point End) guide)
            {
                guide = default;

                double startY = Math.Min(currentRect.Top, targetRect.Top);
                double endY = Math.Max(currentRect.Bottom, targetRect.Bottom);

                if (endY - startY <= 0.01)
                {
                    return false;
                }

                guide = (new Point(x, startY), new Point(x, endY));
                return true;
            }

            if (resizesTop || resizesBottom)
            {
                Rect currentRect = BuildCurrentRect(left, top, right, bottom);
                double sourceY = resizesTop ? currentRect.Top : currentRect.Bottom;
                double bestDistance = snapThreshold + 0.001;
                double bestY = sourceY;
                var bestTargets = new List<EditableComponentHighlight>();

                void ConsiderVerticalCandidate(EditableComponentHighlight other, double candidateY)
                {
                    double distance = Math.Abs(sourceY - candidateY);
                    if (distance > snapThreshold ||
                        IsVerticalPathBlocked(sourceY, candidateY, currentRect, other))
                    {
                        return;
                    }

                    if (!snapOnMatch)
                    {
                        if (distance > guideMatchThreshold)
                        {
                            return;
                        }

                        if (bestTargets.Count == 0)
                        {
                            bestY = candidateY;
                            bestTargets.Add(other);
                            return;
                        }

                        if (Math.Abs(candidateY - bestY) <= guideMatchThreshold &&
                            !bestTargets.Contains(other))
                        {
                            bestTargets.Add(other);
                        }

                        return;
                    }

                    if (distance < bestDistance - epsilon)
                    {
                        bestDistance = distance;
                        bestY = candidateY;
                        bestTargets.Clear();
                        bestTargets.Add(other);
                        return;
                    }

                    if (Math.Abs(distance - bestDistance) <= epsilon &&
                        Math.Abs(candidateY - bestY) <= epsilon &&
                        !bestTargets.Contains(other))
                    {
                        bestTargets.Add(other);
                    }
                }

                foreach (var other in context.WorkingHighlights)
                {
                    if (ReferenceEquals(other, currentHighlight) ||
                        context.IsSelected(other) ||
                        !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var otherRect = new Rect(other.X, other.Y, other.Width, other.Height);
                    if (!IsRectVisibleInCurrentView(otherRect))
                    {
                        continue;
                    }

                    ConsiderVerticalCandidate(other, otherRect.Top);
                    ConsiderVerticalCandidate(other, otherRect.Bottom);
                }

                if (bestTargets.Count > 0)
                {
                    if (snapOnMatch)
                    {
                        if (resizesTop)
                        {
                            top = bestY;
                        }
                        else
                        {
                            bottom = bestY;
                        }
                    }

                    currentRect = BuildCurrentRect(left, top, right, bottom);

                    foreach (var bestTarget in bestTargets)
                    {
                        var targetRect = new Rect(bestTarget.X, bestTarget.Y, bestTarget.Width, bestTarget.Height);

                        if (TryBuildHorizontalGuide(currentRect, targetRect, bestY, out var guide))
                        {
                            snapGuides.Add(guide);
                        }
                    }
                }
            }

            if (resizesLeft || resizesRight)
            {
                Rect currentRect = BuildCurrentRect(left, top, right, bottom);
                double sourceX = resizesLeft ? currentRect.Left : currentRect.Right;
                double bestDistance = snapThreshold + 0.001;
                double bestX = sourceX;
                var bestTargets = new List<EditableComponentHighlight>();

                void ConsiderHorizontalCandidate(EditableComponentHighlight other, double candidateX)
                {
                    double distance = Math.Abs(sourceX - candidateX);
                    if (distance > snapThreshold ||
                        IsHorizontalPathBlocked(sourceX, candidateX, currentRect, other))
                    {
                        return;
                    }

                    if (!snapOnMatch)
                    {
                        if (distance > guideMatchThreshold)
                        {
                            return;
                        }

                        if (bestTargets.Count == 0)
                        {
                            bestX = candidateX;
                            bestTargets.Add(other);
                            return;
                        }

                        if (Math.Abs(candidateX - bestX) <= guideMatchThreshold &&
                            !bestTargets.Contains(other))
                        {
                            bestTargets.Add(other);
                        }

                        return;
                    }

                    if (distance < bestDistance - epsilon)
                    {
                        bestDistance = distance;
                        bestX = candidateX;
                        bestTargets.Clear();
                        bestTargets.Add(other);
                        return;
                    }

                    if (Math.Abs(distance - bestDistance) <= epsilon &&
                        Math.Abs(candidateX - bestX) <= epsilon &&
                        !bestTargets.Contains(other))
                    {
                        bestTargets.Add(other);
                    }
                }

                foreach (var other in context.WorkingHighlights)
                {
                    if (ReferenceEquals(other, currentHighlight) ||
                        context.IsSelected(other) ||
                        !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var otherRect = new Rect(other.X, other.Y, other.Width, other.Height);
                    if (!IsRectVisibleInCurrentView(otherRect))
                    {
                        continue;
                    }

                    ConsiderHorizontalCandidate(other, otherRect.Left);
                    ConsiderHorizontalCandidate(other, otherRect.Right);
                }

                if (bestTargets.Count > 0)
                {
                    if (snapOnMatch)
                    {
                        if (resizesLeft)
                        {
                            left = bestX;
                        }
                        else
                        {
                            right = bestX;
                        }
                    }

                    currentRect = BuildCurrentRect(left, top, right, bottom);

                    foreach (var bestTarget in bestTargets)
                    {
                        var targetRect = new Rect(bestTarget.X, bestTarget.Y, bestTarget.Width, bestTarget.Height);

                        if (TryBuildVerticalGuide(currentRect, targetRect, bestX, out var guide))
                        {
                            snapGuides.Add(guide);
                        }
                    }
                }
            }
        }

        // ###########################################################################################
        // Snaps the moved selection bounds to nearby neighbor edges while preserving the current
        // selection layout. SHIFT still suppresses the snap, matching resize behavior.
        // When snapOnMatch is false, no movement is applied and guides are shown only for exact matches.
        // ###########################################################################################
        public static void ApplyMoveSnap(
            LabelEditorSnapContext context,
            IReadOnlyList<EditableComponentHighlight> selectedHighlights,
            IReadOnlyDictionary<EditableComponentHighlight, Rect> sourceRects,
            ref Rect movedSelectionBounds,
            List<(Point Start, Point End)> snapGuides,
            bool suppressSnap,
            bool snapOnMatch = true)
        {
            const double snapThreshold = 2.0;
            const double epsilon = 0.001;
            const double guideMatchThreshold = 0.5;

            if (suppressSnap || selectedHighlights.Count == 0)
            {
                return;
            }

            string schematicName = context.SchematicName;
            var selectedSet = new HashSet<EditableComponentHighlight>(selectedHighlights);

            static bool RangesOverlap(double a1, double a2, double b1, double b2)
            {
                return Math.Min(a2, b2) > Math.Max(a1, b1);
            }

            Rect GetRect(EditableComponentHighlight highlight)
            {
                if (sourceRects.TryGetValue(highlight, out var sourceRect))
                {
                    return sourceRect;
                }

                return new Rect(highlight.X, highlight.Y, highlight.Width, highlight.Height);
            }

            Rect? visiblePixelRect = context.VisiblePixelRect;

            bool IsRectVisibleInCurrentView(Rect rect)
            {
                if (!visiblePixelRect.HasValue)
                {
                    return true;
                }

                var visibleRect = visiblePixelRect.Value;

                return rect.Right > visibleRect.Left &&
                       rect.Left < visibleRect.Right &&
                       rect.Bottom > visibleRect.Top &&
                       rect.Top < visibleRect.Bottom;
            }

            bool IsVerticalPathBlocked(double sourceY, double targetY, Rect currentRect, EditableComponentHighlight targetHighlight)
            {
                double minY = Math.Min(sourceY, targetY);
                double maxY = Math.Max(sourceY, targetY);

                foreach (var other in context.WorkingHighlights)
                {
                    if (selectedSet.Contains(other) ||
                        ReferenceEquals(other, targetHighlight) ||
                        !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var otherRect = GetRect(other);

                    if (!RangesOverlap(currentRect.Left, currentRect.Right, otherRect.Left, otherRect.Right))
                    {
                        continue;
                    }

                    if (otherRect.Bottom > minY && otherRect.Top < maxY)
                    {
                        return true;
                    }
                }

                return false;
            }

            bool IsHorizontalPathBlocked(double sourceX, double targetX, Rect currentRect, EditableComponentHighlight targetHighlight)
            {
                double minX = Math.Min(sourceX, targetX);
                double maxX = Math.Max(sourceX, targetX);

                foreach (var other in context.WorkingHighlights)
                {
                    if (selectedSet.Contains(other) ||
                        ReferenceEquals(other, targetHighlight) ||
                        !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var otherRect = GetRect(other);

                    if (!RangesOverlap(currentRect.Top, currentRect.Bottom, otherRect.Top, otherRect.Bottom))
                    {
                        continue;
                    }

                    if (otherRect.Right > minX && otherRect.Left < maxX)
                    {
                        return true;
                    }
                }

                return false;
            }

            static bool TryBuildHorizontalGuide(Rect currentRect, Rect targetRect, double y, out (Point Start, Point End) guide)
            {
                guide = default;

                double startX = Math.Min(currentRect.Left, targetRect.Left);
                double endX = Math.Max(currentRect.Right, targetRect.Right);

                if (endX - startX <= 0.01)
                {
                    return false;
                }

                guide = (new Point(startX, y), new Point(endX, y));
                return true;
            }

            static bool TryBuildVerticalGuide(Rect currentRect, Rect targetRect, double x, out (Point Start, Point End) guide)
            {
                guide = default;

                double startY = Math.Min(currentRect.Top, targetRect.Top);
                double endY = Math.Max(currentRect.Bottom, targetRect.Bottom);

                if (endY - startY <= 0.01)
                {
                    return false;
                }

                guide = (new Point(x, startY), new Point(x, endY));
                return true;
            }

            Rect currentMovedSelectionBounds = movedSelectionBounds;

            double bestDeltaY = 0.0;
            double bestDistanceY = snapThreshold + 0.001;
            var bestVerticalTargets = new List<(EditableComponentHighlight Target, double Y)>();

            void ConsiderVerticalCandidate(EditableComponentHighlight other, double sourceY, double candidateY)
            {
                double delta = candidateY - sourceY;
                double distance = Math.Abs(delta);

                if (distance > snapThreshold ||
                    IsVerticalPathBlocked(sourceY, candidateY, currentMovedSelectionBounds, other))
                {
                    return;
                }

                if (!snapOnMatch)
                {
                    if (distance > guideMatchThreshold)
                    {
                        return;
                    }

                    if (bestVerticalTargets.Count == 0)
                    {
                        bestDeltaY = delta;
                        bestVerticalTargets.Add((other, candidateY));
                        return;
                    }

                    if (Math.Abs(delta - bestDeltaY) <= guideMatchThreshold &&
                        !bestVerticalTargets.Any(target =>
                            ReferenceEquals(target.Target, other) &&
                            Math.Abs(target.Y - candidateY) <= guideMatchThreshold))
                    {
                        bestVerticalTargets.Add((other, candidateY));
                    }

                    return;
                }

                if (distance < bestDistanceY - epsilon)
                {
                    bestDistanceY = distance;
                    bestDeltaY = delta;
                    bestVerticalTargets.Clear();
                    bestVerticalTargets.Add((other, candidateY));
                    return;
                }

                if (Math.Abs(distance - bestDistanceY) <= epsilon &&
                    Math.Abs(delta - bestDeltaY) <= epsilon &&
                    !bestVerticalTargets.Any(target =>
                        ReferenceEquals(target.Target, other) &&
                        Math.Abs(target.Y - candidateY) <= epsilon))
                {
                    bestVerticalTargets.Add((other, candidateY));
                }
            }

            foreach (var other in context.WorkingHighlights)
            {
                if (selectedSet.Contains(other) ||
                    !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var otherRect = GetRect(other);
                if (!IsRectVisibleInCurrentView(otherRect))
                {
                    continue;
                }

                ConsiderVerticalCandidate(other, currentMovedSelectionBounds.Top, otherRect.Top);
                ConsiderVerticalCandidate(other, currentMovedSelectionBounds.Top, otherRect.Bottom);
                ConsiderVerticalCandidate(other, currentMovedSelectionBounds.Bottom, otherRect.Top);
                ConsiderVerticalCandidate(other, currentMovedSelectionBounds.Bottom, otherRect.Bottom);
            }

            if (bestVerticalTargets.Count > 0)
            {
                if (snapOnMatch)
                {
                    movedSelectionBounds = new Rect(
                        movedSelectionBounds.X,
                        movedSelectionBounds.Y + bestDeltaY,
                        movedSelectionBounds.Width,
                        movedSelectionBounds.Height);
                }

                currentMovedSelectionBounds = movedSelectionBounds;

                foreach (var target in bestVerticalTargets)
                {
                    var targetRect = GetRect(target.Target);

                    if (TryBuildHorizontalGuide(currentMovedSelectionBounds, targetRect, target.Y, out var guide))
                    {
                        snapGuides.Add(guide);
                    }
                }
            }

            currentMovedSelectionBounds = movedSelectionBounds;

            double bestDeltaX = 0.0;
            double bestDistanceX = snapThreshold + 0.001;
            var bestHorizontalTargets = new List<(EditableComponentHighlight Target, double X)>();

            void ConsiderHorizontalCandidate(EditableComponentHighlight other, double sourceX, double candidateX)
            {
                double delta = candidateX - sourceX;
                double distance = Math.Abs(delta);

                if (distance > snapThreshold ||
                    IsHorizontalPathBlocked(sourceX, candidateX, currentMovedSelectionBounds, other))
                {
                    return;
                }

                if (!snapOnMatch)
                {
                    if (distance > guideMatchThreshold)
                    {
                        return;
                    }

                    if (bestHorizontalTargets.Count == 0)
                    {
                        bestDeltaX = delta;
                        bestHorizontalTargets.Add((other, candidateX));
                        return;
                    }

                    if (Math.Abs(delta - bestDeltaX) <= guideMatchThreshold &&
                        !bestHorizontalTargets.Any(target =>
                            ReferenceEquals(target.Target, other) &&
                            Math.Abs(target.X - candidateX) <= guideMatchThreshold))
                    {
                        bestHorizontalTargets.Add((other, candidateX));
                    }

                    return;
                }

                if (distance < bestDistanceX - epsilon)
                {
                    bestDistanceX = distance;
                    bestDeltaX = delta;
                    bestHorizontalTargets.Clear();
                    bestHorizontalTargets.Add((other, candidateX));
                    return;
                }

                if (Math.Abs(distance - bestDistanceX) <= epsilon &&
                    Math.Abs(delta - bestDeltaX) <= epsilon &&
                    !bestHorizontalTargets.Any(target =>
                        ReferenceEquals(target.Target, other) &&
                        Math.Abs(target.X - candidateX) <= epsilon))
                {
                    bestHorizontalTargets.Add((other, candidateX));
                }
            }

            foreach (var other in context.WorkingHighlights)
            {
                if (selectedSet.Contains(other) ||
                    !string.Equals(other.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var otherRect = GetRect(other);
                if (!IsRectVisibleInCurrentView(otherRect))
                {
                    continue;
                }

                ConsiderHorizontalCandidate(other, currentMovedSelectionBounds.Left, otherRect.Left);
                ConsiderHorizontalCandidate(other, currentMovedSelectionBounds.Left, otherRect.Right);
                ConsiderHorizontalCandidate(other, currentMovedSelectionBounds.Right, otherRect.Left);
                ConsiderHorizontalCandidate(other, currentMovedSelectionBounds.Right, otherRect.Right);
            }

            if (bestHorizontalTargets.Count > 0)
            {
                if (snapOnMatch)
                {
                    movedSelectionBounds = new Rect(
                        movedSelectionBounds.X + bestDeltaX,
                        movedSelectionBounds.Y,
                        movedSelectionBounds.Width,
                        movedSelectionBounds.Height);
                }

                currentMovedSelectionBounds = movedSelectionBounds;

                foreach (var target in bestHorizontalTargets)
                {
                    var targetRect = GetRect(target.Target);

                    if (TryBuildVerticalGuide(currentMovedSelectionBounds, targetRect, target.X, out var guide))
                    {
                        snapGuides.Add(guide);
                    }
                }
            }
        }

        // ###########################################################################################
        // Applies neighbor-edge snapping to a newly drawn rectangle by reusing the existing resize
        // snap logic for all four edges while the rectangle is still only a draft.
        // ###########################################################################################
        public static void ApplyNewRectangleSnap(
            LabelEditorSnapContext context,
            ref Rect rect,
            List<(Point Start, Point End)> snapGuides,
            bool suppressSnap)
        {
            if (suppressSnap ||
                rect.Width <= 0 ||
                rect.Height <= 0 ||
                string.IsNullOrWhiteSpace(context.SchematicName))
            {
                return;
            }

            double left = rect.Left;
            double top = rect.Top;
            double right = rect.Right;
            double bottom = rect.Bottom;

            var draftHighlight = new EditableComponentHighlight
            {
                SchematicName = context.SchematicName,
                X = rect.X,
                Y = rect.Y,
                Width = rect.Width,
                Height = rect.Height
            };

            // Runs the resize snap once per edge. The tab version set this.thisLabelEditorDragMode
            // around each call and restored it in a finally; here each pass gets its own copy of the
            // context carrying that edge's mode. Same four passes, same order, no shared state to
            // save and restore.
            //
            // EdgeSnapOrder is a static field rather than an inline array literal because this runs
            // on the pointer-move path - once per move event while a new highlight is being dragged
            // out - and an inline literal would allocate a four-element array on every one of them.
            foreach (var edgeMode in EdgeSnapOrder)
            {
                ApplyResizeSnap(
                    context.WithDragMode(edgeMode),
                    draftHighlight,
                    ref left,
                    ref top,
                    ref right,
                    ref bottom,
                    snapGuides,
                    suppressSnap: false,
                    snapOnMatch: true);
            }

            rect = new Rect(
                left,
                top,
                Math.Max(1.0, right - left),
                Math.Max(1.0, bottom - top));
        }
    }
}
