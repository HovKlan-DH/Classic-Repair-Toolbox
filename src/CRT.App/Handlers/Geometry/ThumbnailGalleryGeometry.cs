using Avalonia;
using Avalonia.Input;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Auto-fits a fixed set of thumbnails into an arbitrary window client area with no scrolling:
    // given a thumbnail count and the available size, picks the rows x columns grid and uniform
    // cell size that shows every thumbnail as large as possible. Used by the detached schematics
    // thumbnails window (SchematicsThumbnailGallery/ThumbnailGalleryPanel) so a tiny window still
    // shows every thumbnail, just smaller, and a maximized one shows them larger - never scrolled.
    //
    // Every tile gets the SAME uniform cell size, regardless of its own image's aspect ratio - a
    // deliberate simplification over true aspect-ratio-aware bin packing. A schematic and a PCB
    // photo both letterbox via Stretch="Uniform" inside their shared cell. This keeps the layout
    // O(itemCount) and predictable; do not go looking for per-item aspect-aware packing here, it
    // was consciously left out.
    // ###########################################################################################
    public static class ThumbnailGalleryGeometry
    {
        public readonly record struct GridShape(int Rows, int Columns);

        // ###########################################################################################
        // Picks the rows x columns grid that fits itemCount uniform cells into availableSize (after
        // subtracting spacing between cells and margin around the edge) while making each cell as
        // large as possible.
        //
        // Tries every column count from 1..itemCount, derives the row count needed, and keeps the
        // shape whose resulting cell has the largest LIMITING dimension (the smaller of its width
        // and height) - that is the shape that lets a uniform square-ish cell grow the most. O(N),
        // fine for the realistic thumbnail counts here (tens at most).
        //
        // itemCount <= 0 yields (0, 0): an empty gallery, not a crash.
        // ###########################################################################################
        public static GridShape ChooseGridShape(int itemCount, Size availableSize, double spacing, double margin)
        {
            if (itemCount <= 0)
            {
                return new GridShape(0, 0);
            }

            var best = new GridShape(itemCount, 1);
            double bestScore = double.NegativeInfinity;
            int bestImbalance = int.MaxValue;

            for (int columns = 1; columns <= itemCount; columns++)
            {
                int rows = (itemCount + columns - 1) / columns;
                var shape = new GridShape(rows, columns);
                Size cell = ComputeCellSize(shape, availableSize, spacing, margin);
                double score = Math.Min(cell.Width, cell.Height);

                // Ties are broken towards the squarest grid. They arise when an axis is
                // unconstrained, where every shape yields the same fallback cell - without a
                // tie-break the first candidate wins and that is always a single column, which is
                // the one shape an auto-fit gallery should never fall back to.
                int imbalance = Math.Abs(rows - columns);

                bool isBetter = score > bestScore ||
                    (score == bestScore && imbalance < bestImbalance);

                if (isBetter)
                {
                    bestScore = score;
                    bestImbalance = imbalance;
                    best = shape;
                }
            }

            return best;
        }

        // The size a cell falls back to on an axis the caller left unconstrained. An auto-fit
        // layout has nothing to fit INTO on such an axis, so it stops auto-fitting and uses an
        // ordinary thumbnail size instead.
        private const double UnconstrainedCellSize = 160.0;

        // ###########################################################################################
        // The uniform cell size for a chosen grid shape and available size. Clamped to a 1x1 floor
        // so a degenerate (zero, negative, or too-small-for-the-margins) availableSize never
        // produces a NaN or Infinity layout - a tiny window still gets a valid, if useless, size.
        //
        // An INFINITE axis is clamped too, to a fixed fallback rather than to a floor: Math.Max only
        // guards the small side, so an unconstrained measure pass (a ScrollViewer with scrolling
        // enabled, a StackPanel, an auto-sized Window) otherwise produced an Infinity cell - which
        // every child was then measured against, and which propagated out as an Infinity desired
        // size that Avalonia treats as a layout error. It also made ChooseGridShape degenerate:
        // every column count scored Infinity, so the "strictly better" comparison never fired after
        // the first and it always picked a single column.
        // ###########################################################################################
        public static Size ComputeCellSize(GridShape shape, Size availableSize, double spacing, double margin)
        {
            if (shape.Rows <= 0 || shape.Columns <= 0)
            {
                return new Size(0, 0);
            }

            double cellWidth = double.IsInfinity(availableSize.Width) || double.IsNaN(availableSize.Width)
                ? UnconstrainedCellSize
                : Math.Max(1.0, (availableSize.Width - (margin * 2) - (spacing * (shape.Columns - 1))) / shape.Columns);

            double cellHeight = double.IsInfinity(availableSize.Height) || double.IsNaN(availableSize.Height)
                ? UnconstrainedCellSize
                : Math.Max(1.0, (availableSize.Height - (margin * 2) - (spacing * (shape.Rows - 1))) / shape.Rows);

            return new Size(cellWidth, cellHeight);
        }

        // ###########################################################################################
        // Top-left positions for itemCount cells of cellSize, filled row-first left-to-right, top
        // row first - the same reading order the existing vertical list already presents thumbnails
        // in, so switching layouts reflows what's on screen but never reorders it.
        // ###########################################################################################
        public static IReadOnlyList<Point> ArrangeGrid(GridShape shape, Size cellSize, double spacing, double margin, int itemCount)
        {
            var positions = new List<Point>(Math.Max(0, itemCount));

            if (shape.Columns <= 0 || itemCount <= 0)
            {
                return positions;
            }

            for (int index = 0; index < itemCount; index++)
            {
                int row = index / shape.Columns;
                int column = index % shape.Columns;

                double x = margin + (column * (cellSize.Width + spacing));
                double y = margin + (row * (cellSize.Height + spacing));

                positions.Add(new Point(x, y));
            }

            return positions;
        }

        // ###########################################################################################
        // Convenience: grid shape, cell size and positions together from one call, so a caller never
        // computes them separately and risks the three disagreeing about itemCount.
        // ###########################################################################################
        public static (GridShape Shape, Size CellSize, IReadOnlyList<Point> Positions) BuildLayout(
            int itemCount, Size availableSize, double spacing, double margin)
        {
            GridShape shape = ChooseGridShape(itemCount, availableSize, spacing, margin);
            Size cellSize = ComputeCellSize(shape, availableSize, spacing, margin);
            IReadOnlyList<Point> positions = ArrangeGrid(shape, cellSize, spacing, margin, itemCount);

            return (shape, cellSize, positions);
        }

        // ###########################################################################################
        // The index of the cell whose rect contains the point, or whose center is nearest if the
        // point falls in the gap between cells - used while dragging so a drag never "loses" the
        // target between cells the way a naive rect-containment check would while crossing spacing.
        //
        // itemCount <= 0 returns -1 (no valid target).
        // ###########################################################################################
        public static int FindNearestCellIndex(IReadOnlyList<Point> cellTopLefts, Size cellSize, Point point, int itemCount)
        {
            if (itemCount <= 0 || cellTopLefts.Count == 0)
            {
                return -1;
            }

            int nearestIndex = 0;
            double nearestDistanceSquared = double.PositiveInfinity;

            int count = Math.Min(itemCount, cellTopLefts.Count);
            for (int index = 0; index < count; index++)
            {
                Point topLeft = cellTopLefts[index];
                Point center = new(topLeft.X + (cellSize.Width / 2), topLeft.Y + (cellSize.Height / 2));

                double dx = point.X - center.X;
                double dy = point.Y - center.Y;
                double distanceSquared = (dx * dx) + (dy * dy);

                if (distanceSquared < nearestDistanceSquared)
                {
                    nearestDistanceSquared = distanceSquared;
                    nearestIndex = index;
                }
            }

            return nearestIndex;
        }

        // ###########################################################################################
        // The cell the keyboard moves to from cell index, in a grid of itemCount cells and columns
        // columns filled as ArrangeGrid fills it - for the arrow keys, Home/End and Page Up/Down in
        // the gallery (code review, 2026-10-15: the grid said nothing about it, so the keys did
        // nothing). Left and Right follow reading order across rows, as a file manager does; Up and
        // Down move a whole row, and Down from a cell with nothing straight below it in a short last
        // row goes to the last cell. Every cell is on one page (the gallery never scrolls), so Page
        // Up/Down go to the first/last cell. Next/Previous wrap round only when asked to.
        //
        // -1 when there is nowhere to go. From no cell at all (index outside the grid) every
        // direction but Last/End/Page Down starts at the first cell.
        // ###########################################################################################
        public static int NeighbourIndex(int index, int itemCount, int columns, NavigationDirection direction, bool wrap)
        {
            if (itemCount <= 0)
            {
                return -1;
            }

            int last = itemCount - 1;

            switch (direction)
            {
                case NavigationDirection.First:
                case NavigationDirection.PageUp:
                    return 0;

                case NavigationDirection.Last:
                case NavigationDirection.PageDown:
                    return last;
            }

            if (index < 0 || index > last)
            {
                return 0;
            }

            columns = Math.Max(1, columns);

            switch (direction)
            {
                case NavigationDirection.Left:
                    return index > 0 ? index - 1 : -1;

                case NavigationDirection.Right:
                    return index < last ? index + 1 : -1;

                case NavigationDirection.Previous:
                    return index > 0 ? index - 1 : (wrap ? last : -1);

                case NavigationDirection.Next:
                    return index < last ? index + 1 : (wrap ? 0 : -1);

                case NavigationDirection.Up:
                    return index >= columns ? index - columns : -1;

                case NavigationDirection.Down:
                    if (index + columns <= last)
                    {
                        return index + columns;
                    }

                    return (index / columns) < (last / columns) ? last : -1;

                default:
                    return -1;
            }
        }
    }
}
