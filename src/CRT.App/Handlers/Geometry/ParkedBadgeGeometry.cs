using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Arranges the worklog pills that have no marked area to sit on.
    //
    // A worklog entry whose "Show marked area" is unticked draws no coloured rectangle, so its
    // "#N" pill has nothing to anchor to. Rather than hiding the pill - which would make the entry
    // invisible on the board and unreachable without opening the worklog list - it is parked in the
    // top-right corner of the schematic panel, where it is still readable and still clickable.
    //
    // They are laid out as a BLOCK growing down and to the left, not as one long column. A single
    // column of pills runs down over the middle of the board and hides more of it the more entries
    // there are; a compact block keeps them together in the corner and out of the way.
    //
    // Everything here is in the VIEWPORT's own coordinate space and independent of zoom and pan:
    // parked pills stay pinned to the corner while the board moves under them, which is the whole
    // point of parking them.
    // ###########################################################################################
    public static class ParkedBadgeGeometry
    {
        // ###########################################################################################
        // The grid a given number of pills is arranged into: how many COLUMNS wide, and how many
        // rows that actually needs.
        //
        // The progression:
        //
        //     1-2 pills     1 col   -> 1-2 rows    a single short column
        //     3-4 pills     2 cols  -> 2 rows      a second column appears
        //     5-9 pills     3 cols  -> 2-3 rows
        //     10-16 pills   4 cols  -> 3-4 rows
        //     17-25 pills   5 cols  -> 4-5 rows    ... and so on
        //
        // The column count grows square-ish, so the block stays compact in the corner rather than
        // long in either direction: a tall column covers the board, and a wide row runs into the
        // "Netlist names" panel and the thumbnails.
        //
        // One column before a second appears, deliberately - two pills side by side read as two
        // unrelated things, where one above the other reads as a list.
        //
        // Rows are DERIVED from the count and the columns rather than being the square's side,
        // because the block fills row-first and a part-filled last row needs no extra row of space:
        // 5 pills in 3 columns occupy 2 rows, not 3. The two were previously reported independently
        // and disagreed - the caller recomputed rows and discarded the returned value, so the
        // documented shape described a layout the code never produced.
        // ###########################################################################################
        public static (int Rows, int Columns) GetGridShape(int badgeCount)
        {
            if (badgeCount <= 0)
            {
                return (0, 0);
            }

            int columns = GetColumnCount(badgeCount);
            int rows = (badgeCount + columns - 1) / columns;

            return (rows, columns);
        }

        // How many columns wide the block is. Split out because the row count depends on it, so the
        // two can never be derived from different rules.
        private static int GetColumnCount(int badgeCount)
        {
            if (badgeCount <= 2)
            {
                return 1;
            }

            if (badgeCount <= 4)
            {
                return 2;
            }

            return (int)Math.Ceiling(Math.Sqrt(badgeCount));
        }

        // ###########################################################################################
        // Top-left positions for the parked pills, arranged as a block anchored to the top-right
        // corner of the viewport, in the order given.
        //
        // Filled ROW-FIRST, top row first: pills run across the top row and only once it is full
        // does the next row start beneath it. Five pills therefore sit 3-over-2, not 3-down-then-2,
        // which is both what reading order expects and the shorter block - a part-filled LAST row
        // means the block is as flat as it can be, where a part-filled last column would leave it
        // as tall as a full one.
        //
        // Within a row the pills run right to left, so pill #1 stays in the corner the block is
        // anchored to and the block grows towards the board rather than off the edge.
        //
        // Every column is right-aligned to its own edge rather than the pills being left-aligned in
        // a shared column, because the pills differ in width ("#7" against "#12") and a ragged right
        // margin against the panel edge they are pinned to would look like a mistake.
        //
        // reservedRight is how much of the right-hand edge is already spoken for - the "Netlist
        // names" panel's width plus its margin when that panel is open, zero when it is not. The
        // whole block moves left by exactly that much, so it sits beside the panel instead of
        // underneath it, and slides back when the panel closes.
        //
        // A block wider or taller than the space available is pinned to the left and top edges
        // rather than pushed off them: unreadable pills at a sensible position beat correctly-sized
        // ones off-screen. The bottom clamp matters as much as the left one - the parked canvas is
        // inside the clipped schematic container, so a pill placed below the viewport is not merely
        // awkward, it is clipped away entirely and becomes both invisible AND unclickable, which is
        // the exact failure parking exists to prevent. Measured: 13 pills overflow a 120px-tall
        // panel, which a dragged splitter reaches easily.
        // ###########################################################################################
        public static List<Point> ArrangeInTopRightBlock(
            IReadOnlyList<Size> badgeSizes,
            Size viewportSize,
            double margin,
            double spacing,
            double reservedRight)
        {
            var positions = new List<Point>(badgeSizes?.Count ?? 0);

            if (badgeSizes == null || badgeSizes.Count == 0)
            {
                return positions;
            }

            var (rowCount, columnCount) = GetGridShape(badgeSizes.Count);
            if (columnCount <= 0 || rowCount <= 0)
            {
                return positions;
            }

            double rightEdge = viewportSize.Width - margin - Math.Max(0, reservedRight);

            // Each column is as wide as its own widest pill, so a column of narrow "#1".."#9" pills
            // does not reserve the width of a wide "#12" in a different column.
            var columnWidths = new double[columnCount];
            var rowHeights = new double[rowCount];

            for (int i = 0; i < badgeSizes.Count; i++)
            {
                int row = i / columnCount;
                int column = i % columnCount;

                columnWidths[column] = Math.Max(columnWidths[column], badgeSizes[i].Width);
                rowHeights[row] = Math.Max(rowHeights[row], badgeSizes[i].Height);
            }

            // Right edge of each column, walking leftwards from the viewport edge.
            var columnRightEdges = new double[columnCount];
            double edge = rightEdge;
            for (int column = 0; column < columnCount; column++)
            {
                columnRightEdges[column] = edge;
                edge -= columnWidths[column] + spacing;
            }

            // Top of each row, walking down from the margin.
            var rowTops = new double[rowCount];
            double top = margin;
            for (int row = 0; row < rowCount; row++)
            {
                rowTops[row] = top;
                top += rowHeights[row] + spacing;
            }

            for (int i = 0; i < badgeSizes.Count; i++)
            {
                int row = i / columnCount;
                int column = i % columnCount;

                double x = columnRightEdges[column] - badgeSizes[i].Width;
                double y = rowTops[row];

                // Never past the left edge - see the note above about pinning rather than pushing.
                if (x < margin)
                {
                    x = margin;
                }

                // ...and never past the bottom. Clamped per pill rather than by shrinking the whole
                // block, so the rows that DO fit keep their positions and only the overflowing ones
                // pile up on the last visible line - still reachable, where off-canvas is not.
                if (viewportSize.Height > 0)
                {
                    double lowestTop = viewportSize.Height - margin - badgeSizes[i].Height;
                    if (y > lowestTop)
                    {
                        y = Math.Max(margin, lowestTop);
                    }
                }

                positions.Add(new Point(x, y));
            }

            return positions;
        }
    }
}
