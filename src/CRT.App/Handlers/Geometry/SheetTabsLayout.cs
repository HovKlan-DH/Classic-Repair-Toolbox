using Avalonia;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // One item of the sheet tabs' strip: its size as measured (margins included), whether a line starts with
    // it, and whether it is a group's heading rather than a tab.
    internal readonly record struct SheetTabsItem(double Width, double Height, bool StartsLine, bool IsHeading);

    // ###########################################################################################
    // WHERE THE MAINTAINER TAB'S SHEET TABS STAND (owner request, 2026-11-06: "Here-and-now I see all tabs
    // horizontally, and if there are too many, it wraps into next line, which makes it hard to overview ...
    // just have the 'Components' part on a new-line"). The table's strip is its groups' headings and their
    // tabs (BoardTableTabs); laid out as words in a paragraph, a line broke wherever the width ran out -
    // inside a group, or between a heading and its first tab.
    //
    //   - A LINE STARTS AT A GROUP THAT ASKS FOR ONE (BoardTableTabs.StartsLine): "Board" and "KiCad" share
    //     the first line, "Components" has the second.
    //   - THE HEADINGS THAT START A LINE STAND IN ONE COLUMN, as wide as the widest of them, so every line's
    //     tabs begin at the same place and the groups read as a list.
    //   - A LINE TOO LONG FOR THE WIDTH WRAPS UNDER ITS OWN TABS, never back under the headings' column - and
    //     a heading inside a line goes along with its first tab, never left at the end of a line alone.
    //
    // Pure: items in, rectangles out (BoardTableEditor's SheetTabsPanel measures and arranges by it).
    // ###########################################################################################
    internal static class SheetTabsLayout
    {
        public static IReadOnlyList<Rect> Arrange(IReadOnlyList<SheetTabsItem> items, double availableWidth, out Size size)
        {
            ArgumentNullException.ThrowIfNull(items);

            double width = double.IsNaN(availableWidth) || availableWidth <= 0 ? double.PositiveInfinity : availableWidth;

            // The headings' column: as wide as the widest heading a line starts with.
            double column = items
                .Where((item, index) => item.IsHeading && (item.StartsLine || index == 0))
                .Select(item => item.Width)
                .DefaultIfEmpty(0)
                .Max();

            var lineOf = new int[items.Count];
            var xOf = new double[items.Count];
            var lineHeights = new List<double>();

            int line = -1;
            double x = 0;
            double indent = 0;

            void NewLine(double startAt)
            {
                line++;
                lineHeights.Add(0);
                x = startAt;
            }

            for (int i = 0; i < items.Count; i++)
            {
                SheetTabsItem item = items[i];

                if (line < 0 || item.StartsLine)
                {
                    NewLine(0);

                    if (item.IsHeading)
                    {
                        lineOf[i] = line;
                        xOf[i] = 0;
                        lineHeights[line] = Math.Max(lineHeights[line], item.Height);
                        indent = column;
                        x = column;
                        continue;
                    }

                    indent = 0;
                }

                // A heading inside a line keeps its first tab beside it.
                double needs = item.Width + (item.IsHeading && i + 1 < items.Count && !items[i + 1].StartsLine ? items[i + 1].Width : 0);

                if (x + needs > width && x > indent)
                {
                    NewLine(indent);
                }

                lineOf[i] = line;
                xOf[i] = x;
                lineHeights[line] = Math.Max(lineHeights[line], item.Height);
                x += item.Width;
            }

            var tops = new double[lineHeights.Count];
            double top = 0;

            for (int l = 0; l < lineHeights.Count; l++)
            {
                tops[l] = top;
                top += lineHeights[l];
            }

            var rects = new List<Rect>(items.Count);
            double right = 0;

            for (int i = 0; i < items.Count; i++)
            {
                var rect = new Rect(xOf[i], tops[lineOf[i]], items[i].Width, lineHeights[lineOf[i]]);

                rects.Add(rect);
                right = Math.Max(right, rect.Right);
            }

            size = new Size(right, top);

            return rects;
        }
    }
}
