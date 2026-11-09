using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A LIST STANDS, AS ANOTHER LIST OF THE SAME ROWS SHOULD STAND (owner request, 2026-11-09: "Toggling between tabs
    // (e.g. "Contribute" and "Schematics") should keep the exact same components and categories selected, and it should
    // not change position in the lists"; Main.ContributeLists.cs). The Contribute tab and the other tabs each have their
    // lists in the same place, one shown at a time: as the tab changes, the list shown next is put where the one shown
    // until then stood - the same row at its top, as far into it.
    //
    // By ROW, never by offset alone: the two lists may not hold the same rows (the Contribute tab's follows the draft,
    // "With issues" narrows it), and their rows are all one height - a list's extent is its rows' heights added up.
    // ###########################################################################################
    internal static class ListPosition
    {
        // The row at the top of a list of `count` rows, `extent` tall, scrolled to `offset`, and how far into it - as a
        // share of a row, so it holds for a list whose rows are another height. Null for a list with nothing to measure.
        public static (int Row, double Into)? TopOf(double offset, double extent, int count)
        {
            if (count <= 0 || extent <= 0 || double.IsNaN(offset) || double.IsNaN(extent))
            {
                return null;
            }

            double row = extent / count;
            double clamped = Math.Clamp(offset, 0, extent);
            int top = Math.Min(count - 1, (int)Math.Floor(clamped / row));

            return (top, Math.Clamp((clamped - (top * row)) / row, 0, 1));
        }

        // ###########################################################################################
        // Which row of `to` stands for row `row` of `from`, by their keys: the same key, else the first row after it that
        // `to` has too, else the last of `to`. -1 for a list with no rows.
        // ###########################################################################################
        public static int MatchingRow(IReadOnlyList<string> from, int row, IReadOnlyList<string> to)
        {
            ArgumentNullException.ThrowIfNull(from);
            ArgumentNullException.ThrowIfNull(to);

            if (to.Count == 0)
            {
                return -1;
            }

            var places = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int at = to.Count - 1; at >= 0; at--)
            {
                places[to[at]] = at;
            }

            for (int at = Math.Max(0, row); at < from.Count; at++)
            {
                if (places.TryGetValue(from[at], out int found))
                {
                    return found;
                }
            }

            return Math.Min(Math.Max(0, row), to.Count - 1);
        }

        // The offset that shows `row` of a list of `count` rows, `extent` tall, at its top, `into` that far into it - no
        // further than the list can scroll in a `viewport` that tall.
        public static double OffsetOf(int row, double into, double extent, int count, double viewport)
        {
            if (count <= 0 || extent <= 0 || row < 0)
            {
                return 0;
            }

            double height = extent / count;

            return Math.Clamp((row + Math.Clamp(into, 0, 1)) * height, 0, Math.Max(0, extent - viewport));
        }
    }
}
