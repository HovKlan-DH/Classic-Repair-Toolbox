using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // One column of the component popup's pin list: its widest text, its heading included, and - for a column whose texts
    // wrap - its widest word, the least it can narrow to without breaking a word. A column not shown is all 0.
    public readonly record struct PinTableColumn(double Widest, double WidestWord, bool Wraps)
    {
        public bool IsShown => this.Widest > 0;
    }

    // ###########################################################################################
    // THE COMPONENT POPUP'S PIN LIST, COLUMN BY COLUMN (owner request, 2026-11-09: "align on empty spacing between columns -
    // it should be the same for all, and I do like the spacing between 'Pin' and 'Signal'. Should be determined by longest
    // text in each column + that spacing"): every column as wide as its widest text, and Gap between each two - the 20 that
    // stood between "Pin", 16 wide, and "Signal" in the 36 wide column "Pin" had until then.
    //
    // Only when the list has less room than that do the columns that wrap - the signal, the net and the reading - narrow,
    // the widest first, all of them down to one width and none below its widest word; the others keep their texts whole.
    // Too narrow even for every word whole - a popup of its least width holds some 190 for five columns - they come down
    // further together, breaking their words, to Narrowest at the least: every text stays in sight. Narrower still, they
    // stand at that and the list is wider than its room.
    // ###########################################################################################
    public static class PinTableGeometry
    {
        public const double Gap = 20;

        // The least a wrapping column narrows to once its words break - two or three characters.
        public const double Narrowest = 24;

        // ###########################################################################################
        // The pieces a text of the list wraps between, so the widest of them is its column's widest word: split at spaces,
        // and after a "-" or a "/" - where the text's own line breaking breaks it too, a net "Net-(U16-Pad10)" standing as
        // "Net-(U16-" over "Pad10)", a signal "SYNC/LUM" as "SYNC/" over "LUM". Seen as one word, such a net held its
        // column wide and the reading's heading broke "oscilloscope" instead.
        // ###########################################################################################
        public static IEnumerable<string> PiecesOf(string? text)
        {
            int start = 0;
            string line = text ?? string.Empty;

            for (int index = 0; index < line.Length; index++)
            {
                char character = line[index];

                if (char.IsWhiteSpace(character) || character is '-' or '/')
                {
                    int end = char.IsWhiteSpace(character) ? index : index + 1;

                    if (end > start)
                    {
                        yield return line[start..end];
                    }

                    start = index + 1;
                }
            }

            if (start < line.Length)
            {
                yield return line[start..];
            }
        }

        // The list's width with every text whole: the columns shown, and a gap between each two.
        public static double WidthOf(IReadOnlyList<PinTableColumn> columns)
        {
            ArgumentNullException.ThrowIfNull(columns);

            List<PinTableColumn> shown = columns.Where(column => column.IsShown).ToList();

            return shown.Count == 0 ? 0 : shown.Sum(column => column.Widest) + ((shown.Count - 1) * PinTableGeometry.Gap);
        }

        // ###########################################################################################
        // How wide each column's texts may be in `room` - its widest while the list fits, a wrapping column narrower when it
        // does not - in the columns' order, 0 for a column not shown. A room not known yet (0, or no limit) is the widest.
        // ###########################################################################################
        public static IReadOnlyList<double> TextWidths(IReadOnlyList<PinTableColumn> columns, double room)
        {
            ArgumentNullException.ThrowIfNull(columns);

            double[] widths = columns.Select(column => column.IsShown ? column.Widest : 0).ToArray();

            if (room <= 0 || double.IsInfinity(room) || PinTableGeometry.WidthOf(columns) <= room)
            {
                return widths;
            }

            List<int> wrapping = Enumerable.Range(0, columns.Count).Where(index => columns[index].IsShown && columns[index].Wraps).ToList();

            // What the wrapping columns have between them: the room less the whole columns and the gaps. Each narrows
            // between its widest word and its widest text - or, when not even the words fit, between Narrowest and its
            // widest word.
            double budget = room - (PinTableGeometry.WidthOf(columns) - wrapping.Sum(index => columns[index].Widest));
            List<(double Least, double Most)> spans = wrapping
                .Select(index => (Least: Math.Min(columns[index].WidestWord, columns[index].Widest), Most: columns[index].Widest))
                .ToList();

            if (spans.Sum(span => span.Least) > budget)
            {
                spans = spans.Select(span => (Math.Min(span.Least, PinTableGeometry.Narrowest), span.Least)).ToList();
            }

            double level = PinTableGeometry.LevelFor(spans, budget);

            for (int index = 0; index < wrapping.Count; index++)
            {
                widths[wrapping[index]] = PinTableGeometry.At(spans[index], level);
            }

            return widths;
        }

        // A wrapping column at `level`, within its span.
        private static double At((double Least, double Most) span, double level) => Math.Max(span.Least, Math.Min(span.Most, level));

        // ###########################################################################################
        // The one width the wrapping columns come down to so that together they take `budget` - whole pixels, rounded down
        // so they never take more; 0 - each at its least - when not even that fits. What they take grows with the level, in
        // straight pieces between the ends of their spans, so the piece holding the budget gives the level.
        // ###########################################################################################
        private static double LevelFor(IReadOnlyList<(double Least, double Most)> spans, double budget)
        {
            double Taken(double level) => spans.Sum(span => PinTableGeometry.At(span, level));

            List<double> points = spans
                .SelectMany(span => new[] { span.Least, span.Most })
                .Distinct()
                .Order()
                .ToList();

            double before = 0;

            foreach (double point in points)
            {
                if (Taken(point) >= budget)
                {
                    // In this piece every column whose span starts at or under where the piece starts, and ends at or
                    // over where it ends, widens with the level.
                    int widening = spans.Count(span => span.Least <= before && span.Most >= point);

                    return widening == 0 ? before : Math.Floor(before + ((budget - Taken(before)) / widening));
                }

                before = point;
            }

            return before;
        }
    }
}
