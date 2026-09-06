using System;
using System.Collections.Generic;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Turns the decimal hours typed into the "Work done" dialog into the same span said in hours
    // and minutes - "1,25" reading back as "1 hour and 15 minutes".
    //
    // The field is decimal because that is what gets STORED and totalled (WorkbookSummary sums it,
    // the PDF prints it), but decimal hours is not how anyone thinks about time at a bench: 1.25 is
    // an hour and a quarter, 0.75 is three quarters, and 1.4 is an hour and twenty-four minutes,
    // which is the one people get wrong by reading it as an hour and forty. So the number is echoed
    // back in words as it is typed, and a mistyped 1.4-for-1.40 shows itself immediately rather
    // than at invoice time.
    //
    // Returned as PARTS rather than as a finished string, for the same reason WorkbookSummary hands
    // back Stat parts: the numbers are bold and the words are not, and a TextBlock cannot mix
    // weights within one Text. Re-finding the digits in a formatted string would have to guess.
    //
    // Pure string/maths work - no controls, no files - so it is unit tested like the rest of
    // Handlers/.
    // ###########################################################################################
    public static class WorklogDurationFormatter
    {
        // One run of the readback: a number to be drawn bold, and the words that follow it. The
        // separator between two parts (" and ") is the caller's, so the parts stay just the pairs.
        public readonly record struct DurationPart(string Number, string Words);

        // Minutes rather than seconds is the finest unit this is ever asked for - the field's own
        // Increment is 0.25 h - so anything under half a minute rounds away rather than producing a
        // "0 minutes" tail nobody typed.
        private const double MinutesPerHour = 60.0;

        // The largest duration this will render - see BuildParts for why an upper bound is needed at
        // all. A century of hours: past anything a repair could plausibly record, and far enough
        // below long.MaxValue minutes that the conversion cannot overflow.
        public const double MaximumHours = 876_000.0;

        // ###########################################################################################
        // The parts for one duration in decimal hours, e.g. 1.25 -> [("1", " hour and "), ("15",
        // " minutes")]. Empty for zero, for anything negative, and for a value that rounds to no
        // whole minutes at all - there is nothing useful to say about "0 hours and 0 minutes", and a
        // hint line under an untouched field is noise rather than help.
        // ###########################################################################################
        public static IReadOnlyList<DurationPart> BuildParts(double hours)
        {
            var parts = new List<DurationPart>();

            // NaN and infinity cannot come from the NumericUpDown, but they CAN come from a
            // hand-edited entries.json, and rounding either of them throws.
            //
            // The UPPER bound is here for the same reason and is not merely tidiness: the next line
            // casts hours*60 to long, and a cast that overflows is UNCHECKED in C# - it does not
            // throw, it produces an implementation-defined value. So an implausible figure came back
            // either as a negative total (the readback then silently shows NOTHING for a number the
            // user had just typed) or as a saturated long rendering as "153722867280912930 hours and
            // 55 minutes", and the same value flowed on into the summary strip and the exported PDF.
            // The control has its own Maximum now; this guard is what covers the JSON, which no
            // control has ever been between.
            //
            // The limit is generous on purpose - a century of hours is far past any real repair, so
            // nothing a user could legitimately log is rejected by it.
            if (double.IsNaN(hours) || double.IsInfinity(hours) || hours <= 0.0 || hours > MaximumHours)
            {
                return parts;
            }

            // Round to whole minutes ONCE, then split - rounding the two halves separately lets
            // 1.9999 h become "1 hour and 60 minutes".
            long totalMinutes = (long)Math.Round(hours * MinutesPerHour, MidpointRounding.AwayFromZero);

            if (totalMinutes <= 0)
            {
                return parts;
            }

            long wholeHours = totalMinutes / 60;
            long minutes = totalMinutes % 60;

            if (wholeHours > 0)
            {
                string words = wholeHours == 1 ? " hour" : " hours";

                // The joining word belongs to the FIRST part, so the caller can add the parts
                // straight through without knowing whether a second one is coming.
                if (minutes > 0)
                {
                    words += " and ";
                }

                parts.Add(new DurationPart(wholeHours.ToString(CultureInfo.InvariantCulture), words));
            }

            if (minutes > 0)
            {
                parts.Add(new DurationPart(
                    minutes.ToString(CultureInfo.InvariantCulture),
                    minutes == 1 ? " minute" : " minutes"));
            }

            return parts;
        }

        // ###########################################################################################
        // The same duration as WorkbookSummary.Stat parts, for the surfaces that render a LINE of
        // stats - the Workbooks summary strip and the exported PDF both walk a Stat list and bold
        // each Number.
        //
        // Returns TWO stats for a duration carrying both halves, the second marked
        // JoinedToPrevious so the caller's " - " separator does not land between "1 hour and" and
        // "15 minutes" and split one figure into two.
        //
        // Returns NOTHING at all for zero, rather than a "0 minutes" stat: a headline reading
        // "6 worklogs - 0 minutes - 175 CHF" spends a column on the absence of a figure. The
        // surfaces that show it either omit the item (the summary strip, the PDF) or hide their
        // whole row, all of which they already do for other empty values.
        // ###########################################################################################
        public static IReadOnlyList<WorkbookSummary.Stat> BuildStats(double hours)
        {
            var parts = BuildParts(hours);
            var stats = new List<WorkbookSummary.Stat>(parts.Count);

            for (int i = 0; i < parts.Count; i++)
            {
                stats.Add(new WorkbookSummary.Stat(
                    string.Empty, parts[i].Number, parts[i].Words, JoinedToPrevious: i > 0));
            }

            return stats;
        }

        // ###########################################################################################
        // The same readback as one plain string, for anywhere the bold/plain split is not available
        // (a tooltip, a log line, a test's assertion). Empty string where BuildParts is empty.
        // ###########################################################################################
        public static string Format(double hours)
        {
            var parts = BuildParts(hours);

            if (parts.Count == 0)
            {
                return string.Empty;
            }

            var text = new System.Text.StringBuilder();

            foreach (var part in parts)
            {
                text.Append(part.Number).Append(part.Words);
            }

            return text.ToString();
        }
    }
}
