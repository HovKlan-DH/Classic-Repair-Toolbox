using System;
using System.Globalization;

namespace CRT.Review.Handlers
{
    // ###########################################################################################
    // How a queue row reads on screen (NewContributeStrategy.md Phase 5, task 2).
    //
    // *** WHY A QUEUE ROW CARRIES NO CHANGE SUMMARY. *** The window's first design gave each
    // queued item a ReviewChangeSummary, which cannot work and is worth recording rather than
    // quietly fixing: computing that summary needs BOTH the published board and the submitted
    // manifest, and the queue endpoint returns neither - it answers ids, system names and
    // summaries, which is kilobytes. Making it return enough to summarise every row would mean
    // loading two full BoardData per queued submission on the server, for a list the reviewer
    // scrolls past. So the summary belongs to OPENING a submission, and the queue shows what the
    // contributor themselves said about it.
    //
    // That is also the better reviewing experience, not merely the cheaper one: the contributor's
    // own sentence ("Corrected R12.") is what a reviewer scans for, and a row reading "3
    // components changed" tells them nothing about whether it is worth opening next.
    //
    // Pure, so the wording is tested rather than eyeballed.
    // ###########################################################################################
    public static class ReviewQueueDisplay
    {
        // ###########################################################################################
        // The line naming the submission. The SYSTEM leads, because a reviewer works by system -
        // several submissions against one board are reviewed together, and the id alone means
        // nothing to a human.
        // ###########################################################################################
        public static string Title(ReviewQueueRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            string system = string.IsNullOrWhiteSpace(row.SystemId) ? "(unknown system)" : row.SystemId;

            return $"#{row.Id} {system}";
        }

        // ###########################################################################################
        // The line under it: what the contributor said, and how long it has been waiting.
        //
        // *** A SUBMISSION WITH NO SUMMARY SAYS SO. *** Blank is a real possibility - the field is
        // optional - and an empty second line reads as a rendering fault rather than as missing
        // information.
        // ###########################################################################################
        public static string Subtitle(ReviewQueueRow row, DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(row);

            string summary = string.IsNullOrWhiteSpace(row.Summary)
                ? "(no description given)"
                : row.Summary.Trim();

            string waiting = ReviewQueueDisplay.Waiting(row.CreatedUtc, now);

            return string.IsNullOrEmpty(waiting) ? summary : $"{summary}  -  {waiting}";
        }

        // ###########################################################################################
        // How long this has been waiting, in the units a person actually thinks in.
        //
        // *** THIS IS THE NUMBER THAT SHAMES A BACKLOG, so it must not round in the comforting
        // direction. *** Truncating rather than rounding means a submission waiting 13 days reads
        // as "13 days", never "2 weeks", and 29 days never becomes "a month" - the coarser unit
        // always makes a wait sound shorter than it is.
        //
        // An absent timestamp yields an EMPTY string rather than a guess. Rendering a missing date
        // as "waiting 56 years" (the 1970 default) would be worse than saying nothing.
        // ###########################################################################################
        public static string Waiting(DateTimeOffset? createdUtc, DateTimeOffset now)
        {
            if (createdUtc is null)
                return string.Empty;

            TimeSpan waited = now - createdUtc.Value;

            // A clock difference between client and server can produce a small negative. Reading
            // "waiting -3 minutes" looks broken; "just now" is true enough and is what the
            // reviewer would conclude anyway.
            if (waited < TimeSpan.Zero)
                return "just now";

            if (waited.TotalMinutes < 1)
                return "just now";

            if (waited.TotalHours < 1)
                return ReviewQueueDisplay.Plural((int)waited.TotalMinutes, "minute");

            if (waited.TotalDays < 1)
                return ReviewQueueDisplay.Plural((int)waited.TotalHours, "hour");

            return ReviewQueueDisplay.Plural((int)waited.TotalDays, "day");
        }

        private static string Plural(int count, string unit) =>
            count == 1
                ? $"waiting 1 {unit}"
                : $"waiting {count.ToString(CultureInfo.InvariantCulture)} {unit}s";
    }
}
