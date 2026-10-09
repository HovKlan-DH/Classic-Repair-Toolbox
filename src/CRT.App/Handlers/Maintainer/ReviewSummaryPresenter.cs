using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // Turns a ReviewChangeSummary into the lines the review window draws
    // (NewContributeStrategy.md Phase 5, task 3).
    //
    // PURE, AND IN Handlers/ FOR THE USUAL REASON: this is where the wording and the ordering of
    // the maintainer's landing view are decided, and both are things worth pinning down in tests
    // rather than verifying by eye every time the window changes. The window itself then only
    // walks the result and makes TextBlocks.
    //
    // *** THE WORDING IS NOT COSMETIC HERE. *** A maintainer decides whether to look closer from
    // these lines. "2 changed" and "2 removed" carry very different risk, and a line that buries
    // a removal among additions is how a deletion gets approved unseen - which cannot be undone,
    // since the project owner ruled out retained revisions.
    // ###########################################################################################
    public static class ReviewSummaryPresenter
    {
        // ###########################################################################################
        // The per-section lines, in the order the maintainer reads them.
        //
        // SECTIONS WITH NO CHANGES ARE OMITTED. Showing "0 changed" for ten sections buries the
        // one that matters in nine lines of nothing, which is the same failure as opening on the
        // whole board - just smaller.
        //
        // REMOVALS ARE LISTED FIRST WITHIN A SECTION, deliberately. They are the least
        // recoverable thing a submission can do and the easiest to skim past, because a maintainer
        // scanning for "what did they add" is not looking for them.
        // ###########################################################################################
        public static IReadOnlyList<ReviewSummaryLine> BuildLines(ReviewChangeSummaryView summary)
        {
            ArgumentNullException.ThrowIfNull(summary);

            var lines = new List<ReviewSummaryLine>();

            foreach (ReviewSectionView section in summary.Sections)
            {
                var parts = new List<ReviewSummaryPart>();

                if (section.Removed.Count > 0)
                    parts.Add(new ReviewSummaryPart(ReviewChangeKind.Removed, section.Removed.Count, section.Removed));

                if (section.Changed.Count > 0)
                    parts.Add(new ReviewSummaryPart(ReviewChangeKind.Changed, section.Changed.Count, section.Changed));

                if (section.Added.Count > 0)
                    parts.Add(new ReviewSummaryPart(ReviewChangeKind.Added, section.Added.Count, section.Added));

                if (section.Renamed.Count > 0)
                {
                    parts.Add(new ReviewSummaryPart(
                        ReviewChangeKind.Renamed,
                        section.Renamed.Count,
                        [.. section.Renamed.Select(ReviewSummaryPresenter.DescribeRename)]));
                }

                lines.Add(new ReviewSummaryLine(section.Section, parts));
            }

            return lines;
        }

        // ###########################################################################################
        // A rename as a maintainer should read it.
        //
        // The arrow carries the whole meaning, and "(and edited)" is appended only when something
        // BESIDES the name moved - see ReviewSummary.RowsMatchIgnoringKey. Without that
        // distinction "renamed" invites a maintainer not to look further, which is exactly when an
        // edit hidden behind a rename gets through.
        // ###########################################################################################
        public static string DescribeRename(ReviewRenameView renamed) =>
            renamed.AlsoChanged
                ? $"{renamed.From} -> {renamed.To} (and edited)"
                : $"{renamed.From} -> {renamed.To}";

        // ###########################################################################################
        // A field-level change as a maintainer reads it (Phase 5, task 4).
        //
        // *** A BLANK VALUE IS SPELLED OUT, NEVER LEFT EMPTY. *** "Note: -> Measured cold" reads
        // as a rendering fault; "Note: (blank) -> Measured cold" reads as a fact. The case that
        // matters more is the reverse - a field being CLEARED - because deleting information is
        // the edit hardest to notice, and an empty right-hand side would look like the app simply
        // failed to draw it.
        //
        // *** LONG VALUES ARE TRUNCATED IN THE MIDDLE, not at the end. *** A component
        // description can run to a sentence, and two of them on one line push the actual change
        // off screen. Cutting the middle keeps both ends, which is where a real edit usually
        // shows - a corrected part number differs at its end, a corrected prefix at its start.
        // ###########################################################################################
        public static string DescribeFieldChange(ReviewFieldChangeView change)
        {
            ArgumentNullException.ThrowIfNull(change);

            return $"{change.Field}: {ReviewSummaryPresenter.Value(change.Before)} -> {ReviewSummaryPresenter.Value(change.After)}";
        }

        // ###########################################################################################
        // How an image comparison is captioned (task 4).
        //
        // *** THE WORD MATTERS MORE THAN IT LOOKS. *** A maintainer skims these captions to decide
        // which panels to actually study, so each of the three says plainly what happened rather
        // than leaving it to be inferred from which side is blank. "Removed" in particular must
        // never be softened into "changed": a deletion is the least recoverable thing a submission
        // can do and the easiest to approve without noticing.
        // ###########################################################################################
        public static string DescribeImagePair(ReviewImagePair pair)
        {
            ArgumentNullException.ThrowIfNull(pair);

            string what = pair.Change switch
            {
                ReviewImageChange.Added => "Added",
                ReviewImageChange.Replaced => "Replaced",
                // *** SHOUTED, and truthful. *** The board stops citing the file. Whether the
                // publish then REMOVES it from the server (nothing else uses it) or it stays is
                // said in the file list above, from the server's own list (2026-09-25).
                ReviewImageChange.Removed => "NO LONGER USED",
                _ => "Changed"
            };

            return $"{what}: {pair.Path}";
        }

        // ###########################################################################################
        // What to put where a picture cannot be shown.
        //
        // An added image has no "before" and a removed one has no "after". That side must be
        // CAPTIONED rather than left blank - an empty panel beside a full one reads as an image
        // that failed to load, which sends a maintainer looking for a fault that is not there.
        // ###########################################################################################
        public static string DescribeMissingSide(ReviewImageChange change, bool isBeforeSide)
        {
            if (isBeforeSide)
                return change == ReviewImageChange.Added ? "Not in the published board" : string.Empty;

            return change == ReviewImageChange.Removed
                ? "No longer used by the board"
                : string.Empty;
        }

        // The longest a single value is shown at. Beyond this the middle is elided.
        public const int MaximumValueLength = 60;

        private static string Value(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return "(blank)";

            if (raw.Length <= ReviewSummaryPresenter.MaximumValueLength)
                return raw;

            // Half each side of the ellipsis, minus its own width, so the result is never longer
            // than the limit it is enforcing.
            const string ellipsis = "...";
            int keep = (ReviewSummaryPresenter.MaximumValueLength - ellipsis.Length) / 2;

            return string.Concat(raw.AsSpan(0, keep), ellipsis, raw.AsSpan(raw.Length - keep));
        }

        // ###########################################################################################
        // The headline shown above the breakdown.
        //
        // A NEW BOARD IS CALLED OUT EXPLICITLY, because it is the highest-risk submission there
        // is (Phase 6 task 3): unreviewed content, from someone with no track record, establishing
        // a board nobody else knows. A maintainer must never learn that from the row counts alone.
        // ###########################################################################################
        public static string BuildHeadline(ReviewChangeSummaryView summary)
        {
            ArgumentNullException.ThrowIfNull(summary);

            if (summary.IsNewBoard)
            {
                int rows = summary.Sections.Sum(section => section.Added.Count);

                return $"New board, {rows} {(rows == 1 ? "row" : "rows")}";
            }

            if (!summary.HasChanges)
                return "No changes";

            var parts = new List<string>();

            foreach (ReviewSectionView section in summary.Sections)
            {
                string name = section.Section.ToLowerInvariant();

                if (section.Added.Count > 0)
                    parts.Add($"{section.Added.Count} {name} added");

                if (section.Changed.Count > 0)
                    parts.Add($"{section.Changed.Count} {name} changed");

                if (section.Removed.Count > 0)
                    parts.Add($"{section.Removed.Count} {name} removed");

                if (section.Renamed.Count > 0)
                    parts.Add($"{section.Renamed.Count} {name} renamed");
            }

            return string.Join(", ", parts);
        }
    }

    // ###########################################################################################
    // What kind of change a part describes. An enum rather than a string because the window
    // colours by it, and a typo'd colour key fails silently in Avalonia.
    // ###########################################################################################
    public enum ReviewChangeKind
    {
        Removed,
        Changed,
        Added,
        Renamed
    }

    public sealed record ReviewSummaryPart(ReviewChangeKind Kind, int Count, IReadOnlyList<string> Keys)
    {
        // "2 removed", "1 added". The count leads, matching how the rest of this project draws a
        // counted pill (see WorklogInfoPillBuilder).
        public string Describe() => $"{this.Count} {this.Kind.ToString().ToLowerInvariant()}";
    }

    public sealed record ReviewSummaryLine(string Section, IReadOnlyList<ReviewSummaryPart> Parts);
}
