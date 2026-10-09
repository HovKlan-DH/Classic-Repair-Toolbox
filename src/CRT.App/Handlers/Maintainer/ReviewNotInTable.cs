using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHAT A SUBMISSION CHANGES THAT ITS TABLE CANNOT SHOW (2026-09-26).
    //
    // The table became the submission panel's only view (owner request: the change summary "is
    // confusing to look at, so just scrap that information"). The table is the nine workbook
    // sheets, and a submission can change three things besides:
    //
    //   - component HIGHLIGHTS (the rectangles on the schematics, in the .json beside the workbook),
    //   - KiCad CALIBRATION points (the same .json),
    //   - and it can carry the automatic checks' WARNINGS.
    //
    // All three are published by an approval. Dropped with the summary, they would be approved
    // unseen - the very thing ReviewSummary's calibration section was added to stop. So they are
    // kept as a few short lines above the table, and only when there is something to say.
    //
    // *** NO FILE LINES ANY MORE (owner request, 2026-09-30: "Now we have the "Files" button ...
    // I do not think we any longer need to see the files in the yellow highlighted area"). ***
    // "Files: [12] included ([2] new, [1] replaced under the same name)" and "KiCad data included:
    // [3] files" stood here, for files no cell shows as changed. The submission's FILES view says
    // all of it and more - every file new, changed or removed, KiCad data included - and its
    // button carries the count (ReviewFilesCount), so a file replaced under its own name is still
    // said before anything is opened. That count is what those two lines' tests now guard.
    //
    // *** NO "(not in the table)" (owner request, 2026-09-26). *** The lines said so at first, and
    // "Highlights for 2 components (not in the table)" read as if the COMPONENTS were missing from
    // the table - they are on its Components sheet. Where the lines stand says it well enough.
    //
    // *** EVERY COUNT IS "[2]", THE NUMBER BOLD (owner request, 2026-09-26). *** Which is why a line
    // is handed over in PIECES (ReviewNoteRun) rather than as finished text: a TextBlock cannot mix
    // weights within one Text, and finding the digits again in a finished string would have to
    // guess - the same reason WorkbookSummary hands back its Stat parts.
    //
    // Pure, so which lines appear, and how they read, is tested.
    // ###########################################################################################
    public static class ReviewNotInTable
    {
        // At most this many changes are named per section, one line each; the rest are counted.
        public const int MaximumNamedRows = 6;

        private static readonly HashSet<string> TableSheets =
            BoardWorkbookSchema.AllSheets.Select(sheet => sheet.SheetName).ToHashSet(StringComparer.Ordinal);

        public static IReadOnlyList<ReviewNoteLine> Lines(
            ReviewChangeSummaryView? changes,
            IReadOnlyList<ReviewFindingView> findings)
        {
            ArgumentNullException.ThrowIfNull(findings);

            var lines = new List<ReviewNoteLine>();

            if (changes is null)
            {
                // The server could not load the submission to compare it; the table cannot open it
                // either, and saying so beats an empty panel.
                lines.Add(new ReviewNoteLine("The changes in this submission could not be compared.", ReviewNoteKind.Error));
            }
            else if (changes.IsNewBoard)
            {
                foreach (ReviewSectionView section in changes.Sections)
                {
                    if (section.Added.Count > 0 && !ReviewNotInTable.TableSheets.Contains(section.Section))
                        lines.Add(ReviewNoteLine.FromRuns(ReviewNotInTable.DescribeNewBoardSection(section), ReviewNoteKind.Change));
                }
            }
            else
            {
                foreach (ReviewSectionView section in changes.Sections)
                {
                    if (section.TotalChanges > 0 && !ReviewNotInTable.TableSheets.Contains(section.Section))
                        lines.AddRange(ReviewNotInTable.DescribeSection(section));
                }
            }

            foreach (ReviewFindingView finding in findings)
            {
                // The message is written for the contributor and reads for a maintainer too; the
                // subject is added only when there is one.
                string message = string.IsNullOrWhiteSpace(finding.Subject)
                    ? finding.Message
                    : $"{finding.Message} [{finding.Subject}]";

                lines.Add(finding.IsError
                    ? new ReviewNoteLine($"Error from the automatic checks: {message}", ReviewNoteKind.Error)
                    : new ReviewNoteLine($"Warning from the automatic checks: {message}", ReviewNoteKind.Warning));
            }

            return lines;
        }

        // ###########################################################################################
        // *** ONE LINE PER CHANGE, UNDER A COUNTED HEADING (owner request, 2026-10-04). *** It was
        // "Component highlights: [1] removed (Board layout / hest)" - the key's two halves joined by
        // a slash, which a maintainer had to take apart. Now, in the owner's words:
        //
        //     Component highlights have [1] change:
        //       Removed component [hest] from schematic "Board layout"
        //
        // and the KiCad calibration points the same way, being the one other thing the table cannot
        // show. Removals first, as the summary has always put them (ReviewSummaryPresenter.BuildLines
        // says why); at most MaximumNamedRows lines, the rest counted.
        // ###########################################################################################
        private static IEnumerable<ReviewNoteLine> DescribeSection(ReviewSectionView section)
        {
            int count = section.TotalChanges;

            yield return ReviewNoteLine.FromRuns(
                ReviewNotInTable.Counted($"{section.Section} have ", count, count == 1 ? " change:" : " changes:"),
                ReviewNoteKind.Change);

            List<string> described =
            [
                .. section.Removed.Select(key => ReviewNotInTable.DescribeRow(section.Section, ReviewChangeKind.Removed, key, fields: null)),
                .. section.Changed.Select(key => ReviewNotInTable.DescribeRow(
                    section.Section, ReviewChangeKind.Changed, key, section.FieldChanges.TryGetValue(key, out var fields) ? fields : null)),
                .. section.Added.Select(key => ReviewNotInTable.DescribeRow(section.Section, ReviewChangeKind.Added, key, fields: null)),
                .. section.Renamed.Select(rename => ReviewNotInTable.DescribeRename(section.Section, rename))
            ];

            foreach (string line in described.Take(ReviewNotInTable.MaximumNamedRows))
                yield return new ReviewNoteLine(line, ReviewNoteKind.Change) { Indent = 1 };

            int more = described.Count - ReviewNotInTable.MaximumNamedRows;

            if (more > 0)
                yield return ReviewNoteLine.FromRuns(ReviewNotInTable.Counted("and ", more, " more"), ReviewNoteKind.Change, indent: 1);
        }

        // ###########################################################################################
        // One change in words. A highlight's key is its schematic and its component label; a
        // calibration's, its schematic. A key of another shape - a section this does not know -
        // is named as its parts, so nothing is ever left out.
        // ###########################################################################################
        private static string DescribeRow(string section, ReviewChangeKind kind, string key, IReadOnlyList<ReviewFieldChangeView>? fields)
        {
            if (section == ReviewSummary.SectionComponentHighlights && ReviewNotInTable.Highlight(key) is (string schematic, string label))
            {
                return kind switch
                {
                    ReviewChangeKind.Removed => $"Removed component [{label}] from schematic \"{schematic}\"",
                    ReviewChangeKind.Added => $"Added component [{label}] to schematic \"{schematic}\"",
                    _ => $"{ReviewNotInTable.HowMoved(fields)} component [{label}] on schematic \"{schematic}\""
                };
            }

            if (section == ReviewSummary.SectionKiCadCalibrations)
            {
                string calibrated = ReviewScopeBaseline.DescribeRowKey(key);

                return kind switch
                {
                    ReviewChangeKind.Removed => $"Removed the calibration points of schematic \"{calibrated}\"",
                    ReviewChangeKind.Added => $"Added calibration points to schematic \"{calibrated}\"",
                    _ => $"Changed the calibration points of schematic \"{calibrated}\""
                };
            }

            return $"{kind} {ReviewScopeBaseline.DescribeRowKey(key)}";
        }

        // ###########################################################################################
        // A row whose identifying cells changed and nothing else (ReviewSummary pairs them, 2026-10-04)
        // - for a highlight, a component relabelled on its schematic, or its schematic renamed.
        // ###########################################################################################
        private static string DescribeRename(string section, ReviewRenameView rename)
        {
            string also = rename.AlsoChanged ? ", and changed it" : string.Empty;

            if (section == ReviewSummary.SectionComponentHighlights &&
                ReviewNotInTable.Highlight(rename.From) is (string fromSchematic, string fromLabel) &&
                ReviewNotInTable.Highlight(rename.To) is (string toSchematic, string toLabel))
            {
                if (string.Equals(fromSchematic, toSchematic, StringComparison.Ordinal))
                    return $"Renamed component [{fromLabel}] to [{toLabel}] on schematic \"{toSchematic}\"{also}";

                if (string.Equals(fromLabel, toLabel, StringComparison.Ordinal))
                    return $"Moved component [{fromLabel}] from schematic \"{fromSchematic}\" to schematic \"{toSchematic}\"{also}";

                return $"Renamed component [{fromLabel}] on schematic \"{fromSchematic}\" to [{toLabel}] on schematic \"{toSchematic}\"{also}";
            }

            if (section == ReviewSummary.SectionKiCadCalibrations)
            {
                return $"Moved the calibration points of schematic \"{ReviewScopeBaseline.DescribeRowKey(rename.From)}\" " +
                       $"to schematic \"{ReviewScopeBaseline.DescribeRowKey(rename.To)}\"{also}";
            }

            return $"Renamed {ReviewScopeBaseline.DescribeRowKey(rename.From)} to {ReviewScopeBaseline.DescribeRowKey(rename.To)}{also}";
        }

        // A highlight's key as its schematic and component label - null for a key of another shape.
        private static (string Schematic, string Label)? Highlight(string? key)
        {
            string[] parts = (key ?? string.Empty).Split(BoardDraftNaturalKeys.Separator);

            return parts.Length == 2 && parts[1].Trim().Length > 0
                ? (parts[0].Trim(), parts[1].Trim())
                : null;
        }

        // What happened to a highlight that is still there: its rectangle moved (X, Y), was resized
        // (Width, Height), or both - "Changed" when the fields do not say.
        private static string HowMoved(IReadOnlyList<ReviewFieldChangeView>? fields)
        {
            bool moved = fields?.Any(field => field.Field is "X" or "Y") == true;
            bool resized = fields?.Any(field => field.Field is "Width" or "Height") == true;

            return (moved, resized) switch
            {
                (true, true) => "Moved and resized",
                (true, false) => "Moved",
                (false, true) => "Resized",
                _ => "Changed"
            };
        }

        // ###########################################################################################
        // *** A NEW BOARD'S LINE SAYS HOW MUCH THERE IS, NOT WHICH (owner request, 2026-09-26). ***
        // Everything of a new board is added, so "4 added (1N4148 / A1, ...)" listed the whole
        // board's highlights one by one - "it should probably only flag that it has highlights for
        // X components - not WHICH components". So: how many components have highlights, and how
        // many schematics have calibration points - "Highlights included for [2] components", in the
        // owner's words: they come with the submission.
        //
        // A highlight is keyed schematic + label, and one component is often highlighted on more
        // than one schematic, so the components are counted by LABEL (case-insensitive, as keys
        // are). A calibration is keyed by its schematic alone.
        // ###########################################################################################
        private static IReadOnlyList<ReviewNoteRun> DescribeNewBoardSection(ReviewSectionView section)
        {
            if (section.Section == ReviewSummary.SectionComponentHighlights)
            {
                int components = section.Added
                    .Select(key => key.Split(BoardDraftNaturalKeys.Separator).Last())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();

                return ReviewNotInTable.Counted("Highlights included for ", components, components == 1 ? " component" : " components");
            }

            int count = section.Added.Count;

            if (section.Section == ReviewSummary.SectionKiCadCalibrations)
                return ReviewNotInTable.Counted("KiCad calibration points included for ", count, count == 1 ? " schematic" : " schematics");

            return ReviewNotInTable.Counted($"{section.Section}: ", count, count == 1 ? " row" : " rows");
        }

        // "before [N] after", N bold.
        private static IReadOnlyList<ReviewNoteRun> Counted(string before, int count, string after) =>
        [
            new ReviewNoteRun(before + "[", IsCount: false),
            new ReviewNoteRun(count.ToString(System.Globalization.CultureInfo.InvariantCulture), IsCount: true),
            new ReviewNoteRun("]" + after, IsCount: false)
        ];
    }

    public enum ReviewNoteKind
    {
        Change,
        Warning,
        Error
    }

    // ###########################################################################################
    // One line above the table: its whole text, and the same text in pieces, each count marked to
    // be shown bold. A line made from text alone is one plain piece. `Indent` sets a line in under
    // the heading it belongs to - one change of "Component highlights have [2] changes:" (2026-10-04).
    //
    // Equal by text, kind, indent AND pieces - the pieces are a list, which a record would compare
    // by reference.
    // ###########################################################################################
    public sealed record ReviewNoteLine(string Text, ReviewNoteKind Kind)
    {
        public IReadOnlyList<ReviewNoteRun> Runs { get; private init; } = [new ReviewNoteRun(Text, IsCount: false)];

        public int Indent { get; init; }

        public static ReviewNoteLine FromRuns(IReadOnlyList<ReviewNoteRun> runs, ReviewNoteKind kind, int indent = 0)
        {
            ArgumentNullException.ThrowIfNull(runs);

            return new ReviewNoteLine(string.Concat(runs.Select(run => run.Text)), kind) { Runs = runs, Indent = indent };
        }

        public bool Equals(ReviewNoteLine? other) =>
            other is not null &&
            this.Text == other.Text &&
            this.Kind == other.Kind &&
            this.Indent == other.Indent &&
            this.Runs.SequenceEqual(other.Runs);

        public override int GetHashCode() => HashCode.Combine(this.Text, this.Kind, this.Indent);
    }

    public sealed record ReviewNoteRun(string Text, bool IsCount);
}
