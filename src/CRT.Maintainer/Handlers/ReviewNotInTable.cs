using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace CRT.Maintainer.Handlers
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
        // At most this many rows are named per kind of change; the rest are counted.
        public const int MaximumNamedRows = 6;

        private static readonly HashSet<string> TableSheets =
            BoardWorkbookSchema.AllSheets.Select(sheet => sheet.SheetName).ToHashSet(StringComparer.Ordinal);

        public static IReadOnlyList<ReviewNoteLine> Lines(
            ReviewChangeSummaryView? changes,
            IReadOnlyList<ReviewFindingView> findings,

            // The submission's files, as the server states them - where the KiCad data line comes
            // from (owner request, 2026-09-26: "some data is probably fine not to visualize
            // directly - e.g. the component highlights and KiCad data, but I need information about
            // it, and if it is included"). Optional so older callers and tests stand.
            IReadOnlyList<SubmittedFileFact>? submittedFiles = null)
        {
            ArgumentNullException.ThrowIfNull(findings);

            var lines = new List<ReviewNoteLine>();

            if (changes is null)
            {
                // The server could not load the submission to compare it; the table cannot open it
                // either, and saying so beats an empty panel.
                lines.Add(new ReviewNoteLine("The changes in this submission could not be compared.", ReviewNoteKind.Error));
            }
            else if (changes.IsNewSystem)
            {
                foreach (ReviewSectionView section in changes.Sections)
                {
                    if (section.Added.Count > 0 && !ReviewNotInTable.TableSheets.Contains(section.Section))
                        lines.Add(ReviewNoteLine.FromRuns(ReviewNotInTable.DescribeNewSystemSection(section), ReviewNoteKind.Change));
                }
            }
            else
            {
                foreach (ReviewSummaryLine line in ReviewSummaryPresenter.BuildLines(changes))
                {
                    if (line.Parts.Count == 0 || ReviewNotInTable.TableSheets.Contains(line.Section))
                        continue;

                    var runs = new List<ReviewNoteRun> { new($"{line.Section}: ", IsCount: false) };

                    for (int i = 0; i < line.Parts.Count; i++)
                    {
                        if (i > 0)
                            runs.Add(new ReviewNoteRun(", ", IsCount: false));

                        runs.AddRange(ReviewNotInTable.DescribePart(line.Parts[i]));
                    }

                    lines.Add(ReviewNoteLine.FromRuns(runs, ReviewNoteKind.Change));
                }
            }

            if (ReviewNotInTable.FilesLine(changes, submittedFiles) is ReviewNoteLine files)
                lines.Add(files);

            if (ReviewNotInTable.KiCadLine(changes, submittedFiles) is ReviewNoteLine kiCad)
                lines.Add(kiCad);

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

        // "[2] changed (Schematic 1 / U8, Schematic 1 / U9)" - the rows readable, not as raw keys.
        private static IEnumerable<ReviewNoteRun> DescribePart(ReviewSummaryPart part)
        {
            List<string> named = part.Keys
                .Take(ReviewNotInTable.MaximumNamedRows)
                .Select(ReviewScopeBaseline.DescribeRowKey)
                .ToList();

            int more = part.Keys.Count - named.Count;

            string rows = more > 0
                ? $"{string.Join(", ", named)} and {more} more"
                : string.Join(", ", named);

            string kind = part.Kind.ToString().ToLowerInvariant();

            return ReviewNotInTable.Counted(string.Empty, part.Count, rows.Length == 0 ? $" {kind}" : $" {kind} ({rows})");
        }

        // ###########################################################################################
        // *** A NEW SYSTEM'S LINE SAYS HOW MUCH THERE IS, NOT WHICH (owner request, 2026-09-26). ***
        // Everything of a new system is added, so "4 added (1N4148 / A1, ...)" listed the whole
        // board's highlights one by one - "it should probably only flag that it has highlights for
        // X components - not WHICH components". So: how many components have highlights, and how
        // many schematics have calibration points - "Highlights included for [2] components", in the
        // owner's words: they come with the submission.
        //
        // A highlight is keyed schematic + label, and one component is often highlighted on more
        // than one schematic, so the components are counted by LABEL (case-insensitive, as keys
        // are). A calibration is keyed by its schematic alone.
        // ###########################################################################################
        private static IReadOnlyList<ReviewNoteRun> DescribeNewSystemSection(ReviewSectionView section)
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

        // ###########################################################################################
        // *** WHICH FILES THE APPROVAL WOULD WRITE (code review, 2026-09-26). ***
        //
        // The table shows a file cell's TEXT, so a file replaced under its own unchanged path
        // colours nothing: the row reads identically before and after. The hover card compares the
        // bytes, but only for a picture and only while the pointer is on that one cell - so a
        // replaced PDF, text file or scan was approved with nothing on screen ever saying a file
        // changed. That blind spot is what the retired change summary's file list used to close
        // (security review, 2025-09-25), and dropping it took the answer away while the facts
        // stayed: SubmittedFileFact carries the hash published at each path now.
        //
        // KiCad files are left out - they have their own line below, which counts them the same way.
        //
        // "Files: [12] included ([2] new, [1] replaced)" - a submission that changes no file says
        // nothing, because every file it carries is already published exactly as it is.
        // ###########################################################################################
        private static ReviewNoteLine? FilesLine(ReviewChangeSummaryView? changes, IReadOnlyList<SubmittedFileFact>? submittedFiles)
        {
            List<SubmittedFileFact> files = (submittedFiles ?? [])
                .Where(fact => !SubmissionKiCadFiles.IsKiCadDataPath(fact.Path))
                .ToList();

            if (files.Count == 0)
                return null;

            int added = files.Count(fact => !fact.ExistsOnServer);
            int replaced = files.Count(fact => fact.ExistsOnServer && !fact.IsUnchanged);

            // A new system: every file is new, so "new" says nothing the badge does not.
            if (changes is { IsNewSystem: true })
                return ReviewNoteLine.FromRuns(ReviewNotInTable.Counted("Files included: ", files.Count, files.Count == 1 ? " file" : " files"), ReviewNoteKind.Change);

            if (added == 0 && replaced == 0)
                return null;

            var runs = new List<ReviewNoteRun>();
            runs.AddRange(ReviewNotInTable.Counted("Files: ", files.Count, " included ("));

            if (added > 0)
                runs.AddRange(ReviewNotInTable.Counted(string.Empty, added, " new"));

            if (replaced > 0)
            {
                if (added > 0)
                    runs.Add(new ReviewNoteRun(", ", IsCount: false));

                // "replaced", not "changed": the path is the same and the table shows nothing, which
                // is exactly what a maintainer has to be told about this one.
                runs.AddRange(ReviewNotInTable.Counted(string.Empty, replaced, " replaced under the same name"));
            }

            runs.Add(new ReviewNoteRun(")", IsCount: false));

            return ReviewNoteLine.FromRuns(runs, ReviewNoteKind.Change);
        }

        // ###########################################################################################
        // *** THE SUBMISSION'S KiCad DATA, COUNTED (owner request, 2026-09-26). *** KiCad files
        // travel in the submission since the same day, cited by no row - so no sheet shows them, and
        // without this line they would be approved and published unseen, the same failure the
        // highlights line exists for. The facts are the server's SubmittedFileFacts: which files it
        // holds at a "KiCad data" path, and the hash of what is published there now.
        //
        //   - a new system: "KiCad data included: [14] files" - everything is new, so new/changed
        //     would say nothing (the same reasoning as DescribeNewSystemSection);
        //   - a published board: "... ([2] new, [1] changed)", or "(unchanged)" when the folder
        //     travels untouched;
        //   - no KiCad files: no line. The published folder, if any, is left as it is by a publish,
        //     so absence changes nothing worth a sentence.
        // ###########################################################################################
        private static ReviewNoteLine? KiCadLine(ReviewChangeSummaryView? changes, IReadOnlyList<SubmittedFileFact>? submittedFiles)
        {
            List<SubmittedFileFact> kiCad = (submittedFiles ?? [])
                .Where(fact => SubmissionKiCadFiles.IsKiCadDataPath(fact.Path))
                .ToList();

            if (kiCad.Count == 0)
                return null;

            var runs = new List<ReviewNoteRun>();
            runs.AddRange(ReviewNotInTable.Counted("KiCad data included: ", kiCad.Count, kiCad.Count == 1 ? " file" : " files"));

            if (changes is not { IsNewSystem: true })
            {
                int added = kiCad.Count(fact => !fact.ExistsOnServer);
                int replaced = kiCad.Count(fact => fact.ExistsOnServer && !fact.IsUnchanged);

                if (added == 0 && replaced == 0)
                {
                    runs.Add(new ReviewNoteRun(" (unchanged)", IsCount: false));
                }
                else
                {
                    runs.Add(new ReviewNoteRun(" (", IsCount: false));

                    if (added > 0)
                        runs.AddRange(ReviewNotInTable.Counted(string.Empty, added, " new"));

                    if (replaced > 0)
                    {
                        if (added > 0)
                            runs.Add(new ReviewNoteRun(", ", IsCount: false));

                        runs.AddRange(ReviewNotInTable.Counted(string.Empty, replaced, " changed"));
                    }

                    runs.Add(new ReviewNoteRun(")", IsCount: false));
                }
            }

            return ReviewNoteLine.FromRuns(runs, ReviewNoteKind.Change);
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
    // be shown bold. A line made from text alone is one plain piece.
    //
    // Equal by text, kind AND pieces - the pieces are a list, which a record would compare by
    // reference.
    // ###########################################################################################
    public sealed record ReviewNoteLine(string Text, ReviewNoteKind Kind)
    {
        public IReadOnlyList<ReviewNoteRun> Runs { get; private init; } = [new ReviewNoteRun(Text, IsCount: false)];

        public static ReviewNoteLine FromRuns(IReadOnlyList<ReviewNoteRun> runs, ReviewNoteKind kind)
        {
            ArgumentNullException.ThrowIfNull(runs);

            return new ReviewNoteLine(string.Concat(runs.Select(run => run.Text)), kind) { Runs = runs };
        }

        public bool Equals(ReviewNoteLine? other) =>
            other is not null &&
            this.Text == other.Text &&
            this.Kind == other.Kind &&
            this.Runs.SequenceEqual(other.Runs);

        public override int GetHashCode() => HashCode.Combine(this.Text, this.Kind);
    }

    public sealed record ReviewNoteRun(string Text, bool IsCount);
}
