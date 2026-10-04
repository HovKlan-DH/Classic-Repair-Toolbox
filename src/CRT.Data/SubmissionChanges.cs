using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A SUBMISSION CHANGED, AS IT WENT INTO BETA (owner request, 2026-10-04: "Is it possible
    // to summarize each submission change in a textual form ... so it is possible to see what has
    // changed over time? Not necessarily the full details, but at least to get an idea, besides the
    // sometimes vague description from the contributor").
    //
    // *** RECORDED AT THE PUBLISH, BECAUSE THAT IS THE ONLY MOMENT IT CAN BE KNOWN. *** No history of
    // the boards is kept - a publish overwrites BETA in place - so "the board before" exists only
    // while the publish runs. ApprovePublishFlow compares the two there (ReviewSummary, the review's
    // own row comparison) and stores these facts with the submission; the Systems screen's History
    // view reads them back. A submission published before this existed has none, and nothing can
    // make one for it now.
    //
    // FACTS, NOT WORDS - the split every answer on the Systems screen keeps: the server says which
    // rows and fields moved, and the Maintainer tab puts that into sentences (SystemHistoryDisplay).
    //
    // *** BOUNDED. *** Each list keeps at most ListedPerKind entries, and the counts beside it say how
    // many there really were - a new system adds hundreds of rows, and the history is read every
    // time a system is opened. "Not necessarily the full details" is the owner's own brief.
    //
    // Rows are named the way the Drafts tab and the review name them (BoardDraftNaturalKeys.Describe,
    // "U8 / PAL"), sections by their sheet names, and fields by the workbook's column headers, so the
    // history, the table and the workbook all speak of the same things in the same words.
    // ###########################################################################################
    public sealed record SubmissionChanges(
        bool IsNewSystem,
        IReadOnlyList<SectionChanges> Sections,
        FileChanges Files);

    // One sheet's changes. The counts are the whole truth; the lists are the first ListedPerKind.
    public sealed record SectionChanges(
        string Section,
        int AddedCount,
        int ChangedCount,
        int RemovedCount,
        int RenamedCount,
        IReadOnlyList<string> Added,
        IReadOnlyList<ChangedRowFact> Changed,
        IReadOnlyList<string> Removed,
        IReadOnlyList<RenamedRowFact> Renamed);

    // A row that stayed and changed: which row, and which of its columns.
    public sealed record ChangedRowFact(string Row, IReadOnlyList<string> Fields);

    // A row that only changed what identifies it ("C10" became "C51").
    public sealed record RenamedRowFact(string From, string To);

    // ###########################################################################################
    // The files the publish wrote and removed, by their paths from the data root. Added: nothing
    // was at that path before. Replaced: other bytes were. Removed: the board stopped citing it and
    // nothing else used it. The board's own workbook and highlight file are not counted - every
    // publish writes those.
    // ###########################################################################################
    public sealed record FileChanges(
        int AddedCount,
        int ReplacedCount,
        int RemovedCount,
        IReadOnlyList<string> Added,
        IReadOnlyList<string> Replaced,
        IReadOnlyList<string> Removed)
    {
        public static readonly FileChanges None = new(0, 0, 0, [], [], []);

        public bool HasChanges => this.AddedCount + this.ReplacedCount + this.RemovedCount > 0;
    }

    public static class SubmissionChangeFacts
    {
        // How many rows or files of one kind a summary names; the counts carry the rest.
        public const int ListedPerKind = 20;

        // ###########################################################################################
        // The facts of one publish, from the review's own comparison of the board before and after
        // (ReviewSummary.Compare) and the files the publish wrote and removed. Only the sections
        // that changed are kept; the revision date is left out, since the server stamps a new one on
        // every publish and saying so would only repeat the date of the line itself.
        // ###########################################################################################
        public static SubmissionChanges Build(
            ReviewChangeSummary rows,
            IEnumerable<string>? filesAdded = null,
            IEnumerable<string>? filesReplaced = null,
            IEnumerable<string>? filesRemoved = null)
        {
            ArgumentNullException.ThrowIfNull(rows);

            List<SectionChanges> sections = rows.ChangedSections
                .Select(section => new SectionChanges(
                    section.Section,
                    section.Added.Count,
                    section.Changed.Count,
                    section.Removed.Count,
                    section.Renamed.Count,
                    SubmissionChangeFacts.First(section.Added.Select(BoardDraftNaturalKeys.Describe)),
                    section.Changed
                        .Take(SubmissionChangeFacts.ListedPerKind)
                        .Select(key => new ChangedRowFact(
                            BoardDraftNaturalKeys.Describe(key),
                            section.FieldChanges.TryGetValue(key, out IReadOnlyList<ReviewFieldChange>? fields)
                                ? fields.Select(field => field.Field).Distinct(StringComparer.Ordinal).ToList()
                                : []))
                        .ToList(),
                    SubmissionChangeFacts.First(section.Removed.Select(BoardDraftNaturalKeys.Describe)),
                    section.Renamed
                        .Take(SubmissionChangeFacts.ListedPerKind)
                        .Select(rename => new RenamedRowFact(BoardDraftNaturalKeys.Describe(rename.From), BoardDraftNaturalKeys.Describe(rename.To)))
                        .ToList()))
                .ToList();

            return new SubmissionChanges(
                rows.IsNewSystem,
                sections,
                SubmissionChangeFacts.Files(filesAdded, filesReplaced, filesRemoved));
        }

        // The file facts alone - each list sorted and without repeats, then cut to ListedPerKind.
        public static FileChanges Files(IEnumerable<string>? added, IEnumerable<string>? replaced, IEnumerable<string>? removed)
        {
            List<string> Clean(IEnumerable<string>? paths) =>
                (paths ?? [])
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .Select(path => path.Trim().Replace('\\', '/'))
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .ToList();

            List<string> a = Clean(added);
            List<string> r = Clean(replaced);
            List<string> x = Clean(removed);

            return new FileChanges(
                a.Count,
                r.Count,
                x.Count,
                a.Take(SubmissionChangeFacts.ListedPerKind).ToList(),
                r.Take(SubmissionChangeFacts.ListedPerKind).ToList(),
                x.Take(SubmissionChangeFacts.ListedPerKind).ToList());
        }

        // ###########################################################################################
        // The same facts with the removed files filled in - the publish knows them only after it
        // has removed them, once the summary of the rows was taken.
        // ###########################################################################################
        public static SubmissionChanges WithRemovedFiles(SubmissionChanges changes, IEnumerable<string>? removed)
        {
            ArgumentNullException.ThrowIfNull(changes);

            FileChanges files = changes.Files ?? FileChanges.None;

            return changes with
            {
                Files = SubmissionChangeFacts.Files(files.Added, files.Replaced, removed) with
                {
                    // The lists were already cut - keep the real counts of the two the publish
                    // knew before writing.
                    AddedCount = files.AddedCount,
                    ReplacedCount = files.ReplacedCount
                }
            };
        }

        // Whether there is anything to say at all.
        public static bool HasChanges(SubmissionChanges changes)
        {
            ArgumentNullException.ThrowIfNull(changes);

            return changes.IsNewSystem ||
                   (changes.Sections ?? []).Count > 0 ||
                   (changes.Files ?? FileChanges.None).HasChanges;
        }

        // Which of the plan's files are new at their path and which replace what was there - asked
        // of the tree BEFORE the publish writes them.
        public static (IReadOnlyList<string> Added, IReadOnlyList<string> Replaced) SplitWrites(
            IEnumerable<(string RelativePath, string AbsolutePath)> written)
        {
            var added = new List<string>();
            var replaced = new List<string>();

            foreach ((string relative, string absolute) in written ?? [])
            {
                (File.Exists(absolute) ? replaced : added).Add(relative);
            }

            return (added, replaced);
        }

        private static List<string> First(IEnumerable<string> items) =>
            items.Take(SubmissionChangeFacts.ListedPerKind).ToList();
    }
}
