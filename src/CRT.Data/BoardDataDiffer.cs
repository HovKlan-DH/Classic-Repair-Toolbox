using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What one row became, comparing a drafted board against the published one
    // (NewContributeStrategy.md Phase 6 - drafts stored as real board folders).
    //
    // Deliberately the same three words BoardDraft's own DraftRowState uses (Added / Modified /
    // Deleted), because they describe the same three things to the same reader. What has changed
    // is WHERE the answer comes from: DraftRowState was RECORDED at edit time, and this is
    // DERIVED by comparing two boards.
    //
    // *** THAT DIFFERENCE IS THE WHOLE POINT OF THIS CLASS. *** Once a draft is a real board
    // folder the contributor can edit it in Excel, outside the application entirely - and an
    // edit-time ledger cannot survive that, because nothing tells the app the file changed. A
    // derived answer is correct no matter which of the two routes made the change, which is
    // exactly the interchangeability the project owner asked for.
    // ###########################################################################################
    public enum BoardRowChangeKind
    {
        // In the draft, with no published counterpart under the same natural key.
        Added,

        // In both, but at least one field differs.
        Modified,

        // Published, and absent from the draft.
        //
        // *** THIS IS ONLY MEANINGFUL BECAUSE A DRAFT IS SEEDED AS A COMPLETE COPY. *** A draft
        // workbook starts as the published workbook, so a missing row is a row the contributor
        // took out. Were drafts seeded empty (or partially), every untouched row would read as a
        // deletion and this whole report would be noise.
        Deleted
    }

    // ###########################################################################################
    // One changed row, described for a human rather than by its natural key.
    //
    // The shape deliberately matches DraftDriftReport's DraftedRowStandingItem - same Section
    // vocabulary, same "never show the natural key" rule - so the Drafts tab reads identically
    // after the storage change.
    // ###########################################################################################
    public sealed class BoardRowChange
    {
        // The BoardData section, in the words the UI already uses ("Components").
        public string Section { get; init; } = string.Empty;

        public string NaturalKey { get; init; } = string.Empty;

        // The published row's key, for a row whose identifying cells changed (PairRenamedRows) - empty for
        // every other. A pin handed from U1 to U2 changes both components (code review, 2026-11-02), and
        // NaturalKey says only U2.
        public string PreviousNaturalKey { get; init; } = string.Empty;

        // What to actually show. NEVER the natural key: its separator is U+241F, which renders as
        // a box in most fonts.
        public string DisplayLabel { get; init; } = string.Empty;

        public BoardRowChangeKind Kind { get; init; }

        // Which fields differ, for a Modified row - the property names, in declaration order.
        // Empty for Added and Deleted, where "what changed" is the whole row. A row whose identifying
        // cells changed names them here too (BoardDataDiffer.PairRenamedRows).
        //
        // Carried because "U8 changed" is far less useful than "U8: Description, Part number",
        // and the information is free at the point the comparison is made.
        public IReadOnlyList<string> ChangedFields { get; init; } = [];
    }

    // A removed row and the added row it really is, edited - indexes into the two lists handed to
    // BoardDataDiffer.PairRenamedRows.
    public readonly record struct BoardRowRename(int Removed, int Added);

    // ###########################################################################################
    // A row's place in a comparison: its natural key, and WHICH of the rows sharing that key it is - 0 for
    // all but a component's further rectangles (BoardDataDiffer.LineUpRepeatedRows). Capitals ignored, as
    // every key compares (Comparer).
    //
    // *** TWO FIELDS, NEVER ONE STRING (code review, 2026-11-02, Security). *** The place was written into the
    // key, a mark and a number after it: a rectangle labelled "U8" + U+241E + "1" had exactly the key of U8's
    // second rectangle and overwrote it, so a submission could add one that no count and no maintainer saw;
    // and taking the mark off again cut a credit named "Anna" + U+241E + "x" to "Anna" in a board's history.
    // Nothing a contributor types reaches a number held apart from the text.
    // ###########################################################################################
    internal readonly record struct ComparedRow(string NaturalKey, int Occurrence)
    {
        public static IEqualityComparer<ComparedRow> Comparer { get; } = new IgnoringCapitals();

        private sealed class IgnoringCapitals : IEqualityComparer<ComparedRow>
        {
            public bool Equals(ComparedRow first, ComparedRow second) =>
                first.Occurrence == second.Occurrence &&
                string.Equals(first.NaturalKey, second.NaturalKey, StringComparison.OrdinalIgnoreCase);

            public int GetHashCode(ComparedRow row) =>
                HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(row.NaturalKey), row.Occurrence);
        }
    }

    // ###########################################################################################
    // Compares a drafted BoardData against the published one and reports what a contributor has
    // actually changed.
    //
    // *** THIS REPLACES COUNTING DraftRow DELTAS. *** The Drafts tab's "N rows changed" chip and
    // the drift window both used to read a ledger written at edit time. With the draft stored as
    // a real board workbook there is no ledger - and must not be one, because the contributor can
    // edit that workbook in Excel without the application ever running.
    //
    // PURE. Two BoardData in, a report out; no file access, no DataManager, no DraftManager. Every
    // rule here is a unit test.
    //
    // ROW IDENTITY IS BoardDraftNaturalKeys', NOT THIS CLASS'S. That class calls itself "the one
    // place a BoardData row's identity is decided", and a second opinion here would let the report
    // disagree with the merge about which rows are "the same row" - the kind of divergence that is
    // invisible until a submission publishes the wrong thing. ForRow dispatches on row type and
    // throws for an unknown one, so a new section cannot be silently mis-keyed.
    // ###########################################################################################
    public static class BoardDataDiffer
    {
        // ###########################################################################################
        // The report. `published` may be null for a board that exists only as a draft, in which
        // case every drafted row is an addition - which is the literal truth: officially, none of
        // it exists yet.
        //
        // A null `draft` yields an empty report rather than "everything was deleted". A draft that
        // could not be read is a failure to answer the question, and answering it with the most
        // alarming possible verdict would be worse than saying nothing.
        // ###########################################################################################
        public static IReadOnlyList<BoardRowChange> Compare(BoardData? published, BoardData? draft)
        {
            if (draft is null)
            {
                return [];
            }

            var changes = new List<BoardRowChange>();

            // The section vocabulary is DraftDriftReport's, verbatim, so the two surfaces cannot
            // start naming the same section differently.
            BoardDataDiffer.CompareSection(changes, "Board schematics", published?.Schematics, draft.Schematics);
            BoardDataDiffer.CompareSection(changes, "Components", published?.Components, draft.Components);
            BoardDataDiffer.CompareSection(changes, "Component images", published?.ComponentImages, draft.ComponentImages);
            BoardDataDiffer.CompareSection(changes, "Component highlights", published?.ComponentHighlights, draft.ComponentHighlights);
            BoardDataDiffer.CompareSection(changes, "Component local files", published?.ComponentLocalFiles, draft.ComponentLocalFiles);
            BoardDataDiffer.CompareSection(changes, "Component links", published?.ComponentLinks, draft.ComponentLinks);
            BoardDataDiffer.CompareSection(changes, BoardWorkbookSchema.SheetComponentPins, published?.ComponentPins, draft.ComponentPins);
            BoardDataDiffer.CompareSection(changes, BoardWorkbookSchema.SheetComponentReplacements, published?.ComponentReplacements, draft.ComponentReplacements);
            BoardDataDiffer.CompareSection(changes, "Board local files", published?.BoardLocalFiles, draft.BoardLocalFiles);
            BoardDataDiffer.CompareSection(changes, "Board links", published?.BoardLinks, draft.BoardLinks);
            BoardDataDiffer.CompareSection(changes, "Credits", published?.Credits, draft.Credits);
            BoardDataDiffer.CompareSection(changes, "Important signals", published?.KiCadImportantSignals, draft.KiCadImportantSignals);

            return changes;
        }

        // ###########################################################################################
        // How many rows differ - the number the Drafts tab's chip shows.
        //
        // Its own method rather than Compare(...).Count at the call site, so the tab is not
        // building a full report (with display labels and per-field lists) to render one integer.
        // ###########################################################################################
        public static int CountChanges(BoardData? published, BoardData? draft) =>
            BoardDataDiffer.Compare(published, draft).Count;

        // ###########################################################################################
        // Whether ONE section of a draft differs from the published board's ROWS, by this report's
        // rule - a renamed row is one change, rows only in another ORDER are none here. Order is a
        // change too (owner decision, 2026-10-30), counted beside this report by BoardRowOrder: a
        // surface marking one section "Draft" (the component pane's board parts) asks both, as
        // Submit counts both. Null on either side is no rows.
        // ###########################################################################################
        public static bool SectionDiffers<T>(IReadOnlyList<T>? published, IReadOnlyList<T>? draft)
            where T : class
        {
            var changes = new List<BoardRowChange>();
            BoardDataDiffer.CompareSection(changes, string.Empty, published, draft);
            return changes.Count > 0;
        }

        // ###########################################################################################
        // One section, in three passes: rows only in the draft, rows in both, rows only published.
        //
        // *** DUPLICATE KEYS KEEP THE FIRST ROW AND IGNORE THE REST, ON BOTH SIDES. *** A board
        // workbook is a hand-edited file, so two rows CAN share a natural key (two "U8" component
        // rows, say). Grouping rather than ToDictionary is deliberate: ToDictionary throws, and an
        // exception here would take down a board load over a data defect the contributor can see
        // and fix in Excel. First-wins also matches what BoardDraftApplier's own merge does.
        //
        // *** EXCEPT WHERE ROWS SHARE A KEY BY RIGHT (code review, 2026-11-02): a component's rectangles on
        // one schematic. *** There every row counts, whichever it is - LineUpRepeatedRows, below.
        // ###########################################################################################
        //
        // *** A ROW WHOSE IDENTIFYING CELLS CHANGED IS STILL ONE ROW (owner decision, 2026-10-04). ***
        // A draft row with no published row under its key, and a published row with no draft row
        // under its, are paired by PairRenamedRows when nothing BUT their identifying cells differs -
        // "Dennis Helligsoe" made "Dennis Helligsoe 2" in Credits is one row Modified, not a deletion
        // plus an addition. Its NaturalKey is the draft's.
        private static void CompareSection<T>(
            List<BoardRowChange> changes,
            string sectionName,
            IReadOnlyList<T>? publishedRows,
            IReadOnlyList<T>? draftRows)
            where T : class
        {
            (Dictionary<ComparedRow, T> publishedByKey, Dictionary<ComparedRow, T> draftByKey) = BoardDraftNaturalKeys.MayRepeat(typeof(T))
                ? BoardDataDiffer.LineUpRepeatedRows(publishedRows, draftRows)
                : (BoardDataDiffer.IndexByKey(publishedRows), BoardDataDiffer.IndexByKey(draftRows));

            // The rows each side has alone, in their own order - the ones a rename can pair.
            List<KeyValuePair<ComparedRow, T>> draftOnly = draftByKey.Where(drafted => !publishedByKey.ContainsKey(drafted.Key)).ToList();
            List<KeyValuePair<ComparedRow, T>> publishedOnly = publishedByKey.Where(published => !draftByKey.ContainsKey(published.Key)).ToList();

            var renamedFrom = new Dictionary<ComparedRow, KeyValuePair<ComparedRow, T>>(ComparedRow.Comparer);
            var renamedKeys = new HashSet<ComparedRow>(ComparedRow.Comparer);

            foreach (BoardRowRename rename in BoardDataDiffer.PairRenamedRows(
                         publishedOnly.Select(row => (object)row.Value).ToList(),
                         draftOnly.Select(row => (object)row.Value).ToList()))
            {
                renamedFrom[draftOnly[rename.Added].Key] = publishedOnly[rename.Removed];
                renamedKeys.Add(publishedOnly[rename.Removed].Key);
            }

            foreach (KeyValuePair<ComparedRow, T> drafted in draftByKey)
            {
                string naturalKey = drafted.Key.NaturalKey;
                string previousKey = string.Empty;

                if (!publishedByKey.TryGetValue(drafted.Key, out T? publishedRow))
                {
                    if (!renamedFrom.TryGetValue(drafted.Key, out KeyValuePair<ComparedRow, T> renamed))
                    {
                        changes.Add(new BoardRowChange
                        {
                            Section = sectionName,
                            NaturalKey = naturalKey,
                            DisplayLabel = BoardDataDiffer.DescribeRow(drafted.Value, naturalKey),
                            Kind = BoardRowChangeKind.Added,
                        });

                        continue;
                    }

                    publishedRow = renamed.Value;
                    previousKey = renamed.Key.NaturalKey;
                }

                IReadOnlyList<string> differing = BoardDataDiffer.DifferingFields(publishedRow, drafted.Value);
                if (differing.Count == 0)
                {
                    continue;
                }

                changes.Add(new BoardRowChange
                {
                    Section = sectionName,
                    NaturalKey = naturalKey,
                    PreviousNaturalKey = previousKey,
                    DisplayLabel = BoardDataDiffer.DescribeRow(drafted.Value, naturalKey),
                    Kind = BoardRowChangeKind.Modified,
                    ChangedFields = differing,
                });
            }

            foreach (KeyValuePair<ComparedRow, T> removed in publishedOnly)
            {
                if (renamedKeys.Contains(removed.Key))
                {
                    continue;
                }

                string naturalKey = removed.Key.NaturalKey;

                changes.Add(new BoardRowChange
                {
                    Section = sectionName,
                    NaturalKey = naturalKey,
                    DisplayLabel = BoardDataDiffer.DescribeRow(removed.Value, naturalKey),
                    Kind = BoardRowChangeKind.Deleted,
                });
            }
        }

        // ###########################################################################################
        // *** ROWS THAT SHARE A KEY BY RIGHT, LINED UP (code review, 2026-11-02). *** A component can have
        // several rectangles on one schematic (BoardDraftNaturalKeys.MayRepeat - 198 such pairs on the
        // shipped boards), and keyed on schematic + label, first wins, only the first was ever compared. The
        // second one moved, a third added or the second deleted was no change to anybody: a draft holding
        // nothing else was let go of or retired without a word, could not be sent - and sent with other
        // work, was published with no maintainer told.
        //
        // So every row gets a key of its own, the two sides lined up per natural key, as a multiset of ALL
        // their values - the identity DraftRebase uses:
        //
        //   1. a row with the same values on both sides is the same row, wherever each stands - so their
        //      order alone is no change, as BoardRowOrder says of highlights;
        //   2. what is left of a key pairs in order, one row Modified - a rectangle moved, naming X or Y;
        //   3. what is left after that stands on one side alone: the renames PairRenamedRows can pair (a
        //      component relabelled, its schematic renamed), else added or deleted.
        //
        // Each row's place is its natural key and which of that key's rows it is (ComparedRow), so every
        // comparison below works as it does for any other section, and every report still names the key
        // the overlay looks up and a maintainer's lines take apart. Each side's key in its own spelling, as
        // IndexByKey has it.
        //
        // Internal because ReviewSummary lines up the same way: the contributor's count and a maintainer's
        // summary must agree (BoardDataDifferAgreementTests).
        // ###########################################################################################
        internal static (Dictionary<ComparedRow, T> Before, Dictionary<ComparedRow, T> After) LineUpRepeatedRows<T>(
            IEnumerable<T>? beforeRows,
            IEnumerable<T>? afterRows)
            where T : class
        {
            List<T> before = (beforeRows ?? []).Where(row => row is not null).ToList();
            List<T> after = (afterRows ?? []).Where(row => row is not null).ToList();

            ComparedRow[] beforeKeys = new ComparedRow[before.Count];
            ComparedRow[] afterKeys = new ComparedRow[after.Count];

            ILookup<string, int> beforeOfKey = Enumerable.Range(0, before.Count).ToLookup(index => BoardDraftNaturalKeys.ForRow(before[index]), StringComparer.OrdinalIgnoreCase);
            ILookup<string, int> afterOfKey = Enumerable.Range(0, after.Count).ToLookup(index => BoardDraftNaturalKeys.ForRow(after[index]), StringComparer.OrdinalIgnoreCase);

            foreach (string key in beforeOfKey.Select(group => group.Key).Union(afterOfKey.Select(group => group.Key), StringComparer.OrdinalIgnoreCase))
            {
                List<int> leftBefore = beforeOfKey[key].ToList();
                List<int> leftAfter = afterOfKey[key].ToList();
                int occurrence = 0;

                void Name(int? beforeIndex, int? afterIndex)
                {
                    if (beforeIndex is int b)
                        beforeKeys[b] = new ComparedRow(BoardDraftNaturalKeys.ForRow(before[b]), occurrence);

                    if (afterIndex is int a)
                        afterKeys[a] = new ComparedRow(BoardDraftNaturalKeys.ForRow(after[a]), occurrence);

                    occurrence++;
                }

                foreach (int beforeIndex in leftBefore.ToList())
                {
                    int same = leftAfter.FindIndex(afterIndex => BoardDataDiffer.SameValues(before[beforeIndex], after[afterIndex]));

                    if (same >= 0)
                    {
                        Name(beforeIndex, leftAfter[same]);
                        leftBefore.Remove(beforeIndex);
                        leftAfter.RemoveAt(same);
                    }
                }

                for (int index = 0; index < Math.Max(leftBefore.Count, leftAfter.Count); index++)
                {
                    Name(index < leftBefore.Count ? leftBefore[index] : null, index < leftAfter.Count ? leftAfter[index] : null);
                }
            }

            return (BoardDataDiffer.ByKey(before, beforeKeys), BoardDataDiffer.ByKey(after, afterKeys));
        }

        // Each row under its place, in its side's own order.
        private static Dictionary<ComparedRow, T> ByKey<T>(List<T> rows, ComparedRow[] keys)
            where T : class
        {
            var byKey = new Dictionary<ComparedRow, T>(rows.Count, ComparedRow.Comparer);

            for (int index = 0; index < rows.Count; index++)
            {
                byKey.Add(keys[index], rows[index]);
            }

            return byKey;
        }

        // ###########################################################################################
        // *** WHICH REMOVED ROW EACH ADDED ROW REALLY IS (owner decision, 2026-10-04). *** The
        // identifying cells - a credit's Name or handle, a component's Board label, a schematic's
        // name - are its natural key, so changing one used to show the row as removed plus a new row
        // added: "This seems weird to me." Natural keys cannot see a rename, so this does: `removed`
        // are rows that lost their key (published, gone from the draft), `added` rows that gained
        // one, and a removed row and an added row are the SAME ROW, edited, when
        //
        //   - they are the same kind of row, and EVERYTHING BUT THEIR IDENTIFYING CELLS IS EQUAL
        //     (DifferingFields' own rule: trimmed, case-sensitive, blank and null alike), and
        //   - they still have something to be recognised by: at least one cell, identifying or not,
        //     that is filled in and the same in both.
        //
        // The owner chose this over the looser "mostly the same row": changing the name AND the
        // contact of a credit is a different person, and stays a removal plus an addition. And it
        // never pairs two unrelated rows - an important signal (two identifying cells, nothing else)
        // pairs only while one of its halves stays the same.
        //
        // *** ONE ANSWER WHOEVER ASKS. *** The table (BoardTableSheet), this report and the server's
        // review summary (ReviewSummary) all pair through here, so they agree on what is one row.
        // When a row could be either of several, the one sharing the most filled-in identifying
        // cells wins (three renamed credits each keep their own Category), then the earlier removed
        // row, then the added row whose key sorts first - an order no caller's row order can change,
        // since the table holds new rows in screen order and the saved board in saved order.
        //
        // Returns the pairs as indexes into the two lists, each row in at most one pair.
        // ###########################################################################################
        public static IReadOnlyList<BoardRowRename> PairRenamedRows(IReadOnlyList<object> removed, IReadOnlyList<object> added)
        {
            ArgumentNullException.ThrowIfNull(removed);
            ArgumentNullException.ThrowIfNull(added);

            if (removed.Count == 0 || added.Count == 0)
            {
                return [];
            }

            // Only rows with exactly the same non-identifying cells can pair, so the added rows are
            // grouped by those cells first - a sheet relabelled wholesale is then not every removed
            // row compared with every added one.
            var addedBySameRest = new Dictionary<string, List<int>>(StringComparer.Ordinal);

            for (int index = 0; index < added.Count; index++)
            {
                if (BoardDataDiffer.RestOf(added[index]) is not string rest)
                {
                    continue;
                }

                if (!addedBySameRest.TryGetValue(rest, out List<int>? group))
                {
                    group = [];
                    addedBySameRest[rest] = group;
                }

                group.Add(index);
            }

            var candidates = new List<(int Removed, int Added, int Shared, string AddedKey)>();

            for (int removedIndex = 0; removedIndex < removed.Count; removedIndex++)
            {
                if (BoardDataDiffer.RestOf(removed[removedIndex]) is not string rest ||
                    !addedBySameRest.TryGetValue(rest, out List<int>? group))
                {
                    continue;
                }

                bool restIsFilledIn = BoardDataDiffer.HasFilledInRest(removed[removedIndex]);

                foreach (int addedIndex in group)
                {
                    int shared = BoardDataDiffer.SharedKeyCells(removed[removedIndex], added[addedIndex]);

                    if (restIsFilledIn || shared > 0)
                    {
                        candidates.Add((removedIndex, addedIndex, shared, BoardDraftNaturalKeys.ForRow(added[addedIndex])));
                    }
                }
            }

            var takenRemoved = new HashSet<int>();
            var takenAdded = new HashSet<int>();
            var pairs = new List<BoardRowRename>();

            foreach ((int removedIndex, int addedIndex, _, _) in candidates
                         .OrderByDescending(candidate => candidate.Shared)
                         .ThenBy(candidate => candidate.Removed)
                         .ThenBy(candidate => candidate.AddedKey, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(candidate => candidate.AddedKey, StringComparer.Ordinal))
            {
                if (takenRemoved.Contains(removedIndex) || takenAdded.Contains(addedIndex))
                {
                    continue;
                }

                takenRemoved.Add(removedIndex);
                takenAdded.Add(addedIndex);
                pairs.Add(new BoardRowRename(removedIndex, addedIndex));
            }

            return pairs.OrderBy(pair => pair.Removed).ToList();
        }

        // ###########################################################################################
        // A row's NON-identifying cells as one comparable string, led by its type - null for a row
        // type with no key rule, which then pairs with nothing. Each value is read as DifferingFields
        // reads it, so "equal here" is "not a differing field there".
        // ###########################################################################################
        private static string? RestOf(object row)
        {
            if (row is null)
            {
                return null;
            }

            IReadOnlyList<string> keyProperties = BoardDraftNaturalKeys.PropertiesOf(row.GetType());
            if (keyProperties.Count == 0)
            {
                return null;
            }

            IEnumerable<string> rest = BoardDataDiffer.ComparableProperties(row.GetType())
                .Where(property => !keyProperties.Contains(property.Name, StringComparer.Ordinal))
                .Select(property => BoardDataDiffer.ReadValue(property, row));

            return string.Join(BoardDraftNaturalKeys.Separator, rest.Prepend(row.GetType().FullName ?? row.GetType().Name));
        }

        // Whether any non-identifying cell is filled in. RestOf matched, so the other row's are the
        // same values.
        private static bool HasFilledInRest(object row)
        {
            IReadOnlyList<string> keyProperties = BoardDraftNaturalKeys.PropertiesOf(row.GetType());

            return BoardDataDiffer.ComparableProperties(row.GetType())
                .Where(property => !keyProperties.Contains(property.Name, StringComparer.Ordinal))
                .Any(property => BoardDataDiffer.ReadValue(property, row).Length > 0);
        }

        // How many identifying cells are filled in and the same in both - case-insensitively, as a
        // key compares.
        private static int SharedKeyCells(object removed, object added)
        {
            int shared = 0;

            foreach (string name in BoardDraftNaturalKeys.PropertiesOf(removed.GetType()))
            {
                PropertyInfo? property = removed.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (property is null)
                {
                    continue;
                }

                string before = BoardDataDiffer.ReadValue(property, removed);

                if (before.Length > 0 && string.Equals(before, BoardDataDiffer.ReadValue(property, added), StringComparison.OrdinalIgnoreCase))
                {
                    shared++;
                }
            }

            return shared;
        }

        // ###########################################################################################
        // Whether two rows of one type hold the same values - every property, trimmed, ordinal, blank
        // and null alike: exactly "no differing field" (DifferingFields). DraftRebase decides "the
        // contributor left this row alone" with it, so the merge and the count agree.
        // ###########################################################################################
        internal static bool SameValues(object first, object second)
        {
            ArgumentNullException.ThrowIfNull(first);
            ArgumentNullException.ThrowIfNull(second);

            return first.GetType() == second.GetType() && BoardDataDiffer.DifferingFields(first, second).Count == 0;
        }

        // ###########################################################################################
        // Keys one side of a section. Case-INSENSITIVE, matching BoardDraftApplier: a board label
        // typed "u8" in Excel is the same row as "U8", and treating them as two rows would report
        // a deletion plus an addition for what is a capitalisation fix at worst.
        // ###########################################################################################
        private static Dictionary<ComparedRow, T> IndexByKey<T>(IReadOnlyList<T>? rows)
            where T : class
        {
            var index = new Dictionary<ComparedRow, T>(ComparedRow.Comparer);

            if (rows is null)
            {
                return index;
            }

            foreach (T row in rows)
            {
                if (row is null)
                {
                    continue;
                }

                // First wins - see CompareSection's header.
                index.TryAdd(new ComparedRow(BoardDraftNaturalKeys.ForRow(row), 0), row);
            }

            return index;
        }

        // ###########################################################################################
        // Which properties differ between two rows of the same type.
        //
        // *** BY REFLECTION, AND THAT IS THE POINT. *** Every BoardData row type is a flat bag of
        // string properties, and they gain columns over time. A hand-written field-by-field
        // comparison would silently stop noticing a column the moment one was added - a report
        // that quietly under-counts is worse than no report, because it is believed. Walking the
        // properties means a new column is compared the day it exists.
        //
        // Values are compared TRIMMED and case-SENSITIVELY. Trimmed because a workbook round-trip
        // can pick up or lose trailing whitespace and that is not an edit the contributor made;
        // case-sensitively because changing "u8" to "U8" in a DESCRIPTION is a real edit, even
        // though the same change in a KEY is not (see IndexByKey).
        // ###########################################################################################
        private static IReadOnlyList<string> DifferingFields(object published, object drafted)
        {
            // The very same row, which is what most of two boards compared share by now: a draft held in
            // memory keeps the rows no edit touched (code review, 2026-11-02 - a reflection walk per row, on
            // the UI thread, at every typing pause).
            if (ReferenceEquals(published, drafted))
            {
                return [];
            }

            var differing = new List<string>();

            foreach (PropertyInfo property in BoardDataDiffer.ComparableProperties(drafted.GetType()))
            {
                string publishedValue = BoardDataDiffer.ReadValue(property, published);
                string draftedValue = BoardDataDiffer.ReadValue(property, drafted);

                if (!string.Equals(publishedValue, draftedValue, StringComparison.Ordinal))
                {
                    differing.Add(property.Name);
                }
            }

            return differing;
        }

        // ###########################################################################################
        // The properties worth comparing on a row type: every readable public one.
        //
        // EVERY property counts, with no exclusion list. There used to be one, for the retired
        // UuidV4 column - a value that differed between a published row and its drafted copy
        // while meaning nothing - but the column was dropped from the data entirely on
        // 2026-09-23, so there is nothing left to exclude. Do not reintroduce an exclusion list
        // without a reason as concrete as that one was: a field quietly left out of the
        // comparison is an edit the contributor never gets told about.
        //
        // Not cached. A board load compares a few thousand rows at most and reflection on a type's
        // properties is fast; a static cache here would be a mutable shared dictionary guarding
        // against a cost nobody has measured.
        // ###########################################################################################
        private static IEnumerable<PropertyInfo> ComparableProperties(Type rowType) =>
            rowType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => property.CanRead);

        // ###########################################################################################
        // One property's value as a comparable string.
        //
        // Null and blank are the SAME THING here, deliberately: a workbook round-trip writes a
        // blank cell as an empty cell (BoardWorkbookWriter.WriteTextCell returns early on a blank)
        // and the reader answers "" for both, so a row that went through a save would otherwise
        // read as Modified against its own unchanged self.
        //
        // Non-string properties are formatted invariantly. Every BoardData row field is a string
        // today, but KiCadCalibrationEntry's are doubles and bools - and it reaches this class
        // through BoardDraftNaturalKeys.ForRow, so the case is real rather than hypothetical. A
        // culture-formatted double would make the same value compare unequal between a
        // comma-decimal and a point-decimal machine.
        // ###########################################################################################
        private static string ReadValue(PropertyInfo property, object row)
        {
            object? value = property.GetValue(row);

            return value switch
            {
                null => string.Empty,
                string text => text.Trim(),
                IFormattable formattable => formattable.ToString(
                    null,
                    System.Globalization.CultureInfo.InvariantCulture),
                _ => value.ToString()?.Trim() ?? string.Empty
            };
        }

        // ###########################################################################################
        // A row described for a human. The same shapes DraftDriftReport.BuildDisplayLabel produces,
        // so the two surfaces label the same row identically.
        //
        // Falls back to the natural key's own parts, joined readably, for a row type this method
        // has not been taught - never the raw key, whose U+241F separator renders as a box.
        // ###########################################################################################
        private static string DescribeRow(object row, string naturalKey)
        {
            string fromRow = row switch
            {
                BoardSchematicEntry schematic => schematic.SchematicName,
                // With its region, when it has one - U1 for PAL and U1 for NTSC are two rows.
                ComponentEntry component => BoardDataDiffer.Join(component.BoardLabel, component.Region),
                ComponentImageEntry image => BoardDataDiffer.Join(image.BoardLabel, image.Pin, image.Name),
                ComponentHighlightEntry highlight => BoardDataDiffer.Join(highlight.SchematicName, highlight.BoardLabel),
                ComponentLocalFileEntry file => BoardDataDiffer.Join(file.BoardLabel, file.Name),
                ComponentLinkEntry link => BoardDataDiffer.Join(link.BoardLabel, link.Name),
                BoardLocalFileEntry file => BoardDataDiffer.Join(file.Category, file.Name),
                BoardLinkEntry link => BoardDataDiffer.Join(link.Category, link.Name),
                CreditEntry credit => BoardDataDiffer.Join(credit.Category, credit.NameOrHandle),
                // With its net: "9VAC" names both the 9VAC and the 9VAC~ net, two rows.
                KiCadImportantSignalEntry signal => BoardDataDiffer.Join(signal.DisplayName, signal.KiCadNetName),
                KiCadCalibrationEntry calibration => calibration.SchematicName,
                _ => string.Empty,
            };

            if (!string.IsNullOrWhiteSpace(fromRow))
            {
                return fromRow.Trim();
            }

            string fromKey = string.Join(
                " / ",
                (naturalKey ?? string.Empty).Split(BoardDraftNaturalKeys.Separator));

            return string.IsNullOrWhiteSpace(fromKey) ? "(unnamed)" : fromKey;
        }

        private static string Join(params string?[] parts) =>
            string.Join(" / ", parts
                .Select(part => part?.Trim() ?? string.Empty)
                .Where(part => part.Length > 0));
    }
}
