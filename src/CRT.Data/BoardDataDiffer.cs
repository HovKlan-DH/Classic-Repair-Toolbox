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
    // exactly the interchangeability the maintainer asked for.
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

        // What to actually show. NEVER the natural key: its separator is U+241F, which renders as
        // a box in most fonts.
        public string DisplayLabel { get; init; } = string.Empty;

        public BoardRowChangeKind Kind { get; init; }

        // Which fields differ, for a Modified row - the property names, in declaration order.
        // Empty for Added and Deleted, where "what changed" is the whole row.
        //
        // Carried because "U8 changed" is far less useful than "U8: Description, Part number",
        // and the information is free at the point the comparison is made.
        public IReadOnlyList<string> ChangedFields { get; init; } = [];
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
        // The report. `published` may be null for a system that exists only as a draft, in which
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
        // One section, in three passes: rows only in the draft, rows in both, rows only published.
        //
        // *** DUPLICATE KEYS KEEP THE FIRST ROW AND IGNORE THE REST, ON BOTH SIDES. *** A board
        // workbook is a hand-edited file, so two rows CAN share a natural key (two "U8" component
        // rows, say). Grouping rather than ToDictionary is deliberate: ToDictionary throws, and an
        // exception here would take down a board load over a data defect the contributor can see
        // and fix in Excel. First-wins also matches what BoardDraftApplier's own merge does.
        // ###########################################################################################
        private static void CompareSection<T>(
            List<BoardRowChange> changes,
            string sectionName,
            IReadOnlyList<T>? publishedRows,
            IReadOnlyList<T>? draftRows)
            where T : class
        {
            Dictionary<string, T> publishedByKey = BoardDataDiffer.IndexByKey(publishedRows);
            Dictionary<string, T> draftByKey = BoardDataDiffer.IndexByKey(draftRows);

            foreach (KeyValuePair<string, T> drafted in draftByKey)
            {
                if (!publishedByKey.TryGetValue(drafted.Key, out T? publishedRow))
                {
                    changes.Add(new BoardRowChange
                    {
                        Section = sectionName,
                        NaturalKey = drafted.Key,
                        DisplayLabel = BoardDataDiffer.DescribeRow(drafted.Value, drafted.Key),
                        Kind = BoardRowChangeKind.Added,
                    });

                    continue;
                }

                IReadOnlyList<string> differing = BoardDataDiffer.DifferingFields(publishedRow, drafted.Value);
                if (differing.Count == 0)
                {
                    continue;
                }

                changes.Add(new BoardRowChange
                {
                    Section = sectionName,
                    NaturalKey = drafted.Key,
                    DisplayLabel = BoardDataDiffer.DescribeRow(drafted.Value, drafted.Key),
                    Kind = BoardRowChangeKind.Modified,
                    ChangedFields = differing,
                });
            }

            foreach (KeyValuePair<string, T> publishedOnly in publishedByKey)
            {
                if (draftByKey.ContainsKey(publishedOnly.Key))
                {
                    continue;
                }

                changes.Add(new BoardRowChange
                {
                    Section = sectionName,
                    NaturalKey = publishedOnly.Key,
                    DisplayLabel = BoardDataDiffer.DescribeRow(publishedOnly.Value, publishedOnly.Key),
                    Kind = BoardRowChangeKind.Deleted,
                });
            }
        }

        // ###########################################################################################
        // Keys one side of a section. Case-INSENSITIVE, matching BoardDraftApplier: a board label
        // typed "u8" in Excel is the same row as "U8", and treating them as two rows would report
        // a deletion plus an addition for what is a capitalisation fix at worst.
        // ###########################################################################################
        private static Dictionary<string, T> IndexByKey<T>(IReadOnlyList<T>? rows)
            where T : class
        {
            var index = new Dictionary<string, T>(StringComparer.OrdinalIgnoreCase);

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

                string key = BoardDraftNaturalKeys.ForRow(row);

                // First wins - see CompareSection's header.
                if (!index.ContainsKey(key))
                {
                    index[key] = row;
                }
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
                KiCadImportantSignalEntry signal => signal.DisplayName,
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
