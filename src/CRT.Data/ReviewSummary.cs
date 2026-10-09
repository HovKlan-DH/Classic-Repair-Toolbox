using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What CHANGED between the published board and a submitted one, counted per section
    // (NewContributeStrategy.md Phase 5, task 3).
    //
    // *** THE MAINTAINER TAB OPENS ON THIS, NEVER ON A WHOLE BOARD. *** That is task 3's instruction
    // and it is the difference between reviewing and reading: a C64 board carries hundreds of
    // components and over a thousand files, and a maintainer shown all of it has to find the change
    // themselves. "3 components changed, 1 added, 2 images added, 1 highlight moved" is a
    // reviewable sentence; a board is not.
    //
    // *** ROWS PAIR ON NATURAL KEYS, NOT ON POSITION. *** Phase 4 retired UuidV4 as an identity,
    // so a row is identified by what it IS - a component by its board label, a highlight by its
    // schematic plus label. Comparing by list position would report every row below an insertion
    // as changed, which is the classic diff bug and would make the summary useless on exactly the
    // submissions it matters most for.
    //
    // *** A RENAME IS CARRIED EXPLICITLY, because natural keys cannot see one. *** Renaming U8 to
    // U9 reads as a delete plus an add, which is technically true and tells the maintainer nothing.
    // A client that knows it was a rename can say so in the manifest, and this honours that - and
    // since 2026-10-04 a rename nobody declared is recognised too, by the table's own rule
    // (BoardDataDiffer.PairRenamedRows).
    //
    // Pure, so the whole thing is unit tested with no UI, no database and no files - which is the
    // point of putting the Maintainer tab's central screen in CRT.Data rather than in its code-behind.
    // ###########################################################################################
    public static class ReviewSummary
    {
        // ###########################################################################################
        // Compares the published board against the submitted one.
        //
        // published may be null: that is a NEW BOARD, where every row is an addition. It is a
        // real and important case - a new board is the highest-risk submission there is (Phase 6
        // task 3) - so it must produce a summary rather than an error.
        // ###########################################################################################
        public static ReviewChangeSummary Compare(
            BoardData? published,
            BoardData submitted,
            IEnumerable<SubmissionRename>? renames = null,

            // ###########################################################################################
            // The KiCad CALIBRATIONS, which cannot travel inside either BoardData.
            //
            // *** WITHOUT THESE, CALIBRATIONS PUBLISH UNREVIEWED. *** They began travelling with
            // submissions on 2026-09-22; BoardData has no calibration section, so comparing only
            // the two boards would let them reach the published tree with no maintainer ever having
            // seen them. Getting one wrong breaks nothing visibly - the trace overlay simply lands
            // in the wrong place - so it is exactly the change a maintainer must be TOLD about.
            //
            // Optional and trailing, so every existing caller keeps working and reports no
            // calibration changes, which is right for one that does not know about them.
            // ###########################################################################################
            IEnumerable<KiCadCalibrationEntry>? publishedCalibrations = null,
            IEnumerable<KiCadCalibrationEntry>? submittedCalibrations = null)
        {
            ArgumentNullException.ThrowIfNull(submitted);

            BoardData before = published ?? new BoardData();

            var sections = new List<ReviewSectionChange>();

            foreach (ReviewSection section in ReviewSummary.AllSections)
            {
                sections.Add(ReviewSummary.CompareSection(section, before, submitted, renames));
            }

            // The eleventh section, compared the same way as the other ten but from its own pair
            // of lists rather than from the boards.
            sections.Add(ReviewSummary.CompareCalibrations(
                publishedCalibrations, submittedCalibrations, renames));

            bool isNewBoard = published == null;

            return new ReviewChangeSummary(
                isNewBoard,
                ReviewSummary.RevisionDateChange(before, submitted),
                sections);
        }

        // ###########################################################################################
        // The KiCad calibrations, compared.
        //
        // Their own method rather than an eleventh entry in AllSections, because AllSections maps
        // BoardData to rows and a calibration is not on BoardData at all. Everything below the key
        // lookup is shared with the other ten - the same natural keys, the same field diff - so a
        // calibration change reads exactly like any other change on the review screen.
        // ###########################################################################################
        private static ReviewSectionChange CompareCalibrations(
            IEnumerable<KiCadCalibrationEntry>? before,
            IEnumerable<KiCadCalibrationEntry>? after,
            IEnumerable<SubmissionRename>? renames)
        {
            Dictionary<string, object> beforeRows = ReviewSummary.KeyCalibrations(before);
            Dictionary<string, object> afterRows = ReviewSummary.KeyCalibrations(after);

            return ReviewSummary.CompareKeyedRows(
                ReviewSummary.SectionKiCadCalibrations, beforeRows, afterRows, renames);
        }

        // Invariant, always - see the KiCadCalibrationEntry arm of FieldsOf. "R" round-trips, so a
        // value read back from the sidecar formats identically to the one that was written.
        private static string Number(double value) =>
            value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

        private static Dictionary<string, object> KeyCalibrations(IEnumerable<KiCadCalibrationEntry>? entries)
        {
            var rows = new Dictionary<string, object>(StringComparer.Ordinal);

            foreach (KiCadCalibrationEntry entry in entries ?? [])
            {
                string key = BoardDraftNaturalKeys.ForRow(entry);

                // Last one wins, matching KeyRows - a duplicate key in contributed data is the
                // client's problem to report, not a reason to fail the whole summary.
                if (!string.IsNullOrEmpty(key))
                    rows[key] = entry;
            }

            return rows;
        }

        // ###########################################################################################
        // One section's changes.
        //
        // A rename is recorded as a RENAME and removed from both the added and removed lists, so
        // it is never double-counted. Doing that here rather than in the caller means every
        // consumer - the summary line, the drill-down, the exported report - agrees about what
        // happened.
        // ###########################################################################################
        private static ReviewSectionChange CompareSection(
            ReviewSection section,
            BoardData before,
            BoardData after,
            IEnumerable<SubmissionRename>? renames)
        {
            Dictionary<string, object> beforeRows = ReviewSummary.KeyRows(section, before);
            Dictionary<string, object> afterRows = ReviewSummary.KeyRows(section, after);

            return ReviewSummary.CompareKeyedRows(section.Name, beforeRows, afterRows, renames);
        }

        // ###########################################################################################
        // The comparison itself, once the rows have been keyed.
        //
        // Split out of CompareSection so the KiCad calibrations - which are keyed from their own
        // list rather than from a BoardData - go through the SAME logic rather than a second copy
        // of it. Two implementations of "what changed" would eventually disagree, and the one that
        // drifted would be the one nobody reads.
        // ###########################################################################################
        private static ReviewSectionChange CompareKeyedRows(
            string sectionName,
            Dictionary<string, object> beforeRows,
            Dictionary<string, object> afterRows,
            IEnumerable<SubmissionRename>? renames)
        {
            // Renames declared for THIS section only. A rename in another section must not move a
            // row here that happens to share a key.
            List<SubmissionRename> sectionRenames = (renames ?? [])
                .Where(rename => string.Equals(rename.Section, sectionName, StringComparison.OrdinalIgnoreCase))
                .Where(rename => !string.IsNullOrWhiteSpace(rename.From) && !string.IsNullOrWhiteSpace(rename.To))
                .ToList();

            var renamed = new List<ReviewRenamedRow>();
            var handledBefore = new HashSet<string>(StringComparer.Ordinal);
            var handledAfter = new HashSet<string>(StringComparer.Ordinal);

            foreach (SubmissionRename rename in sectionRenames)
            {
                // *** A DECLARED RENAME IS UNTRUSTED INPUT - it arrives inside the submission. ***
                // It is honoured only when it actually describes what is there, and that takes
                // THREE conditions, not two:
                //
                //   - the old key existed before;
                //   - the new key exists now;
                //   - the old key is GONE now.
                //
                // The third is the one a first implementation omitted, and the omission matters:
                // with U8 still present, a submission declaring "U8 became U9" would have had U9
                // absorbed as a rename and never shown to the maintainer as the ADDITION it really
                // is. Hiding an added row from the person approving it is precisely what this
                // screen exists to prevent. Caught by its own test.
                if (!beforeRows.ContainsKey(rename.From) ||
                    !afterRows.ContainsKey(rename.To) ||
                    afterRows.ContainsKey(rename.From))
                {
                    continue;
                }

                // *** "ALSO CHANGED" MEANS "BESIDES THE RENAME". *** The renamed value is itself
                // one of the compared fields - a component's BoardLabel IS its key - so a
                // straight row comparison reports every rename as also changed, which is true and
                // useless: it is the rename, counted twice. What a maintainer needs to know is
                // whether anything OTHER than the name moved, because "renamed" invites them not
                // to look further. So the key's own parts are excluded from this comparison.
                //
                // Found by a test asserting a clean rename reports AlsoChanged == false. The
                // first implementation compared whole rows and reported true.
                renamed.Add(new ReviewRenamedRow(
                    rename.From,
                    rename.To,
                    !ReviewSummary.RowsMatchIgnoringKey(beforeRows[rename.From], afterRows[rename.To])));

                handledBefore.Add(rename.From);
                handledAfter.Add(rename.To);
            }

            // ###########################################################################################
            // *** AND A RENAME NOBODY DECLARED (owner decision, 2026-10-04). *** No client declares
            // one, so a highlight whose label changed, or a calibration whose schematic was renamed,
            // read as one removed and one added. The rows left over on each side are paired by the
            // rule the table and BoardDataDiffer pair by - nothing but their identifying cells
            // different (BoardDataDiffer.PairRenamedRows) - so the maintainer's summary counts one
            // change where the contributor's table showed one. That rule IS "nothing else changed",
            // so such a pair is never "also changed" - and is not put through RowsMatchIgnoringKey,
            // whose drop-by-value cannot line the fields up when the key gains a part (a component
            // given a region).
            // ###########################################################################################
            List<string> lostKey = [.. beforeRows.Keys.Where(key => !handledBefore.Contains(key) && !afterRows.ContainsKey(key))];
            List<string> gainedKey = [.. afterRows.Keys.Where(key => !handledAfter.Contains(key) && !beforeRows.ContainsKey(key))];

            foreach (BoardRowRename pair in BoardDataDiffer.PairRenamedRows(
                         [.. lostKey.Select(key => beforeRows[key])],
                         [.. gainedKey.Select(key => afterRows[key])]))
            {
                string from = lostKey[pair.Removed];
                string to = gainedKey[pair.Added];

                renamed.Add(new ReviewRenamedRow(from, to, AlsoChanged: false));
                handledBefore.Add(from);
                handledAfter.Add(to);
            }

            var added = new List<string>();
            var changed = new List<string>();
            var fieldChanges = new Dictionary<string, IReadOnlyList<ReviewFieldChange>>(StringComparer.Ordinal);

            foreach ((string key, object row) in afterRows)
            {
                if (handledAfter.Contains(key))
                    continue;

                if (!beforeRows.TryGetValue(key, out object? existing))
                {
                    added.Add(key);
                    continue;
                }

                if (!ReviewSummary.RowsMatch(existing, row))
                {
                    changed.Add(key);
                    fieldChanges[key] = ReviewSummary.FieldChanges(existing, row);
                }
            }

            List<string> removed = [.. beforeRows.Keys
                .Where(key => !handledBefore.Contains(key) && !afterRows.ContainsKey(key))];

            return new ReviewSectionChange(
                sectionName,
                [.. added.OrderBy(key => key, StringComparer.Ordinal)],
                [.. removed.OrderBy(key => key, StringComparer.Ordinal)],
                [.. changed.OrderBy(key => key, StringComparer.Ordinal)],
                renamed,
                fieldChanges);
        }

        // ###########################################################################################
        // Whether two rows of the same section are field-for-field identical.
        //
        // *** UuidV4 IS DELIBERATELY IGNORED. *** Phase 4 retired it as an identity and stopped
        // writing new ones, so an old row carrying one and a resubmitted row without it are the
        // SAME row as far as a maintainer is concerned. Comparing it would report a change nobody
        // made, on every row of every board authored before the transition.
        //
        // Comparison is ORDINAL and case-SENSITIVE, like every other comparison in this project:
        // changing "u8" to "U8" is a real edit a maintainer should see, and on the Linux server a
        // file name's case genuinely matters.
        // ###########################################################################################
        private static bool RowsMatch(object before, object after) =>
            ReviewSummary.FieldsOf(before).Select(field => field.Value)
                .SequenceEqual(ReviewSummary.FieldsOf(after).Select(field => field.Value), StringComparer.Ordinal);

        // ###########################################################################################
        // The same comparison with the fields that make up the natural key removed - used only for
        // a declared rename, where those fields are EXPECTED to differ because that difference is
        // the rename.
        //
        // The key's parts are dropped by VALUE rather than by index: the fields that form a key
        // are not always the leading ones (a component image's key is label + region + pin + name,
        // which are positions 1, 2, 3 and 4 of ten), and an index list would silently rot the
        // first time a field was inserted. Dropping by value is slightly blunter - a non-key field
        // that happens to hold the same text as a key part is dropped too - but the cost is at
        // worst an unreported edit to a field that duplicates the name, whereas a stale index list
        // silently compares the wrong columns.
        // ###########################################################################################
        private static bool RowsMatchIgnoringKey(object before, object after)
        {
            HashSet<string> keyParts =
            [
                .. BoardDraftNaturalKeys.ForRow(before).Split(BoardDraftNaturalKeys.Separator),
                .. BoardDraftNaturalKeys.ForRow(after).Split(BoardDraftNaturalKeys.Separator)
            ];

            return ReviewSummary.FieldsOf(before)
                .Where(field => !keyParts.Contains(field.Value))
                .Select(field => field.Value)
                .SequenceEqual(
                    ReviewSummary.FieldsOf(after)
                        .Where(field => !keyParts.Contains(field.Value))
                        .Select(field => field.Value),
                    StringComparer.Ordinal);
        }

        // ###########################################################################################
        // One row's fields, as NAME/VALUE pairs.
        //
        // *** THE NAMES ARE THE WORKBOOK'S OWN COLUMN HEADERS, from BoardWorkbookSchema. *** A
        // maintainer reading "Part-number changed" and someone opening the .xlsx to look must
        // be talking about the same column. Inventing friendlier labels here would create a second
        // vocabulary for the same data, and the two would drift.
        //
        // UuidV4 is absent on purpose - see RowsMatch.
        // ###########################################################################################
        private static IReadOnlyList<ReviewField> FieldsOf(object row) => row switch
        {
            BoardSchematicEntry entry =>
            [
                new(BoardWorkbookSchema.ColSchematicName, entry.SchematicName),
                new(BoardWorkbookSchema.ColCadName, entry.CadName),
                new(BoardWorkbookSchema.ColSchematicImageFile, entry.SchematicImageFile),
                new(BoardWorkbookSchema.ColSchematicHighlightColor, entry.SchematicHighlightColor),
                new(BoardWorkbookSchema.ColSchematicHighlightOpacity, entry.SchematicHighlightOpacity),
                new(BoardWorkbookSchema.ColOppositeTraceHighlightColor, entry.OppositeTraceHighlightColor),
                new(BoardWorkbookSchema.ColThumbnailHighlightColor, entry.ThumbnailHighlightColor),
                new(BoardWorkbookSchema.ColThumbnailHighlightOpacity, entry.ThumbnailHighlightOpacity)
            ],
            ComponentEntry entry =>
            [
                new(BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                new(BoardWorkbookSchema.ColFriendlyName, entry.FriendlyName),
                new(BoardWorkbookSchema.ColTechnicalNameOrValue, entry.TechnicalNameOrValue),
                new(BoardWorkbookSchema.ColPartNumber, entry.PartNumber),
                new(BoardWorkbookSchema.ColCategory, entry.Category),
                new(BoardWorkbookSchema.ColRegion, entry.Region),
                new(BoardWorkbookSchema.ColDescription, entry.Description)
            ],
            ComponentImageEntry entry =>
            [
                new(BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                new(BoardWorkbookSchema.ColRegion, entry.Region),
                new(BoardWorkbookSchema.ColPin, entry.Pin),
                new(BoardWorkbookSchema.ColName, entry.Name),
                new(BoardWorkbookSchema.ColExpectedOscilloscopeReading, entry.ExpectedOscilloscopeReading),
                new(BoardWorkbookSchema.ColFile, entry.File),
                new(BoardWorkbookSchema.ColNote, entry.Note),
                new(BoardWorkbookSchema.ColTimeDiv, entry.TimeDiv),
                new(BoardWorkbookSchema.ColVoltsDiv, entry.VoltsDiv),
                new(BoardWorkbookSchema.ColTriggerLevelVolts, entry.TriggerLevelVolts)
            ],
            // The coordinates are STRINGS here, like every other BoardData field - they are
            // parsed invariant-culture where they are used, not on the way in. So they compare as
            // text, which is also what a maintainer wants: "100" becoming "100.0" is a real edit to
            // the stored data even though the number is the same.
            ComponentHighlightEntry entry =>
            [
                new(BoardWorkbookSchema.ColSchematicName, entry.SchematicName),
                new(BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                new("X", entry.X),
                new("Y", entry.Y),
                new("Width", entry.Width),
                new("Height", entry.Height)
            ],
            ComponentLocalFileEntry entry =>
            [
                new(BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                new(BoardWorkbookSchema.ColName, entry.Name),
                new(BoardWorkbookSchema.ColFile, entry.File)
            ],
            ComponentLinkEntry entry =>
            [
                new(BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                new(BoardWorkbookSchema.ColName, entry.Name),
                new(BoardWorkbookSchema.ColUrl, entry.Url)
            ],
            BoardLocalFileEntry entry =>
            [
                new(BoardWorkbookSchema.ColCategory, entry.Category),
                new(BoardWorkbookSchema.ColName, entry.Name),
                new(BoardWorkbookSchema.ColFile, entry.File)
            ],
            BoardLinkEntry entry =>
            [
                new(BoardWorkbookSchema.ColCategory, entry.Category),
                new(BoardWorkbookSchema.ColName, entry.Name),
                new(BoardWorkbookSchema.ColUrl, entry.Url)
            ],
            CreditEntry entry =>
            [
                new(BoardWorkbookSchema.ColCategory, entry.Category),
                new(BoardWorkbookSchema.ColSubCategory, entry.SubCategory),
                new(BoardWorkbookSchema.ColNameOrHandle, entry.NameOrHandle),
                new(BoardWorkbookSchema.ColContact, entry.Contact)
            ],
            KiCadImportantSignalEntry entry =>
            [
                new(BoardWorkbookSchema.ColDisplayName, entry.DisplayName),
                new(BoardWorkbookSchema.ColKiCadNetName, entry.KiCadNetName)
            ],
            // ###########################################################################################
            // The KiCad calibration's fields. Unlike every other row type these are DOUBLES and
            // BOOLS rather than strings, so they are FORMATTED here - and formatted INVARIANT.
            //
            // A culture-sensitive format would render 0.75 as "0,75" on a Danish server and report
            // a change on every field of every calibration, forever, against a board nobody had
            // touched. Fifth encounter with that class of bug in this project.
            //
            // The names are the JSON sidecar's own field names, so a maintainer reading "ScaleX" and
            // someone opening the `.json` by hand are talking about the same thing.
            // ###########################################################################################
            KiCadCalibrationEntry entry =>
            [
                new("CadName", entry.CadName),
                new("OffsetX", ReviewSummary.Number(entry.OffsetX)),
                new("OffsetY", ReviewSummary.Number(entry.OffsetY)),
                new("ScaleX", ReviewSummary.Number(entry.ScaleX)),
                new("ScaleY", ReviewSummary.Number(entry.ScaleY)),
                new("MirrorX", entry.MirrorX ? "yes" : "no"),
                new("MirrorY", entry.MirrorY ? "yes" : "no")
            ],
            _ => throw new ArgumentOutOfRangeException(
                nameof(row),
                row.GetType().Name,
                "No field list for this row type - add one alongside its section.")
        };

        // ###########################################################################################
        // Which FIELDS differ between two versions of the same row (Phase 5, task 4's "text rows
        // as a field-level diff").
        //
        // *** THIS IS WHAT MAKES A "changed" ROW REVIEWABLE. *** Knowing that U8 changed tells a
        // maintainer nothing they can act on; knowing its Part-number went from 906114 to 251715-01
        // is the whole decision. Without it they would have to open the board in CRT and hunt.
        //
        // A cleared field is reported with an EMPTY "after" rather than omitted, because deleting
        // information is a real edit and the one hardest to notice - arguably the most important
        // thing on this screen.
        // ###########################################################################################
        private static IReadOnlyList<ReviewFieldChange> FieldChanges(object before, object after)
        {
            IReadOnlyList<ReviewField> from = ReviewSummary.FieldsOf(before);
            IReadOnlyList<ReviewField> to = ReviewSummary.FieldsOf(after);

            var changes = new List<ReviewFieldChange>();

            // Both sides are the same row TYPE, so the lists are the same length and in the same
            // order by construction. Guarded anyway: a mismatch would mean a field list was edited
            // on one arm of the switch only, and silently comparing the wrong columns is worse
            // than reporting nothing.
            if (from.Count != to.Count)
                return changes;

            for (int index = 0; index < from.Count; index++)
            {
                if (!string.Equals(from[index].Value, to[index].Value, StringComparison.Ordinal))
                    changes.Add(new ReviewFieldChange(from[index].Name, from[index].Value, to[index].Value));
            }

            return changes;
        }


        // ###########################################################################################
        // A section's rows keyed by natural key.
        //
        // A DUPLICATE KEY KEEPS THE FIRST ROW rather than throwing. Duplicates are a validation
        // error SubmissionValidator already reports, and a maintainer opening a submission to find
        // out what is wrong with it must not be met with a crash instead of the answer.
        // ###########################################################################################
        private static Dictionary<string, object> KeyRows(ReviewSection section, BoardData data)
        {
            var rows = new Dictionary<string, object>(StringComparer.Ordinal);

            foreach (object row in section.RowsOf(data))
            {
                string key = BoardDraftNaturalKeys.ForRow(row);

                if (!string.IsNullOrEmpty(key))
                    rows.TryAdd(key, row);
            }

            return rows;
        }

        private static ReviewFieldChange? RevisionDateChange(BoardData before, BoardData after) =>
            string.Equals(before.RevisionDate, after.RevisionDate, StringComparison.Ordinal)
                ? null
                : new ReviewFieldChange("Revision date", before.RevisionDate, after.RevisionDate);

        // ###########################################################################################
        // The sections, in the order a maintainer reads them: what the board IS first (schematics,
        // components), then what is attached to it, then metadata.
        //
        // Section names are the SHEET names, so a finding that names a section and a workbook
        // someone opens by hand agree about what it is called.
        // ###########################################################################################
        // ###########################################################################################
        // The highlights section's name, as a constant because the MAINTAINER TAB matches on it.
        //
        // Every other section is named by a BoardWorkbookSchema sheet constant, but highlights do
        // not live in the workbook at all - they are in the `.json` sidecar beside it
        // (BoardComponentHighlightStorage), whose own property name this matches exactly. So there
        // is no sheet constant to borrow, and it was a bare literal until the Maintainer tab needed to
        // find this section to draw a moved highlight. A literal on both sides of a process
        // boundary is a rename waiting to break the one screen that renders geometry.
        // ###########################################################################################
        public const string SectionComponentHighlights = "Component highlights";

        // ###########################################################################################
        // The KiCad calibration section's name. Like highlights, this does NOT come from a workbook
        // sheet - calibrations live in the `.json` sidecar under their own root - so there is no
        // BoardWorkbookSchema constant to borrow. Named to match that root, so a maintainer reading
        // this and someone opening the file by hand are talking about the same thing.
        // ###########################################################################################
        public const string SectionKiCadCalibrations = "KiCad calibration points";

        public static readonly IReadOnlyList<ReviewSection> AllSections =
        [
            new(BoardWorkbookSchema.SheetBoardSchematics, data => data.Schematics.Cast<object>()),
            new(BoardWorkbookSchema.SheetComponents, data => data.Components.Cast<object>()),
            new(BoardWorkbookSchema.SheetComponentImages, data => data.ComponentImages.Cast<object>()),
            new(ReviewSummary.SectionComponentHighlights, data => data.ComponentHighlights.Cast<object>()),
            new(BoardWorkbookSchema.SheetComponentLocalFiles, data => data.ComponentLocalFiles.Cast<object>()),
            new(BoardWorkbookSchema.SheetComponentLinks, data => data.ComponentLinks.Cast<object>()),
            new(BoardWorkbookSchema.SheetBoardLocalFiles, data => data.BoardLocalFiles.Cast<object>()),
            new(BoardWorkbookSchema.SheetBoardLinks, data => data.BoardLinks.Cast<object>()),
            new(BoardWorkbookSchema.SheetCredits, data => data.Credits.Cast<object>()),
            new(BoardWorkbookSchema.SheetKiCadImportantSignals, data => data.KiCadImportantSignals.Cast<object>())
        ];
    }

    public sealed record ReviewSection(string Name, Func<BoardData, IEnumerable<object>> RowsOf);

    public sealed record ReviewFieldChange(string Field, string Before, string After);

    // One field of one row: the workbook's own column name, and the value in it.
    public sealed record ReviewField(string Name, string Value);

    public sealed record ReviewRenamedRow(string From, string To, bool AlsoChanged);

    public sealed record ReviewSectionChange(
        string Section,
        IReadOnlyList<string> Added,
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Changed,
        IReadOnlyList<ReviewRenamedRow> Renamed,

        // Which FIELDS moved, per changed row key (task 4's field-level diff). Keyed rather than
        // parallel to Changed, so a caller cannot pair them up wrongly by index.
        IReadOnlyDictionary<string, IReadOnlyList<ReviewFieldChange>> FieldChanges)
    {
        public bool HasChanges =>
            this.Added.Count > 0 ||
            this.Removed.Count > 0 ||
            this.Changed.Count > 0 ||
            this.Renamed.Count > 0;

        public int TotalChanges =>
            this.Added.Count + this.Removed.Count + this.Changed.Count + this.Renamed.Count;
    }

    // ###########################################################################################
    // The whole summary. Sections with no changes are KEPT rather than filtered out, so a caller
    // can show "no changes" for a section if it wants to - but Describe() omits them, because the
    // opening line must be the changes and nothing else.
    // ###########################################################################################
    public sealed record ReviewChangeSummary(
        bool IsNewBoard,
        ReviewFieldChange? RevisionDate,
        IReadOnlyList<ReviewSectionChange> Sections)
    {
        public IReadOnlyList<ReviewSectionChange> ChangedSections =>
            [.. this.Sections.Where(section => section.HasChanges)];

        public int TotalChanges => this.Sections.Sum(section => section.TotalChanges);

        public bool HasChanges => this.TotalChanges > 0 || this.RevisionDate != null;

        // ###########################################################################################
        // The one-line summary task 3 asks for: "3 components changed, 1 added, 2 images added".
        //
        // Written for a human at a glance, so it names the SECTION and the verb and nothing else -
        // the keys are in the drill-down. A new board says so instead of listing every row as an
        // addition, which would be a true but useless sentence hundreds of items long.
        // ###########################################################################################
        public string Describe()
        {
            if (this.IsNewBoard)
            {
                int rows = this.Sections.Sum(section => section.Added.Count);
                return $"New board, {rows} {ReviewChangeSummary.Plural(rows, "row", "rows")}";
            }

            if (!this.HasChanges)
                return "No changes";

            var parts = new List<string>();

            foreach (ReviewSectionChange section in this.ChangedSections)
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

            if (this.RevisionDate != null)
                parts.Add("revision date changed");

            return string.Join(", ", parts);
        }

        private static string Plural(int count, string one, string many) => count == 1 ? one : many;
    }
}
