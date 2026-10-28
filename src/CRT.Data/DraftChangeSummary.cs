using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // What became of a component in the draft - only the board's own group has none.
    public enum DraftComponentChange
    {
        Changed,
        Added,
        Removed
    }

    // ###########################################################################################
    // One group of a draft's summary: a component (`Label` its board label), or the board itself
    // (`Label` empty) - with a title to show and one plain line per thing that changed.
    // ###########################################################################################
    public sealed record DraftChangeGroup(string Label, string Title, DraftComponentChange Change, IReadOnlyList<string> Lines)
    {
        public bool IsComponent => this.Label.Length > 0;
    }

    // ###########################################################################################
    // WHAT A DRAFT CHANGES, IN WORDS A PERSON READS (CRT 4.0 step 1, 2026-10-25 - the plan: "Review
    // changes shows a readable summary, and the Submit dialog shows the same").
    //
    // The Drafts tab's row says "12 rows changed", and the table colours each cell - neither says
    // what the work IS. This does, grouped as a contributor thinks of it: per component ("IC1 - ULA
    // 6C001E-7: description changed, 2 pictures added"), and the board's own parts ("Schematic image
    // [Top] renamed [Front]", "KiCad calibration of [Top] changed").
    //
    // *** THE DIFFERENCES ARE BoardDataDiffer'S. *** Every section is compared by the differ the
    // table and the row count use (BoardDraftNaturalKeys, PairRenamedRows), per component - so what
    // this calls one change is what they call one row. A relabelled component is one component
    // (ComponentRelabel, the rule its highlights follow), a renamed schematic one schematic
    // (SubmissionRowsBoard.FollowSchematics). The files the board lists are compared too when given:
    // a picture replaced under its own name changes no row, and is still work.
    //
    // Pure: two boards in, a summary out. The words are here, once, for every window that shows it.
    // ###########################################################################################
    public sealed class DraftChangeSummary
    {
        public const string BoardTitle = "The board";

        private DraftChangeSummary(IReadOnlyList<DraftChangeGroup> groups)
        {
            this.Groups = groups;
        }

        // The board's own group first, then the components in label order.
        public IReadOnlyList<DraftChangeGroup> Groups { get; }

        public bool IsEmpty => this.Groups.Count == 0;

        public static DraftChangeSummary Empty { get; } = new([]);

        // ###########################################################################################
        // The summary as single lines - "IC1 - ULA: Description changed" - for a dialog that only
        // lists what goes (Discard draft). At most `max`, the last then saying how many more.
        // ###########################################################################################
        public IReadOnlyList<string> AsLines(int max = int.MaxValue)
        {
            List<string> lines = this.Groups
                .SelectMany(group => group.Lines.Select(line => $"{group.Title}: {line}"))
                .ToList();

            if (max <= 0 || lines.Count <= max)
                return lines;

            int more = lines.Count - (max - 1);
            return [.. lines.Take(max - 1), $"... and {more.ToString(CultureInfo.InvariantCulture)} more"];
        }

        // ###########################################################################################
        // The summary of `draft` against `published` (null for a board that exists only as a draft:
        // every component is then added). `publishedFiles` and `draftFiles` are the two board files'
        // lists of files, by name - when given, a file whose bytes changed under its own name, and
        // the board's KiCad data, are said too.
        // ###########################################################################################
        public static DraftChangeSummary Build(
            BoardContent? published,
            BoardContent draft,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? publishedFiles = null,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draftFiles = null)
        {
            ArgumentNullException.ThrowIfNull(draft);

            BoardData before = published?.Board ?? DraftChangeSummary.NoBoard();
            BoardData after = draft.Board;

            IReadOnlyDictionary<string, string> relabelled = ComponentRelabel.Between(before.Components, after.Components);
            List<string> replaced = DraftChangeSummary.ReplacedFiles(publishedFiles, draftFiles);

            var groups = new List<DraftChangeGroup>();

            List<string> boardLines = DraftChangeSummary.BoardLines(before, after, published?.Calibrations ?? [], draft.Calibrations);
            boardLines.AddRange(DraftChangeSummary.KiCadLines(publishedFiles, draftFiles));

            // A replaced file cited by no component row is the board's.
            var citedByComponents = new HashSet<string>(
                DraftChangeSummary.ComponentFileNames(after).Select(cited => cited.Name),
                StringComparer.OrdinalIgnoreCase);

            boardLines.AddRange(replaced
                .Where(name => !citedByComponents.Contains(name) && !SubmissionKiCadFiles.IsInKiCadFolder(name))
                .Select(name => $"File [{name}] replaced"));

            if (boardLines.Count > 0)
                groups.Add(new DraftChangeGroup(string.Empty, DraftChangeSummary.BoardTitle, DraftComponentChange.Changed, boardLines));

            ILookup<string, (string Name, string Label)> replacedByLabel = DraftChangeSummary.ComponentFileNames(after)
                .Where(cited => replaced.Contains(cited.Name, StringComparer.OrdinalIgnoreCase))
                .Distinct()
                .ToLookup(cited => cited.Label, StringComparer.OrdinalIgnoreCase);

            ComponentRows beforeRows = ComponentRows.Of(before, relabelled);
            ComponentRows afterRows = ComponentRows.Of(after, relabels: null);

            IEnumerable<string> labels = beforeRows.Labels
                .Concat(afterRows.Labels)
                .Concat(replacedByLabel.Select(group => group.Key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Order(NaturalLabelComparer.Instance);

            var relabelledFrom = relabelled.ToLookup(pair => pair.Value, pair => pair.Key, StringComparer.OrdinalIgnoreCase);

            foreach (string label in labels)
            {
                if (DraftChangeSummary.ComponentGroup(label, beforeRows, afterRows, relabelledFrom[label], replacedByLabel[label].Select(cited => cited.Name)) is DraftChangeGroup group)
                    groups.Add(group);
            }

            return new DraftChangeSummary(groups);
        }

        // ---- the components ---------------------------------------------------------------------

        private static DraftChangeGroup? ComponentGroup(
            string label,
            ComponentRows beforeRows,
            ComponentRows afterRows,
            IEnumerable<string> relabelledFrom,
            IEnumerable<string> replacedFiles)
        {
            BoardData before = beforeRows.For(label);
            BoardData after = afterRows.For(label);

            var lines = new List<string>();
            List<string> oldLabels = relabelledFrom.ToList();
            lines.AddRange(oldLabels.Select(old => $"Relabelled from [{old}]"));

            DraftComponentChange change = (before.Components.Count, after.Components.Count) switch
            {
                (0, > 0) => DraftComponentChange.Added,
                (> 0, 0) => DraftComponentChange.Removed,
                _ => DraftComponentChange.Changed,
            };

            if (change == DraftComponentChange.Added)
                lines.Add("New component");
            else if (change == DraftComponentChange.Removed)
                lines.Add("Removed");
            else
                lines.AddRange(DraftChangeSummary.ComponentRowLines(before, after, relabelled: oldLabels.Count > 0));

            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(PicturesOf(before), PicturesOf(after)), "picture", "pictures"));
            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(BaselinesOf(before), BaselinesOf(after)), "scope baseline", "scope baselines"));
            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(Section(before, highlights: true), Section(after, highlights: true)), "mark on the schematics", "marks on the schematics"));
            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(Section(before, files: true), Section(after, files: true)), "file", "files"));
            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(Section(before, links: true), Section(after, links: true)), "link", "links"));
            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(Section(before, pins: true), Section(after, pins: true)), "pin", "pins"));
            lines.AddRange(DraftChangeSummary.Counted(
                BoardDataDiffer.Compare(Section(before, replacements: true), Section(after, replacements: true)), "replacement part", "replacement parts"));

            lines.AddRange(replacedFiles.Distinct(StringComparer.OrdinalIgnoreCase).Select(name => $"File [{name}] replaced"));

            if (lines.Count == 0)
                return null;

            ComponentEntry? shown = after.Components.FirstOrDefault() ?? before.Components.FirstOrDefault();
            string name = shown is null
                ? string.Empty
                : DraftChangeSummary.FirstFilled(shown.FriendlyName, shown.TechnicalNameOrValue);

            return new DraftChangeGroup(
                label,
                name.Length == 0 ? label : $"{label} - {name}",
                change,
                lines);
        }

        // The component rows of one label that changed: "Description changed", with the region when
        // the label has more than one ("Description changed (PAL)"), and regions added or removed. A
        // relabelled component's new label is its "Relabelled from" line, not a field changed too.
        private static IEnumerable<string> ComponentRowLines(BoardData before, BoardData after, bool relabelled)
        {
            bool regional = before.Components.Concat(after.Components).Select(row => (row.Region ?? string.Empty).Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1;

            foreach (BoardRowChange row in BoardDataDiffer.Compare(Section(before, components: true), Section(after, components: true)))
            {
                string region = DraftChangeSummary.RegionOf(row, before, after);
                string where = regional && region.Length > 0 ? $" ({region})" : string.Empty;

                if (row.Kind == BoardRowChangeKind.Modified)
                {
                    List<string> fields = row.ChangedFields
                        .Where(field => !relabelled || field != nameof(ComponentEntry.BoardLabel))
                        .ToList();

                    if (fields.Count > 0)
                        yield return DraftChangeSummary.Capitalised(DraftChangeSummary.ChangedList(fields)) + where;

                    continue;
                }

                yield return row.Kind == BoardRowChangeKind.Added
                    ? $"Added for region [{region}]"
                    : $"Removed for region [{region}]";
            }
        }

        private static string RegionOf(BoardRowChange row, BoardData before, BoardData after) =>
            after.Components.Concat(before.Components)
                .Where(component => string.Equals(BoardDraftNaturalKeys.ForRow(component), row.NaturalKey, StringComparison.OrdinalIgnoreCase))
                .Select(component => (component.Region ?? string.Empty).Trim())
                .FirstOrDefault() ?? string.Empty;

        // "description and part number changed" - the fields in plain words.
        private static string ChangedList(IReadOnlyList<string> fields)
        {
            List<string> words = fields.Select(DraftChangeSummary.FieldWord).Distinct().ToList();

            string list = BoardDataChecks.ListOf(words);
            return list.Length == 0 ? "changed" : list + " changed";
        }

        private static string Capitalised(string text) =>
            text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text[1..];

        private static string FieldWord(string field) => field switch
        {
            nameof(ComponentEntry.BoardLabel) => "board label",
            nameof(ComponentEntry.FriendlyName) => "name",
            nameof(ComponentEntry.TechnicalNameOrValue) => "technical name or value",
            nameof(ComponentEntry.PartNumber) => "part number",
            nameof(ComponentEntry.Category) => "category",
            nameof(ComponentEntry.Region) => "region",
            nameof(ComponentEntry.Description) => "description",
            nameof(BoardSchematicEntry.SchematicName) => "name",
            nameof(BoardSchematicEntry.CadName) => "CAD name",
            nameof(BoardSchematicEntry.SchematicImageFile) => "image",
            nameof(BoardSchematicEntry.SchematicHighlightColor) or nameof(BoardSchematicEntry.SchematicHighlightOpacity) or
                nameof(BoardSchematicEntry.OppositeTraceHighlightColor) or nameof(BoardSchematicEntry.ThumbnailHighlightColor) or
                nameof(BoardSchematicEntry.ThumbnailHighlightOpacity) => "highlight colours",
            "File" => "file",
            "Url" => "address",
            "Name" => "name",
            "Contact" => "contact",
            "NameOrHandle" => "name",
            "SubCategory" => "sub-category",
            "DisplayName" => "name",
            "KiCadNetName" => "KiCad net",
            _ => DraftChangeSummary.Spaced(field),
        };

        // "TriggerLevelVolts" -> "trigger level volts".
        private static string Spaced(string field)
        {
            var words = new List<char>();

            for (int index = 0; index < field.Length; index++)
            {
                if (index > 0 && char.IsUpper(field[index]))
                    words.Add(' ');

                words.Add(char.ToLowerInvariant(field[index]));
            }

            return new string(words.ToArray());
        }

        // "2 pictures added", "1 picture changed", "3 pictures removed".
        private static IEnumerable<string> Counted(IReadOnlyList<BoardRowChange> changes, string one, string many)
        {
            foreach ((BoardRowChangeKind kind, string word) in new[]
                     {
                         (BoardRowChangeKind.Added, "added"),
                         (BoardRowChangeKind.Modified, "changed"),
                         (BoardRowChangeKind.Deleted, "removed"),
                     })
            {
                int count = changes.Count(change => change.Kind == kind);

                if (count > 0)
                    yield return $"{count.ToString(CultureInfo.InvariantCulture)} {(count == 1 ? one : many)} {word}";
            }
        }

        // ---- the board's own parts ----------------------------------------------------------------

        private static List<string> BoardLines(
            BoardData before,
            BoardData after,
            IReadOnlyList<KiCadCalibrationEntry> beforeCalibrations,
            IReadOnlyList<KiCadCalibrationEntry> afterCalibrations)
        {
            var lines = new List<string>();

            // Schematic images: a rename said as one (same image file, new name).
            IReadOnlyDictionary<string, string?> follow = SubmissionRowsBoard.FollowSchematics(before.Schematics, after.Schematics);
            var renamedTo = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach ((string from, string? to) in follow)
            {
                if (to is { Length: > 0 } && !string.Equals(from, to, StringComparison.Ordinal))
                {
                    lines.Add($"Picture [{from}] renamed [{to}]");
                    renamedTo.Add(to);
                }
            }

            foreach (BoardRowChange row in BoardDataDiffer.Compare(Section(before, schematics: true), Section(after, schematics: true)))
            {
                List<string> otherFields = row.ChangedFields.Where(field => field != nameof(BoardSchematicEntry.SchematicName)).ToList();

                if (row.Kind == BoardRowChangeKind.Modified && renamedTo.Contains(row.DisplayLabel) && otherFields.Count == 0)
                    continue;

                lines.Add(row.Kind switch
                {
                    BoardRowChangeKind.Added => $"Picture [{row.DisplayLabel}] added",
                    BoardRowChangeKind.Deleted => $"Picture [{row.DisplayLabel}] removed",
                    _ => $"Picture [{row.DisplayLabel}]: {DraftChangeSummary.ChangedList(otherFields.Count > 0 ? otherFields : row.ChangedFields)}",
                });
            }

            lines.AddRange(DraftChangeSummary.Named(BoardDataDiffer.Compare(Section(before, boardFiles: true), Section(after, boardFiles: true)), "Board file"));
            lines.AddRange(DraftChangeSummary.Named(BoardDataDiffer.Compare(Section(before, boardLinks: true), Section(after, boardLinks: true)), "Board link"));
            lines.AddRange(DraftChangeSummary.Named(BoardDataDiffer.Compare(Section(before, credits: true), Section(after, credits: true)), "Credit"));
            lines.AddRange(DraftChangeSummary.Named(BoardDataDiffer.Compare(Section(before, signals: true), Section(after, signals: true)), "Important signal"));

            lines.AddRange(DraftChangeSummary.CalibrationLines(beforeCalibrations, afterCalibrations, follow));

            return lines;
        }

        // "Credit [Software / Dennis] added" - one line per row, named as the differ names it.
        private static IEnumerable<string> Named(IReadOnlyList<BoardRowChange> changes, string what) =>
            changes.Select(row => row.Kind switch
            {
                BoardRowChangeKind.Added => $"{what} [{row.DisplayLabel}] added",
                BoardRowChangeKind.Deleted => $"{what} [{row.DisplayLabel}] removed",
                _ => $"{what} [{row.DisplayLabel}]: {DraftChangeSummary.ChangedList(row.ChangedFields)}",
            });

        // One line per schematic whose KiCad calibration was added, changed or removed - by its name,
        // followed through a rename (a calibration that only moved with its schematic is no change).
        private static IEnumerable<string> CalibrationLines(
            IReadOnlyList<KiCadCalibrationEntry> before,
            IReadOnlyList<KiCadCalibrationEntry> after,
            IReadOnlyDictionary<string, string?> follow)
        {
            Dictionary<string, KiCadCalibrationEntry> beforeByName = SchematicRename.FollowCalibrations(before, follow)
                .GroupBy(entry => (entry.SchematicName ?? string.Empty).Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            Dictionary<string, KiCadCalibrationEntry> afterByName = after
                .GroupBy(entry => (entry.SchematicName ?? string.Empty).Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            foreach (string name in beforeByName.Keys.Concat(afterByName.Keys).Distinct(StringComparer.OrdinalIgnoreCase).Order(StringComparer.OrdinalIgnoreCase))
            {
                bool had = beforeByName.TryGetValue(name, out KiCadCalibrationEntry? was);
                bool has = afterByName.TryGetValue(name, out KiCadCalibrationEntry? now);

                if (had && has && DraftChangeSummary.SameCalibration(was!, now!))
                    continue;

                yield return (had, has) switch
                {
                    (false, true) => $"KiCad calibration of [{name}] added",
                    (true, false) => $"KiCad calibration of [{name}] removed",
                    _ => $"KiCad calibration of [{name}] changed",
                };
            }
        }

        private static bool SameCalibration(KiCadCalibrationEntry a, KiCadCalibrationEntry b) =>
            string.Equals((a.CadName ?? string.Empty).Trim(), (b.CadName ?? string.Empty).Trim(), StringComparison.Ordinal) &&
            a.OffsetX.Equals(b.OffsetX) && a.OffsetY.Equals(b.OffsetY) &&
            a.ScaleX.Equals(b.ScaleX) && a.ScaleY.Equals(b.ScaleY) &&
            a.MirrorX == b.MirrorX && a.MirrorY == b.MirrorY;

        // ---- the files ----------------------------------------------------------------------------

        // The names both board files list with other bytes - a file replaced under its own name.
        private static List<string> ReplacedFiles(
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? published,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draft)
        {
            if (published is null || draft is null)
                return [];

            return draft.Values
                .Where(file => published.TryGetValue(file.Name, out BoardDocument.StoredFileRow? theirs) &&
                               !string.Equals(theirs.Sha256, file.Sha256, StringComparison.OrdinalIgnoreCase))
                .Select(file => file.Name)
                .Order(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // "KiCad data: 3 files added, 1 replaced" - the board's KiCad files, which no row cites.
        private static IEnumerable<string> KiCadLines(
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? published,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draft)
        {
            if (published is null && draft is null)
                yield break;

            Dictionary<string, BoardDocument.StoredFileRow> before = DraftChangeSummary.KiCadFiles(published);
            Dictionary<string, BoardDocument.StoredFileRow> after = DraftChangeSummary.KiCadFiles(draft);

            int added = after.Keys.Count(name => !before.ContainsKey(name));
            int removed = before.Keys.Count(name => !after.ContainsKey(name));
            int replaced = after.Count(file => before.TryGetValue(file.Key, out BoardDocument.StoredFileRow? was) &&
                                               !string.Equals(was.Sha256, file.Value.Sha256, StringComparison.OrdinalIgnoreCase));

            var parts = new List<string>();

            if (added > 0)
                parts.Add($"{added.ToString(CultureInfo.InvariantCulture)} {(added == 1 ? "file" : "files")} added");
            if (replaced > 0)
                parts.Add($"{replaced.ToString(CultureInfo.InvariantCulture)} replaced");
            if (removed > 0)
                parts.Add($"{removed.ToString(CultureInfo.InvariantCulture)} removed");

            if (parts.Count > 0)
                yield return "KiCad data: " + string.Join(", ", parts);
        }

        private static Dictionary<string, BoardDocument.StoredFileRow> KiCadFiles(IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? files) =>
            (files?.Values ?? [])
                .Where(DraftFiles.IsKiCad)
                .GroupBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

        // Every file a component's row cites, with its label.
        private static IEnumerable<(string Name, string Label)> ComponentFileNames(BoardData board) =>
            board.ComponentImages.Select(row => (row.File, row.BoardLabel))
                .Concat(board.ComponentLocalFiles.Select(row => (row.File, row.BoardLabel)))
                .Where(cited => !string.IsNullOrWhiteSpace(cited.File) && !string.IsNullOrWhiteSpace(cited.BoardLabel))
                .Select(cited => (BoardFileNames.Normalise(cited.File), cited.BoardLabel.Trim()));

        // ---- boards of one section ----------------------------------------------------------------

        private static BoardData NoBoard() => new() { HardwareName = string.Empty, BoardName = string.Empty };

        private static BoardData Section(
            BoardData board,
            bool components = false,
            bool highlights = false,
            bool files = false,
            bool links = false,
            bool schematics = false,
            bool boardFiles = false,
            bool boardLinks = false,
            bool credits = false,
            bool signals = false,
            bool pins = false,
            bool replacements = false) => new()
        {
            RevisionDate = string.Empty,
            HardwareName = string.Empty,
            BoardName = string.Empty,
            Components = components ? board.Components : [],

            // Pictures and baselines are compared apart (PicturesOf, BaselinesOf).
            ComponentImages = [],
            ComponentHighlights = highlights ? board.ComponentHighlights : [],
            ComponentLocalFiles = files ? board.ComponentLocalFiles : [],
            ComponentLinks = links ? board.ComponentLinks : [],
            Schematics = schematics ? board.Schematics : [],
            BoardLocalFiles = boardFiles ? board.BoardLocalFiles : [],
            BoardLinks = boardLinks ? board.BoardLinks : [],
            Credits = credits ? board.Credits : [],
            KiCadImportantSignals = signals ? board.KiCadImportantSignals : [],
            ComponentPins = pins ? board.ComponentPins : [],
            ComponentReplacements = replacements ? board.ComponentReplacements : [],
        };

        private static BoardData PicturesOf(BoardData board) => new()
        {
            HardwareName = string.Empty,
            BoardName = string.Empty,
            ComponentImages = board.ComponentImages.Where(image => !ComponentImageKinds.IsScopeBaseline(image)).ToList(),
        };

        private static BoardData BaselinesOf(BoardData board) => new()
        {
            HardwareName = string.Empty,
            BoardName = string.Empty,
            ComponentImages = board.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline).ToList(),
        };

        private static string FirstFilled(params string?[] values) =>
            values.Select(value => value?.Trim() ?? string.Empty).FirstOrDefault(value => value.Length > 0) ?? string.Empty;

        // ###########################################################################################
        // A board's component rows by label (trimmed, capitals ignored) - so each component is
        // compared on its own. `relabels` (old -> new) files a relabelled component's rows and
        // highlights under its new label, as the draft holds them.
        // ###########################################################################################
        private sealed class ComponentRows
        {
            private readonly Dictionary<string, BoardData> thisByLabel = new(StringComparer.OrdinalIgnoreCase);

            public IEnumerable<string> Labels => this.thisByLabel.Keys;

            public BoardData For(string label) =>
                this.thisByLabel.TryGetValue(label, out BoardData? rows) ? rows : DraftChangeSummary.NoBoard();

            public static ComponentRows Of(BoardData board, IReadOnlyDictionary<string, string>? relabels)
            {
                var rows = new ComponentRows();

                string LabelOf(string? label)
                {
                    string trimmed = (label ?? string.Empty).Trim();
                    return relabels is not null && relabels.TryGetValue(trimmed, out string? to) ? to : trimmed;
                }

                foreach (ComponentEntry row in board.Components)
                    rows.Get(LabelOf(row.BoardLabel)).Components.Add(row);

                foreach (ComponentImageEntry row in board.ComponentImages)
                    rows.Get((row.BoardLabel ?? string.Empty).Trim()).ComponentImages.Add(row);

                foreach (ComponentHighlightEntry row in board.ComponentHighlights)
                    rows.Get(LabelOf(row.BoardLabel)).ComponentHighlights.Add(relabels is not null && LabelOf(row.BoardLabel) != (row.BoardLabel ?? string.Empty).Trim()
                        ? row.WithBoardLabel(LabelOf(row.BoardLabel))
                        : row);

                foreach (ComponentLocalFileEntry row in board.ComponentLocalFiles)
                    rows.Get((row.BoardLabel ?? string.Empty).Trim()).ComponentLocalFiles.Add(row);

                foreach (ComponentLinkEntry row in board.ComponentLinks)
                    rows.Get((row.BoardLabel ?? string.Empty).Trim()).ComponentLinks.Add(row);

                foreach (ComponentPinEntry row in board.ComponentPins)
                    rows.Get((row.BoardLabel ?? string.Empty).Trim()).ComponentPins.Add(row);

                foreach (ComponentReplacementEntry row in board.ComponentReplacements)
                    rows.Get((row.BoardLabel ?? string.Empty).Trim()).ComponentReplacements.Add(row);

                rows.thisByLabel.Remove(string.Empty);
                return rows;
            }

            private BoardData Get(string label)
            {
                if (!this.thisByLabel.TryGetValue(label, out BoardData? rows))
                {
                    rows = DraftChangeSummary.NoBoard();
                    this.thisByLabel[label] = rows;
                }

                return rows;
            }
        }
    }
}
