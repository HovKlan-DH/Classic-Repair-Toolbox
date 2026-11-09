using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // SEVERAL COMPONENTS EDITED AS ONE (owner request, 2026-11-08: "I want to be able to mark multiple components and when
    // doing that I am bulk editing whatever data I changes. E.g. selecting components "R21" and "R22" ... if I e.g. change
    // "Category" or "Friendly name" fields, then both components will get this new value ... If I go to "Replacement
    // parts", then I think it should remove all replacements line which are not part of all of them"; the proposal agreed
    // the same day: "OK for your 3 proposals of editing multiple components - do it").
    //
    // What the component pane shows and changes for several components, pure, so its tests need no pane:
    //
    //   - A FIELD: the value every one has, or none when they differ (ValuesOf says each one's); typed, it is set on every
    //     one (WithField). The label is never one - a label names ONE component.
    //   - A LIST - replacement parts, pins, pictures, files, links (the Section of each): the rows EVERY one has, compared
    //     by all they hold but the label (CommonRows). A common row changed or removed is changed or removed on every one,
    //     a row added is added to every one, and the rows only some have stay as they are (WithCommonRows; RowsOfSome
    //     lists them, with whose they are).
    //   - "MAKE ALL LIKE": one component's field values (FieldsMadeLike), or one page's rows (MadeLike), given to the
    //     others - one explicit act, never a mode that silently changes what later edits do.
    //
    // Each is one board in and one board out, built on ComponentPaneEdits - the very rules an edit of one component
    // follows, regions included - so the pane makes it ONE step of the draft's history, and a change that changes nothing
    // returns the same board, no step at all. Oscilloscope baselines are not here: a baseline is measured on one component.
    // ###########################################################################################
    public static class ComponentBulk
    {
        // ###########################################################################################
        // One of a component's lists, as the component pane reads and writes it for ONE component - and how two of its
        // rows are told apart, the label left out.
        // ###########################################################################################
        public sealed class Section<T>
        {
            private readonly Func<BoardData, string, string?, List<T>> thisRead;
            private readonly Func<BoardData, string, string?, IReadOnlyList<T>, BoardData> thisWrite;
            private readonly Func<T, string> thisKey;
            private readonly Func<T, string, T> thisFiledAs;
            private readonly Func<BoardData, string, List<T>> thisReadAll;
            private readonly Func<BoardData, string, IReadOnlyList<T>, BoardData> thisWriteAll;
            private readonly Func<BoardData, T, string, string?, T>? thisFiledNew;

            internal Section(
                Func<BoardData, string, string?, List<T>> read,
                Func<BoardData, string, string?, IReadOnlyList<T>, BoardData> write,
                Func<T, string> key,
                Func<T, string, T> filedAs,
                Func<BoardData, string, List<T>> readAll,
                Func<BoardData, string, IReadOnlyList<T>, BoardData> writeAll,
                Func<BoardData, T, string, string?, T>? filedNew = null)
            {
                this.thisRead = read;
                this.thisWrite = write;
                this.thisKey = key;
                this.thisFiledAs = filedAs;
                this.thisReadAll = readAll;
                this.thisWriteAll = writeAll;
                this.thisFiledNew = filedNew;
            }

            // A row ADDED to every one, filed for `label` of `board`, CRT showing `region` - FiledAs, unless the list files a
            // new row by the component it goes to (a pin joins that component's own set).
            public T FiledNew(BoardData board, T row, string label, string? region) =>
                this.thisFiledNew is { } filed ? filed(board, row, label, region) : this.thisFiledAs(row, label);

            public List<T> Read(BoardData board, string label, string? region) => this.thisRead(board, label, region);

            public BoardData Write(BoardData board, string label, string? region, IReadOnlyList<T> rows) => this.thisWrite(board, label, region, rows);

            // The component's rows of this list in EVERY region, each as it is filed - and all of them replaced (MadeLike).
            public List<T> ReadAll(BoardData board, string label) => this.thisReadAll(board, label);

            public BoardData WriteAll(BoardData board, string label, IReadOnlyList<T> rows) => this.thisWriteAll(board, label, rows);

            // What the row holds, its label left out.
            public string Key(T row) => this.thisKey(row);

            public T FiledAs(T row, string label) => this.thisFiledAs(row, label);
        }

        public static Section<ComponentReplacementEntry> Replacements { get; } = new(
            ComponentPaneEdits.ReplacementsOf,
            ComponentPaneEdits.WithReplacements,
            row => ComponentPaneEdits.KeyOf(row.WithBoardLabel(string.Empty)),
            (row, label) => row.WithBoardLabel(label),
            (board, label) => board.ComponentReplacements.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).ToList(),
            (board, label, rows) => ComponentBulk.Replaced(board.ComponentReplacements, row => row.BoardLabel, label, rows, list => board.WithComponentSections(componentReplacements: list)));

        public static Section<ComponentPinEntry> Pins { get; } = new(
            ComponentPaneEdits.PinsOf,
            ComponentPaneEdits.WithPins,
            row => ComponentPaneEdits.KeyOf(row.WithBoardLabel(string.Empty)),
            (row, label) => row.WithBoardLabel(label),
            (board, label) => board.ComponentPins.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).ToList(),
            (board, label, rows) => ComponentBulk.Replaced(board.ComponentPins, row => row.BoardLabel, label, rows, list => board.WithComponentSections(componentPins: list)),

            // A pin added joins each one's own set (ComponentPaneEdits.RegionForNewPin; code review, 2026-11-09: the first one's
            // region went to all - an every-region set got a pin of one region, and the other way round).
            (board, row, label, region) => row.WithBoardLabel(label).WithRegion(ComponentPaneEdits.RegionForNewPin(board, label, region)));

        // A picture is a component image that is no scope baseline (ComponentImageKinds) - the baselines stay as they are.
        public static Section<ComponentImageEntry> Pictures { get; } = new(
            ComponentPaneEdits.PicturesOf,
            ComponentPaneEdits.WithPictures,
            row => ComponentPaneEdits.KeyOf(row.WithBoardLabel(string.Empty)),
            (row, label) => row.WithBoardLabel(label),
            (board, label) => board.ComponentImages.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label) && !ComponentImageKinds.IsScopeBaseline(row)).ToList(),
            (board, label, rows) => ComponentBulk.Replaced(
                board.ComponentImages,
                row => ComponentImageKinds.IsScopeBaseline(row) ? string.Empty : row.BoardLabel,
                label,
                rows,
                list => board.WithComponentSections(componentImages: list)));

        // A component's files and links are its whole lists, whatever region CRT shows (as its pages list them).
        public static Section<ComponentLocalFileEntry> Files { get; } = new(
            (board, label, _) => ComponentPaneEdits.AllFilesOf(board, label),
            (board, label, _, rows) => ComponentPaneEdits.WithAllFiles(board, label, rows),
            row => ComponentPaneEdits.KeyOf(row.WithBoardLabel(string.Empty)),
            (row, label) => row.WithBoardLabel(label),
            ComponentPaneEdits.AllFilesOf,
            ComponentPaneEdits.WithAllFiles);

        public static Section<ComponentLinkEntry> Links { get; } = new(
            (board, label, _) => ComponentPaneEdits.AllLinksOf(board, label),
            (board, label, _, rows) => ComponentPaneEdits.WithAllLinks(board, label, rows),
            row => ComponentPaneEdits.KeyOf(row.WithBoardLabel(string.Empty)),
            (row, label) => row.WithBoardLabel(label),
            ComponentPaneEdits.AllLinksOf,
            ComponentPaneEdits.WithAllLinks);

        // ###########################################################################################
        // `rows` - the component `label`'s rows of a list in every region - in place of the ones `all` has under its label
        // (`labelOf` empty for a row that is not of the list), the places they stood at taken in turn. `with` puts the list
        // into the board.
        // ###########################################################################################
        private static BoardData Replaced<T>(List<T> all, Func<T, string> labelOf, string label, IReadOnlyList<T> rows, Func<List<T>, BoardData> with)
            where T : class =>
            with(ComponentPaneEdits.ReplaceInScope(all, row => ComponentPaneEdits.IsOfComponent(labelOf(row), label), rows));

        // ---- fields ---------------------------------------------------------------------------------

        // Each one's value of `field`, in the order of `labels` - its row as shown in `region`, empty for one the board lacks.
        public static IReadOnlyList<string> ValuesOf(BoardData board, IReadOnlyList<string> labels, string? region, ComponentField field)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(labels);

            return labels
                .Select(label => ComponentPaneEdits.RowShown(board, label, region) is { } row ? ComponentPaneEdits.ValueOf(row, field) : string.Empty)
                .ToList();
        }

        // ###########################################################################################
        // The value every one of `labels` has - null when they differ. Compared as an edit keeps them (Normalised), a category
        // without regard to capitals, as CRT groups categories everywhere (code review, 2026-11-09: "10k 1%" and "10k 1% " read
        // as two values, the box empty and its list naming the same text twice; "IC" and "ic" as two categories).
        // ###########################################################################################
        public static string? CommonValue(BoardData board, IReadOnlyList<string> labels, string? region, ComponentField field)
        {
            IReadOnlyList<string> values = ComponentBulk.ValuesOf(board, labels, region, field);

            if (values.Count == 0)
            {
                return null;
            }

            StringComparison comparison = field == ComponentField.Category ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            string first = ComponentPaneEdits.Normalised(field, values[0]);

            return values.All(value => string.Equals(ComponentPaneEdits.Normalised(field, value), first, comparison)) ? values[0] : null;
        }

        // `field` set to `value` on every one of `labels`, each as an edit of that one component would set it.
        public static BoardData WithField(BoardData board, IReadOnlyList<string> labels, string? region, ComponentField field, string? value)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(labels);

            return labels.Aggregate(board, (current, label) => ComponentPaneEdits.WithField(current, label, region, field, value));
        }

        // ###########################################################################################
        // Every field of "Basics" of `source` given to the others of `labels` - its label excepted. Nothing for a source the
        // board lacks. NOT the package, which "Pinout" shows and changes (code review, 2026-11-09: "Make all like" on "Basics",
        // whose note says only that page changes, gave the others its package - their pinout, part drawing and pin checks
        // with it).
        // ###########################################################################################
        public static BoardData FieldsMadeLike(BoardData board, string source, IReadOnlyList<string> labels, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(labels);

            if (ComponentPaneEdits.RowShown(board, source, region) is not { } row)
            {
                return board;
            }

            List<string> others = ComponentBulk.Others(source, labels);

            return ComponentBulk.BasicsFields
                .Aggregate(board, (current, field) => ComponentBulk.WithField(current, others, region, field, ComponentPaneEdits.ValueOf(row, field)));
        }

        // The fields "Basics" shows - every one but the package.
        public static IReadOnlyList<ComponentField> BasicsFields { get; } =
            Enum.GetValues<ComponentField>().Where(field => field != ComponentField.Package).ToList();

        // ---- lists ----------------------------------------------------------------------------------

        // ###########################################################################################
        // The rows EVERY one of `labels` has, in the first one's order - by all they hold but the label. A row the first
        // has twice counts twice only where every other has it twice too.
        // ###########################################################################################
        public static List<T> CommonRows<T>(BoardData board, IReadOnlyList<string> labels, string? region, Section<T> section)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(labels);
            ArgumentNullException.ThrowIfNull(section);

            if (labels.Count == 0)
            {
                return [];
            }

            List<Dictionary<string, int>> others = labels.Skip(1).Select(label => ComponentBulk.Counts(section.Read(board, label, region), section)).ToList();
            var common = new List<T>();

            foreach (T row in section.Read(board, labels[0], region))
            {
                string key = section.Key(row);

                if (others.All(counts => counts.GetValueOrDefault(key) > 0))
                {
                    others.ForEach(counts => counts[key]--);
                    common.Add(row);
                }
            }

            return common;
        }

        // The rows only some of `labels` have - what is left of each once the common rows are taken out - with whose they are.
        public static List<(string Label, T Row)> RowsOfSome<T>(BoardData board, IReadOnlyList<string> labels, string? region, Section<T> section)
        {
            Dictionary<string, int> common = ComponentBulk.Counts(ComponentBulk.CommonRows(board, labels, region, section), section);
            var some = new List<(string Label, T Row)>();

            foreach (string label in labels)
            {
                var left = new Dictionary<string, int>(common, StringComparer.Ordinal);

                foreach (T row in section.Read(board, label, region))
                {
                    string key = section.Key(row);

                    if (left.GetValueOrDefault(key) > 0)
                    {
                        left[key]--;
                    }
                    else
                    {
                        some.Add((label, row));
                    }
                }
            }

            return some;
        }

        // ###########################################################################################
        // THE COMMON ROWS CHANGED, ON EVERY ONE OF `labels`. `before` is the common rows as they were shown, `after` as they
        // are now, each with the place in `before` it was shown at - -1 for a row added. On each component a row that is
        // one of `before` (by what it holds, in order) becomes what that one is now, or goes when it is gone; the rows
        // added come after its own; a row only it has stays where it is.
        // ###########################################################################################
        public static BoardData WithCommonRows<T>(
            BoardData board,
            IReadOnlyList<string> labels,
            string? region,
            Section<T> section,
            IReadOnlyList<T> before,
            IReadOnlyList<(T Row, int From)> after)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(labels);
            ArgumentNullException.ThrowIfNull(section);
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);

            foreach (string label in labels)
            {
                List<T> own = section.Read(board, label, region);
                int[] shownAt = ComponentBulk.ShownAt(own, before, section);
                var rows = new List<T>();

                for (int index = 0; index < own.Count; index++)
                {
                    if (shownAt[index] < 0)
                    {
                        rows.Add(own[index]);
                        continue;
                    }

                    rows.AddRange(after.Where(each => each.From == shownAt[index]).Select(each => section.FiledAs(each.Row, label)));
                }

                BoardData now = board;
                rows.AddRange(after.Where(each => each.From < 0).Select(each => section.FiledNew(now, each.Row, label, region)));
                board = section.Write(board, label, region, rows);
            }

            return board;
        }

        // ###########################################################################################
        // `source`'s rows of `section` given to the others of `labels`, in place of theirs - IN EVERY REGION, each row filed as
        // `source` files it (second code review, 2026-11-09: read in the region shown and written over what the others showed
        // there, a list one files per region and another for every region mixed - U2's pins for every region became PAL's
        // alone, gone from NTSC, or NTSC listed both). So the page shows on each of them what it shows on `source`, whichever
        // region CRT shows. One already like it is left as it is; none changed, the same board.
        // ###########################################################################################
        public static BoardData MadeLike<T>(BoardData board, string source, IReadOnlyList<string> labels, Section<T> section)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(labels);
            ArgumentNullException.ThrowIfNull(section);

            List<T> rows = section.ReadAll(board, source);
            List<string> keys = rows.Select(section.Key).ToList();

            return ComponentBulk.Others(source, labels)
                .Aggregate(board, (current, label) =>
                    section.ReadAll(current, label).Select(section.Key).SequenceEqual(keys, StringComparer.Ordinal)
                        ? current
                        : section.WriteAll(current, label, rows.Select(row => section.FiledAs(row, label)).ToList()));
        }

        private static List<string> Others(string source, IReadOnlyList<string> labels) =>
            labels.Where(label => !ComponentPaneEdits.IsOfComponent(label, source)).ToList();

        private static Dictionary<string, int> Counts<T>(IEnumerable<T> rows, Section<T> section)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (T row in rows)
            {
                string key = section.Key(row);
                counts[key] = counts.GetValueOrDefault(key) + 1;
            }

            return counts;
        }

        // For each of `own`, the place in `before` it is - each of `before` matched once, to the first of `own` holding the same.
        private static int[] ShownAt<T>(List<T> own, IReadOnlyList<T> before, Section<T> section)
        {
            int[] shownAt = Enumerable.Repeat(-1, own.Count).ToArray();

            for (int at = 0; at < before.Count; at++)
            {
                string key = section.Key(before[at]);

                for (int index = 0; index < own.Count; index++)
                {
                    if (shownAt[index] < 0 && string.Equals(section.Key(own[index]), key, StringComparison.Ordinal))
                    {
                        shownAt[index] = at;
                        break;
                    }
                }
            }

            return shownAt;
        }
    }
}
