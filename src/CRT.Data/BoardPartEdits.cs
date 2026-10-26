using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE BOARD PARTS CHANGE IN A DRAFT (CRT 4.0 step 2, slice 4a, 2026-10-26 - the plan: "Board >
    // Files and links, grouped by category ... Board > Credits: a list by category ... Board > Important
    // signals: pick nets from a searchable KiCad list").
    //
    // The component pane shows the board's own lists too - its files, its links, its credits and its
    // important signals - and every change to one is a transform of the draft's board, ONE step of its
    // history, as ComponentPaneEdits' are:
    //
    //   - A LIST is replaced as a whole: every value trimmed, a row with nothing in it left out.
    //   - Files, links and credits are kept GROUPED BY CATEGORY (Grouped - the categories in the order
    //     they first appear, a row's place inside its category kept), as the pane shows them. Order is
    //     no change the differ counts, and a list that differs only by its order returns THE SAME board,
    //     so opening and closing a part changes nothing.
    //   - An edit that changes nothing returns THE SAME board, which DraftSession takes as no step.
    //
    // PURE: takes a board, returns a board, touches no files.
    // ###########################################################################################
    public static class BoardPartEdits
    {
        // ###########################################################################################
        // `rows` grouped by their category - trimmed, capitals ignored - the categories in the order
        // they first appear and every row's place inside its category kept. What the pane lists, and
        // the order a list is written in.
        // ###########################################################################################
        public static List<T> Grouped<T>(IEnumerable<T> rows, Func<T, string?> categoryOf)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(categoryOf);

            var groups = new List<(string Key, List<T> Rows)>();

            foreach (T row in rows)
            {
                string key = categoryOf(row)?.Trim().ToUpperInvariant() ?? string.Empty;
                int index = groups.FindIndex(group => string.Equals(group.Key, key, StringComparison.Ordinal));

                if (index < 0)
                {
                    groups.Add((key, [row]));
                }
                else
                {
                    groups[index].Rows.Add(row);
                }
            }

            return groups.SelectMany(group => group.Rows).ToList();
        }

        // The categories of `rows`, once each whatever their capitals, in the order they first appear -
        // offered as a category is typed, so a new row joins a group rather than starting a twin of it.
        public static List<string> CategoriesOf<T>(IEnumerable<T> rows, Func<T, string?> categoryOf)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(categoryOf);

            return rows
                .Select(row => categoryOf(row)?.Trim() ?? string.Empty)
                .Where(category => category.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ---- the board's files and links ----------------------------------------------------------

        public static BoardData WithBoardFiles(BoardData board, IReadOnlyList<BoardLocalFileEntry> files)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(files);

            List<BoardLocalFileEntry> rows = BoardPartEdits.Grouped(
                files
                    .Select(row => new BoardLocalFileEntry
                    {
                        Category = row.Category?.Trim() ?? string.Empty,
                        Name = row.Name?.Trim() ?? string.Empty,
                        File = row.File?.Trim() ?? string.Empty,
                    })
                    .Where(row => row.Category.Length > 0 || row.Name.Length > 0 || row.File.Length > 0),
                row => row.Category);

            return BoardPartEdits.SameAsGrouped(board.BoardLocalFiles, rows, row => row.Category, row => string.Join('\u001F', row.Category, row.Name, row.File))
                ? board
                : board.WithBoardSections(boardLocalFiles: rows);
        }

        public static BoardData WithBoardLinks(BoardData board, IReadOnlyList<BoardLinkEntry> links)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(links);

            List<BoardLinkEntry> rows = BoardPartEdits.Grouped(
                links
                    .Select(row => new BoardLinkEntry
                    {
                        Category = row.Category?.Trim() ?? string.Empty,
                        Name = row.Name?.Trim() ?? string.Empty,
                        Url = row.Url?.Trim() ?? string.Empty,
                    })
                    .Where(row => row.Category.Length > 0 || row.Name.Length > 0 || row.Url.Length > 0),
                row => row.Category);

            return BoardPartEdits.SameAsGrouped(board.BoardLinks, rows, row => row.Category, row => string.Join('\u001F', row.Category, row.Name, row.Url))
                ? board
                : board.WithBoardSections(boardLinks: rows);
        }

        // ---- credits ------------------------------------------------------------------------------

        public static BoardData WithCredits(BoardData board, IReadOnlyList<CreditEntry> credits)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(credits);

            List<CreditEntry> rows = BoardPartEdits.Grouped(
                credits
                    .Select(row => new CreditEntry
                    {
                        Category = row.Category?.Trim() ?? string.Empty,
                        SubCategory = row.SubCategory?.Trim() ?? string.Empty,
                        NameOrHandle = row.NameOrHandle?.Trim() ?? string.Empty,
                        Contact = row.Contact?.Trim() ?? string.Empty,
                    })
                    .Where(row => row.Category.Length > 0 || row.SubCategory.Length > 0 || row.NameOrHandle.Length > 0 || row.Contact.Length > 0),
                row => row.Category);

            return BoardPartEdits.SameAsGrouped(board.Credits, rows, row => row.Category, row => string.Join('\u001F', row.Category, row.SubCategory, row.NameOrHandle, row.Contact))
                ? board
                : board.WithBoardSections(credits: rows);
        }

        // ---- the pictures' colours (slice 4b) -------------------------------------------------------

        // ###########################################################################################
        // The board with each picture named in `colours` - by its name AND its image file, as two may
        // share a name (the "Schematic images" window's rule) - drawn in its new colours; every other
        // column of the picture, and every other picture, as it was. The same board when nothing
        // changed.
        // ###########################################################################################
        public static BoardData WithPictureColours(BoardData board, IReadOnlyList<(string Name, string File, PictureColours Colours)> colours)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(colours);

            bool changed = false;
            var pictures = colours.ToList();

            List<BoardSchematicEntry> schematics = board.Schematics
                .Select(row =>
                {
                    int index = pictures.FindIndex(picture =>
                        string.Equals(picture.Name?.Trim(), row.SchematicName?.Trim(), StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(picture.File?.Trim(), row.SchematicImageFile?.Trim(), StringComparison.OrdinalIgnoreCase));

                    if (index < 0)
                    {
                        return row;
                    }

                    PictureColours wanted = pictures[index].Colours.Trimmed();

                    if (wanted == row.ReadColours().Trimmed())
                    {
                        return row;
                    }

                    changed = true;
                    return row.WithColours(wanted);
                })
                .ToList();

            return changed ? board.WithSchematics(schematics) : board;
        }

        // ---- important signals --------------------------------------------------------------------

        // ###########################################################################################
        // The board's important signals replaced by `signals`, in their order - a signal is its
        // display name PLUS its KiCad net (the table's identity), so two rows may share either.
        // ###########################################################################################
        public static BoardData WithImportantSignals(BoardData board, IReadOnlyList<KiCadImportantSignalEntry> signals)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(signals);

            List<KiCadImportantSignalEntry> rows = signals
                .Select(row => new KiCadImportantSignalEntry
                {
                    DisplayName = row.DisplayName?.Trim() ?? string.Empty,
                    KiCadNetName = row.KiCadNetName?.Trim() ?? string.Empty,
                })
                .Where(row => row.DisplayName.Length > 0 || row.KiCadNetName.Length > 0)
                .ToList();

            static string Key(KiCadImportantSignalEntry row) => string.Join('\u001F', row.DisplayName?.Trim(), row.KiCadNetName?.Trim());

            return board.KiCadImportantSignals.Select(Key).SequenceEqual(rows.Select(Key), StringComparer.Ordinal)
                ? board
                : board.WithBoardSections(kiCadImportantSignals: rows);
        }

        // ###########################################################################################
        // Whether `rows` are the board's `before` as the pane shows them - grouped, every value trimmed
        // - so a list the pane did not change is no change, however the board had it ordered.
        // ###########################################################################################
        private static bool SameAsGrouped<T>(IEnumerable<T> before, IReadOnlyList<T> rows, Func<T, string?> categoryOf, Func<T, string> key)
        {
            List<string> shown = BoardPartEdits.Grouped(before, categoryOf)
                .Select(key)
                .Select(BoardPartEdits.Trimmed)
                .ToList();

            return shown.SequenceEqual(rows.Select(key), StringComparer.Ordinal);
        }

        // A key's parts trimmed, as the rows written are.
        private static string Trimmed(string key) => string.Join('\u001F', key.Split('\u001F').Select(part => part.Trim()));
    }
}
