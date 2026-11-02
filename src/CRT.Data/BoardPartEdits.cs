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
    //   - A LIST IS WRITTEN IN THE ORDER THE PANE SHOWS IT, and the pane shows it as the board has it
    //     (owner request, 2026-10-30: "please have the entire line moving up/down, and not only within
    //     its own category ... I also want to be able to rearrange the categories"). Files, links and
    //     credits were kept grouped by category until then, a row moving only inside its group; now a
    //     row moves past any other (Moved), so a category is rearranged by moving its rows. The order
    //     is a change (owner decision, the same day - BoardRowOrder): Submit sends it.
    //   - An edit that changes nothing returns THE SAME board, which DraftSession takes as no step.
    //
    // PURE: takes a board, returns a board, touches no files.
    // ###########################################################################################
    public static class BoardPartEdits
    {
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

        // ###########################################################################################
        // MOVING A ROW (owner requests, 2026-10-29: "I need a possibility to reorder the list elements";
        // 2026-10-30: "the entire line moving up/down, and not only within its own category") - files,
        // links, credits and important signals. A row moves one place past the row next to it, whatever
        // its category. `rows` in the order the pane shows them; `by` -1 up, +1 down.
        // ###########################################################################################
        public static bool CanMove<T>(IReadOnlyList<T> rows, int index, int by)
        {
            ArgumentNullException.ThrowIfNull(rows);

            int target = index + by;

            return by is -1 or 1 &&
                index >= 0 && index < rows.Count &&
                target >= 0 && target < rows.Count;
        }

        // `rows` with the row at `index` moved one place by `by` - as they were when it cannot move
        // (CanMove).
        public static List<T> Moved<T>(IReadOnlyList<T> rows, int index, int by)
        {
            List<T> moved = rows.ToList();

            if (BoardPartEdits.CanMove(rows, index, by))
            {
                (moved[index], moved[index + by]) = (moved[index + by], moved[index]);
            }

            return moved;
        }

        // ###########################################################################################
        // A row with nothing in it - left out of the list written, and so not counted where the list
        // is (the pane's title over the part; code review, 2026-10-30: "Add credit" made "Credits (4)"
        // of a board with three, an empty row never written).
        // ###########################################################################################
        public static bool IsBlank(BoardLocalFileEntry row) =>
            string.IsNullOrWhiteSpace(row.Category) && string.IsNullOrWhiteSpace(row.Name) && string.IsNullOrWhiteSpace(row.File);

        public static bool IsBlank(BoardLinkEntry row) =>
            string.IsNullOrWhiteSpace(row.Category) && string.IsNullOrWhiteSpace(row.Name) && string.IsNullOrWhiteSpace(row.Url);

        public static bool IsBlank(CreditEntry row) =>
            string.IsNullOrWhiteSpace(row.Category) && string.IsNullOrWhiteSpace(row.SubCategory) &&
            string.IsNullOrWhiteSpace(row.NameOrHandle) && string.IsNullOrWhiteSpace(row.Contact);

        public static bool IsBlank(KiCadImportantSignalEntry row) =>
            string.IsNullOrWhiteSpace(row.DisplayName) && string.IsNullOrWhiteSpace(row.KiCadNetName);

        // ---- the board's files and links ----------------------------------------------------------

        public static BoardData WithBoardFiles(BoardData board, IReadOnlyList<BoardLocalFileEntry> files)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(files);

            List<BoardLocalFileEntry> rows = files
                .Select(row => new BoardLocalFileEntry
                {
                    Category = row.Category?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                })
                .Where(row => !BoardPartEdits.IsBlank(row))
                .ToList();

            return BoardPartEdits.SameRows(board.BoardLocalFiles, rows, row => string.Join('\u001F', row.Category, row.Name, row.File, row.Region))
                ? board
                : board.WithBoardSections(boardLocalFiles: rows);
        }

        public static BoardData WithBoardLinks(BoardData board, IReadOnlyList<BoardLinkEntry> links)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(links);

            List<BoardLinkEntry> rows = links
                .Select(row => new BoardLinkEntry
                {
                    Category = row.Category?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                })
                .Where(row => !BoardPartEdits.IsBlank(row))
                .ToList();

            return BoardPartEdits.SameRows(board.BoardLinks, rows, row => string.Join('\u001F', row.Category, row.Name, row.Url, row.Region))
                ? board
                : board.WithBoardSections(boardLinks: rows);
        }

        // ---- credits ------------------------------------------------------------------------------

        public static BoardData WithCredits(BoardData board, IReadOnlyList<CreditEntry> credits)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(credits);

            List<CreditEntry> rows = credits
                .Select(row => new CreditEntry
                {
                    Category = row.Category?.Trim() ?? string.Empty,
                    SubCategory = row.SubCategory?.Trim() ?? string.Empty,
                    NameOrHandle = row.NameOrHandle?.Trim() ?? string.Empty,
                    Contact = row.Contact?.Trim() ?? string.Empty,
                })
                .Where(row => !BoardPartEdits.IsBlank(row))
                .ToList();

            return BoardPartEdits.SameRows(board.Credits, rows, row => string.Join('\u001F', row.Category, row.SubCategory, row.NameOrHandle, row.Contact))
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
                .Where(row => !BoardPartEdits.IsBlank(row))
                .ToList();

            static string Key(KiCadImportantSignalEntry row) => string.Join('\u001F', row.DisplayName?.Trim(), row.KiCadNetName?.Trim());

            return board.KiCadImportantSignals.Select(Key).SequenceEqual(rows.Select(Key), StringComparer.Ordinal)
                ? board
                : board.WithBoardSections(kiCadImportantSignals: rows);
        }

        // ###########################################################################################
        // Whether `rows` are the board's `before` - in the same order, every value trimmed - so a list
        // the pane did not change is no change.
        // ###########################################################################################
        private static bool SameRows<T>(IEnumerable<T> before, IReadOnlyList<T> rows, Func<T, string> key)
        {
            List<string> shown = before
                .Select(key)
                .Select(BoardPartEdits.Trimmed)
                .ToList();

            return shown.SequenceEqual(rows.Select(key), StringComparer.Ordinal);
        }

        // A key's parts trimmed, as the rows written are.
        private static string Trimmed(string key) => string.Join('\u001F', key.Split('\u001F').Select(part => part.Trim()));
    }
}
