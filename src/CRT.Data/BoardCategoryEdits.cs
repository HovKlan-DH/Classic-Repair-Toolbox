using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // One category of a board's components as "Board: Categories" lists it. A blank Name is the
    // components with no category, which the checks refuse at submit.
    public sealed record BoardCategoryRow(string Name, int Components);

    // ###########################################################################################
    // THE BOARD'S COMPONENT CATEGORIES (owner request, 2026-11-01: "a page where the categories can be
    // managed, meaning that it should be possible to rename, add and delete. No components must have
    // an empty category").
    //
    // NO LIST OF THEM IS STORED (owner decision, 2026-11-01: "kept until first used"): a category is
    // the word its components carry, so renaming and deleting are changes of those components' rows,
    // and one nothing uses is the page's to hold (ComponentPane.RegionsAndCategories.cs) until something does.
    // Rows stay where they stand - only the word changes (ComponentPlacement's "an edited one stays
    // put"), so no order changes with it.
    //
    // PURE: takes a board, returns a board, touches no files. The same board when nothing changes.
    // ###########################################################################################
    public static class BoardCategoryEdits
    {
        // ###########################################################################################
        // Every category with how many components carry it - each component once, however many regions
        // it has a row in. In first-seen order, spelt as first seen; the components with none first,
        // when there are any, as that is what needs doing.
        // ###########################################################################################
        public static List<BoardCategoryRow> Of(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            var names = new List<string>();
            var labels = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (ComponentEntry row in board.Components)
            {
                string label = row.BoardLabel?.Trim() ?? string.Empty;

                if (label.Length == 0)
                {
                    continue;
                }

                string category = row.Category?.Trim() ?? string.Empty;

                if (!labels.TryGetValue(category, out HashSet<string>? carrying))
                {
                    carrying = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    labels[category] = carrying;
                    names.Add(category);
                }

                carrying.Add(label);
            }

            return names
                .OrderBy(name => name.Length == 0 ? 0 : 1)
                .Select(name => new BoardCategoryRow(name, labels[name].Count))
                .ToList();
        }

        // ###########################################################################################
        // Every component of the category `from` ("" for those with none) under `to`. A name another
        // category has already, in any capitals, joins it under that spelling - "capacitor" never
        // stands beside "Capacitor". The same board when `to` is blank or says what `from` does.
        // ###########################################################################################
        public static BoardData Renamed(BoardData board, string? from, string? to)
        {
            ArgumentNullException.ThrowIfNull(board);

            string source = from?.Trim() ?? string.Empty;
            string wanted = BoardCategoryEdits.Spelling(board, source, to);

            if (wanted.Length == 0 || string.Equals(source, wanted, StringComparison.Ordinal))
            {
                return board;
            }

            bool changed = false;

            List<ComponentEntry> components = board.Components
                .Select(row =>
                {
                    if (!string.Equals(row.Category?.Trim() ?? string.Empty, source, StringComparison.OrdinalIgnoreCase) ||
                        string.IsNullOrWhiteSpace(row.BoardLabel))
                    {
                        return row;
                    }

                    changed = true;
                    return row.WithCategory(wanted);
                })
                .ToList();

            return changed ? board.WithComponentSections(components: components) : board;
        }

        // ###########################################################################################
        // How `to` is written once it lands: as a category it joins spells it, else as it was typed.
        // A change of capitals alone, from `from` itself, is taken as typed.
        // ###########################################################################################
        public static string Spelling(BoardData board, string? from, string? to)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = to?.Trim() ?? string.Empty;
            string source = from?.Trim() ?? string.Empty;

            if (string.Equals(wanted, source, StringComparison.OrdinalIgnoreCase))
            {
                return wanted;
            }

            return BoardCategoryEdits.Of(board)
                .Select(row => row.Name)
                .FirstOrDefault(name => string.Equals(name, wanted, StringComparison.OrdinalIgnoreCase)) ?? wanted;
        }
    }
}
