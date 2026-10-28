using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHETHER A DRAFT HOLDS ITS ROWS IN ANOTHER ORDER THAN THE PUBLISHED BOARD - the one change
    // BoardDataDiffer does not count (making order a counted change is an owner decision not yet
    // taken), but a change the contributor MADE: the pictures' Move up and Move down, a row dragged in
    // the table. A draft "made only to look" holding nothing else must not be thrown away with it
    // (code review, 2026-10-28: a picture moved, the pane closed, and the draft went without a word).
    //
    // Every sheet of BoardData, read by reflection so a new sheet is covered without a change here;
    // rows are told apart by BoardDraftNaturalKeys.ForRow, the pairing rule every comparison uses.
    // Only the rows BOTH boards have are compared - an added or removed row is the differ's to count.
    // ###########################################################################################
    public static class BoardRowOrder
    {
        private static readonly PropertyInfo[] Sheets = typeof(BoardData)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.PropertyType.IsGenericType &&
                               property.PropertyType.GetGenericTypeDefinition() == typeof(List<>) &&
                               property.PropertyType.GetGenericArguments()[0].IsClass &&
                               property.PropertyType.GetGenericArguments()[0] != typeof(string))
            .ToArray();

        public static bool DiffersFrom(BoardData published, BoardData draft)
        {
            ArgumentNullException.ThrowIfNull(published);
            ArgumentNullException.ThrowIfNull(draft);

            foreach (PropertyInfo sheet in Sheets)
            {
                if (KeysOf(sheet.GetValue(published)) is not { } before || KeysOf(sheet.GetValue(draft)) is not { } after)
                {
                    continue;
                }

                var shared = new HashSet<string>(before, StringComparer.OrdinalIgnoreCase);
                shared.IntersectWith(after);

                if (!before.Where(shared.Contains).SequenceEqual(after.Where(shared.Contains), StringComparer.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        // A sheet's row keys in order, the first of rows sharing a key only - or null for a list that
        // holds no board rows (ForRow knows no such row).
        private static List<string>? KeysOf(object? rows)
        {
            var keys = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (object? row in rows as IEnumerable ?? Array.Empty<object>())
            {
                if (row is null)
                {
                    continue;
                }

                string key;

                try
                {
                    key = BoardDraftNaturalKeys.ForRow(row);
                }
                catch (NotSupportedException)
                {
                    return null;
                }

                if (seen.Add(key))
                {
                    keys.Add(key);
                }
            }

            return keys;
        }
    }
}
