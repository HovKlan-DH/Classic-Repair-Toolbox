using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // KEEPING A SHOWN LIST'S ITEMS WHILE ITS CONTENT CHANGES - two lists found that replacing their
    // items breaks what is on screen (2026-10-10):
    //
    //   - a ComboBox whose items are replaced lets go of its selection, and its binding writes
    //     "nothing" back (the "Schematic images" window's CAD name drop-downs - Sync);
    //   - a row replaced between a click's press and its release throws the click away (the Drafts
    //     tab's list, read again whenever CRT's window comes to the front - KeepUnchanged).
    //
    // Pure list work, kept out of the tabs (code review, 2026-10-10) and tested on plain lists.
    // ###########################################################################################
    internal static class InPlaceList
    {
        // ###########################################################################################
        // Makes `items` show `wanted`, in order, changing it in place: an item already showing a
        // wanted entry keeps its object, a missing one is made (`create`), an unwanted one is
        // removed - never `keep`, the item on screen as selected, while a wanted entry still shows
        // it: a new item is inserted before it rather than it being removed and made again.
        // ###########################################################################################
        public static void Sync<TItem, TWanted>(
            IList<TItem> items,
            IReadOnlyList<TWanted> wanted,
            Func<TItem, TWanted, bool> shows,
            Func<TWanted, TItem> create,
            TItem? keep = null)
            where TItem : class
        {
            for (int i = 0; i < wanted.Count; i++)
            {
                int found = -1;

                for (int j = i; j < items.Count; j++)
                {
                    if (shows(items[j], wanted[i]))
                    {
                        found = j;
                        break;
                    }
                }

                int keptAt = InPlaceList.IndexOfObject(items, keep);
                bool wouldDropKept = found > i && keptAt >= i && keptAt < found;

                if (found < 0 || wouldDropKept)
                {
                    items.Insert(i, create(wanted[i]));
                    continue;
                }

                while (found > i)
                {
                    items.RemoveAt(i);
                    found--;
                }
            }

            while (items.Count > wanted.Count)
            {
                items.RemoveAt(items.Count - 1);
            }
        }

        // ###########################################################################################
        // Puts `fresh` on `rows`, keeping every row that `showsTheSame` as its fresh counterpart at
        // the same place - buttons and all - and replacing only the ones that changed; the list
        // grows or shrinks at its end.
        // ###########################################################################################
        public static void KeepUnchanged<T>(IList<T> rows, IReadOnlyList<T> fresh, Func<T, T, bool> showsTheSame)
        {
            for (int i = 0; i < fresh.Count; i++)
            {
                if (i >= rows.Count)
                {
                    rows.Add(fresh[i]);
                }
                else if (!showsTheSame(rows[i], fresh[i]))
                {
                    rows[i] = fresh[i];
                }
            }

            while (rows.Count > fresh.Count)
            {
                rows.RemoveAt(rows.Count - 1);
            }
        }

        private static int IndexOfObject<TItem>(IList<TItem> items, TItem? item)
            where TItem : class
        {
            if (item is null)
            {
                return -1;
            }

            for (int i = 0; i < items.Count; i++)
            {
                if (ReferenceEquals(items[i], item))
                {
                    return i;
                }
            }

            return -1;
        }
    }
}
