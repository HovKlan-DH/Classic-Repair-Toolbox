using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // KEEPING A SHOWN LIST'S ITEMS WHILE ITS CONTENT CHANGES - two lists found that replacing their
    // items breaks what is on screen (2026-10-10):
    //
    //   - a ComboBox whose items are replaced lets go of its selection, and its binding writes
    //     "nothing" back (the picture cards' "Associate KiCad file" drop-downs - Sync);
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
        //
        // A `keep` NO wanted entry shows any more is no item to keep (code review, 2026-10-30): the
        // component open renamed in the pane, every item after it was made again around it, and it was
        // let go of at the end all the same - the list's selection with it.
        //
        // `itemKey` and `wantedKey` (2026-11-16, performance walkthrough), when given: a key per item and per
        // entry, the same for any two that `shows` pairs (equal keys may still not pair; different keys never
        // do). With them a list sharing NO key with what is wanted - another board's components - is emptied
        // once and filled (one Reset, then the new items at its end), where every entry was looked for in every
        // item and put in at the top one by one; and an entry whose key no remaining item has is put in without
        // being looked for. The result is the same list, the same objects kept.
        // ###########################################################################################
        public static void Sync<TItem, TWanted>(
            IList<TItem> items,
            IReadOnlyList<TWanted> wanted,
            Func<TItem, TWanted, bool> shows,
            Func<TWanted, TItem> create,
            TItem? keep = null,
            Func<TItem, string>? itemKey = null,
            Func<TWanted, string>? wantedKey = null)
            where TItem : class
        {
            if (keep is not null && !InPlaceList.AnyShows(keep, wanted, shows))
            {
                keep = null;
            }

            // How many of the items not yet passed have each key - null without keys.
            Dictionary<string, int>? remaining = null;

            if (itemKey is not null && wantedKey is not null)
            {
                var wantedKeys = new HashSet<string>(StringComparer.Ordinal);

                for (int i = 0; i < wanted.Count; i++)
                {
                    wantedKeys.Add(wantedKey(wanted[i]));
                }

                remaining = new Dictionary<string, int>(StringComparer.Ordinal);
                bool anyShared = false;

                for (int i = 0; i < items.Count; i++)
                {
                    string key = itemKey(items[i]);
                    remaining[key] = remaining.GetValueOrDefault(key) + 1;
                    anyShared |= wantedKeys.Contains(key);
                }

                if (!anyShared)
                {
                    items.Clear();

                    for (int i = 0; i < wanted.Count; i++)
                    {
                        items.Add(create(wanted[i]));
                    }

                    return;
                }
            }

            for (int i = 0; i < wanted.Count; i++)
            {
                int found = -1;

                if (remaining is null || remaining.GetValueOrDefault(wantedKey!(wanted[i])) > 0)
                {
                    for (int j = i; j < items.Count; j++)
                    {
                        if (shows(items[j], wanted[i]))
                        {
                            found = j;
                            break;
                        }
                    }
                }

                // Where the kept item is matters only when items are about to be taken out above the one found.
                bool wouldDropKept = false;

                if (found > i && keep is not null)
                {
                    int keptAt = InPlaceList.IndexOfObject(items, keep);
                    wouldDropKept = keptAt >= i && keptAt < found;
                }

                if (found < 0 || wouldDropKept)
                {
                    items.Insert(i, create(wanted[i]));
                    continue;
                }

                while (found > i)
                {
                    InPlaceList.Passed(remaining, itemKey, items[i]);
                    items.RemoveAt(i);
                    found--;
                }

                InPlaceList.Passed(remaining, itemKey, items[i]);
            }

            while (items.Count > wanted.Count)
            {
                items.RemoveAt(items.Count - 1);
            }
        }

        // An item left behind - kept in its place or taken out - is no longer among those remaining.
        private static void Passed<TItem>(Dictionary<string, int>? remaining, Func<TItem, string>? itemKey, TItem item)
        {
            if (remaining is not null && itemKey is not null)
            {
                string key = itemKey(item);
                remaining[key] = remaining.GetValueOrDefault(key) - 1;
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

        private static bool AnyShows<TItem, TWanted>(TItem item, IReadOnlyList<TWanted> wanted, Func<TItem, TWanted, bool> shows)
        {
            for (int i = 0; i < wanted.Count; i++)
            {
                if (shows(item, wanted[i]))
                {
                    return true;
                }
            }

            return false;
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
