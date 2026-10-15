using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A small, thread-safe cache of values that each belong to one KEY (a set of files) as it was
    // at one STAMP (their write times) - at most `capacity` keys, the least recently used going
    // first. Storing a key's value at a new stamp drops its older one: a file that changed never
    // leaves its old parse behind.
    //
    // Made for the parsed KiCad projects (KiCadProjectLoader, code review 2026-10-13): their cache
    // was a static Dictionary that grew by one whole parsed board - tens of megabytes - for every
    // board shown and every time its files changed, and it was written from pool threads with no
    // lock at all. Main's board load keeps the last boards' schematic thumbnails in one too
    // (SchematicThumbnailSet), disposing what Store hands back once nothing shows it.
    // ###########################################################################################
    internal sealed class StampedLruCache<TValue>
    {
        private readonly object thisSync = new();
        private readonly int thisCapacity;

        // Most recently used first.
        private readonly LinkedList<(string Key, string Stamp, TValue Value)> thisOrder = new();
        private readonly Dictionary<string, LinkedListNode<(string Key, string Stamp, TValue Value)>> thisByKey;

        internal StampedLruCache(int capacity, StringComparer keyComparer)
        {
            this.thisCapacity = Math.Max(1, capacity);
            this.thisByKey = new Dictionary<string, LinkedListNode<(string Key, string Stamp, TValue Value)>>(keyComparer);
        }

        internal int Count
        {
            get
            {
                lock (this.thisSync)
                {
                    return this.thisByKey.Count;
                }
            }
        }

        // The value stored for `key` at exactly `stamp` - which then counts as just used.
        internal bool TryGet(string key, string stamp, out TValue value)
        {
            lock (this.thisSync)
            {
                if (this.thisByKey.TryGetValue(key, out var node) &&
                    string.Equals(node.Value.Stamp, stamp, StringComparison.Ordinal))
                {
                    this.thisOrder.Remove(node);
                    this.thisOrder.AddFirst(node);
                    value = node.Value.Value;
                    return true;
                }

                value = default!;
                return false;
            }
        }

        // Stores `value` for `key` at `stamp`, replacing whatever the key held; the least recently
        // used key goes when there are more than the capacity. Returns the values that left - the
        // one replaced and any evicted - for a caller whose values need disposing, once nothing
        // shows them any more.
        internal List<TValue> Store(string key, string stamp, TValue value)
        {
            var removed = new List<TValue>();

            lock (this.thisSync)
            {
                if (this.thisByKey.Remove(key, out var old))
                {
                    this.thisOrder.Remove(old);
                    removed.Add(old.Value.Value);
                }

                this.thisByKey[key] = this.thisOrder.AddFirst((key, stamp, value));

                while (this.thisByKey.Count > this.thisCapacity && this.thisOrder.Last is { } last)
                {
                    this.thisOrder.RemoveLast();
                    this.thisByKey.Remove(last.Value.Key);
                    removed.Add(last.Value.Value);
                }
            }

            return removed;
        }
    }
}
