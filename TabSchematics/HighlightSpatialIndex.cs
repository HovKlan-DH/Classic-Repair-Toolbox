using Avalonia;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TabSchematics
{
    // ###########################################################################################
    // Spatial index for fast "visible rect" queries of many component highlight rectangles.
    // Uses a fixed-size grid in bitmap pixel coordinates and a stamp-based de-dupe per query.
    // ###########################################################################################
    public sealed class HighlightSpatialIndex
    {
        private readonly int _cellSize;
        private readonly Rect[] _rects;
        private readonly Dictionary<long, List<int>> _cells;
        private readonly int[] _seenStamp;
        private int _stamp;

        public int Count => this._rects.Length;

        public HighlightSpatialIndex(IReadOnlyList<Rect> rects, int cellSize = 512)
        {
            this._cellSize = Math.Max(32, cellSize);
//            this._rects = rects is Rect[] arr ? arr : [.. rects];
            this._rects = rects is Rect[] arr ? arr : rects.ToArray(); // .NET6 compliant
            this._cells = new Dictionary<long, List<int>>(capacity: Math.Max(16, this._rects.Length / 4));
            this._seenStamp = new int[this._rects.Length];
            this._stamp = 1;

            for (int i = 0; i < this._rects.Length; i++)
                this.AddToCells(i, this._rects[i]);
        }

        public Rect GetRect(int index) => this._rects[index];

        // ###########################################################################################
        // Queries all highlight indices that intersect the given pixelRect (bitmap pixel coordinates).
        // Results are appended to the provided list (which is cleared first).
        // ###########################################################################################
        public void Query(in Rect pixelRect, List<int> results)
        {
            results.Clear();

            if (this._rects.Length == 0 || pixelRect.Width <= 0 || pixelRect.Height <= 0)
                return;

            int stamp = this.NextStamp();

            int minCx = (int)Math.Floor(pixelRect.Left / this._cellSize);
            int maxCx = (int)Math.Floor(pixelRect.Right / this._cellSize);
            int minCy = (int)Math.Floor(pixelRect.Top / this._cellSize);
            int maxCy = (int)Math.Floor(pixelRect.Bottom / this._cellSize);

            for (int cy = minCy; cy <= maxCy; cy++)
            {
                for (int cx = minCx; cx <= maxCx; cx++)
                {
                    var key = MakeKey(cx, cy);

                    if (!this._cells.TryGetValue(key, out var bucket))
                        continue;

                    for (int b = 0; b < bucket.Count; b++)
                    {
                        int idx = bucket[b];

                        if (this._seenStamp[idx] == stamp)
                            continue;

                        this._seenStamp[idx] = stamp;

                        if (this._rects[idx].Intersects(pixelRect))
                            results.Add(idx);
                    }
                }
            }
        }

        private void AddToCells(int index, Rect rect)
        {
            int minCx = (int)Math.Floor(rect.Left / this._cellSize);
            int maxCx = (int)Math.Floor(rect.Right / this._cellSize);
            int minCy = (int)Math.Floor(rect.Top / this._cellSize);
            int maxCy = (int)Math.Floor(rect.Bottom / this._cellSize);

            for (int cy = minCy; cy <= maxCy; cy++)
            {
                for (int cx = minCx; cx <= maxCx; cx++)
                {
                    var key = MakeKey(cx, cy);
                    if (!this._cells.TryGetValue(key, out var list))
                    {
                        list = new List<int>(8);
                        this._cells[key] = list;
                    }

                    list.Add(index);
                }
            }
        }

        private int NextStamp()
        {
            this._stamp++;

            if (this._stamp == int.MaxValue)
            {
                Array.Clear(this._seenStamp);
                this._stamp = 1;
            }

            return this._stamp;
        }

        private static long MakeKey(int cx, int cy)
            => (unchecked((long)cx) << 32) | unchecked((uint)cy);
    }
}