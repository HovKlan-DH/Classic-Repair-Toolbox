using Avalonia;
using Avalonia.Media.Imaging;
using System;
using System.Collections.Generic;
using System.IO;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT PANE'S PICTURES, READ ONCE (code review, 2026-10-26). A component's thumbnails and
    // the Pictures part's previews are decoded off the UI thread; the part's rows are built again at
    // every re-read (an undo, a change made elsewhere), and each re-read decoded every picture again.
    //
    // KEPT BY WHERE THE BYTES LIE, NOT BY NAME (`KeyOf`: the file found, its length, its write time,
    // and the width decoded to). By name, another board's file of the same name - or the published
    // bytes "Undo changes" put back under a name already shown - took the picture read for the other.
    // A load still running when the cache is let go (another board, another component) keeps nothing:
    // it carries the `Generation` it started in. Bitmaps are never disposed by hand - one still drawn
    // takes the render thread down - so letting go is all `Clear` does. UI thread only.
    // ###########################################################################################
    internal sealed class PaneImageCache
    {
        private readonly Dictionary<string, PaneImage> thisImages = new(StringComparer.Ordinal);

        // Moves on at every Clear: a load that started before keeps nothing.
        public int Generation { get; private set; }

        internal int CountForTests => this.thisImages.Count;

        // ###########################################################################################
        // The key a picture is kept under - null for no file there. Runs anywhere (it asks the disk).
        // ###########################################################################################
        public static string? KeyOf(string? path, int width)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            try
            {
                var file = new FileInfo(path);

                return file.Exists
                    ? string.Join('\u001F', file.FullName, file.Length, file.LastWriteTimeUtc.Ticks, width)
                    : null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
            {
                return null;
            }
        }

        public bool TryGet(string key, out PaneImage image)
        {
            if (this.thisImages.TryGetValue(key, out PaneImage? found))
            {
                image = found;
                return true;
            }

            image = PaneImage.None;
            return false;
        }

        // Kept - unless the cache was let go since `generation`. True when kept.
        public bool Keep(string key, PaneImage image, int generation)
        {
            if (generation != this.Generation)
            {
                return false;
            }

            this.thisImages[key] = image;
            return true;
        }

        public void Clear()
        {
            this.thisImages.Clear();
            this.Generation++;
        }
    }

    // A picture as decoded - null when it could not be - and its own pixel size, when known.
    internal sealed record PaneImage(Bitmap? Bitmap, Size OriginalSize)
    {
        public static PaneImage None { get; } = new(null, default);
    }
}
