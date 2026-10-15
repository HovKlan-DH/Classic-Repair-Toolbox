using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // The sums behind a board's schematic thumbnails (code review, 2026-10-13): how wide to decode
    // an image, and what a set of thumbnails was made from - so a board shown again reuses its set
    // while its images are unchanged (Main's board load, SchematicThumbnailSet).
    // ###########################################################################################
    internal static class ThumbnailDecodeGeometry
    {
        // ###########################################################################################
        // The width to decode an image `originalWidth` pixels wide at for a thumbnail at most
        // `maxWidth` wide: the thumbnail's width, never wider than the image itself (no upscaling -
        // the old render capped its scale at 1 the same way). At least 1.
        // ###########################################################################################
        internal static int DecodeWidth(int originalWidth, int maxWidth) =>
            Math.Max(1, Math.Min(originalWidth, maxWidth));

        // ###########################################################################################
        // The height of that decode, by the image's own aspect ratio. At least 1.
        // ###########################################################################################
        internal static int DecodeHeight(int originalWidth, int originalHeight, int decodeWidth) =>
            originalWidth <= 0
                ? 1
                : Math.Max(1, (int)Math.Round((double)originalHeight * decodeWidth / originalWidth));

        // ###########################################################################################
        // What a set of thumbnails was made from: each image's name, file, length and write time,
        // in order. Two sets with the same signature show the same pictures under the same names.
        // A missing file is in it too (length -1), so it appearing later makes a new set.
        // ###########################################################################################
        internal static string SignatureOf(IEnumerable<(string Name, string FullPath, long Length, long LastWriteUtcTicks)> images) =>
            string.Join(
                "\u001E",
                images.Select(image => string.Join(
                    "\u001F",
                    image.Name,
                    image.FullPath,
                    image.Length.ToString(CultureInfo.InvariantCulture),
                    image.LastWriteUtcTicks.ToString(CultureInfo.InvariantCulture))));
    }
}
