using Avalonia;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Tabs.TabSchematics
{
    // One schematic image of a board, as its thumbnail: the reduced bitmap (null when the file is
    // missing or unreadable) and the full image's size, which the highlights are placed against.
    internal sealed record SchematicThumbnailImage(string Name, string FullPath, Bitmap? Thumbnail, PixelSize OriginalPixelSize);

    // ###########################################################################################
    // A board's schematic thumbnails, decoded once and kept (code review, 2026-10-13).
    //
    // A board switch used to decode every schematic image FULL SIZE (a 4220x2941 scan is ~50 MB)
    // only to render it into an 800-pixel thumbnail through an Image control on a pool thread - a
    // UI object off the UI thread - and a load superseded by the next click ran to its end. Now
    // each image is decoded straight to the thumbnail's width (Bitmap.DecodeToWidth; the full size
    // comes from the file's header), the decode stops between images when the board load is
    // cancelled, and Main keeps the last few boards' sets (StampedLruCache, by
    // ThumbnailDecodeGeometry.SignatureOf) so going back to a board decodes nothing.
    //
    // The set owns its bitmaps: Main disposes it when the cache lets it go, never while it is shown.
    // ###########################################################################################
    internal sealed class SchematicThumbnailSet : IDisposable
    {
        internal SchematicThumbnailSet(string signature, IReadOnlyList<SchematicThumbnailImage> images)
        {
            this.Signature = signature;
            this.Images = images;
        }

        internal string Signature { get; }

        internal IReadOnlyList<SchematicThumbnailImage> Images { get; }

        public void Dispose()
        {
            foreach (var image in this.Images)
            {
                image.Thumbnail?.Dispose();
            }
        }

        // ###########################################################################################
        // The signature of these images as they are on disk now (ThumbnailDecodeGeometry).
        // Off the UI thread: it reads every file's length and time.
        // ###########################################################################################
        internal static string SignatureOf(IReadOnlyList<(string Name, string FullPath)> images)
        {
            var stamped = new List<(string, string, long, long)>(images.Count);

            foreach (var (name, fullPath) in images)
            {
                long length = -1;
                long ticks = 0;

                try
                {
                    var file = new FileInfo(fullPath);

                    if (file.Exists)
                    {
                        length = file.Length;
                        ticks = file.LastWriteTimeUtc.Ticks;
                    }
                }
                catch (Exception)
                {
                    // Unreadable is "not there" for the signature - the decode says why.
                }

                stamped.Add((name, fullPath, length, ticks));
            }

            return ThumbnailDecodeGeometry.SignatureOf(stamped);
        }

        // ###########################################################################################
        // Decodes every image to a thumbnail at most `maxWidth` wide. Off the UI thread. Checks
        // `cancellationToken` between images; cancelled, it disposes what it decoded and throws
        // OperationCanceledException.
        // ###########################################################################################
        internal static SchematicThumbnailSet Decode(
            IReadOnlyList<(string Name, string FullPath)> images,
            string signature,
            int maxWidth,
            CancellationToken cancellationToken)
        {
            var decoded = new List<SchematicThumbnailImage>(images.Count);

            try
            {
                foreach (var (name, fullPath) in images)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var (thumbnail, originalSize) = SchematicThumbnailSet.DecodeOne(fullPath, maxWidth);
                    decoded.Add(new SchematicThumbnailImage(name, fullPath, thumbnail, originalSize));
                }
            }
            catch (OperationCanceledException)
            {
                foreach (var image in decoded)
                {
                    image.Thumbnail?.Dispose();
                }

                throw;
            }

            return new SchematicThumbnailSet(signature, decoded);
        }

        private static (Bitmap? Thumbnail, PixelSize OriginalSize) DecodeOne(string fullPath, int maxWidth)
        {
            if (!File.Exists(fullPath))
            {
                return (null, default);
            }

            try
            {
                // The full size from the header: no full decode just to learn two numbers.
                if (WorkbookPdfExporter.TryReadImageSize(fullPath) is { } size)
                {
                    using var stream = File.OpenRead(fullPath);
                    var thumbnail = Bitmap.DecodeToWidth(
                        stream,
                        ThumbnailDecodeGeometry.DecodeWidth(size.Width, maxWidth),
                        BitmapInterpolationMode.HighQuality);

                    return (thumbnail, new PixelSize(size.Width, size.Height));
                }

                // A format whose header is not read here (neither PNG nor JPEG): decoded whole, as
                // every image used to be, then scaled.
                using var full = new Bitmap(fullPath);
                int width = ThumbnailDecodeGeometry.DecodeWidth(full.PixelSize.Width, maxWidth);
                int height = ThumbnailDecodeGeometry.DecodeHeight(full.PixelSize.Width, full.PixelSize.Height, width);

                return (full.CreateScaledBitmap(new PixelSize(width, height), BitmapInterpolationMode.HighQuality), full.PixelSize);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not load schematic image [{fullPath}] - [{ex.Message}]");
                return (null, default);
            }
        }
    }
}
