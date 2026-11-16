using Avalonia;
using Avalonia.Media.Imaging;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A schematic image as the "Workbooks" tab's board pane and the worklog editor opened from it show it: decoded at
    // most Width wide, with the size of the picture ITSELF beside it (2026-11-16, performance walkthrough).
    //
    // A worklog's marked area is kept in the picture's own pixels, so everything mapping one - the pane's overlay and
    // badges, the editor's preview, the default area the editor places - goes by OriginalPixelSize, never the
    // decoded bitmap's PixelSize. The pane showed each picture at most 420 high from a full-size decode: some 47 MB of
    // native memory for a 4220x2941 image, decoded on the UI thread at every board switch (62-77 ms each, measured)
    // and again at every visit to the tab, as leaving it disposed them.
    // ###########################################################################################
    internal sealed record SchematicPreviewPicture(Bitmap Bitmap, PixelSize OriginalPixelSize);

    // ###########################################################################################
    // EACH PICTURE READ ONCE, kept by where it lies, its length and its write time - a picture replaced is a new one -
    // and NEVER DISPOSED BY HAND, as DevicePhotoThumbnails: one still drawn takes the render thread down, so a read let
    // go of is freed once nothing shows it (the pane, or an editor opened from it). At most Kept are held. Two asking
    // at once share one read. A picture that cannot be read has none, said in the log.
    // ###########################################################################################
    internal static class SchematicPreviewPictures
    {
        // The width each is decoded to at most: the pane shows a picture at most 420 high, and a 4220x2941 schematic at
        // 1600 wide is 1115 high - room for a screen at twice the pixels. A narrower picture is decoded as it is.
        internal const int Width = 1600;

        // How many are held at most - some 7 MB each at the widest.
        internal const int Kept = 24;

        private static readonly object Lock = new();

        private static readonly StampedLruCache<Task<SchematicPreviewPicture?>> Reads = new(SchematicPreviewPictures.Kept, StringComparer.OrdinalIgnoreCase);

        private static readonly Task<SchematicPreviewPicture?> NoPicture = Task.FromResult<SchematicPreviewPicture?>(null);

        private static int ReadsMade;

        // ###########################################################################################
        // The picture at `path`: the read made already or under way, else a new one on the pool. A file that is not
        // there reads as nothing, at once.
        // ###########################################################################################
        public static Task<SchematicPreviewPicture?> Read(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || SchematicPreviewPictures.StampOf(path) is not { } stamp)
            {
                return SchematicPreviewPictures.NoPicture;
            }

            lock (SchematicPreviewPictures.Lock)
            {
                if (!SchematicPreviewPictures.Reads.TryGet(path, stamp, out Task<SchematicPreviewPicture?> reading))
                {
                    reading = Task.Run(() => SchematicPreviewPictures.Decode(path));
                    SchematicPreviewPictures.Reads.Store(path, stamp, reading);
                    SchematicPreviewPictures.ReadsMade++;

                    // ###########################################################################################
                    // *** A READ THAT CAME TO NOTHING IS NOT KEPT (code review, 2026-11-16), as PictureReads'. *** A picture
                    // held a moment by a virus scanner as the pane first read it stayed "could not be found" - and a worklog
                    // opened from the pane had no preview - until the file changed or 24 others were read. The next ask
                    // reads it again.
                    // ###########################################################################################
                    Task<SchematicPreviewPicture?> made = reading;
                    _ = made.ContinueWith(
                        read =>
                        {
                            if (read.Result is null)
                            {
                                SchematicPreviewPictures.Reads.Forget(path, made);
                            }
                        },
                        CancellationToken.None,
                        TaskContinuationOptions.OnlyOnRanToCompletion,
                        TaskScheduler.Default);
                }

                return reading;
            }
        }

        // What a picture is kept by besides where it lies - its length and write time, by the one stamp rule
        // (BoardDataReader.FileStampIfThere; code review, 2026-11-16: a copy of it stood here) - or null for no file there.
        // A file that cannot be asked about has a stamp of its own each time, so it is never taken as read.
        internal static string? StampOf(string? path) => BoardDataReader.FileStampIfThere(path);

        // How many pictures have been read since CRT started - for the tests.
        internal static int ReadsMadeForTests
        {
            get
            {
                lock (SchematicPreviewPictures.Lock)
                {
                    return SchematicPreviewPictures.ReadsMade;
                }
            }
        }

        // ###########################################################################################
        // Decoded at most Width wide, its own size read off its header (ImagePixelSize - the one reader of picture
        // headers); one whose header CRT cannot read is decoded whole, and its size is the bitmap's. Never throws.
        // ###########################################################################################
        internal static SchematicPreviewPicture? Decode(string path)
        {
            try
            {
                (int Width, int Height)? header = ImagePixelSize.TryRead(path);

                using FileStream stream = File.OpenRead(path);

                Bitmap bitmap = header is { Width: > SchematicPreviewPictures.Width }
                    ? Bitmap.DecodeToWidth(stream, SchematicPreviewPictures.Width)
                    : new Bitmap(stream);

                PixelSize original = header is { Width: > 0, Height: > 0 } known
                    ? new PixelSize(known.Width, known.Height)
                    : bitmap.PixelSize;

                return new SchematicPreviewPicture(bitmap, original);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Workbooks tab: could not load schematic image [{path}] - [{ex.Message}]");
                return null;
            }
        }
    }
}
