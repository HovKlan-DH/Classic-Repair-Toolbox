using Avalonia.Media.Imaging;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A DEVICE'S PHOTOS, DECODED SMALL OFF THE UI THREAD - ONE READ FOR THE DEVICE'S PAGE AND THE DEVICE FORM ALIKE (code
    // review, 2026-11-08: each kept a copy of this, and the same photo was decoded once for each).
    //
    // Each photo is read ONCE, kept by where it lies, its length and its write time - a photo replaced is a new one, and
    // its older read is let go of - and never disposed by hand: one still drawn takes the render thread down, so a read let
    // go of is freed once nothing shows it. At most Kept photos are held, the one used longest ago going first (code
    // review, 2026-11-08: a plain dictionary held every version of every photo, deleted devices' too, for the run). A
    // photo that cannot be read has none, said in the log. The caller shows the read once it is in
    // (TabWorkbooks.ShowDevicePhoto, OwnedDeviceWindow.ShowThumbnail).
    // ###########################################################################################
    internal static class DevicePhotoThumbnails
    {
        // The width each is decoded to: the page shows a photo 160 wide, the form at most 200 - room for a screen at twice
        // the pixels.
        internal const int Width = 320;

        // How many photos are held at most - some 0.3 MB each, so tens of megabytes at the most, and room for every photo
        // of a well documented device's page.
        internal const int Kept = 200;

        private static readonly object Lock = new();

        private static readonly StampedLruCache<Task<Bitmap?>> Reads = new(DevicePhotoThumbnails.Kept, StringComparer.OrdinalIgnoreCase);

        private static int ReadsMade;

        // What a photo that is not there reads as: done, and nothing.
        private static readonly Task<Bitmap?> NoPhoto = Task.FromResult<Bitmap?>(null);

        // ###########################################################################################
        // The photo at `path`, read: the read made already or under way, else a new one on the pool - `whileRead` run there
        // first, for a test to hold it. A file that is not there reads as nothing, at once.
        // ###########################################################################################
        public static Task<Bitmap?> Read(string? path, Action? whileRead = null)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return DevicePhotoThumbnails.NoPhoto;
            }

            string stamp;

            try
            {
                var file = new FileInfo(path);
                stamp = $"{file.Length}|{file.LastWriteTimeUtc.Ticks}";
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return DevicePhotoThumbnails.NoPhoto;
            }

            // The look-up and the read made are one step: two asking at once share one read.
            lock (DevicePhotoThumbnails.Lock)
            {
                if (!DevicePhotoThumbnails.Reads.TryGet(path, stamp, out Task<Bitmap?> reading))
                {
                    reading = Task.Run(() =>
                    {
                        whileRead?.Invoke();
                        return DevicePhotoThumbnails.Decode(path);
                    });

                    DevicePhotoThumbnails.Reads.Store(path, stamp, reading);
                    DevicePhotoThumbnails.ReadsMade++;
                }

                return reading;
            }
        }

        // How many photos are held now - for the tests.
        internal static int HeldForTests => DevicePhotoThumbnails.Reads.Count;

        // How many photos have been read since CRT started - for the tests.
        internal static int ReadsMadeForTests
        {
            get
            {
                lock (DevicePhotoThumbnails.Lock)
                {
                    return DevicePhotoThumbnails.ReadsMade;
                }
            }
        }

        // Never throws - a photo that cannot be read has no thumbnail.
        private static Bitmap? Decode(string path)
        {
            try
            {
                using FileStream stream = File.OpenRead(path);
                return Bitmap.DecodeToWidth(stream, DevicePhotoThumbnails.Width);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not show a device photo: [{path}] [{ex.Message}]");
                return null;
            }
        }
    }
}
