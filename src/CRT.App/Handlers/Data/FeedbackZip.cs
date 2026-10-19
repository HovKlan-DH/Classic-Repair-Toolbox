using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // How the Feedback tab's zip is made (taken out of TabFeedback, code review 2026-10-16): the
    // name each entry goes in under, how hard each is compressed, and the pictures of the text put
    // in at the zip's root under the names their markers say (FeedbackComposition).
    // ###########################################################################################
    internal static class FeedbackZip
    {
        // ###########################################################################################
        // A picture is compressed already - PNG, JPEG, GIF, WebP - and deflating it again costs time
        // under the waiting overlay for next to nothing (code review, 2026-10-16: twenty 4K
        // screenshots went through the slowest deflate), so it is stored as it is. A BMP, which is
        // not compressed, and every other file are deflated as before.
        // ###########################################################################################
        public static CompressionLevel CompressionFor(string entryName) =>
            ImageFileTypes.IsDisplayable(entryName) && !string.Equals(Path.GetExtension(entryName), ".bmp", StringComparison.OrdinalIgnoreCase)
                ? CompressionLevel.NoCompression
                : CompressionLevel.Optimal;

        // ###########################################################################################
        // *** NO TWO ZIP ENTRIES OF ONE NAME (code review, 2026-10-16). *** The pictures go in under
        // the names their markers say; a file named as one of them, or as a file before it, goes in
        // as "name (2).ext" (FeedbackComposition.UniqueName - a folder's files keep their folder).
        // Two entries of one name and the server took the first for the picture at the marker, or
        // saved one file and could not save the other.
        // ###########################################################################################
        public static List<(string Source, string ZipEntryName)> WithUniqueEntryNames(
            IEnumerable<string> pictureNames, IEnumerable<(string Source, string ZipEntryName)> files)
        {
            var taken = new HashSet<string>(pictureNames, StringComparer.OrdinalIgnoreCase);
            var named = new List<(string Source, string ZipEntryName)>();

            foreach ((string source, string entryName) in files)
            {
                // Asked only on a clash: a folder of thousands of files is otherwise one look each.
                string unique = taken.Contains(entryName) ? FeedbackComposition.UniqueName(entryName, taken) : entryName;
                taken.Add(unique);
                named.Add((source, unique));
            }

            return named;
        }

        // ###########################################################################################
        // A picture of the text into the zip, at its root under the name its marker says: a pasted
        // one's PNG from memory, a dropped one STREAMED from its file as an attachment is (code
        // review, 2026-10-16: twenty 15 MB photos were 300 MB of arrays at once). The file is opened
        // before its entry is made, so one that has gone throws and leaves the zip as it was.
        // `added` hears the bytes as they go in, for the progress.
        // ###########################################################################################
        public static void AddPicture(ZipArchive archive, FeedbackZipPicture picture, Action<int>? added = null)
        {
            ArgumentNullException.ThrowIfNull(archive);
            ArgumentNullException.ThrowIfNull(picture);

            if (picture.Bytes is { } bytes)
            {
                using Stream entryStream = archive.CreateEntry(picture.Name, FeedbackZip.CompressionFor(picture.Name)).Open();
                entryStream.Write(bytes, 0, bytes.Length);
                added?.Invoke(bytes.Length);
                return;
            }

            using var file = new FileStream(picture.Path!, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using Stream target = archive.CreateEntry(picture.Name, FeedbackZip.CompressionFor(picture.Name)).Open();

            var buffer = new byte[81920];
            int read;

            while ((read = file.Read(buffer, 0, buffer.Length)) > 0)
            {
                target.Write(buffer, 0, read);
                added?.Invoke(read);
            }
        }

        // What a picture brings to the zip, in bytes, for the progress: a file's length as it is
        // now, 0 for one that has gone (its absence is the zip's to find).
        public static long SizeOf(FeedbackZipPicture picture)
        {
            ArgumentNullException.ThrowIfNull(picture);

            if (picture.Bytes is { } bytes)
                return bytes.Length;

            try
            {
                return new FileInfo(picture.Path!).Length;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return 0;
            }
        }
    }

    // ###########################################################################################
    // A picture as it goes into the zip: its name, and either its bytes (a pasted picture's PNG)
    // or the file they are read from (a dropped one) - exactly one of the two.
    // ###########################################################################################
    internal sealed class FeedbackZipPicture
    {
        private FeedbackZipPicture(string name, byte[]? bytes, string? path)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);

            this.Name = name;
            this.Bytes = bytes;
            this.Path = path;
        }

        public string Name { get; }

        public byte[]? Bytes { get; }

        public string? Path { get; }

        public static FeedbackZipPicture FromBytes(string name, byte[] bytes) =>
            new(name, bytes ?? throw new ArgumentNullException(nameof(bytes)), null);

        public static FeedbackZipPicture FromFile(string name, string path)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            return new(name, null, path);
        }
    }
}
