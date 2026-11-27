using System.IO.Compression;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // "DOWNLOAD ALL" ON A COLLECTION'S PAGE (owner request, 2026-11-27: the page after the owner's capture gallery, whose header
    // offered every picture in one go): every capture in one ZIP, in the page's order, each named as the viewer's "Download"
    // names it - "001-U7-12.png" (CollectionPages.CaptureFileName).
    //
    // *** KESTREL REFUSES A SYNCHRONOUS WRITE TO A RESPONSE, AND ZipArchive MAKES ONE WHATEVER IS ASKED OF IT. *** Measured on
    // .NET 10 (CollectionFlowsTests' write-only, async-only stream): ZipArchive.CreateAsync and ZipArchiveEntry.OpenAsync, the
    // entry closed with DisposeAsync, still wrote each entry's data descriptor synchronously (ZipArchiveEntry
    // .DirectToArchiveWriterStream.Dispose). So the ZIP is written into ASP.NET's own buffer - in memory up to
    // `InMemoryAtMost`, past it in a file of the service's own /tmp (PrivateTmp in the unit) - and that is sent to the response
    // asynchronously. The pictures are PNGs already, so they are STORED, not compressed again. A picture gone from the folder
    // meanwhile is left out, the rest still sent.
    // ###########################################################################################
    public static class CollectionDownload
    {
        public const string ContentType = "application/zip";

        public sealed record Entry(string Name, string Sha256);

        public static IReadOnlyList<Entry> EntriesOf(CollectionRecord collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            IReadOnlyList<string> anchors = CollectionRules.AnchorsOf(collection.Images.Select(image => image.Entry).ToList());

            return collection.Images
                .Select((image, index) => new Entry(CollectionPages.CaptureFileName(index, anchors[index]), image.Entry.Sha256))
                .ToList();
        }

        public static string FileNameOf(CollectionRecord collection) => $"collection-{collection.Id}.zip";

        // A collection of a few dozen captures is sent from memory; a larger one goes through a file.
        public const int InMemoryAtMost = 4 * 1024 * 1024;

        public static async Task WriteAsync(Stream target, IReadOnlyList<Entry> entries, Func<string, Stream?> open, CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(target);
            ArgumentNullException.ThrowIfNull(entries);
            ArgumentNullException.ThrowIfNull(open);

            await using var buffer = new Microsoft.AspNetCore.WebUtilities.FileBufferingWriteStream(CollectionDownload.InMemoryAtMost);

            using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (Entry entry in entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    using Stream? source = open(entry.Sha256);

                    if (source is null)
                        continue;

                    ZipArchiveEntry zipped = zip.CreateEntry(entry.Name, CompressionLevel.NoCompression);

                    using Stream into = zipped.Open();
                    source.CopyTo(into);
                }
            }

            await buffer.DrainBufferAsync(target, cancellationToken);
        }
    }
}
