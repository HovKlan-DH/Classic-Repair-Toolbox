using System.IO.Compression;
using System.Text;
using CRT.Server.Handlers.Feedback;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // ZipCentralDirectory - counting a zip's entries without ZipArchive, so a zip of millions of
    // empty records is refused before ZipArchive builds an object for each (code review, 2026-10-04).
    // ###########################################################################################
    public sealed class ZipCentralDirectoryTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-zip-directory-tests", Guid.NewGuid().ToString("N"));

        public ZipCentralDirectoryTests()
        {
            Directory.CreateDirectory(this.thisRoot);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private string Zip(int entries, string? comment = null)
        {
            string path = Path.Combine(this.thisRoot, $"{Guid.NewGuid():N}.zip");

            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                for (int index = 0; index < entries; index++)
                {
                    using Stream entry = archive.CreateEntry($"folder/file {index}.txt").Open();
                    entry.Write(Encoding.UTF8.GetBytes($"file {index}"));
                }

                if (comment is not null)
                    archive.Comment = comment;
            }

            return path;
        }

        [Fact]
        public void Every_record_is_counted_while_the_count_is_within_the_limit()
        {
            Assert.Equal(7L, ZipCentralDirectory.CountRecords(this.Zip(7), limit: 10));
            Assert.Equal(7L, ZipCentralDirectory.CountRecords(this.Zip(7), limit: 7));
        }

        // The walk stops one past the limit - which is how a caller tells "too many" - rather than
        // reading the rest of a directory that may hold millions.
        [Fact]
        public void Counting_stops_one_past_the_limit()
        {
            Assert.Equal(4L, ZipCentralDirectory.CountRecords(this.Zip(50), limit: 3));
        }

        // The end record is found behind a zip comment too - searched back from the end, as
        // ZipArchive searches for it.
        [Fact]
        public void A_zip_with_a_comment_is_counted_too()
        {
            Assert.Equal(5L, ZipCentralDirectory.CountRecords(this.Zip(5, comment: "Sent from CRT."), limit: 100));
        }

        [Fact]
        public void An_empty_zip_holds_nothing()
        {
            Assert.Equal(0L, ZipCentralDirectory.CountRecords(this.Zip(0), limit: 100));
        }

        // Not a zip, or no file at all: no answer, so ZipArchive says what is wrong in its own words.
        [Fact]
        public void A_file_that_is_not_a_zip_or_is_not_there_has_no_count()
        {
            string notAZip = Path.Combine(this.thisRoot, "not.zip");
            File.WriteAllText(notAZip, "not a zip at all, and long enough to hold an end record's worth of bytes");

            Assert.Null(ZipCentralDirectory.CountRecords(notAZip, limit: 100));
            Assert.Null(ZipCentralDirectory.CountRecords(Path.Combine(this.thisRoot, "missing.zip"), limit: 100));
        }

        // A record that does not begin with a header's signature ends the walk, as it ends
        // ZipArchive's - everything after it is not read as a record.
        [Fact]
        public void A_damaged_record_ends_the_count()
        {
            string path = this.Zip(3);
            byte[] bytes = File.ReadAllBytes(path);
            int directory = BitConverter.ToInt32(bytes, bytes.Length - 22 + 16);
            bytes[directory] = 0;
            File.WriteAllBytes(path, bytes);

            Assert.Equal(0L, ZipCentralDirectory.CountRecords(path, limit: 100));
        }
    }
}
