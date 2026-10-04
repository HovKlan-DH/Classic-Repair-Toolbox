using System.IO.Compression;
using System.Net.Http.Headers;
using System.Text;
using CRT.Server.Handlers.Feedback;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The feedback form as the server reads it - fed the very form CRT builds
    // (FeedbackContract.BuildForm), so a field renamed on one side fails here, and the form older
    // CRTs send, written out by hand with the PHP page's names, so the contract cannot drift away
    // from what is already installed (Apache forwards their /app-feedback/ posts here).
    // ###########################################################################################
    public sealed class FeedbackFormReaderTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-feedback-form-tests", Guid.NewGuid().ToString("N"));

        public FeedbackFormReaderTests()
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

        private string Incoming => Path.Combine(this.thisRoot, "incoming.zip");

        private async Task<FeedbackFormRead> ReadAsync(HttpContent content)
        {
            await using Stream body = await content.ReadAsStreamAsync();
            return await FeedbackFormReader.ReadAsync(content.Headers.ContentType?.ToString(), body, this.Incoming, CancellationToken.None);
        }

        internal static byte[] Zip(params (string Name, string Text)[] entries)
        {
            using var stream = new MemoryStream();

            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach ((string name, string text) in entries)
                {
                    using Stream entry = archive.CreateEntry(name).Open();
                    entry.Write(Encoding.UTF8.GetBytes(text));
                }
            }

            return stream.ToArray();
        }

        [Fact]
        public async Task The_form_CRT_builds_is_read_field_for_field_with_its_zip_saved_where_asked()
        {
            byte[] zip = FeedbackFormReaderTests.Zip(("Classic-Repair-Toolbox.log", "line one"));

            using MultipartFormDataContent form = FeedbackContract.BuildForm("user@example.com", "The thumbnails flicker.", "2.6.0", new MemoryStream(zip));

            FeedbackFormRead read = await this.ReadAsync(form);

            Assert.False(read.IsRefused);
            Assert.Equal("user@example.com", read.Form!.Email);
            Assert.Equal("The thumbnails flicker.", read.Form.Feedback);
            Assert.Equal("2.6.0", read.Form.Version);
            Assert.Equal(this.Incoming, read.Form.AttachmentPath);
            Assert.Equal(zip, await File.ReadAllBytesAsync(this.Incoming));
        }

        [Fact]
        public async Task Without_files_attached_there_is_no_zip()
        {
            using MultipartFormDataContent form = FeedbackContract.BuildForm(string.Empty, "Just a thought.", "2.6.0", attachmentZip: null);

            FeedbackFormRead read = await this.ReadAsync(form);

            Assert.Null(read.Form!.AttachmentPath);
            Assert.False(File.Exists(this.Incoming));
        }

        // The PHP page's names and file name, written out literally: this is what every CRT already
        // installed sends, and it must keep working once Apache forwards it here.
        [Fact]
        public async Task The_form_older_CRTs_send_is_still_read()
        {
            byte[] zip = FeedbackFormReaderTests.Zip(("photo.jpg", "jpeg"));

            using var form = new MultipartFormDataContent
            {
                { new StringContent("old@example.com"), "email" },
                { new StringContent("Sent from 2.4.0."), "feedback" },
                { new StringContent("2.4.0"), "version" }
            };

            var file = new ByteArrayContent(zip);
            file.Headers.ContentType = new MediaTypeHeaderValue("application/zip");
            form.Add(file, "attachmentFile", "FeedbackPayload.zip");

            FeedbackFormRead read = await this.ReadAsync(form);

            Assert.Equal("old@example.com", read.Form!.Email);
            Assert.Equal("Sent from 2.4.0.", read.Form.Feedback);
            Assert.Equal("2.4.0", read.Form.Version);
            Assert.Equal(zip, await File.ReadAllBytesAsync(this.Incoming));
            Assert.Equal("Success", FeedbackContract.SuccessAnswer);
        }

        [Fact]
        public async Task A_feedback_text_is_kept_up_to_its_limit()
        {
            string longText = new string('x', FeedbackContract.MaximumFeedbackCharacters + 500);

            using MultipartFormDataContent form = FeedbackContract.BuildForm(string.Empty, longText, "2.6.0", attachmentZip: null);

            FeedbackFormRead read = await this.ReadAsync(form);

            Assert.Equal(FeedbackContract.MaximumFeedbackCharacters, read.Form!.Feedback.Length);
        }

        [Fact]
        public async Task A_body_that_is_not_a_form_is_refused_with_415()
        {
            using var json = new StringContent("{\"feedback\":\"hi\"}", Encoding.UTF8, "application/json");

            FeedbackFormRead read = await this.ReadAsync(json);

            Assert.True(read.IsRefused);
            Assert.Equal(415, read.RefusalStatus);
        }
    }
}
