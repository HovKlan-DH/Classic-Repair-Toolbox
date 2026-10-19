using System.Text;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Feedback;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // FEEDBACK (owner request, 2026-10-03: the service takes over feedback; the mail as HTML with
    // the log "as monospace"; the files "just keep exact same behaviour" - a feedback-<random>
    // folder, opened from the network share). Against a temporary folder and a fake mailer.
    // ###########################################################################################
    public sealed class FeedbackFlowTests : IDisposable
    {
        private const string Owner = "dennis@classic-repair-toolbox.dk";
        private const string Reference = "feedback-TestReference0001";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-feedback-flow-tests", Guid.NewGuid().ToString("N"));
        private readonly FakeEmailSender thisMailer = new();

        public FeedbackFlowTests()
        {
            Directory.CreateDirectory(Path.Combine(this.thisRoot, "user-feedback"));
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

        private string FeedbackRoot => Path.Combine(this.thisRoot, "user-feedback");

        private string Folder => Path.Combine(this.FeedbackRoot, FeedbackFlowTests.Reference);

        private FeedbackDestination Destination(long? freeBytes = null) =>
            new(this.FeedbackRoot, FeedbackFlowTests.Owner, () => freeBytes, 1024, () => FeedbackFlowTests.Reference);

        private string ZipFile(params (string Name, string Text)[] entries)
        {
            string path = Path.Combine(this.thisRoot, $"{Guid.NewGuid():N}.zip");
            File.WriteAllBytes(path, FeedbackFormReaderTests.Zip(entries));
            return path;
        }

        private Task<FeedbackOutcome> SendAsync(string email, string feedback, string? zip = null, string version = "2.6.0", FeedbackDestination? destination = null, string pictureWidths = "") =>
            FeedbackFlow.HandleAsync(new FeedbackForm(email, feedback, version, zip, pictureWidths), destination ?? this.Destination(), this.thisMailer, CancellationToken.None);

        // A "PNG" of the given size: the signature the content check reads, then nothing.
        private static byte[] LargePng(int size)
        {
            byte[] bytes = new byte[size];
            FeedbackFlowTests.OnePixelPng.AsSpan(0, 8).CopyTo(bytes);
            return bytes;
        }

        // A zip of binary entries - the pictures CRT's Feedback tab puts at the zip's root.
        private string ZipFileOfBytes(params (string Name, byte[] Bytes)[] entries)
        {
            string path = Path.Combine(this.thisRoot, $"{Guid.NewGuid():N}.zip");

            using (var stream = new FileStream(path, FileMode.CreateNew))
            using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create))
            {
                foreach ((string name, byte[] bytes) in entries)
                {
                    using Stream entry = archive.CreateEntry(name).Open();
                    entry.Write(bytes, 0, bytes.Length);
                }
            }

            return path;
        }

        // A minimal valid 1x1 PNG - the signature check reads the first bytes of a real one.
        private static readonly byte[] OnePixelPng =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        ];

        // ###########################################################################################
        // *** PICTURES IN THE TEXT ARE SHOWN IN THE MAIL (owner request, 2026-10-15). *** The text
        // names a picture by a marker line - what CRT's Feedback tab writes (CRT.Data's
        // FeedbackComposition, used here too, so a marker CRT writes is one the server reads) - and
        // the zip holds the file at its root: the HTML shows it INSIDE the sender's quoted words,
        // where the line sat (owner report, 2026-10-15: each run of words was a box of its own,
        // the pictures between the boxes), and the plain text keeps the marker. A picture the mail
        // shows is NOT saved as well (owner report, the same day: "they are inline so they should
        // not be uploaded to server") - the other files are, as before.
        // ###########################################################################################
        [Fact]
        public async Task A_picture_the_text_names_is_shown_inside_the_senders_quote_and_not_saved()
        {
            string marker = FeedbackComposition.Marker(FeedbackComposition.PastedName(1));
            string zip = this.ZipFileOfBytes(("pasted-image-1.png", FeedbackFlowTests.OnePixelPng), ("notes.txt", "some notes"u8.ToArray()));

            FeedbackOutcome outcome = await this.SendAsync("user@example.com", $"Black screen after warm-up:\n{marker}\nThe PLA gets hot.", zip);

            Assert.True(outcome.IsDelivered);
            EmailMessage mail = Assert.Single(this.thisMailer.Sent);

            MailLinkedImage picture = Assert.Single(mail.LinkedImages ?? []);
            Assert.Equal("pasted-image-1.png", picture.FileName);
            Assert.Equal("image/png", picture.ContentType);
            Assert.Equal(FeedbackFlowTests.OnePixelPng, picture.Bytes);

            // ONE quote panel, holding the words above, the picture, and the words below - in that order.
            string html = mail.HtmlBody!;
            int quote = html.IndexOf(MailBody.QuotePanelStart, StringComparison.Ordinal);
            int above = html.IndexOf("Black screen after warm-up:", StringComparison.Ordinal);
            int image = html.IndexOf($"<img src=\"cid:{picture.ContentId}\"", StringComparison.Ordinal);
            int below = html.IndexOf("The PLA gets hot.", StringComparison.Ordinal);
            int quoteEnd = html.IndexOf(MailBody.QuotePanelEnd, below, StringComparison.Ordinal);
            Assert.True(quote >= 0 && quote < above && above < image && image < below && below < quoteEnd, html);
            Assert.Equal(1, FeedbackFlowTests.Count(html, MailBody.QuotePanelStart));

            // The plain text reads as the sender typed it.
            Assert.Contains("[pasted-image-1.png]", mail.Body, StringComparison.Ordinal);

            // Shown, so not saved; the other file is.
            Assert.False(File.Exists(Path.Combine(this.Folder, "pasted-image-1.png")));
            Assert.True(File.Exists(Path.Combine(this.Folder, "notes.txt")));
            Assert.Equal(1, outcome.FilesSaved);
        }

        // The owner's own feedback of 2026-10-15: words and two pasted screenshots, nothing else.
        // Nothing is left to save, so there is no folder, no reference and no list of files.
        [Fact]
        public async Task Pictures_alone_save_nothing_so_the_mail_has_no_reference_and_no_files()
        {
            string zip = this.ZipFileOfBytes(("pasted-image-1.png", FeedbackFlowTests.OnePixelPng), ("pasted-image-2.png", FeedbackFlowTests.OnePixelPng));

            FeedbackOutcome outcome = await this.SendAsync(
                "user@example.com", "Hej Dennis.\n[pasted-image-1.png]\nDet går fint:\n[pasted-image-2.png]\n// Dennis", zip);

            Assert.True(outcome.IsDelivered);
            Assert.Null(outcome.Reference);
            Assert.Equal(0, outcome.FilesSaved);
            Assert.False(Directory.Exists(this.Folder));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Equal(2, (mail.LinkedImages ?? []).Count);
            Assert.Equal(1, FeedbackFlowTests.Count(mail.HtmlBody!, MailBody.QuotePanelStart));
            Assert.DoesNotContain("Internal reference", mail.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("Files associated", mail.Body, StringComparison.Ordinal);
        }

        // A feedback of pictures and no words at all (owner report, 2026-10-16: then the pictures
        // were not in the panel the sender's words have): the pictures are still the sender's
        // quotation - one panel the mail draws itself (MailBody's QuotePanelStart, never a
        // <blockquote>, which Fastmail painted only when it held words), holding them, and nothing
        // standing in for missing words.
        [Fact]
        public async Task Pictures_without_any_words_are_shown_inside_the_senders_panel()
        {
            string zip = this.ZipFileOfBytes(("pasted-image-1.png", FeedbackFlowTests.OnePixelPng), ("pasted-image-2.png", FeedbackFlowTests.OnePixelPng));

            await this.SendAsync("user@example.com", "[pasted-image-1.png]\n[pasted-image-2.png]", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            string html = mail.HtmlBody!;

            int panel = html.IndexOf(MailBody.QuotePanelStart, StringComparison.Ordinal);
            int first = html.IndexOf("<img", StringComparison.Ordinal);
            int last = html.LastIndexOf("<img", StringComparison.Ordinal);
            int panelEnd = html.IndexOf(MailBody.QuotePanelEnd, last, StringComparison.Ordinal);

            Assert.True(panel >= 0 && panel < first && first < last && last < panelEnd, html);
            Assert.Equal(1, FeedbackFlowTests.Count(html, MailBody.QuotePanelStart));
            Assert.DoesNotContain("<blockquote", html, StringComparison.Ordinal);
            Assert.DoesNotContain("(no text given)", html, StringComparison.Ordinal);
        }

        private static int Count(string text, string value)
        {
            int count = 0;

            for (int at = text.IndexOf(value, StringComparison.Ordinal); at >= 0; at = text.IndexOf(value, at + value.Length, StringComparison.Ordinal))
                count++;

            return count;
        }

        [Fact]
        public async Task A_marker_naming_no_file_in_the_zip_stays_a_line_of_the_senders_text()
        {
            string zip = this.ZipFile(("notes.txt", "some notes"));

            await this.SendAsync("user@example.com", "See:\n[pasted-image-1.png]\nmore", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Null(mail.LinkedImages);
            Assert.DoesNotContain("<img", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("[pasted-image-1.png]", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("[pasted-image-1.png]", mail.Body, StringComparison.Ordinal);
        }

        // The mail may be so large; a picture past what the text leaves of it is saved and said, not
        // shown - and the pictures before it still are.
        [Fact]
        public async Task A_picture_past_the_mails_room_is_saved_and_said_not_shown()
        {
            FeedbackDestination destination = this.Destination() with { MaxMailBytes = FeedbackFlow.MailOverheadBytes + 300 * 1024 };
            string zip = this.ZipFileOfBytes(("pasted-image-1.png", FeedbackFlowTests.OnePixelPng), ("pasted-image-2.png", FeedbackFlowTests.LargePng(400 * 1024)));

            await this.SendAsync("user@example.com", "[pasted-image-1.png]\n[pasted-image-2.png]", zip, destination: destination);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            MailLinkedImage shown = Assert.Single(mail.LinkedImages ?? []);
            Assert.Equal("pasted-image-1.png", shown.FileName);
            Assert.Contains("pasted-image-2.png", mail.Body, StringComparison.Ordinal);
            Assert.Contains("too large to show in the mail", mail.Body, StringComparison.Ordinal);

            // The one the mail could not show is kept, or it would reach nobody; the shown one is not.
            Assert.True(File.Exists(Path.Combine(this.Folder, "pasted-image-2.png")));
            Assert.False(File.Exists(Path.Combine(this.Folder, "pasted-image-1.png")));
        }

        // The pictures get the mail's room in the order the TEXT names them, not the zip's order - a
        // picture early in the text is never pushed out by one further down.
        [Fact]
        public async Task The_pictures_get_the_mails_room_in_the_order_the_text_names_them()
        {
            FeedbackDestination destination = this.Destination() with { MaxMailBytes = FeedbackFlow.MailOverheadBytes + 640 * 1024 };
            string zip = this.ZipFileOfBytes(("small.png", FeedbackFlowTests.LargePng(100 * 1024)), ("big.png", FeedbackFlowTests.LargePng(400 * 1024)));

            await this.SendAsync("user@example.com", "[big.png]\n[small.png]", zip, destination: destination);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Equal("big.png", Assert.Single(mail.LinkedImages ?? []).FileName);
            Assert.True(File.Exists(Path.Combine(this.Folder, "small.png")));
        }

        // A width the sender gave a picture in CRT's field is its width in the mail.
        [Fact]
        public async Task A_picture_is_shown_at_the_width_its_sender_gave_it()
        {
            string zip = this.ZipFileOfBytes(("pasted-image-1.png", FeedbackFlowTests.OnePixelPng), ("pasted-image-2.png", FeedbackFlowTests.OnePixelPng));

            await this.SendAsync(
                "user@example.com", "[pasted-image-1.png]\n[pasted-image-2.png]", zip,
                pictureWidths: FeedbackComposition.PictureWidthsText([("pasted-image-1.png", 480)]));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Equal([480, null], (mail.LinkedImages ?? []).Select(picture => picture.Width));
            Assert.Contains("width=\"480\"", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Equal(1, FeedbackFlowTests.Count(mail.HtmlBody!, "width=\""));
        }

        // ###########################################################################################
        // *** THE WHOLE MAIL STAYS WITHIN ITS LIMIT (owner decision, 2026-10-15: up to 20 MB). *** The
        // pictures' room, base64-encoded in lines, plus both text bodies base64-encoded, plus the
        // overhead, never comes to more than the limit - and comes close to it.
        // ###########################################################################################
        [Theory]
        [InlineData(0)]
        [InlineData(1024)]
        [InlineData(3 * 1024 * 1024)]
        [InlineData(3 * 1024 * 1024 + 100_000)]
        public void The_pictures_room_keeps_the_whole_mail_within_its_limit(long textBytesPerBody)
        {
            long room = FeedbackFlow.PictureRoom(FeedbackFlow.MaxMailBytes, textBytesPerBody);
            long pictures = room * 4 / 3 * 78 / 76;
            long text = 2 * textBytesPerBody * 4 / 3;

            Assert.True(pictures + text + FeedbackFlow.MailOverheadBytes <= FeedbackFlow.MaxMailBytes);
            Assert.True(pictures + text + FeedbackFlow.MailOverheadBytes > FeedbackFlow.MaxMailBytes - 16);
        }

        [Fact]
        public void A_mail_of_words_alone_leaves_the_pictures_about_fourteen_MB_and_more_text_leaves_less()
        {
            Assert.InRange(FeedbackFlow.PictureRoom(FeedbackFlow.MaxMailBytes, 0), 14L * 1024 * 1024, 15L * 1024 * 1024);
            Assert.True(FeedbackFlow.PictureRoom(FeedbackFlow.MaxMailBytes, 3 * 1024 * 1024) < FeedbackFlow.PictureRoom(FeedbackFlow.MaxMailBytes, 0));
            Assert.Equal(0, FeedbackFlow.PictureRoom(FeedbackFlow.MaxMailBytes, FeedbackFlow.MaxMailBytes));
        }

        // A zip is what the sender made it: a file named as a picture whose bytes are not one is
        // never put into the mail as an image.
        [Fact]
        public async Task A_file_named_as_a_picture_whose_contents_are_not_one_is_not_shown()
        {
            string zip = this.ZipFileOfBytes(("pasted-image-1.png", "<html>not a picture</html>"u8.ToArray()));

            await this.SendAsync("user@example.com", "[pasted-image-1.png]", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Null(mail.LinkedImages);
            Assert.Contains("not the image its name says", mail.Body, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(this.Folder, "pasted-image-1.png")));
        }

        // Only a file at the zip's ROOT is a picture of the text; a marker names a file name, never a path.
        [Fact]
        public async Task A_picture_in_a_folder_of_the_zip_is_not_the_texts_picture()
        {
            string zip = this.ZipFileOfBytes(("board/pasted-image-1.png", FeedbackFlowTests.OnePixelPng));

            await this.SendAsync("user@example.com", "[pasted-image-1.png]", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Null(mail.LinkedImages);
            Assert.True(File.Exists(Path.Combine(this.Folder, "board", "pasted-image-1.png")));
        }

        // A feedback without markers - every CRT released before pictures - is mailed as before.
        [Fact]
        public async Task A_feedback_without_markers_carries_no_pictures()
        {
            string zip = this.ZipFileOfBytes(("shot.png", FeedbackFlowTests.OnePixelPng));

            await this.SendAsync("user@example.com", "Just words.", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Null(mail.LinkedImages);
            Assert.DoesNotContain("<img", mail.HtmlBody, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(this.Folder, "shot.png")));
        }

        // ###########################################################################################
        // *** A MARKER LINE REPEATED SHOWS ITS PICTURE ONCE (code review, 2026-10-16). *** The text
        // is the sender's, and every repeat of a marker line put the picture into the mail again -
        // 12,000 lines naming one 14 MB picture made SmtpClient encode 220 GB. The picture is shown
        // at its first line and carried once; a later line naming it stays as written.
        // ###########################################################################################
        [Fact]
        public async Task A_marker_line_repeated_shows_its_picture_once_and_carries_it_once()
        {
            string zip = this.ZipFileOfBytes(("a.png", FeedbackFlowTests.OnePixelPng));

            await this.SendAsync("user@example.com", "[a.png]\nagain:\n[a.png]\n[A.PNG]", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Single(mail.LinkedImages ?? []);
            Assert.Equal(1, FeedbackFlowTests.Count(mail.HtmlBody!, "<img"));
            Assert.Contains("[a.png]", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("[A.PNG]", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Equal(3, FeedbackFlowTests.Count(mail.Body, "[a.png]") + FeedbackFlowTests.Count(mail.Body, "[A.PNG]"));
        }

        // A line naming the sender's own attachment that is no picture refers to it: the file is
        // saved as ever, and the mail never calls it a picture it could not show.
        [Fact]
        public async Task A_marker_naming_a_file_that_is_no_picture_leaves_it_an_attachment()
        {
            string zip = this.ZipFile(("notes.txt", "some notes"));

            await this.SendAsync("user@example.com", "My notes:\n[notes.txt]", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Null(mail.LinkedImages);
            Assert.DoesNotContain("picture", mail.Body, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("[notes.txt]", mail.Body, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(this.Folder, "notes.txt")));
        }

        // Each picture's type is the one the review routes serve the same file as - the server's
        // one allowlist (ReviewAssetLocator).
        [Theory]
        [InlineData("shot.png")]
        [InlineData("photo.JPEG")]
        public async Task A_pictures_type_is_the_one_the_server_serves_it_as(string name)
        {
            byte[] bytes = name.EndsWith(".png", StringComparison.Ordinal)
                ? FeedbackFlowTests.OnePixelPng
                : [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01];
            string zip = this.ZipFileOfBytes((name, bytes));

            await this.SendAsync("user@example.com", $"[{name}]", zip);

            MailLinkedImage picture = Assert.Single(Assert.Single(this.thisMailer.Sent).LinkedImages ?? []);
            Assert.Equal(CRT.Server.Handlers.Submissions.ReviewAssetLocator.ContentTypeFor(name), picture.ContentType);
            Assert.StartsWith("image/", picture.ContentType, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** THE TEXT'S SHARE OF THE MAIL IS WHAT THE MAIL MAKES OF IT (code review, 2026-10-16). ***
        // The quotation indents every line by four spaces in the plain text and ends it with "<br>"
        // in the HTML, so a text of short lines is far larger in the mail than as typed - and the
        // pictures were given the difference: a picture as large as the room counted from the typed
        // text took the whole mail past its limit.
        // ###########################################################################################
        [Fact]
        public async Task A_text_of_many_short_lines_leaves_the_pictures_only_what_the_mail_leaves()
        {
            string feedback = string.Join("\n", Enumerable.Repeat("a", 50_000)) + "\n[big.png]";
            FeedbackDestination destination = this.Destination() with { MaxMailBytes = FeedbackFlow.MailOverheadBytes + 1024 * 1024 };
            long roomAsTyped = FeedbackFlow.PictureRoom(destination.MaxMailBytes, FeedbackFlow.MailBytes(feedback));
            string zip = this.ZipFileOfBytes(("big.png", FeedbackFlowTests.LargePng((int)roomAsTyped - 2048)));

            await this.SendAsync("user@example.com", feedback, zip, destination: destination);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.True(FeedbackFlowTests.MailSize(mail) <= destination.MaxMailBytes, $"{FeedbackFlowTests.MailSize(mail)} > {destination.MaxMailBytes}");
        }

        // A zip of thousands of entries the server will not save made a note each - written after the
        // pictures' room was fixed, megabytes of them. The mail lists them within a room of their
        // own and says how many more there were.
        [Fact]
        public async Task Thousands_of_notes_are_listed_within_their_room_and_the_rest_counted()
        {
            var entries = Enumerable.Range(0, 3000).Select(i => ($"../escaped-{i}.txt", "evil")).Append(("fine.txt", "ok")).ToArray();
            string zip = this.ZipFile(entries);

            await this.SendAsync("user@example.com", "Words.", zip);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.InRange(Encoding.UTF8.GetByteCount(mail.Body), 1, 64 * 1024);
            Assert.InRange(Encoding.UTF8.GetByteCount(mail.HtmlBody!), 1, 96 * 1024);
            Assert.Contains("more note(s) not listed", mail.Body, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(this.Folder, "fine.txt")));
        }

        // The text's share is measured as the quotation draws it - never less than either body takes.
        [Theory]
        [InlineData("One line.")]
        [InlineData("a\nb\nc\n\n\nd")]
        [InlineData("\"quoted\" <b>&amp;</b>\r\nsecond")]
        public void The_texts_share_is_what_its_quotation_takes_in_either_body(string text)
        {
            MailBody quoted = new MailBody(MailBody.WideWidth).Quote(text, string.Empty);
            long share = FeedbackFlow.QuotedMailBytes(text);

            Assert.True(share >= Encoding.UTF8.GetByteCount(quoted.ToText()));
            Assert.True(share >= Encoding.UTF8.GetByteCount(quoted.ToHtml(string.Empty)));
            Assert.True(share >= FeedbackFlow.MailBytes(text));
        }

        // A list goes in, in order, while it fits its count and its room; the rest is counted.
        [Fact]
        public void A_list_is_listed_while_it_fits_and_the_rest_counted()
        {
            string[] items = ["first", "second", "third"];

            AssertListed(["first", "second", "third"], 0, FeedbackFlow.ListWithin(items, 10, 10_000));
            AssertListed(["first", "second"], 1, FeedbackFlow.ListWithin(items, 2, 10_000));
            AssertListed(["first"], 2, FeedbackFlow.ListWithin(items, 10, FeedbackFlow.MailBytes("first") + 40));
            AssertListed([], 3, FeedbackFlow.ListWithin(items, 10, 0));

            // A name full of what HTML encodes counts as the mail carries it.
            string quotes = new('"', 1000);
            Assert.Equal(1, FeedbackFlow.ListWithin([quotes], 10, 2000).NotListed);
        }

        private static void AssertListed(string[] listed, int notListed, (List<string> Listed, int NotListed) list)
        {
            Assert.Equal(listed, list.Listed);
            Assert.Equal(notListed, list.NotListed);
        }

        // What a mail comes to as it is sent: both bodies base64-encoded, and each picture too, in
        // lines of 76, with about a kilobyte of its own part's headers.
        private static long MailSize(EmailMessage mail) =>
            (Encoding.UTF8.GetByteCount(mail.Body) + Encoding.UTF8.GetByteCount(mail.HtmlBody ?? string.Empty)) * 4L / 3 +
            (mail.LinkedImages ?? []).Sum(picture => picture.Bytes.Length * 4L / 3 * 78 / 76 + 1024);

        [Fact]
        public async Task Text_alone_is_mailed_to_the_owner_from_the_service_with_reply_to_the_sender()
        {
            FeedbackOutcome outcome = await this.SendAsync("user@example.com", "The <b>zoom</b> jumps.");

            Assert.True(outcome.IsDelivered);
            Assert.Equal(FeedbackContract.SuccessAnswer, outcome.Answer);
            Assert.Null(outcome.Reference);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Equal(FeedbackFlowTests.Owner, mail.ToAddress);
            Assert.Equal("Feedback from CRT 2.6.0", mail.Subject);
            Assert.Equal("user@example.com", mail.ReplyToAddress);

            // The sender's words, never their markup.
            Assert.Contains("The &lt;b&gt;zoom&lt;/b&gt; jumps.", mail.HtmlBody, StringComparison.Ordinal);
            Assert.DoesNotContain("<b>zoom</b>", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("The <b>zoom</b> jumps.", mail.Body, StringComparison.Ordinal);

            // No files, so no folder and no reference.
            Assert.Empty(Directory.EnumerateFileSystemEntries(this.FeedbackRoot));
            Assert.DoesNotContain("Internal reference", mail.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task Without_a_usable_address_there_is_no_reply_to_and_the_mail_says_so()
        {
            await this.SendAsync("not an address", "Hello.");

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Null(mail.ReplyToAddress);
            Assert.Contains("no email address given", mail.Body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // CRT's own text files are SHOWN, monospaced ("logfile should still show as monospace as
        // this is kind of quoted text"); everything else is SAVED in the feedback's folder, listed
        // in the mail under the folder's name - the "Internal reference" feedback mails have always
        // given.
        // ###########################################################################################
        [Fact]
        public async Task The_log_and_settings_are_shown_and_the_other_files_are_saved_under_the_reference()
        {
            string zip = this.ZipFile(
                ("Classic-Repair-Toolbox.log", "12:00 Started <CRT>\n12:01 Board loaded"),
                ("Classic-Repair-Toolbox.settings.json", "{ \"Theme\": \"Dark\" }"),
                ("photo of board.jpg", "jpeg bytes"),
                ("My notes/part list.txt", "U1 6510"));

            FeedbackOutcome outcome = await this.SendAsync("user@example.com", "See the log.", zip);

            Assert.True(outcome.IsDelivered);
            Assert.Equal(FeedbackFlowTests.Reference, outcome.Reference);
            Assert.Equal(2, outcome.FilesSaved);

            Assert.Equal("jpeg bytes", File.ReadAllText(Path.Combine(this.Folder, "photo of board.jpg")));
            Assert.Equal("U1 6510", File.ReadAllText(Path.Combine(this.Folder, "My notes", "part list.txt")));
            Assert.False(File.Exists(Path.Combine(this.Folder, "Classic-Repair-Toolbox.log")));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Contains($"Internal reference: {FeedbackFlowTests.Reference}", mail.Body, StringComparison.Ordinal);
            Assert.Contains("Logfile:", mail.Body, StringComparison.Ordinal);
            Assert.Contains("12:01 Board loaded", mail.Body, StringComparison.Ordinal);
            Assert.Contains("Settings:", mail.Body, StringComparison.Ordinal);
            Assert.Contains("My notes/part list.txt", mail.Body, StringComparison.Ordinal);

            // Monospaced and escaped in the HTML.
            Assert.Contains("<pre style=", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("monospace", mail.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("12:00 Started &lt;CRT&gt;", mail.HtmlBody, StringComparison.Ordinal);
        }

        [Fact]
        public async Task A_zip_holding_only_CRTs_own_files_makes_no_folder()
        {
            string zip = this.ZipFile(("Classic-Repair-Toolbox.log", "log"), ("Classic-Repair-Toolbox.crash.log", "crash"));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Crashed.", zip);

            Assert.Null(outcome.Reference);
            Assert.Empty(Directory.EnumerateFileSystemEntries(this.FeedbackRoot));
            Assert.Contains("Crash log:", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        // SubmissionPathRules, the one containment rule: a zip entry climbing out is never written.
        [Fact]
        public async Task An_entry_climbing_out_of_the_folder_is_not_saved_and_the_mail_says_so()
        {
            string zip = this.ZipFile(("../escaped.txt", "evil"), ("fine.txt", "ok"));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Two files.", zip);

            Assert.Equal(1, outcome.FilesSaved);
            Assert.False(File.Exists(Path.Combine(this.FeedbackRoot, "escaped.txt")));
            Assert.True(File.Exists(Path.Combine(this.Folder, "fine.txt")));
            Assert.Contains("Not saved:", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task A_zip_that_cannot_be_opened_is_still_mailed_with_the_reason()
        {
            string notAZip = Path.Combine(this.thisRoot, "broken.zip");
            File.WriteAllText(notAZip, "not a zip at all");

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Attached something.", notAZip);

            Assert.True(outcome.IsDelivered);
            Assert.Contains("could not be opened", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        // The mail carries a shown file twice (HTML and plain text), and postfix refuses a large
        // mail - so a log beyond the limit is saved like any other file, and the mail says so.
        [Fact]
        public async Task A_log_too_large_to_show_is_saved_instead_and_the_mail_says_so()
        {
            string bigLog = new string('x', (int)FeedbackFlow.ShownBytesLimit + 10);
            string zip = this.ZipFile(("Classic-Repair-Toolbox.log", bigLog));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Long session.", zip);

            Assert.Equal(1, outcome.FilesSaved);
            Assert.True(File.Exists(Path.Combine(this.Folder, "Classic-Repair-Toolbox.log")));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Contains("too large to show here", mail.Body, StringComparison.Ordinal);
            Assert.True(mail.Body.Length < FeedbackFlow.ShownBytesLimit);
        }

        [Fact]
        public async Task With_the_disk_too_full_nothing_is_saved_and_the_mail_says_so()
        {
            string zip = this.ZipFile(("photo.jpg", "jpeg bytes"));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Photo attached.", zip, destination: this.Destination(freeBytes: 100));

            Assert.Null(outcome.Reference);
            Assert.Empty(Directory.EnumerateFileSystemEntries(this.FeedbackRoot));
            Assert.Contains("disk is too full", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        // The mail IS the delivery: one postfix refused reached nobody, and CRT must not say "thank you".
        [Fact]
        public async Task A_mail_that_could_not_be_sent_is_a_failure_CRT_sees()
        {
            this.thisMailer.Delivers = false;

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Hello.");

            Assert.False(outcome.IsDelivered);
            Assert.Equal(502, outcome.StatusCode);
            Assert.False(FeedbackContract.IsSuccess(outcome.StatusCode, outcome.Answer));
        }

        // ###########################################################################################
        // No mail names the folder of a feedback whose mail failed, and CRT keeps the files to send
        // again - so a folder left behind would be a full copy per retry that nobody is told about
        // (code review, 2026-10-04).
        // ###########################################################################################
        [Fact]
        public async Task The_saved_files_are_removed_again_when_the_mail_could_not_be_sent()
        {
            this.thisMailer.Delivers = false;
            string zip = this.ZipFile(("photo.jpg", "jpeg bytes"), ("notes/part list.txt", "U1"));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Photo attached.", zip);

            Assert.Equal(502, outcome.StatusCode);
            Assert.Null(outcome.Reference);
            Assert.Equal(0, outcome.FilesSaved);
            Assert.Empty(Directory.EnumerateFileSystemEntries(this.FeedbackRoot));
        }

        // ###########################################################################################
        // The feedback folder has a TOTAL (code review, 2026-10-04): anonymous senders must not fill
        // the disk down to the reserve the blob store shares. Full, the text still reaches the
        // project owner, with the reason, and nothing is saved.
        // ###########################################################################################
        [Fact]
        public async Task With_the_feedback_folder_full_the_text_is_mailed_and_nothing_is_saved()
        {
            string zip = this.ZipFile(("photo.jpg", "jpeg bytes"));
            FeedbackDestination full = this.Destination() with { MaxStoredBytes = 100, StoredBytes = () => 95 };

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Photo attached.", zip, destination: full);

            Assert.True(outcome.IsDelivered);
            Assert.Null(outcome.Reference);
            Assert.Empty(Directory.EnumerateFileSystemEntries(this.FeedbackRoot));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Contains("FEEDBACK FOLDER FULL", mail.Body, StringComparison.Ordinal);
            Assert.Contains("Photo attached.", mail.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task Below_the_total_the_files_are_saved_as_usual()
        {
            string zip = this.ZipFile(("photo.jpg", "jpeg bytes"));
            FeedbackDestination roomy = this.Destination() with { MaxStoredBytes = 1_000, StoredBytes = () => 95 };

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Photo attached.", zip, destination: roomy);

            Assert.Equal(FeedbackFlowTests.Reference, outcome.Reference);
            Assert.True(File.Exists(Path.Combine(this.Folder, "photo.jpg")));
        }

        // A zip that lies about its sizes is stopped at the total's room, not at the 2 GB a request
        // may otherwise unpack to.
        [Fact]
        public async Task A_zip_unpacking_past_the_room_left_is_stopped_there()
        {
            string zip = this.ZipFile(("first.txt", new string('a', 40)), ("second.txt", new string('b', 40)));
            FeedbackDestination tight = this.Destination() with { MaxStoredBytes = 150, StoredBytes = () => 70 };

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Two files.", zip, destination: tight);

            // 80 bytes declared fits the 80 left, so it is taken - and both fit exactly.
            Assert.Equal(2, outcome.FilesSaved);

            FeedbackDestination tighter = this.Destination() with { MaxStoredBytes = 150, StoredBytes = () => 71 };
            this.thisMailer.Sent.Clear();
            Directory.Delete(this.Folder, recursive: true);

            FeedbackOutcome refused = await this.SendAsync(string.Empty, "Two files.", zip, destination: tighter);

            Assert.Null(refused.Reference);
            Assert.Contains("FEEDBACK FOLDER FULL", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        // The total counts the saved feedback folders - never the zip being read beside them, nor
        // anything else that happens to sit in the folder.
        [Fact]
        public void The_total_counts_only_the_feedback_folders()
        {
            Directory.CreateDirectory(Path.Combine(this.FeedbackRoot, "feedback-aaa", "sub"));
            File.WriteAllBytes(Path.Combine(this.FeedbackRoot, "feedback-aaa", "one.bin"), new byte[30]);
            File.WriteAllBytes(Path.Combine(this.FeedbackRoot, "feedback-aaa", "sub", "two.bin"), new byte[20]);
            Directory.CreateDirectory(Path.Combine(this.FeedbackRoot, "feedback-bbb"));
            File.WriteAllBytes(Path.Combine(this.FeedbackRoot, "feedback-bbb", "three.bin"), new byte[5]);
            File.WriteAllBytes(Path.Combine(this.FeedbackRoot, ".incoming-123.zip"), new byte[1_000]);
            Directory.CreateDirectory(Path.Combine(this.FeedbackRoot, "other"));
            File.WriteAllBytes(Path.Combine(this.FeedbackRoot, "other", "x.bin"), new byte[1_000]);

            Assert.Equal(55L, FeedbackFlow.StoredBytesUnder(this.FeedbackRoot));
        }

        [Fact]
        public void A_folder_that_cannot_be_read_has_an_unknown_total()
        {
            Assert.Null(FeedbackFlow.StoredBytesUnder(Path.Combine(this.thisRoot, "no such folder")));
        }

        // ###########################################################################################
        // A zip whose end record is fine but whose central directory is damaged opens, and throws
        // only when its entries are read - which must be the same "could not be opened" note in the
        // mail, never a 500 on the anonymous route (code review, 2026-10-04).
        // ###########################################################################################
        [Fact]
        public async Task A_zip_with_a_damaged_central_directory_is_still_mailed_with_the_reason()
        {
            byte[] bytes = FeedbackFormReaderTests.Zip(("photo.jpg", "jpeg bytes"));
            int directory = FeedbackFlowTests.CentralDirectoryOffset(bytes);
            bytes[directory] = 0;
            bytes[directory + 1] = 0;
            string zip = Path.Combine(this.thisRoot, "damaged.zip");
            File.WriteAllBytes(zip, bytes);

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Attached something.", zip);

            Assert.True(outcome.IsDelivered);
            Assert.Null(outcome.Reference);
            Assert.Contains("could not be opened", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        // A shown file the zip cannot give back (here an unknown compression method) is said in the
        // mail, and the zip's other files are still saved.
        [Fact]
        public async Task A_log_that_cannot_be_read_is_said_in_the_mail_and_the_other_files_are_still_saved()
        {
            byte[] bytes = FeedbackFormReaderTests.Zip(("Classic-Repair-Toolbox.log", "log text"), ("photo.jpg", "jpeg bytes"));
            int directory = FeedbackFlowTests.CentralDirectoryOffset(bytes);
            bytes[directory + 10] = 99;
            bytes[directory + 11] = 0;
            string zip = Path.Combine(this.thisRoot, "odd-method.zip");
            File.WriteAllBytes(zip, bytes);

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Log attached.", zip);

            Assert.True(outcome.IsDelivered);
            Assert.Equal(1, outcome.FilesSaved);
            Assert.True(File.Exists(Path.Combine(this.Folder, "photo.jpg")));
            Assert.Contains("could not be read from the zip file", Assert.Single(this.thisMailer.Sent).Body, StringComparison.Ordinal);
        }

        // The first central directory header's offset, from the end record (the last 22 bytes of a
        // zip with no comment).
        private static int CentralDirectoryOffset(byte[] zip) =>
            BitConverter.ToInt32(zip, zip.Length - 22 + 16);

        // ###########################################################################################
        // *** A ZIP OF MORE ENTRIES THAN THE SERVER READS IS REFUSED BEFORE IT IS OPENED (code review,
        // 2026-10-04). *** ZipArchive builds an object per central directory record before any of
        // the flow's limits apply, so millions of empty records exhausted the service's memory. The
        // records are counted first; past the limit nothing is shown or saved and the mail says why.
        // ###########################################################################################
        [Fact]
        public async Task A_zip_holding_more_entries_than_the_server_reads_is_refused_and_the_text_still_mailed()
        {
            string zip = this.ZipFile(
                ("Classic-Repair-Toolbox.log", "log"), ("a.txt", "a"), ("b.txt", "b"), ("c.txt", "c"), ("d.txt", "d"));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Many files.", zip, destination: this.Destination() with { MaxEntries = 4 });

            Assert.True(outcome.IsDelivered);
            Assert.Null(outcome.Reference);
            Assert.Empty(Directory.EnumerateFileSystemEntries(this.FeedbackRoot));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Contains("UPLOAD REFUSED: the zip holds more than 4 entries", mail.Body, StringComparison.Ordinal);
            Assert.Contains("Many files.", mail.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("Logfile:", mail.Body, StringComparison.Ordinal);
        }

        // At the limit exactly, the zip is unpacked as usual.
        [Fact]
        public async Task A_zip_holding_exactly_the_entries_the_server_reads_is_unpacked()
        {
            string zip = this.ZipFile(("a.txt", "a"), ("b.txt", "b"), ("c.txt", "c"), ("d.txt", "d"));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Four files.", zip, destination: this.Destination() with { MaxEntries = 4 });

            Assert.Equal(4, outcome.FilesSaved);
        }

        // ###########################################################################################
        // *** THE SHOWN TEXT IS MEASURED BY WHAT IS READ (code review, 2026-10-04). *** Four logs in
        // different folders, each 1 MB but each DECLARING 0 bytes: the declared total never grew, so
        // every one was read whole and the mail carried all of them, twice (HTML and text). Now the
        // first three fill the 3 MB the shown files may take and the fourth is saved instead.
        // ###########################################################################################
        [Fact]
        public async Task Logs_lying_about_their_size_are_shown_no_further_than_the_limit()
        {
            int each = (int)(FeedbackFlow.ShownBytesLimit / 3);
            string zip = Path.Combine(this.thisRoot, "lying.zip");

            File.WriteAllBytes(zip, FeedbackFlowTests.StoredZipDeclaringNoSize(
                Enumerable.Range(1, 4).Select(index => ($"run{index}/Classic-Repair-Toolbox.log", new string('x', each))).ToArray()));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Four sessions.", zip);

            Assert.True(outcome.IsDelivered);
            Assert.Equal(1, outcome.FilesSaved);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.True(mail.Body.Length < FeedbackFlow.ShownBytesLimit + (64 * 1024), $"The mail's text is {mail.Body.Length} characters.");
            Assert.Contains("too large to show here", mail.Body, StringComparison.Ordinal);

            // Three really were read and shown - ZipArchive does not refuse a size that lies, so
            // the limit above is the flow's doing.
            Assert.Equal(3, mail.Body.Split("Logfile:").Length - 1);
        }

        // ###########################################################################################
        // A zip of Stored (uncompressed) entries whose central directory says each holds 0 bytes -
        // what a crafted upload sends. ZipArchive reads a Stored entry's data by its compressed size,
        // which is left true, and reports its Length from the declared size.
        // ###########################################################################################
        private static byte[] StoredZipDeclaringNoSize(params (string Name, string Text)[] entries)
        {
            byte[] bytes;

            using (var stream = new MemoryStream())
            {
                using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
                {
                    foreach ((string name, string text) in entries)
                    {
                        using Stream entry = archive.CreateEntry(name, System.IO.Compression.CompressionLevel.NoCompression).Open();
                        entry.Write(Encoding.UTF8.GetBytes(text));
                    }
                }

                bytes = stream.ToArray();
            }

            int at = FeedbackFlowTests.CentralDirectoryOffset(bytes);

            while (BitConverter.ToUInt32(bytes, at) == 0x02014b50)
            {
                BitConverter.GetBytes(0u).CopyTo(bytes, at + 24);

                at += 46 + BitConverter.ToUInt16(bytes, at + 28) + BitConverter.ToUInt16(bytes, at + 30) + BitConverter.ToUInt16(bytes, at + 32);
            }

            return bytes;
        }

        // ###########################################################################################
        // What an unpack kept is handed on, so FeedbackStorage's remembered folder total grows without
        // walking the folder again - and handed back for a mail that failed, whose files are deleted
        // again. (Until the code review of 2026-10-04 nothing was handed on until the mail had gone;
        // see the next test for why that was too late.)
        // ###########################################################################################
        [Fact]
        public async Task The_bytes_an_unpack_kept_are_handed_on_and_handed_back_for_a_failed_mail()
        {
            var kept = new List<long>();
            var removed = new List<long>();
            FeedbackDestination destination = this.Destination() with { Saved = kept.Add, Removed = removed.Add };

            await this.SendAsync(string.Empty, "Files.", this.ZipFile(("a.txt", "12345"), ("b.txt", "123")), destination: destination);

            Assert.Equal([8L], kept);
            Assert.Empty(removed);

            Directory.Delete(this.Folder, recursive: true);
            this.thisMailer.Delivers = false;

            await this.SendAsync(string.Empty, "Files.", this.ZipFile(("a.txt", "12345")), destination: destination);

            Assert.Equal([8L, 5L], kept);
            Assert.Equal([5L], removed);
            Assert.False(Directory.Exists(this.Folder));
        }

        // ###########################################################################################
        // *** A FEEDBACK WAITING ON ITS MAIL ALREADY COUNTS AGAINST THE FOLDER'S TOTAL (code review,
        // 2026-10-04). *** Its bytes were added only after the mail - outside the one-unpack gate - so
        // the next feedback could take the gate while the first waited on the mail server, read the
        // remembered total without them, and pass the cap too. Here the first feedback's mail is held
        // until the second has been decided.
        // ###########################################################################################
        [Fact]
        public async Task A_feedback_waiting_on_its_mail_already_counts_against_the_folder_total()
        {
            var storage = new FeedbackStorage(() => 0);
            var mailer = new HeldMailer();
            int made = 0;

            FeedbackDestination destination = this.Destination() with
            {
                NewReference = () => $"feedback-Held{++made}",
                MaxStoredBytes = 10,
                StoredBytes = storage.StoredBytes,
                UnpackGate = storage.UnpackGate,
                Saved = storage.Saved,
                Removed = storage.Removed
            };

            Task<FeedbackOutcome> first = FeedbackFlow.HandleAsync(
                new FeedbackForm(string.Empty, "First.", "3.0.0", this.ZipFile(("a.txt", "12345678"))), destination, mailer, CancellationToken.None);

            await mailer.FirstArrived.WaitAsync(TimeSpan.FromSeconds(10));

            FeedbackOutcome second = await FeedbackFlow.HandleAsync(
                new FeedbackForm(string.Empty, "Second.", "3.0.0", this.ZipFile(("b.txt", "12345"))), destination, mailer, CancellationToken.None)
                .WaitAsync(TimeSpan.FromSeconds(10));

            mailer.Release();
            FeedbackOutcome firstOutcome = await first.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.Equal(1, firstOutcome.FilesSaved);
            Assert.Equal(0, second.FilesSaved);
            Assert.Contains("FEEDBACK FOLDER FULL", mailer.Sent[^1].Body, StringComparison.Ordinal);
            Assert.Equal(8L, storage.StoredBytes());
        }

        // Holds the first mail until released, so a test can act while a feedback waits on postfix.
        private sealed class HeldMailer : IEmailSender
        {
            private readonly TaskCompletionSource thisFirstArrived = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly TaskCompletionSource thisRelease = new(TaskCreationOptions.RunContinuationsAsynchronously);
            private int thisCount;

            public List<EmailMessage> Sent { get; } = [];

            public Task FirstArrived => this.thisFirstArrived.Task;

            public void Release() => this.thisRelease.TrySetResult();

            public async Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
            {
                lock (this.Sent)
                    this.Sent.Add(message);

                if (Interlocked.Increment(ref this.thisCount) == 1)
                {
                    this.thisFirstArrived.TrySetResult();
                    await this.thisRelease.Task;
                }

                return true;
            }
        }

        // ###########################################################################################
        // *** A SHOWN FILE IS MEASURED BY WHAT IT TAKES IN THE MAIL (code review, 2026-10-04). *** A
        // settings file full of double quotes stays well under the limit by characters, but the HTML
        // body carries each one as &quot; - six bytes - and the mail went past postfix's 10 MB, so
        // the feedback answered 502 on every try. Now it is saved instead, and the mail says so.
        // ###########################################################################################
        [Fact]
        public async Task A_file_too_large_to_show_once_encoded_for_the_mail_is_saved_instead()
        {
            string quotes = new string('"', (int)(FeedbackFlow.ShownBytesLimit / 2));
            string zip = this.ZipFile(("Classic-Repair-Toolbox.settings.json", quotes));

            FeedbackOutcome outcome = await this.SendAsync(string.Empty, "Settings attached.", zip);

            Assert.True(outcome.IsDelivered);
            Assert.Equal(1, outcome.FilesSaved);
            Assert.True(File.Exists(Path.Combine(this.Folder, "Classic-Repair-Toolbox.settings.json")));

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Contains("too large to show here", mail.Body, StringComparison.Ordinal);
            Assert.True(Encoding.UTF8.GetByteCount(mail.HtmlBody!) < 64 * 1024, $"The HTML is {mail.HtmlBody!.Length} characters.");
        }

        // The larger of the two bodies' bytes: the HTML's entities, or the plain text's CRLF.
        [Theory]
        [InlineData("abc", 3)]
        [InlineData("a\"b", 8)]
        [InlineData("x\ny", 4)]
        [InlineData("x\r\ny", 4)]
        [InlineData("<\n", 5)]
        public void A_shown_texts_mail_bytes_are_the_larger_of_its_HTML_and_its_plain_text(string text, long bytes)
        {
            Assert.Equal(bytes, FeedbackFlow.MailBytes(text));
        }

        // ###########################################################################################
        // *** ONE UNPACK AT A TIME (code review, 2026-10-04). *** Each held a thread for up to 2 GB of
        // synchronous file work, so a few at once tied up a small box. A feedback with files waits
        // for the one before it; text alone never touches the gate.
        // ###########################################################################################
        [Fact]
        public async Task A_feedback_with_files_waits_for_the_unpack_before_it()
        {
            var gate = new SemaphoreSlim(1, 1);
            FeedbackDestination destination = this.Destination() with { UnpackGate = gate };

            await gate.WaitAsync();

            Task<FeedbackOutcome> withFiles = this.SendAsync(string.Empty, "Files.", this.ZipFile(("a.txt", "a")), destination: destination);
            FeedbackOutcome textAlone = await this.SendAsync(string.Empty, "Just text.", destination: destination).WaitAsync(TimeSpan.FromSeconds(10));

            await Task.Delay(100);
            Assert.False(withFiles.IsCompleted);
            Assert.True(textAlone.IsDelivered);

            gate.Release();

            FeedbackOutcome outcome = await withFiles.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(1, outcome.FilesSaved);
            Assert.Equal(1, gate.CurrentCount);
        }

        [Fact]
        public async Task Nothing_at_all_is_refused_and_mails_nothing()
        {
            FeedbackOutcome outcome = await this.SendAsync("user@example.com", "   ");

            Assert.Equal(400, outcome.StatusCode);
            Assert.Empty(this.thisMailer.Sent);
        }

        // The version reaches the subject line, where a line break would be a header injection.
        [Fact]
        public async Task The_version_is_cleaned_before_it_reaches_the_subject()
        {
            await this.SendAsync(string.Empty, "Hello.", version: "2.6.0\r\nBcc: someone@example.com");

            string subject = Assert.Single(this.thisMailer.Sent).Subject;

            Assert.DoesNotContain('\r', subject);
            Assert.DoesNotContain('\n', subject);
            Assert.StartsWith("Feedback from CRT 2.6.0", subject, StringComparison.Ordinal);
        }

        // A 250 MB zip lands on the disk before it is unpacked, so room is asked for first - and an
        // unknown length or free space is room (the request limit still bounds it).
        [Theory]
        [InlineData(100L, 10_000L, 1_000L, true)]
        [InlineData(9_500L, 10_000L, 1_000L, false)]
        [InlineData(null, 10L, 1_000L, true)]
        [InlineData(100L, null, 1_000L, true)]
        public void An_upload_is_taken_only_while_the_disk_keeps_its_reserve(long? length, long? free, long reserve, bool expected)
        {
            Assert.Equal(expected, FeedbackFlow.HasRoomForUpload(length, free, reserve));
        }

        // ###########################################################################################
        // The project owner deletes old feedback through a network share (owner answer, 2026-10-03),
        // so what the service saves is group-writable, folders setgid so they keep the share's group.
        // Linux only - the CI runner checks it; Windows has no such modes.
        // ###########################################################################################
        [Fact]
        public async Task Saved_folders_and_files_are_group_writable_so_the_share_can_delete_them()
        {
            // An explicit guard, which the platform analyzer (CA1416) understands; SkipWhen it does not.
            if (OperatingSystem.IsWindows())
            {
                Assert.Skip("Unix file modes - checked on the Linux CI runner.");
                return;
            }

            string zip = this.ZipFile(("Photos/board.jpg", "jpeg"), ("note.txt", "hello"));

            await this.SendAsync(string.Empty, "Files.", zip);

            Assert.Equal(FeedbackFlow.SharedFolderMode, File.GetUnixFileMode(this.Folder));
            Assert.Equal(FeedbackFlow.SharedFolderMode, File.GetUnixFileMode(Path.Combine(this.Folder, "Photos")));
            Assert.Equal(FeedbackFlow.SharedFileMode, File.GetUnixFileMode(Path.Combine(this.Folder, "Photos", "board.jpg")));
            Assert.Equal(FeedbackFlow.SharedFileMode, File.GetUnixFileMode(Path.Combine(this.Folder, "note.txt")));
            Assert.True(FeedbackFlow.SharedFileMode.HasFlag(UnixFileMode.GroupWrite));
        }

        [Fact]
        public void A_reference_has_the_shape_it_has_always_had()
        {
            string reference = FeedbackFlow.NewReference();

            Assert.StartsWith(FeedbackFlow.ReferencePrefix, reference, StringComparison.Ordinal);
            Assert.Equal(FeedbackFlow.ReferencePrefix.Length + 16, reference.Length);
            Assert.NotEqual(reference, FeedbackFlow.NewReference());
        }
    }
}
