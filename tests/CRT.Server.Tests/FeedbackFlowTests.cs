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

        private Task<FeedbackOutcome> SendAsync(string email, string feedback, string? zip = null, string version = "2.6.0", FeedbackDestination? destination = null) =>
            FeedbackFlow.HandleAsync(new FeedbackForm(email, feedback, version, zip), destination ?? this.Destination(), this.thisMailer, CancellationToken.None);

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
