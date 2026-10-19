using System.IO.Compression;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Feedback
{
    // ###########################################################################################
    // *** FEEDBACK FROM CRT'S FEEDBACK TAB (owner request, 2026-10-03: "have the new backend
    // server handle that"). *** What a feedback has always done, decided here and tested against a
    // temporary folder and a fake mailer:
    //
    //   - CRT's own text files in the zip (FeedbackContract.InlineFiles - the log above all) are
    //     SHOWN in the mail rather than saved;
    //   - so are the pictures pasted or dropped into the text (2026-10-15, TryReadPicture), inside
    //     the sender's quoted words where each sat - one the mail cannot show is saved instead;
    //   - every other file is saved under "<FeedbackRoot>/feedback-<16 random characters>", the
    //     folder name feedback has always been saved under, which the mail gives as its "Internal
    //     reference" - the project owner opens it from the network share, as before ("just keep
    //     exact same behaviour");
    //   - one mail goes to the project owner, now as HTML (EmailTemplates.Feedback).
    //
    // *** WHAT IS DELIBERATELY DIFFERENT FROM BEFORE. ***
    //   - The mail is FROM the service, with Reply-To the sender. It used to be sent from the
    //     sender's own address, which fails SPF at the sender's provider - mail that lands in spam.
    //   - Every saved path goes through SubmissionPathRules (the one containment rule): a zip
    //     entry naming "../" or an absolute path is skipped and said in the mail, never written.
    //   - Limits that used to be left to the web server's defaults: how many entries a zip may hold
    //     before it is opened at all (ZipCentralDirectory), how many files and how many bytes it may
    //     unpack to (a 60 MB zip can claim gigabytes), and the disk reserve the blob store keeps.
    //   - The shown files together are capped (ShownBytesLimit), counted by what is READ rather
    //     than what the zip declares, and by the bytes each takes in the mail ONCE ENCODED for it
    //     (MailBytes - in the HTML a " is six bytes); a log beyond it is SAVED instead, and the
    //     mail says so. The mail carries the HTML and the plain text, so a large log is in it
    //     twice, transfer-encoded on top, and postfix refuses a mail over 10 MB by default.
    //   - One unpack at a time, and the folder's total remembered between walks (FeedbackStorage).
    //   - A mail that could not be sent is a failure the sender sees (IEmailSender's answer), so
    //     the text stays in CRT to send again. That was said before too ("Mail sending failed").
    // ###########################################################################################
    public static class FeedbackFlow
    {
        // How many feedbacks one address may send an hour (in memory - FeedbackEndpoints).
        public const int MaxPerAddressPerHour = 10;

        // How much of CRT's own text files the mail shows, all together - in the bytes they take in
        // EACH of the mail's two bodies (MailBytes), so both bodies together stay under twice this,
        // and under postfix's 10 MB once transfer-encoded.
        public const long ShownBytesLimit = 3L * 1024 * 1024;

        // What a zip may unpack to - room for a whole board's folder, or several (the largest shipped
        // board is ~2,200 files, the whole data tree ~11,000), and a 250 MB zip of pictures, which
        // barely compress. A zip claiming more is refused before anything is written; one lying
        // about its sizes is stopped while it is written.
        public const int MaxSavedFiles = 20_000;
        public const long MaxSavedBytes = 2L * 1024 * 1024 * 1024;

        // How many central directory records a zip may hold before ZipArchive is let near it
        // (ZipCentralDirectory): room for MaxSavedFiles plus a folder entry each, and a bound on the
        // objects ZipArchive builds - ~10 MB of them rather than the gigabytes a zip of nothing but
        // empty records would make (code review, 2026-10-04).
        public const int MaxZipEntries = 2 * FeedbackFlow.MaxSavedFiles;

        // How many saved files the mail lists by name (200, as it always has) - as many of them as
        // fit ListRoomBytes.
        public const int MaxListedFiles = 200;

        // ###########################################################################################
        // *** A FEEDBACK MAIL MAY BE UP TO 20 MB (owner decision, 2026-10-15: "Please allow for a
        // total up to 20MB in the mail - I have received many mails with this size, and if my local
        // Postfix should change something, then I can do that"). *** The text goes in first; the
        // pictures CRT put into the text share what is left (PictureRoom), and one that no longer
        // fits stays a marker line and is saved with the files. Postfix's own default is 10 MB, so
        // its message_size_limit is raised to match (INSTALLING.md, "Mail") - a mail it refuses
        // reaches nobody, and CRT says the server could not pass it on.
        // ###########################################################################################
        public const long MaxMailBytes = 20L * 1024 * 1024;

        // What a mail carries besides its two text bodies and its pictures - headers, the HTML's
        // frame, the notes, the list of saved files - kept free of both. The notes and the list
        // take ListRoomBytes each in each body (ListWithin), transfer-encoded: 171 KB of it.
        public const long MailOverheadBytes = 256L * 1024;

        // What the notes, and the list of saved files, may each take in ONE body (MailBytes, with a
        // line's markup). Both come after the pictures' room is fixed, from a zip the sender made:
        // thousands of refused entries made megabytes of notes (code review, 2026-10-16).
        public const long ListRoomBytes = 32L * 1024;

        // What a picture takes in the mail besides its bytes, in the pictures' room's own units: its
        // <img> in the HTML and its own part's headers - generously (code review, 2026-10-16).
        public const long PictureOverheadBytes = 1024;

        // A listed line's markup in the HTML: its paragraph's <p style="..."> and </p>.
        private const long ListLineBytes = 40;

        public const string ReferencePrefix = "feedback-";

        // The alphabet references have always used: no 0/O, 1/l/I or v/V, so a reference read aloud
        // is unambiguous.
        private const string ReferenceAlphabet = "23456789abcdefghkmnpqrstuwxyzABCDEFGHJKMNPQRSTUWXYZ";

        // ###########################################################################################
        // Is there room to take an upload of `contentLength` bytes and still keep the disk reserve
        // the blob store keeps (ServerOptions.MinimumFreeDiskBytes)? Asked before reading a byte -
        // a 250 MB zip lands on the disk before it is unpacked. An unknown length or free space is
        // room: the request limit still bounds it, and a probe error must not refuse everybody.
        // ###########################################################################################
        public static bool HasRoomForUpload(long? contentLength, long? freeBytes, long minimumFreeBytes) =>
            contentLength is not long length || freeBytes is not long free || free - length >= minimumFreeBytes;

        public static string NewReference() =>
            FeedbackFlow.ReferencePrefix + RandomNumberGenerator.GetString(FeedbackFlow.ReferenceAlphabet, 16);

        // ###########################################################################################
        // How many bytes of pictures fit in a mail of `maxMailBytes` whose text takes
        // `textBytesPerBody` in EACH of its two bodies (MailBytes). Everything travels transfer-
        // encoded: the text counted as base64 (a third larger, the most the mail system makes of
        // it), the pictures as base64 in lines of 76 (4/3, and 78/76 for the line ends). So the
        // pictures, the text and MailOverheadBytes together come to no more than `maxMailBytes`.
        // ###########################################################################################
        public static long PictureRoom(long maxMailBytes, long textBytesPerBody)
        {
            long text = 2 * Math.Max(0, textBytesPerBody) * 4 / 3;
            long left = maxMailBytes - FeedbackFlow.MailOverheadBytes - text;

            return left <= 0 ? 0 : left * 3 / 4 * 76 / 78;
        }

        // ###########################################################################################
        // What the saved feedback under `root` takes: every file in its "feedback-..." folders (not
        // the zip being read, which sits beside them). Null when the folder cannot be read - an
        // unknown total is room, like an unknown free space, so a probe error refuses nobody.
        // ###########################################################################################
        public static long? StoredBytesUnder(string root)
        {
            try
            {
                long total = 0;

                foreach (string folder in Directory.EnumerateDirectories(root, FeedbackFlow.ReferencePrefix + "*"))
                {
                    foreach (string file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
                        total += new FileInfo(file).Length;
                }

                return total;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }

        public static async Task<FeedbackOutcome> HandleAsync(
            FeedbackForm form,
            FeedbackDestination destination,
            IEmailSender mailer,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(form);
            ArgumentNullException.ThrowIfNull(destination);
            ArgumentNullException.ThrowIfNull(mailer);

            string sender = FeedbackFlow.CleanAddress(form.Email);
            string feedback = form.Feedback.Trim();
            string version = FeedbackFlow.CleanVersion(form.Version);

            var shown = new List<FeedbackShownFile>();
            var pictures = new List<MailLinkedImage>();
            var saved = new List<string>();
            var notes = new List<string>();
            string? reference = null;
            long bytesSaved = 0;

            if (form.AttachmentPath is not null)
            {
                // One unpack at a time (FeedbackStorage), on the pool rather than the request's own
                // thread - it is synchronous file work, up to MaxSavedBytes of it.
                if (destination.UnpackGate is SemaphoreSlim gate)
                    await gate.WaitAsync(cancellationToken);

                try
                {
                    IReadOnlyDictionary<string, int> widths = FeedbackComposition.ReadPictureWidths(form.PictureWidths);

                    (reference, bytesSaved) = await Task.Run(
                        () => FeedbackFlow.Unpack(form.AttachmentPath, destination, feedback, widths, shown, pictures, saved, notes),
                        CancellationToken.None);

                    // *** COUNTED BEFORE THE GATE OPENS (code review, 2026-10-04). *** Counted after
                    // the mail, the next feedback could take the gate while this one waited on the
                    // mail server, read a folder total without these bytes, and pass the folder's
                    // cap as well - two unpacks over it at once. Taken off again below if no mail
                    // names them.
                    if (reference is not null)
                        destination.Saved?.Invoke(bytesSaved);
                }
                finally
                {
                    destination.UnpackGate?.Release();
                }
            }

            if (feedback.Length == 0 && shown.Count == 0 && pictures.Count == 0 && saved.Count == 0 && notes.Count == 0)
                return new FeedbackOutcome(StatusCodes.Status400BadRequest, "There is nothing to send.", null, 0);

            (List<string> listedFiles, int filesNotListed) = FeedbackFlow.ListWithin(saved, FeedbackFlow.MaxListedFiles, FeedbackFlow.ListRoomBytes);
            (List<string> listedNotes, int notesNotListed) = FeedbackFlow.ListWithin(notes, int.MaxValue, FeedbackFlow.ListRoomBytes);

            if (notesNotListed > 0)
                listedNotes.Add($"... and {notesNotListed} more note(s) not listed.");

            EmailMessage message = EmailTemplates.Feedback(
                destination.ToAddress,
                version,
                sender,
                feedback,
                reference,
                pictures,
                shown,
                listedFiles,
                filesNotListed,
                listedNotes) with
            {
                // Not kept in the outbox (2026-10-13): CRT is told it failed and sends it again, so a
                // copy going out later as well would arrive twice. See EmailMessage.
                KeepUntilSent = false
            };

            if (!await mailer.SendAsync(message, cancellationToken))
            {
                // No mail names the folder, and CRT keeps the files to send again - so a folder left
                // here would be a full copy per retry that nobody is ever told about (code review,
                // 2026-10-04).
                // Taken off the folder's total BEFORE the folder goes: a walk in between then counts
                // the folder that is about to go, which is too much for a few minutes, never too
                // little.
                if (reference is not null)
                {
                    destination.Removed?.Invoke(bytesSaved);
                    FeedbackFlow.DeleteQuietly(Path.Combine(destination.Root, reference));
                }

                return new FeedbackOutcome(
                    StatusCodes.Status502BadGateway,
                    "Mail sending failed - the feedback did not reach anybody. Please try again later.",
                    null,
                    0);
            }

            return new FeedbackOutcome(StatusCodes.Status200OK, FeedbackContract.SuccessAnswer, reference, saved.Count);
        }

        // ###########################################################################################
        // Opens the zip, shows CRT's own files, reads the pictures the text names into the mail, and
        // saves the rest. Returns the reference (the folder's name) when anything was saved, with
        // the bytes saved, and fills in what the mail says.
        // ###########################################################################################
        private static (string? Reference, long BytesSaved) Unpack(
            string zipPath,
            FeedbackDestination destination,
            string feedback,
            IReadOnlyDictionary<string, int> widths,
            List<FeedbackShownFile> shown,
            List<MailLinkedImage> pictures,
            List<string> saved,
            List<string> notes)
        {
            // *** COUNTED BEFORE ZipArchive SEES IT (code review, 2026-10-04). *** Reading Entries
            // builds an object per central directory record before any limit below applies - a zip
            // of millions of empty records would exhaust the service's memory first.
            if (ZipCentralDirectory.CountRecords(zipPath, destination.MaxEntries) is long records && records > destination.MaxEntries)
            {
                notes.Add($"UPLOAD REFUSED: the zip holds more than {destination.MaxEntries} entries, more than the server reads. None of its files were shown or saved.");
                return (null, 0);
            }

            ZipArchive archive;

            try
            {
                archive = ZipFile.OpenRead(zipPath);
            }
            catch (InvalidDataException ex)
            {
                notes.Add($"SERVER ERROR: the attached zip file could not be opened ({ex.Message}).");
                return (null, 0);
            }

            using (archive)
            {
                var toSave = new List<ZipArchiveEntry>();
                long shownBytes = 0;

                // The pictures the text names (marker lines), in the order it names them, and the
                // zip's entry for each - read once the text's share of the mail is known.
                IReadOnlyList<string> named = FeedbackComposition.MarkedNames(feedback);
                var wanted = new HashSet<string>(named, StringComparer.OrdinalIgnoreCase);
                var candidates = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);

                // The central directory is read here, not by OpenRead - so a damaged one throws
                // here, and is the same note rather than a 500 (code review, 2026-10-04).
                IReadOnlyCollection<ZipArchiveEntry> entries;

                try
                {
                    entries = archive.Entries;
                }
                catch (InvalidDataException ex)
                {
                    notes.Add($"SERVER ERROR: the attached zip file could not be opened ({ex.Message}).");
                    return (null, 0);
                }

                foreach (ZipArchiveEntry entry in entries)
                {
                    // A folder entry carries nothing; its files arrive as their own entries.
                    if (entry.FullName.EndsWith('/') || entry.FullName.EndsWith('\\'))
                        continue;

                    FeedbackInlineFile? inline = FeedbackContract.InlineFiles.FirstOrDefault(file =>
                        string.Equals(file.FileName, entry.Name, StringComparison.Ordinal));

                    if (inline is not null)
                    {
                        // ###########################################################################
                        // *** MEASURED BY WHAT IS READ, NOT BY WHAT THE ZIP DECLARES (code review,
                        // 2026-10-04). *** A size the zip states can be 0 for 3 MB of text, and eighty
                        // such logs in different folders were each read whole while the total stayed
                        // at nothing. Each is read no further than the room the shown files have
                        // left; one that runs past it is saved instead.
                        //
                        // *** AND BY WHAT IT TAKES IN THE MAIL (code review, 2026-10-04). *** What is
                        // read is characters, but the HTML body carries each " as &quot; - so a
                        // settings file full of quotes, under the room by characters, made a mail
                        // postfix refused (502, every retry alike). Held to the room by MailBytes.
                        // ###########################################################################
                        long room = FeedbackFlow.ShownBytesLimit - shownBytes;
                        string size = FeedbackFlow.Megabytes(entry.Length);

                        if (entry.Length <= room)
                        {
                            try
                            {
                                if (FeedbackFlow.TryReadText(entry, room, out string text, out _))
                                {
                                    long inMail = FeedbackFlow.MailBytes(text);

                                    if (inMail <= room)
                                    {
                                        shownBytes += inMail;
                                        shown.Add(new FeedbackShownFile(inline.Title, text));
                                        continue;
                                    }

                                    size = FeedbackFlow.Megabytes(inMail) + " in the mail";
                                }
                                else
                                {
                                    size = "over " + FeedbackFlow.Megabytes(room);
                                }
                            }
                            catch (Exception ex) when (ex is InvalidDataException or IOException)
                            {
                                notes.Add($"{inline.Title} could not be read from the zip file ({ex.Message}).");
                                continue;
                            }
                        }

                        notes.Add($"{inline.Title} ({size}) was too large to show here - it is saved with the files instead.");
                    }
                    else if (FeedbackFlow.IsPictureOfTheText(entry, wanted) && candidates.TryAdd(entry.Name, entry))
                    {
                        continue;
                    }

                    toSave.Add(entry);
                }

                // ###################################################################################
                // The pictures, in the order the text names them, share what the text leaves of the
                // mail (PictureRoom) - so a picture early in the text is never pushed out by one
                // further down. One shown is NOT saved as well (owner report, 2026-10-15: "they are
                // inline so they should not be uploaded to server"), so a feedback of words and
                // pictures alone leaves nothing behind - no folder, no reference. One not shown is
                // saved, and the mail says why.
                // ###################################################################################
                long pictureRoom = FeedbackFlow.PictureRoom(destination.MaxMailBytes, FeedbackFlow.QuotedMailBytes(feedback) + shownBytes);

                foreach (string name in named)
                {
                    if (!candidates.Remove(name, out ZipArchiveEntry? entry))
                        continue;

                    if (FeedbackFlow.TryReadPicture(entry, pictureRoom - FeedbackFlow.PictureOverheadBytes, pictures.Count, widths, notes) is { } picture)
                    {
                        pictures.Add(picture);
                        pictureRoom -= picture.Bytes.Length + FeedbackFlow.PictureOverheadBytes;
                    }
                    else
                    {
                        toSave.Add(entry);
                    }
                }

                if (toSave.Count == 0)
                    return (null, 0);

                if (toSave.Count > FeedbackFlow.MaxSavedFiles)
                {
                    notes.Add($"UPLOAD REFUSED: the zip holds {toSave.Count} files, more than the {FeedbackFlow.MaxSavedFiles} the server unpacks. None of them were saved.");
                    return (null, 0);
                }

                long declared = toSave.Sum(entry => entry.Length);

                if (declared > FeedbackFlow.MaxSavedBytes)
                {
                    notes.Add($"UPLOAD REFUSED: the files unpack to {FeedbackFlow.Megabytes(declared)}, more than the {FeedbackFlow.Megabytes(FeedbackFlow.MaxSavedBytes)} the server unpacks. None of them were saved.");
                    return (null, 0);
                }

                long? free = destination.FreeBytes();

                if (free is long freeBytes && freeBytes - declared < destination.MinimumFreeBytes)
                {
                    notes.Add("SERVER ERROR: the server's disk is too full to save the attached files. None of them were saved.");
                    return (null, 0);
                }

                // *** THE FEEDBACK FOLDER HAS A TOTAL (code review, 2026-10-04). *** Without one,
                // anonymous senders could fill the disk down to the reserve the blob store shares,
                // and contributors' uploads would then be refused until somebody cleaned up by hand.
                // Full, nothing is saved and the mail says so - the project owner reads that mail,
                // and is the one who deletes old feedback.
                long limit = FeedbackFlow.MaxSavedBytes;

                if (destination.MaxStoredBytes > 0 && destination.StoredBytes?.Invoke() is long stored)
                {
                    long room = destination.MaxStoredBytes - stored;

                    if (declared > room)
                    {
                        notes.Add(
                            $"FEEDBACK FOLDER FULL: saved feedback already takes {FeedbackFlow.Megabytes(stored)} of the " +
                            $"{FeedbackFlow.Megabytes(destination.MaxStoredBytes)} allowed (FeedbackMaxStoredBytes). None of the " +
                            "attached files were saved - delete old feedback folders to make room.");
                        return (null, 0);
                    }

                    limit = Math.Min(limit, room);
                }

                string reference = FeedbackFlow.CreateFolder(destination);
                string folder = Path.Combine(destination.Root, reference);
                long written = 0;

                foreach (ZipArchiveEntry entry in toSave)
                {
                    string relative = entry.FullName.Replace('\\', '/');

                    if (!SubmissionPathRules.TryResolve(folder, relative, out string target, out string reason))
                    {
                        notes.Add($"Not saved: {reason}");
                        continue;
                    }

                    try
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(target)!);

                        long? copied = FeedbackFlow.CopyCapped(entry, target, limit - written);

                        if (copied is null)
                        {
                            File.Delete(target);
                            notes.Add($"UPLOAD STOPPED: the files unpack to more than {FeedbackFlow.Megabytes(limit)}; the rest were not saved.");
                            break;
                        }

                        written += copied.Value;
                        saved.Add(relative);
                    }
                    catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException)
                    {
                        notes.Add($"Not saved: [{relative}] ({ex.Message})");
                    }
                }

                if (saved.Count > 0)
                {
                    FeedbackFlow.ShareWithGroup(folder);
                    return (reference, written);
                }

                FeedbackFlow.DeleteQuietly(folder);
                return (null, 0);
            }
        }

        // ###########################################################################################
        // *** THE PROJECT OWNER DELETES OLD FEEDBACK THROUGH A NETWORK SHARE (owner answer,
        // 2026-10-03). *** The service's default modes (rwxr-xr-x, rw-r--r--) let only crt-server
        // delete what it saved, so every folder and file of a feedback is made GROUP-writable here,
        // and the folder they sit in is the service's, group crt-server, with the setgid bit
        // (INSTALLING.md, "Folders and permissions") - so they inherit that group, which the
        // share's user is put in when it is not root (owner decision, 2026-10-03: no group of its
        // own). The service owns the folder but runs as group crt-data, so as it sets these modes
        // the kernel may drop the setgid bit from SUBfolders - harmless, every file is in place by
        // then. Linux only; Windows has no such modes. A file left as it was is still readable,
        // so a failure here is not a failure of the feedback.
        // ###########################################################################################
        internal const UnixFileMode SharedFolderMode =
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |
            UnixFileMode.GroupRead | UnixFileMode.GroupWrite | UnixFileMode.GroupExecute |
            UnixFileMode.OtherRead | UnixFileMode.OtherExecute |
            UnixFileMode.SetGroup;

        internal const UnixFileMode SharedFileMode =
            UnixFileMode.UserRead | UnixFileMode.UserWrite |
            UnixFileMode.GroupRead | UnixFileMode.GroupWrite |
            UnixFileMode.OtherRead;

        internal static void ShareWithGroup(string folder)
        {
            if (OperatingSystem.IsWindows())
                return;

            try
            {
                File.SetUnixFileMode(folder, FeedbackFlow.SharedFolderMode);

                foreach (string directory in Directory.EnumerateDirectories(folder, "*", SearchOption.AllDirectories))
                    File.SetUnixFileMode(directory, FeedbackFlow.SharedFolderMode);

                foreach (string file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
                    File.SetUnixFileMode(file, FeedbackFlow.SharedFileMode);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Readable as it is - only deleting it through the share needs the change.
            }
        }

        // The folder for one feedback's files - a new name every time.
        private static string CreateFolder(FeedbackDestination destination)
        {
            for (int attempt = 0; ; attempt++)
            {
                string reference = destination.NewReference();
                string folder = Path.Combine(destination.Root, reference);

                if (!Directory.Exists(folder) || attempt >= 5)
                {
                    Directory.CreateDirectory(folder);
                    return reference;
                }
            }
        }

        // ###########################################################################################
        // *** A PICTURE IN THE TEXT IS SHOWN IN THE MAIL (owner request, 2026-10-15). *** CRT's
        // Feedback tab puts a marker line "[pasted-image-1.png]" where a picture was pasted or
        // dropped into the text, and the picture itself at the zip's root under that name
        // (CRT.Data's FeedbackComposition, the one rule). Such an entry goes into the mail as a
        // linked image, shown where its line sits at the width its sender gave it - when its
        // contents are the image its name says (the submission rule, by signature - a zip is what
        // the sender made it), and it fits what is left of the mail's room for pictures. One shown
        // is not saved; one not shown is saved with the files and the mail says why, so it reaches
        // somebody either way. Only an image's name comes here (IsPictureOfTheText).
        // ###########################################################################################
        private static MailLinkedImage? TryReadPicture(ZipArchiveEntry entry, long room, int index, IReadOnlyDictionary<string, int> widths, List<string> notes)
        {
            string name = entry.Name;

            if (entry.Length > room)
            {
                notes.Add($"{name} ({FeedbackFlow.Megabytes(entry.Length)}) is too large to show in the mail - it is saved with the files instead.");
                return null;
            }

            byte[]? bytes;

            try
            {
                using Stream source = entry.Open();
                using var target = new MemoryStream();
                bytes = FeedbackFlow.CopyCapped(source, target, room) is null ? null : target.ToArray();
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException)
            {
                notes.Add($"{name} could not be read from the zip file ({ex.Message}) - it is not shown in the mail.");
                return null;
            }

            if (bytes is null)
            {
                notes.Add($"{name} is too large to show in the mail - it is saved with the files instead.");
                return null;
            }

            if (!SubmissionContentRules.Matches(name, bytes, out _))
            {
                notes.Add($"{name} is not the image its name says - it is saved with the files, not shown in the mail.");
                return null;
            }

            // The type served for the same file from the review routes - the server's one allowlist.
            return new MailLinkedImage(
                $"picture-{index + 1}@crt-feedback", name, ReviewAssetLocator.ContentTypeFor(name), bytes,
                widths.TryGetValue(name, out int width) ? width : null);
        }

        // ###########################################################################################
        // An IMAGE at the zip's ROOT that the text names - a marker names a file name, never a path.
        // Any other file a marker names is the sender's own attachment, saved as ever and never
        // called a picture the mail could not show (code review, 2026-10-16: "[notes.txt]" on a
        // line of its own did).
        // ###########################################################################################
        private static bool IsPictureOfTheText(ZipArchiveEntry entry, HashSet<string> wanted) =>
            !entry.FullName.Contains('/') && !entry.FullName.Contains('\\') &&
            wanted.Contains(entry.Name) && ImageFileTypes.IsDisplayable(entry.Name);

        // An entry's bytes into a NEW file, at most `limit` of them; null when there were more.
        private static long? CopyCapped(ZipArchiveEntry entry, string target, long limit)
        {
            using Stream source = entry.Open();
            using var file = new FileStream(target, FileMode.CreateNew, FileAccess.Write);

            return FeedbackFlow.CopyCapped(source, file, limit);
        }

        // At most `limit` bytes of a stream into `target`, whatever the zip claims - how many, or
        // null when there were more.
        private static long? CopyCapped(Stream source, Stream target, long limit)
        {
            var buffer = new byte[81920];
            long total = 0;
            int read;

            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;

                if (total > limit)
                    return null;

                target.Write(buffer, 0, read);
            }

            return total;
        }

        // ###########################################################################################
        // A shown file's text: UTF-8, its control characters other than tabs and line breaks dropped.
        // False when it runs past `limit` characters - read no further than one buffer beyond it,
        // whatever size the zip claims - with `read` how far it got. Characters stand in for bytes,
        // which they are for CRT's own (ASCII) logs.
        // ###########################################################################################
        private static bool TryReadText(ZipArchiveEntry entry, long limit, out string text, out long read)
        {
            using Stream source = entry.Open();
            using var reader = new StreamReader(source, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);

            var builder = new StringBuilder();
            var buffer = new char[8192];
            int count;
            read = 0;

            while ((count = reader.Read(buffer, 0, buffer.Length)) > 0)
            {
                read += count;

                if (read > limit)
                {
                    text = string.Empty;
                    return false;
                }

                for (int i = 0; i < count; i++)
                {
                    char character = buffer[i];

                    if (!char.IsControl(character) || character is '\t' or '\n' or '\r')
                        builder.Append(character);
                }
            }

            text = builder.ToString().Trim();
            return true;
        }

        // ###########################################################################################
        // What a shown text takes in ONE of the mail's bodies, in UTF-8 bytes: the larger of the plain
        // text (line breaks as CRLF, as MailBody writes them) and the HTML (HTML-encoded, so a " is
        // six bytes and a < four). Transfer encoding comes on top of both, which ShownBytesLimit
        // leaves room for.
        // ###########################################################################################
        internal static long MailBytes(string text)
        {
            ArgumentNullException.ThrowIfNull(text);

            string lines = text.Replace("\r\n", "\n", StringComparison.Ordinal);
            long plain = Encoding.UTF8.GetByteCount(lines) + lines.Count(character => character == '\n');
            long html = Encoding.UTF8.GetByteCount(WebUtility.HtmlEncode(lines));

            return Math.Max(plain, html);
        }

        // ###########################################################################################
        // What the sender's words take in EACH of the mail's bodies as its quotation draws them (code
        // review, 2026-10-16): four spaces before every line of the plain text and "<br>" between the
        // HTML's make a text of short lines far larger than typed, and the pictures were given the
        // difference. Measured by drawing it, so the count is the renderer's own; a marker line is
        // counted as words, a shown picture's own markup by PictureOverheadBytes.
        // ###########################################################################################
        internal static long QuotedMailBytes(string text)
        {
            ArgumentNullException.ThrowIfNull(text);

            MailBody quoted = new MailBody(MailBody.WideWidth).Quote(text, string.Empty);

            return Math.Max(Encoding.UTF8.GetByteCount(quoted.ToText()), Encoding.UTF8.GetByteCount(quoted.ToHtml(string.Empty)));
        }

        // ###########################################################################################
        // The first of `items` the mail lists - at most `maxItems`, taking no more than `room` bytes
        // in each body (MailBytes, with a line's markup) - and how many are left unlisted. The notes
        // and the saved files' names come after the pictures' room is fixed, so they are held to a
        // room of their own inside MailOverheadBytes (code review, 2026-10-16).
        // ###########################################################################################
        internal static (List<string> Listed, int NotListed) ListWithin(IReadOnlyList<string> items, int maxItems, long room)
        {
            ArgumentNullException.ThrowIfNull(items);

            var listed = new List<string>();
            long used = 0;

            foreach (string item in items)
            {
                long bytes = FeedbackFlow.MailBytes(item) + FeedbackFlow.ListLineBytes;

                if (listed.Count >= maxItems || used + bytes > room)
                    break;

                listed.Add(item);
                used += bytes;
            }

            return (listed, items.Count - listed.Count);
        }

        // The sender's address when it is one, else empty - it becomes the mail's Reply-To.
        internal static string CleanAddress(string? email)
        {
            string trimmed = email?.Trim() ?? string.Empty;

            if (trimmed.Length == 0 || trimmed.Any(char.IsControl) || !trimmed.Contains('@', StringComparison.Ordinal))
                return string.Empty;

            return MailAddress.TryCreate(trimmed, out MailAddress? address) && address.Address == trimmed
                ? trimmed
                : string.Empty;
        }

        // The version as a subject line can carry it: no control characters, short, never blank.
        internal static string CleanVersion(string? version)
        {
            string cleaned = new string((version ?? string.Empty).Where(character => !char.IsControl(character)).ToArray()).Trim();

            if (cleaned.Length > 40)
                cleaned = cleaned[..40];

            return cleaned.Length == 0 ? "Unknown" : cleaned;
        }

        private static string Megabytes(long bytes) =>
            $"{bytes / (1024.0 * 1024.0):0.#} MB";

        private static void DeleteQuietly(string folder)
        {
            try
            {
                Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // An empty folder left behind is harmless.
            }
        }
    }

    // ###########################################################################################
    // Where feedback goes: the folder its files are saved in, the address its mail goes to, the
    // disk reserve and the folder's total (ServerOptions; 0 is no total). The free-space probe, the
    // total probe and the reference maker are seams, so a test can fill the disk or the folder, or
    // know the folder's name. MaxEntries is the most central directory records a zip may hold
    // (lowered by a test). UnpackGate lets one unpack run at a time, Saved hears how many bytes an
    // unpack kept (before the gate opens) and Removed the same bytes again when no mail names them
    // and they are deleted - all FeedbackStorage's, the service's one copy. MaxMailBytes is how
    // large the mail may be (lowered by a test).
    // ###########################################################################################
    public sealed record FeedbackDestination(
        string Root,
        string ToAddress,
        Func<long?> FreeBytes,
        long MinimumFreeBytes,
        Func<string> NewReference,
        long MaxStoredBytes = 0,
        Func<long?>? StoredBytes = null,
        int MaxEntries = FeedbackFlow.MaxZipEntries,
        SemaphoreSlim? UnpackGate = null,
        Action<long>? Saved = null,
        Action<long>? Removed = null,
        long MaxMailBytes = FeedbackFlow.MaxMailBytes);

    // What became of one feedback: the status and text CRT is answered with, and what was saved.
    public sealed record FeedbackOutcome(int StatusCode, string Answer, string? Reference, int FilesSaved)
    {
        public bool IsDelivered => this.StatusCode == StatusCodes.Status200OK;
    }
}
