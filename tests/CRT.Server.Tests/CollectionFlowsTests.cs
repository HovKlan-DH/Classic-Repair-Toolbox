using System.Security.Cryptography;
using CRT.Data.Tests;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using CRT.Server.Handlers.Email;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Sharing an oscilloscope collection (CollectionFlows): who may, what a collection may show, the limits, and the hourly
    // sweep - on the fake store, the pictures in a folder of the test's own.
    // ###########################################################################################
    public sealed class CollectionFlowsTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private const string Site = "https://classic-repair-toolbox.dk";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-collections-" + Guid.NewGuid().ToString("N"));

        private readonly FakeCollectionStore thisStore = new();

        private readonly FakeAccountStore thisAccounts = new();

        private readonly FakeEmailSender thisMailer = new();

        private readonly CollectionFiles thisFiles;

        private static readonly AccountRecord Anna = Account(1, "anna@example.com", "Anna");

        private static readonly AccountRecord Bo = Account(2, "bo@example.com", "Bo");

        public CollectionFlowsTests()
        {
            Directory.CreateDirectory(this.thisRoot);
            this.thisFiles = new CollectionFiles(this.thisRoot);
            this.thisAccounts.Accounts[Anna.Id] = Anna;
            this.thisAccounts.Accounts[Bo.Id] = Bo;
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        private static AccountRecord Account(long id, string email, string name, bool verified = true, bool administrator = false) =>
            new(id, email, email, "hash", name, verified, administrator, IsLocked: false, Now.AddDays(-10), null);

        private static string ShaOf(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

        private static CollectionImageEntry Image(string sha, string label = "U7", string pin = "12") =>
            new(sha, label, pin, "/RAS", "Pulsing", "50us", "2V", "1V", null, CapturedUtc: Now);

        private static CollectionContentRequest Content(params CollectionImageEntry[] images) =>
            new("Commodore 64", "250407", "PAL", "No picture", "It went black.", ["https://www.lemon64.com/forum/viewtopic.php?t=1"], CollectionRules.DeleteAt(Now, "1y"), images);

        private CollectionStorage Storage(long maximum = 0, long? free = null, long reserve = 0) => new(maximum, () => free, reserve);

        private Task<CollectionOutcome> UploadAsync(AccountRecord? account, byte[] bytes, CollectionStorage? storage = null, DateTimeOffset? when = null) =>
            CollectionFlows.UploadImageAsync(account, ShaOf(bytes), bytes, this.thisStore, this.thisFiles, storage ?? this.Storage(), when ?? Now, CancellationToken.None);

        private Task<CollectionOutcome> CreateAsync(AccountRecord? account, CollectionContentRequest content, DateTimeOffset? when = null, string? clientKey = null) =>
            CollectionFlows.CreateAsync(account, content, clientKey, this.thisStore, this.thisFiles, Site, when ?? Now, CancellationToken.None);

        // A collection of one capture, shared by `account` - uploaded, then sent.
        private async Task<CollectionAnswer> SharedAsync(AccountRecord account, int seed = 1, DateTimeOffset? when = null, string lifetime = "1y")
        {
            byte[] png = TestPng.Rgb(32, 20, seed);

            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(account, png, when: when)).Verdict);

            CollectionOutcome created = await this.CreateAsync(account, Content(Image(ShaOf(png))) with { DeleteAtUtc = CollectionRules.DeleteAt(when ?? Now, lifetime) }, when);

            Assert.Equal(CollectionVerdict.Created, created.Verdict);
            return created.Answer!;
        }

        // ---- Uploading a capture --------------------------------------------------------

        [Fact]
        public async Task A_capture_uploaded_is_kept_once_and_its_size_answered()
        {
            byte[] png = TestPng.Rgb(800, 480);

            CollectionOutcome outcome = await this.UploadAsync(Anna, png);

            Assert.Equal(CollectionVerdict.Done, outcome.Verdict);
            Assert.Equal(new CollectionImageUploadAnswer(ShaOf(png), 800, 480), outcome.Upload);
            Assert.Equal(png, File.ReadAllBytes(this.thisFiles.PathOf(ShaOf(png))));
            Assert.Single(this.thisStore.Uploads);

            // The same again is done, and counts nothing more.
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(Anna, png)).Verdict);
            Assert.Single(this.thisStore.Uploads);
        }

        [Fact]
        public async Task Nobody_signed_in_shares_nothing()
        {
            Assert.Equal(CollectionVerdict.Unauthorized, (await this.UploadAsync(null, TestPng.Rgb(8, 8))).Verdict);
            Assert.Equal(CollectionVerdict.Unauthorized, (await this.CreateAsync(null, Content(Image(new string('a', 64))))).Verdict);
        }

        // An address not confirmed cannot sign in since server 14.0.0 - a session from before is told to confirm it.
        [Fact]
        public async Task An_account_whose_address_is_not_confirmed_is_told_to_confirm_it()
        {
            AccountRecord unconfirmed = Account(3, "new@example.com", "New", verified: false);

            CollectionOutcome outcome = await this.UploadAsync(unconfirmed, TestPng.Rgb(8, 8));

            Assert.Equal(CollectionVerdict.Forbidden, outcome.Verdict);
            Assert.Equal([CollectionFlows.AddressNotConfirmed], outcome.Errors);
            Assert.Empty(this.thisStore.Uploads);
        }

        [Fact]
        public async Task Bytes_that_are_not_their_checksum_are_refused_and_nothing_is_kept()
        {
            byte[] png = TestPng.Rgb(8, 8);

            CollectionOutcome outcome = await CollectionFlows.UploadImageAsync(
                Anna, new string('a', 64), png, this.thisStore, this.thisFiles, this.Storage(), Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal([CollectionFlows.ChecksumMismatch], outcome.Errors);
            Assert.Empty(Directory.GetFiles(this.thisRoot, "*", SearchOption.AllDirectories));
        }

        [Fact]
        public async Task A_file_that_is_no_capture_is_refused_and_nothing_is_kept()
        {
            byte[] notPng = "this is a zip file, honestly"u8.ToArray();

            CollectionOutcome outcome = await this.UploadAsync(Anna, notPng);

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal([CollectionPng.NotAPicture], outcome.Errors);
            Assert.Empty(this.thisStore.Uploads);
            Assert.Empty(Directory.GetFiles(this.thisRoot, "*", SearchOption.AllDirectories));
        }

        [Fact]
        public async Task An_address_naming_no_checksum_leads_nowhere()
        {
            CollectionOutcome outcome = await CollectionFlows.UploadImageAsync(
                Anna, "../../etc/passwd", TestPng.Rgb(8, 8), this.thisStore, this.thisFiles, this.Storage(), Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
        }

        [Fact]
        public async Task An_account_uploads_its_day_s_captures_and_no_more()
        {
            for (int i = 0; i < CollectionRules.MaximumUploadsPerDay; i++)
                this.thisStore.Uploads.Add(new CollectionUploadRecord(Anna.Id, i.ToString("x64", System.Globalization.CultureInfo.InvariantCulture), 10, 8, 8, Now.AddHours(-2)));

            CollectionOutcome outcome = await this.UploadAsync(Anna, TestPng.Rgb(8, 8));

            Assert.Equal(CollectionVerdict.TooMany, outcome.Verdict);

            // Another account's day is its own.
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(Bo, TestPng.Rgb(8, 8))).Verdict);
        }

        [Fact]
        public async Task With_no_room_left_a_capture_is_refused_saying_so()
        {
            Assert.Equal(CollectionVerdict.NoRoom, (await this.UploadAsync(Anna, TestPng.Rgb(8, 8), this.Storage(maximum: 10))).Verdict);
            Assert.Equal(CollectionVerdict.NoRoom, (await this.UploadAsync(Anna, TestPng.Rgb(8, 8), this.Storage(free: 1000, reserve: 5000))).Verdict);

            // A disk that cannot be measured is taken as having room, as the blob store takes it.
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(Anna, TestPng.Rgb(8, 8), this.Storage(free: null, reserve: 5000))).Verdict);
        }

        // Bytes kept already take no more room - another account sending the same capture is not refused for a full disk.
        [Fact]
        public async Task A_capture_kept_already_needs_no_room()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(Anna, png);

            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(Bo, png, this.Storage(maximum: 1))).Verdict);
            Assert.Equal(2, this.thisStore.Uploads.Count);
        }

        // ---- Sharing a collection ----------------------------------------------------------

        [Fact]
        public async Task A_collection_naming_captures_not_uploaded_is_answered_with_them()
        {
            byte[] first = TestPng.Rgb(8, 8, seed: 1);
            byte[] second = TestPng.Rgb(8, 8, seed: 2);

            await this.UploadAsync(Anna, first);

            CollectionOutcome outcome = await this.CreateAsync(Anna, Content(Image(ShaOf(first)), Image(ShaOf(second), "U8", "3")));

            Assert.Equal(CollectionVerdict.MissingImages, outcome.Verdict);
            Assert.Equal([ShaOf(second)], outcome.Missing);
            Assert.Empty(this.thisStore.Collections);
        }

        [Fact]
        public async Task A_collection_with_its_captures_uploaded_is_shared_under_a_new_id()
        {
            byte[] png = TestPng.Rgb(800, 480);

            await this.UploadAsync(Anna, png);

            CollectionOutcome outcome = await this.CreateAsync(Anna, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.Created, outcome.Verdict);
            Assert.True(CollectionRules.IsId(outcome.Answer!.Id));
            Assert.Equal($"{Site}/collections/{outcome.Answer.Id}", outcome.Answer.Url);
            Assert.Equal(1, outcome.Answer.ImageCount);

            CollectionRecord kept = this.thisStore.Collections[outcome.Answer.Id];

            Assert.Equal(Anna.Id, kept.AccountId);
            Assert.Equal(Now, kept.CreatedUtc);
            Assert.Equal((800, 480), (kept.Images[0].Width, kept.Images[0].Height));
        }

        // *** A CHECKSUM READ OFF SOMEBODY ELSE'S PAGE IS NO WAY TO SHOW THEIR PICTURE. ***
        [Fact]
        public async Task A_collection_shows_only_captures_its_own_account_uploaded()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(Anna, png);

            CollectionOutcome outcome = await this.CreateAsync(Bo, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.MissingImages, outcome.Verdict);
            Assert.Equal([ShaOf(png)], outcome.Missing);
        }

        // The bytes gone from disk (moved to another disk and one lost, say): the capture is asked for again, not shown broken.
        [Fact]
        public async Task A_capture_whose_bytes_are_gone_is_asked_for_again()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(Anna, png);
            File.Delete(this.thisFiles.PathOf(ShaOf(png)));

            CollectionOutcome outcome = await this.CreateAsync(Anna, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.MissingImages, outcome.Verdict);

            // Sent again, the bytes are kept again.
            await this.UploadAsync(Anna, png);
            Assert.Equal(CollectionVerdict.Created, (await this.CreateAsync(Anna, Content(Image(ShaOf(png))))).Verdict);
        }

        [Fact]
        public async Task A_collection_breaking_a_rule_is_refused_in_the_rules_words()
        {
            CollectionOutcome outcome = await this.CreateAsync(Anna, Content(Image(ShaOf(TestPng.Rgb(8, 8)))) with { HardwareName = "" });

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal(["The collection does not say the hardware it is about."], outcome.Errors);
        }

        [Fact]
        public async Task An_account_keeps_so_many_collections_and_shares_so_many_a_day()
        {
            for (int i = 0; i < CollectionRules.MaximumCollectionsPerDay; i++)
                await this.SharedAsync(Anna, seed: i);

            byte[] png = TestPng.Rgb(8, 8, seed: 999);

            await this.UploadAsync(Anna, png);

            CollectionOutcome today = await this.CreateAsync(Anna, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.TooMany, today.Verdict);
            Assert.Contains("a day", today.Errors[0], StringComparison.Ordinal);

            // A day later the day's limit is past - but not how many are kept at once.
            for (int i = 0; i < CollectionRules.MaximumLiveCollectionsPerAccount - CollectionRules.MaximumCollectionsPerDay; i++)
            {
                var record = this.thisStore.Collections.Values.First() with { Id = CollectionRules.NewId(), CreatedUtc = Now.AddDays(-30) };
                this.thisStore.Collections[record.Id] = record;
            }

            CollectionOutcome kept = await this.CreateAsync(Anna, Content(Image(ShaOf(png))), Now.AddDays(2));

            Assert.Equal(CollectionVerdict.TooMany, kept.Verdict);
            Assert.Contains("delete one", kept.Errors[0], StringComparison.Ordinal);
        }

        // ---- Changing and deleting one -----------------------------------------------------

        [Fact]
        public async Task A_collection_changed_keeps_its_id_and_creation_and_takes_its_new_captures()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            byte[] added = TestPng.Rgb(16, 16, seed: 5);

            await this.UploadAsync(Anna, added);

            CollectionImageEntry first = this.thisStore.Collections[shared.Id].Images[0].Entry;
            CollectionOutcome outcome = await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(first, Image(ShaOf(added), "U8", "1")) with { Title = "Found it" },
                this.thisStore, this.thisFiles, Site, Now.AddDays(3), CancellationToken.None);

            Assert.Equal(CollectionVerdict.Done, outcome.Verdict);
            Assert.Equal(shared.Id, outcome.Answer!.Id);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal("Found it", kept.Title);
            Assert.Equal(Now, kept.CreatedUtc);
            Assert.Equal(Now.AddDays(3), kept.UpdatedUtc);
            Assert.Equal(["U7", "U8"], kept.Images.Select(image => image.Entry.ComponentLabel));
        }

        // Whether a collection exists is nobody else's business.
        [Fact]
        public async Task Another_account_finds_no_such_collection_to_change_or_delete()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);

            CollectionOutcome changed = await CollectionFlows.UpdateAsync(
                Bo, shared.Id, Content(Image(new string('a', 64))), this.thisStore, this.thisFiles, Site, Now, CancellationToken.None);
            CollectionOutcome deleted = await CollectionFlows.DeleteAsync(Bo, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.NotFound, changed.Verdict);
            Assert.Equal(CollectionVerdict.NotFound, deleted.Verdict);
            Assert.Equal([CollectionFlows.NoSuchCollection], deleted.Errors);
            Assert.False(this.thisStore.Collections[shared.Id].IsDeleted);
        }

        [Fact]
        public async Task The_owner_deletes_a_collection_and_it_can_no_longer_be_changed()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);

            Assert.Equal(CollectionVerdict.Removed, (await CollectionFlows.DeleteAsync(Anna, shared.Id, this.thisStore, Now, CancellationToken.None)).Verdict);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal(CollectionDeletedReasons.Owner, kept.DeletedReason);
            Assert.Null(kept.Title);
            Assert.Empty(kept.Images);

            // Deleting it again is done; changing it is told it went.
            Assert.Equal(CollectionVerdict.Removed, (await CollectionFlows.DeleteAsync(Anna, shared.Id, this.thisStore, Now, CancellationToken.None)).Verdict);

            CollectionOutcome changed = await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(Image(new string('a', 64))), this.thisStore, this.thisFiles, Site, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Gone, changed.Verdict);
            Assert.Equal(["That collection was deleted on 2026-November-26."], changed.Errors);
        }

        [Fact]
        public async Task The_administrator_may_delete_anybody_s_collection()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            AccountRecord administrator = Account(9, "admin@example.com", "Admin", administrator: true);

            Assert.Equal(CollectionVerdict.Removed, (await CollectionFlows.DeleteAsync(administrator, shared.Id, this.thisStore, Now, CancellationToken.None)).Verdict);
            Assert.Equal(CollectionDeletedReasons.Administrator, this.thisStore.Collections[shared.Id].DeletedReason);
        }

        [Fact]
        public async Task An_account_lists_its_own_collections_not_deleted()
        {
            CollectionAnswer kept = await this.SharedAsync(Anna, seed: 1);
            CollectionAnswer gone = await this.SharedAsync(Anna, seed: 2);
            await this.SharedAsync(Bo, seed: 3);
            await CollectionFlows.DeleteAsync(Anna, gone.Id, this.thisStore, Now, CancellationToken.None);

            CollectionOutcome outcome = await CollectionFlows.ListMineAsync(Anna, this.thisStore, Site, Now, CancellationToken.None);

            CollectionSummaryAnswer only = Assert.Single(outcome.Mine!.Collections);

            Assert.Equal(kept.Id, only.Id);
            Assert.Equal(kept.Url, only.Url);
            Assert.Equal(1, only.ImageCount);
            Assert.Equal(("Commodore 64", "250407", "PAL"), (only.HardwareName, only.BoardName, only.Region));
        }

        // ---- The hourly sweep ----------------------------------------------------------------

        private Task<CollectionSweepResult> SweepAsync(DateTimeOffset when) =>
            CollectionFlows.SweepAsync(this.thisStore, this.thisFiles, this.thisAccounts, this.thisMailer, Site, when, CancellationToken.None);

        [Fact]
        public async Task A_collection_whose_time_is_up_is_deleted_and_its_capture_goes_a_day_later()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1m");
            string sha = this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256;

            Assert.Equal(new CollectionSweepResult(0, 0, 0), await this.SweepAsync(Now.AddDays(20)));

            CollectionSweepResult result = await this.SweepAsync(shared.DeleteAtUtc.AddMinutes(1));

            Assert.Equal(1, result.Expired);
            Assert.Equal(CollectionDeletedReasons.Expired, this.thisStore.Collections[shared.Id].DeletedReason);

            // The upload came a month ago, and nothing shows it now: it goes, and with no other upload of it, its bytes.
            Assert.Equal(1, result.FilesDeleted);
            Assert.False(this.thisFiles.Has(sha));
        }

        // The rows are gone once the sweep has them, so a file that cannot be deleted is named and the others still go (code
        // review, 2026-11-27: the first refusal stopped the loop, and every file after it stayed for good). Only Windows refuses
        // to delete a file held open, which is how this test makes one.
        [Fact]
        public async Task A_capture_that_cannot_be_deleted_is_named_and_the_rest_still_go()
        {
            Assert.SkipUnless(OperatingSystem.IsWindows(), "Only Windows refuses to delete a file held open.");

            byte[] held = TestPng.Rgb(8, 8, seed: 1);
            byte[] free = TestPng.Rgb(8, 8, seed: 2);

            await this.UploadAsync(Anna, held);
            await this.UploadAsync(Anna, free);

            CollectionSweepResult result;

            using (new FileStream(this.thisFiles.PathOf(ShaOf(held)), FileMode.Open, FileAccess.Read, FileShare.Read))
                result = await this.SweepAsync(Now.AddDays(2));

            Assert.Equal([ShaOf(held)], result.FilesNotDeleted);
            Assert.Equal(1, result.FilesDeleted);
            Assert.False(this.thisFiles.Has(ShaOf(free)));
        }

        // The database failing one ask names that picture, and the sweep goes on with the rest - their rows are gone already, so
        // a picture left now would be left for good (code review, 2026-11-27).
        [Fact]
        public async Task A_picture_the_database_cannot_be_asked_about_is_named_and_the_rest_still_go()
        {
            byte[] failing = TestPng.Rgb(8, 8, seed: 5);
            byte[] fine = TestPng.Rgb(8, 8, seed: 6);

            await this.UploadAsync(Anna, failing);
            await this.UploadAsync(Anna, fine);
            this.thisStore.IsUploadedFails.Add(ShaOf(failing));

            CollectionSweepResult result = await this.SweepAsync(Now.AddDays(2));

            Assert.Equal([ShaOf(failing)], result.FilesNotDeleted);
            Assert.Equal(1, result.FilesDeleted);
            Assert.False(this.thisFiles.Has(ShaOf(fine)));
        }

        // Another account sending the very same picture as the sweep lets go of it keeps the bytes - its upload wrote none, as
        // they were here (code review, 2026-11-27).
        [Fact]
        public async Task A_picture_sent_again_by_another_account_as_it_is_swept_stays()
        {
            byte[] png = TestPng.Rgb(8, 8, seed: 4);

            await this.UploadAsync(Anna, png);
            this.thisStore.AfterDeleteUnusedUploads = () =>
                this.thisStore.Uploads.Add(new CollectionUploadRecord(Bo.Id, ShaOf(png), png.Length, 8, 8, Now.AddDays(2)));

            CollectionSweepResult result = await this.SweepAsync(Now.AddDays(2));

            Assert.Equal(0, result.FilesDeleted);
            Assert.True(this.thisFiles.Has(ShaOf(png)));
        }

        // A capture uploaded for a collection about to be sent is not taken from under it.
        [Fact]
        public async Task A_capture_uploaded_today_and_shown_nowhere_yet_stays_a_day()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(Anna, png);

            await this.SweepAsync(Now.AddHours(23));
            Assert.True(this.thisFiles.Has(ShaOf(png)));

            await this.SweepAsync(Now.AddHours(25));
            Assert.False(this.thisFiles.Has(ShaOf(png)));
        }

        // A picture sent again - its old upload shown by no collection, its day long past - is kept a day from that send, so the
        // collection it was sent for finds it (code review, 2026-11-27: the next sweep took the old row and the bytes, and the
        // share was answered "missing" again).
        [Fact]
        public async Task A_capture_sent_again_stays_a_day_from_its_last_send()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(Anna, png);
            await this.UploadAsync(Anna, png, when: Now.AddDays(10));

            await this.SweepAsync(Now.AddDays(10).AddHours(23));

            Assert.True(this.thisFiles.Has(ShaOf(png)));
            Assert.Equal(Now.AddDays(10), Assert.Single(this.thisStore.Uploads).UploadedUtc);
        }

        // Two accounts sent the same picture: one collection gone takes only its own account's upload, never the bytes the
        // other still shows.
        [Fact]
        public async Task A_capture_another_collection_still_shows_stays()
        {
            CollectionAnswer annas = await this.SharedAsync(Anna, seed: 7, lifetime: "1m");
            await this.SharedAsync(Bo, seed: 7, lifetime: "5y");

            string sha = this.thisStore.Collections[annas.Id].Images[0].Entry.Sha256;

            CollectionSweepResult result = await this.SweepAsync(annas.DeleteAtUtc.AddMinutes(1));

            Assert.Equal(1, result.Expired);
            Assert.Equal(0, result.FilesDeleted);
            Assert.True(this.thisFiles.Has(sha));
            Assert.Single(this.thisStore.Uploads);
        }

        [Fact]
        public async Task The_owner_is_reminded_once_a_month_before_and_a_month_long_collection_never()
        {
            CollectionAnswer year = await this.SharedAsync(Anna, seed: 1, lifetime: "1y");
            CollectionAnswer month = await this.SharedAsync(Anna, seed: 2, lifetime: "1m");

            // The month-long one is within the month at once - and told nothing.
            await this.SweepAsync(Now.AddHours(1));
            Assert.Empty(this.thisMailer.Sent);

            CollectionSweepResult result = await this.SweepAsync(year.DeleteAtUtc.AddDays(-29));

            Assert.Equal(1, result.Reminded);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);

            Assert.Equal("anna@example.com", mail.ToAddress);
            Assert.Contains(year.Url, mail.Body, StringComparison.Ordinal);
            Assert.Contains("No picture", mail.Body, StringComparison.Ordinal);

            // Once.
            await this.SweepAsync(year.DeleteAtUtc.AddDays(-28));
            Assert.Single(this.thisMailer.Sent);
            Assert.NotNull(this.thisStore.Collections[month.Id].RemindedUtc ?? this.thisStore.Collections[month.Id].DeletedUtc);
        }

        // Kept longer after the reminder, the next reminder comes before the new date.
        [Fact]
        public async Task A_collection_kept_longer_is_reminded_again_before_its_new_date()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1y");
            DateTimeOffset reminderDay = shared.DeleteAtUtc.AddDays(-29);

            await this.SweepAsync(reminderDay);
            Assert.Single(this.thisMailer.Sent);

            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;

            await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(image) with { DeleteAtUtc = CollectionRules.DeleteAt(reminderDay, "1y") },
                this.thisStore, this.thisFiles, Site, reminderDay, CancellationToken.None);

            Assert.Null(this.thisStore.Collections[shared.Id].RemindedUtc);
        }

        // A collection shared again with "1 month" left is told nothing, however long ago it was made - the month counts from the
        // share that set it (code review, 2026-11-27: measured from its making, the reminder came the day after).
        [Fact]
        public async Task A_collection_shared_again_with_a_month_left_is_not_reminded()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1y");
            DateTimeOffset later = shared.DeleteAtUtc.AddDays(-40);
            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;

            await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(image) with { DeleteAtUtc = CollectionRules.DeleteAt(later, "1m") },
                this.thisStore, this.thisFiles, Site, later, CancellationToken.None);

            await this.SweepAsync(later.AddDays(1));

            Assert.Empty(this.thisMailer.Sent);
        }

        // Whether a reminder is due is decided as the date is set: a share keeping its date - a caption changed 45 days before
        // it - still gets the reminder (code review, 2026-11-27: measured from the last share, it never came).
        [Fact]
        public async Task A_share_keeping_its_date_keeps_the_reminder_coming()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1y");
            DateTimeOffset edited = shared.DeleteAtUtc.AddDays(-45);
            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;

            await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(image) with { DeleteAtUtc = shared.DeleteAtUtc, Title = "Changed" },
                this.thisStore, this.thisFiles, Site, edited, CancellationToken.None);

            CollectionSweepResult result = await this.SweepAsync(shared.DeleteAtUtc.AddDays(-29));

            Assert.Equal(1, result.Reminded);
            Assert.Single(this.thisMailer.Sent);
        }

        [Fact]
        public async Task An_owner_whose_account_is_locked_or_gone_is_mailed_nothing()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1y");
            this.thisAccounts.Accounts[Anna.Id] = Anna with { IsLocked = true };

            await this.SweepAsync(shared.DeleteAtUtc.AddDays(-29));

            Assert.Empty(this.thisMailer.Sent);
            Assert.NotNull(this.thisStore.Collections[shared.Id].RemindedUtc);
        }

        // The deletion date as the database keeps it, whole milliseconds - else the same date shared again read as later than
        // the stored one, and the month's reminder was mailed again (code review, 2026-11-27).
        [Fact]
        public async Task The_deletion_date_is_kept_to_the_millisecond_so_the_same_date_again_is_no_later()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1y");
            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;
            DateTimeOffset deleteAt = this.thisStore.Collections[shared.Id].DeleteAtUtc.AddTicks(1234);

            await CollectionFlows.UpdateAsync(Anna, shared.Id, Content(image) with { DeleteAtUtc = deleteAt }, this.thisStore, this.thisFiles, Site, Now, CancellationToken.None);

            Assert.Equal(0, this.thisStore.Collections[shared.Id].DeleteAtUtc.UtcTicks % TimeSpan.TicksPerMillisecond);
            Assert.Equal(CollectionFlows.WholeMilliseconds(deleteAt), this.thisStore.Collections[shared.Id].DeleteAtUtc);
        }

        // A first share the server kept whose answer was lost is sent again under the same key: the page it made is changed,
        // never a second one made that only the administrator could take down (code review, 2026-11-27).
        [Fact]
        public async Task A_first_share_sent_again_under_its_key_changes_the_page_it_made()
        {
            const string key = "0123456789abcdef0123456789abcdef";
            byte[] png = TestPng.Rgb(32, 20, seed: 3);
            await this.UploadAsync(Anna, png);

            CollectionOutcome first = await this.CreateAsync(Anna, Content(Image(ShaOf(png))), clientKey: key);
            CollectionOutcome again = await this.CreateAsync(Anna, Content(Image(ShaOf(png))) with { Title = "Typed since" }, clientKey: key);

            Assert.Equal(first.Answer!.Id, again.Answer!.Id);
            Assert.Equal("Typed since", Assert.Single(this.thisStore.Collections.Values).Title);

            // Another account's key, a key of another shape, or none: each first share is a page of its own.
            await this.UploadAsync(Bo, png);
            CollectionOutcome bos = await this.CreateAsync(Bo, Content(Image(ShaOf(png))), clientKey: key);
            CollectionOutcome noKey = await this.CreateAsync(Anna, Content(Image(ShaOf(png))), clientKey: "../../x");

            Assert.NotEqual(first.Answer.Id, bos.Answer!.Id);
            Assert.NotEqual(first.Answer.Id, noKey.Answer!.Id);
        }

        // A first share retried while the first is still being written waits for it, then finds its page under the key - looked
        // up before the lock, it found none and made a second page (code review, 2026-11-27).
        [Fact]
        public async Task A_first_share_retried_while_the_first_is_written_makes_no_second_page()
        {
            const string key = "0123456789abcdef0123456789abcdef";
            byte[] png = TestPng.Rgb(32, 20, seed: 5);
            await this.UploadAsync(Anna, png);

            Task<CollectionOutcome>? retried = null;
            this.thisStore.WhileCountingCreated = () => retried = this.CreateAsync(Anna, Content(Image(ShaOf(png))), clientKey: key);

            CollectionOutcome first = await this.CreateAsync(Anna, Content(Image(ShaOf(png))), clientKey: key);
            CollectionOutcome again = await retried!;

            Assert.Equal(first.Answer!.Id, again.Answer!.Id);
            Assert.Single(this.thisStore.Collections.Values);
        }

        // A page taken down is no page to change: shared again under its key, a new one is made.
        [Fact]
        public async Task A_key_whose_page_was_taken_down_makes_a_new_page()
        {
            const string key = "0123456789abcdef0123456789abcdef";
            byte[] png = TestPng.Rgb(32, 20, seed: 4);
            await this.UploadAsync(Anna, png);

            CollectionOutcome first = await this.CreateAsync(Anna, Content(Image(ShaOf(png))), clientKey: key);
            await CollectionFlows.DeleteAsync(Anna, first.Answer!.Id, this.thisStore, Now, CancellationToken.None);

            CollectionOutcome again = await this.CreateAsync(Anna, Content(Image(ShaOf(png))), clientKey: key);

            Assert.Equal(CollectionVerdict.Created, again.Verdict);
            Assert.NotEqual(first.Answer.Id, again.Answer!.Id);
        }

        // Its time up, a collection is gone for CRT too - its list, and "Share", which then makes a new page rather than bring
        // the old one back - as its page says (code review, 2026-11-27).
        [Fact]
        public async Task A_collection_whose_time_is_up_is_gone_for_CRT_s_list_and_share_too()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1m");
            DateTimeOffset after = shared.DeleteAtUtc.AddMinutes(1);
            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;

            Assert.Empty((await CollectionFlows.ListMineAsync(Anna, this.thisStore, Site, after, CancellationToken.None)).Mine!.Collections);

            CollectionOutcome update = await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(image) with { DeleteAtUtc = CollectionRules.DeleteAt(after, "1y") }, this.thisStore, this.thisFiles, Site, after, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Gone, update.Verdict);
        }

        // ---- Shown to a reader (code review, 2026-11-27) --------------------------------------------

        private static readonly AccountRecord TheAdministrator = Account(9, "admin@example.com", "Dennis@CRT", administrator: true);

        private Task<ShownCollection?> ShownToAsync(string id, AccountRecord? viewer, Action? asked = null) =>
            CollectionFlows.FindShownToAsync(
                id,
                _ =>
                {
                    asked?.Invoke();
                    return Task.FromResult(viewer);
                },
                this.thisStore,
                this.thisAccounts,
                Now,
                CancellationToken.None);

        // Its time up, a collection is deleted from that moment - not only once the hourly sweep makes it so (code review,
        // 2026-11-27: for up to an hour its page took comments under "Deleted on" a date passed).
        [Fact]
        public async Task A_collection_whose_time_is_up_is_shown_as_deleted_before_the_sweep()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1m");
            DateTimeOffset after = shared.DeleteAtUtc.AddMinutes(1);

            Assert.Null(await CollectionFlows.FindShownAsync(shared.Id, this.thisStore, this.thisAccounts, after, CancellationToken.None));

            CollectionRecord? asDeleted = await CollectionFlows.FindShownAsync(shared.Id, this.thisStore, this.thisAccounts, after, CancellationToken.None, evenIfDeleted: true);

            Assert.Equal((true, CollectionDeletedReasons.Expired), (asDeleted?.IsDeleted, asDeleted?.DeletedReason));
            Assert.NotNull(await CollectionFlows.FindShownAsync(shared.Id, this.thisStore, this.thisAccounts, shared.DeleteAtUtc.AddMinutes(-1), CancellationToken.None));
        }

        // The administrator's "Every collection" links a locked owner's collection - its page must open for the administrator,
        // to delete it or hide its comments, and say it is offline; for anybody else it stays "not found".
        [Fact]
        public async Task A_locked_owner_s_collection_is_shown_to_the_administrator_alone_as_offline()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            this.thisAccounts.Accounts[Anna.Id] = Anna with { IsLocked = true };

            ShownCollection? toAdministrator = await this.ShownToAsync(shared.Id, TheAdministrator);

            Assert.Equal((shared.Id, CollectionOffline.OwnerLocked), (toAdministrator?.Collection.Id, toAdministrator?.Offline));
            Assert.Null(await this.ShownToAsync(shared.Id, Bo));
            Assert.Null(await this.ShownToAsync(shared.Id, null));
            Assert.Null(await CollectionFlows.FindShownAsync(shared.Id, this.thisStore, this.thisAccounts, Now, CancellationToken.None));
        }

        [Fact]
        public async Task A_collection_whose_owner_has_no_account_is_shown_to_the_administrator_as_offline()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            this.thisAccounts.Accounts.Remove(Anna.Id);

            Assert.Equal(CollectionOffline.OwnerGone, (await this.ShownToAsync(shared.Id, TheAdministrator))?.Offline);
            Assert.Null(await this.ShownToAsync(shared.Id, Bo));
        }

        // A picture's every request asks this - a collection shown to all reads no session.
        [Fact]
        public async Task A_collection_shown_to_all_asks_nothing_of_the_reader()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            bool asked = false;

            ShownCollection? shown = await this.ShownToAsync(shared.Id, Bo, () => asked = true);

            Assert.Equal((shared.Id, false), (shown?.Collection.Id, shown?.IsOffline));
            Assert.False(asked);
        }

        // The owner deletes their own in CRT, whose window keeps in step with the server - the website's delete is the
        // administrator's alone.
        [Fact]
        public async Task The_website_s_delete_is_the_administrator_s_alone()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);

            CollectionOutcome byOwner = await CollectionFlows.DeleteFromWebAsync(Anna, shared.Id, this.thisStore, Now, CancellationToken.None);
            CollectionOutcome signedOut = await CollectionFlows.DeleteFromWebAsync(null, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal((CollectionVerdict.Forbidden, CollectionFlows.OnlyTheAdministratorDeletesHere), (byOwner.Verdict, byOwner.Errors.Single()));
            Assert.Equal(CollectionVerdict.Unauthorized, signedOut.Verdict);
            Assert.False(this.thisStore.Collections[shared.Id].IsDeleted);

            CollectionOutcome byAdministrator = await CollectionFlows.DeleteFromWebAsync(TheAdministrator, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Removed, byAdministrator.Verdict);
            Assert.Equal(CollectionDeletedReasons.Administrator, this.thisStore.Collections[shared.Id].DeletedReason);
        }

        // ---- "Stop sharing" (owner decision, 2026-11-27: "Offline, same link back"; migration 0028) ---------------------------

        private Task<CollectionOutcome> SetSharedAsync(AccountRecord? account, string id, bool shared, DateTimeOffset? when = null) =>
            CollectionFlows.SetSharedAsync(account, id, shared, this.thisStore, Site, when ?? Now, CancellationToken.None);

        // Stopped, its captures and comments stay on the server and its date keeps counting - nothing of it is deleted.
        [Fact]
        public async Task Stopping_sharing_takes_the_page_offline_and_keeps_all_of_it()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            await this.thisStore.AddCommentAsync(shared.Id, this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256, Bo.Id, "Check R7.", Now.AddHours(1));

            CollectionOutcome stopped = await this.SetSharedAsync(Anna, shared.Id, false, Now.AddDays(2));

            Assert.Equal(CollectionVerdict.Done, stopped.Verdict);
            Assert.Equal(shared.Url, stopped.Answer!.Url);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal(Now.AddDays(2), kept.UnsharedUtc);
            Assert.False(kept.IsDeleted);
            Assert.Single(kept.Images);
            Assert.Equal(shared.DeleteAtUtc, kept.DeleteAtUtc);
            Assert.Single(await this.thisStore.ListCommentsAsync(shared.Id));
        }

        // To everybody else only that it is not shared at the moment - to its owner and the administrator the page still, said so.
        [Fact]
        public async Task A_collection_whose_sharing_stopped_is_withheld_from_all_but_its_owner_and_the_administrator()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            await this.SetSharedAsync(Anna, shared.Id, false);

            ShownCollection? toOthers = await this.ShownToAsync(shared.Id, Bo);
            ShownCollection? signedOut = await this.ShownToAsync(shared.Id, null);
            ShownCollection? toOwner = await this.ShownToAsync(shared.Id, Anna);
            ShownCollection? toAdministrator = await this.ShownToAsync(shared.Id, TheAdministrator);

            Assert.Equal((true, CollectionOffline.SharingStopped), (toOthers!.IsWithheld, toOthers.Offline));
            Assert.True(signedOut!.IsWithheld);
            Assert.Equal((false, CollectionOffline.SharingStopped), (toOwner!.IsWithheld, toOwner.Offline));
            Assert.Equal((false, CollectionOffline.SharingStopped), (toAdministrator!.IsWithheld, toAdministrator.Offline));
        }

        // Its pictures go with the page: none for anybody else, and nobody comments on it meanwhile.
        [Fact]
        public async Task A_collection_whose_sharing_stopped_shows_nobody_else_its_pictures_and_takes_no_comment()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            string sha = this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256;
            await this.SetSharedAsync(Anna, shared.Id, false);

            Task<ShownCollection?> ImageTo(AccountRecord? viewer) =>
                CollectionFlows.FindShownImageToAsync(shared.Id, sha, _ => Task.FromResult(viewer), this.thisStore, this.thisAccounts, Now, CancellationToken.None);

            Assert.Null(await ImageTo(Bo));
            Assert.Null(await ImageTo(null));
            Assert.NotNull(await ImageTo(Anna));
            Assert.Null(await CollectionFlows.FindShownAsync(shared.Id, this.thisStore, this.thisAccounts, Now, CancellationToken.None));
        }

        // "Share" - the whole collection sent again - brings it back under the SAME link; and so does sharing it again alone.
        [Fact]
        public async Task Sharing_again_brings_the_page_back_under_the_same_link()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            CollectionImageEntry first = this.thisStore.Collections[shared.Id].Images[0].Entry;

            await this.SetSharedAsync(Anna, shared.Id, false);

            CollectionOutcome sentAgain = await CollectionFlows.UpdateAsync(
                Anna, shared.Id, Content(first), this.thisStore, this.thisFiles, Site, Now.AddDays(1), CancellationToken.None);

            Assert.Equal((CollectionVerdict.Done, shared.Url), (sentAgain.Verdict, sentAgain.Answer!.Url));
            Assert.Null(this.thisStore.Collections[shared.Id].UnsharedUtc);

            await this.SetSharedAsync(Anna, shared.Id, false);
            Assert.Equal(CollectionVerdict.Done, (await this.SetSharedAsync(Anna, shared.Id, true)).Verdict);
            Assert.Null(this.thisStore.Collections[shared.Id].UnsharedUtc);
            Assert.False((await this.ShownToAsync(shared.Id, Bo))!.IsWithheld);
        }

        // The owner's alone, as sharing it is; a deleted one is gone, its page saying so.
        [Fact]
        public async Task Only_its_owner_stops_sharing_a_collection_and_a_deleted_one_is_gone()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);

            Assert.Equal(CollectionVerdict.NotFound, (await this.SetSharedAsync(Bo, shared.Id, false)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await this.SetSharedAsync(TheAdministrator, shared.Id, false)).Verdict);
            Assert.Equal(CollectionVerdict.Unauthorized, (await this.SetSharedAsync(null, shared.Id, false)).Verdict);
            Assert.Null(this.thisStore.Collections[shared.Id].UnsharedUtc);

            await CollectionFlows.DeleteAsync(Anna, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Gone, (await this.SetSharedAsync(Anna, shared.Id, false)).Verdict);
        }

        // Stopped twice keeps the first stop's time - the page has been offline since then.
        [Fact]
        public async Task Stopping_sharing_again_keeps_when_it_first_stopped()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);

            await this.SetSharedAsync(Anna, shared.Id, false, Now.AddDays(1));
            await this.SetSharedAsync(Anna, shared.Id, false, Now.AddDays(5));

            Assert.Equal(Now.AddDays(1), this.thisStore.Collections[shared.Id].UnsharedUtc);
        }

        // "You commented on" leaves out another's collection its owner stopped sharing, as its page does.
        [Fact]
        public async Task Your_collections_leaves_out_one_you_commented_on_whose_sharing_stopped()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            await this.thisStore.AddCommentAsync(shared.Id, this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256, Bo.Id, "Check R7.", Now.AddHours(1));
            await this.SetSharedAsync(Anna, shared.Id, false);

            CollectionListView toBo = await CollectionListFlows.ListForAsync(Bo, this.thisStore, this.thisAccounts, Now, CancellationToken.None);
            CollectionListView toAnna = await CollectionListFlows.ListForAsync(Anna, this.thisStore, this.thisAccounts, Now, CancellationToken.None);

            Assert.Empty(toBo.CommentedOn);
            Assert.Single(toAnna.Own);
        }

        // ---- One capture deleted from the page at once (owner request, 2026-11-28; server 16.2.0) -------------------------------

        private Task<CollectionOutcome> RemoveImageAsync(AccountRecord? account, string id, string? sha256, DateTimeOffset? when = null) =>
            CollectionFlows.RemoveImageAsync(account, id, sha256, this.thisStore, Site, when ?? Now, CancellationToken.None);

        // A collection of two captures, shared by Anna - and the two checksums.
        private async Task<(CollectionAnswer Shared, string First, string Second)> SharedWithTwoAsync()
        {
            byte[] first = TestPng.Rgb(32, 20, 1);
            byte[] second = TestPng.Rgb(32, 20, 2);
            await this.UploadAsync(Anna, first);
            await this.UploadAsync(Anna, second);

            CollectionOutcome created = await this.CreateAsync(Anna, Content(Image(ShaOf(first)), Image(ShaOf(second), pin: "13")));

            Assert.Equal(CollectionVerdict.Created, created.Verdict);
            return (created.Answer!, ShaOf(first), ShaOf(second));
        }

        // "This date should only change when the USER does something": the capture goes, and the date, the sharing, the words and
        // the other capture stay as the share left them.
        [Fact]
        public async Task Deleting_a_capture_takes_it_off_the_page_and_leaves_the_deletion_date_and_the_rest_alone()
        {
            (CollectionAnswer shared, string first, string second) = await this.SharedWithTwoAsync();
            CollectionRecord before = this.thisStore.Collections[shared.Id];

            CollectionOutcome removed = await this.RemoveImageAsync(Anna, shared.Id, first, Now.AddDays(40));

            Assert.Equal(CollectionVerdict.Done, removed.Verdict);
            Assert.Equal(1, removed.Answer!.ImageCount);
            Assert.Equal(shared.DeleteAtUtc, removed.Answer.DeleteAtUtc);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal([second], kept.Images.Select(image => image.Entry.Sha256));
            Assert.Equal(before.DeleteAtUtc, kept.DeleteAtUtc);
            Assert.Equal(before.RemindedUtc, kept.RemindedUtc);
            Assert.Equal(before.UnsharedUtc, kept.UnsharedUtc);
            Assert.Equal(before.Title, kept.Title);
            Assert.Equal(Now.AddDays(40), kept.UpdatedUtc);
        }

        // The page reads the collection as it is now - the capture is not on it any more.
        [Fact]
        public async Task A_deleted_capture_is_on_the_page_no_more()
        {
            (CollectionAnswer shared, string first, _) = await this.SharedWithTwoAsync();

            await this.RemoveImageAsync(Anna, shared.Id, first);

            ShownCollection? page = await this.ShownToAsync(shared.Id, null);

            Assert.DoesNotContain(page!.Collection.Images, image => image.Entry.Sha256 == first);
        }

        // The last one may go too - the page then says it has none - and one deleted twice, or never there, is done.
        [Fact]
        public async Task The_last_capture_may_go_and_one_not_there_is_done()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna);
            string only = this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256;

            Assert.Equal(CollectionVerdict.Done, (await this.RemoveImageAsync(Anna, shared.Id, only)).Verdict);
            Assert.Equal(CollectionVerdict.Done, (await this.RemoveImageAsync(Anna, shared.Id, only)).Verdict);
            Assert.Empty(this.thisStore.Collections[shared.Id].Images);
            Assert.False(this.thisStore.Collections[shared.Id].IsDeleted);
        }

        // Its owner's alone - anybody else's "not found", the administrator's too (who deletes a whole collection on the website);
        // a deleted collection is gone; an address naming no checksum is refused.
        [Fact]
        public async Task Only_its_owner_deletes_a_capture_and_a_deleted_collection_is_gone()
        {
            (CollectionAnswer shared, string first, _) = await this.SharedWithTwoAsync();

            Assert.Equal(CollectionVerdict.NotFound, (await this.RemoveImageAsync(Bo, shared.Id, first)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await this.RemoveImageAsync(TheAdministrator, shared.Id, first)).Verdict);
            Assert.Equal(CollectionVerdict.Unauthorized, (await this.RemoveImageAsync(null, shared.Id, first)).Verdict);
            Assert.Equal(CollectionVerdict.BadRequest, (await this.RemoveImageAsync(Anna, shared.Id, first.ToUpperInvariant())).Verdict);
            Assert.Equal(2, this.thisStore.Collections[shared.Id].Images.Count);

            await CollectionFlows.DeleteAsync(Anna, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Gone, (await this.RemoveImageAsync(Anna, shared.Id, first)).Verdict);
        }

        // ---- The oscilloscopes the captures were taken with (owner request, 2026-11-28; migration 0030) --------------------------

        [Fact]
        public async Task The_oscilloscopes_sent_are_kept_trimmed_each_once()
        {
            byte[] capture = TestPng.Rgb(32, 20, 1);
            await this.UploadAsync(Anna, capture);

            CollectionOutcome created = await this.CreateAsync(
                Anna,
                Content(Image(ShaOf(capture))) with { Oscilloscopes = [" Rigol DS1054Z ", "rigol ds1054z", "Siglent SDS1104X-E"] });

            Assert.Equal(CollectionVerdict.Created, created.Verdict);
            Assert.Equal(["Rigol DS1054Z", "Siglent SDS1104X-E"], this.thisStore.Collections[created.Answer!.Id].Oscilloscopes);
        }

        // ---- A capture's component and pinout drawing (owner request, 2026-11-27; migration 0029) ------------------------------

        // The drawing is a picture of the account's as a capture is: asked for when not uploaded, then kept with the capture, its
        // component's facts with the collection - and served as one of the collection's pictures.
        [Fact]
        public async Task A_capture_s_pinout_drawing_is_a_picture_of_the_collection_and_its_component_is_kept()
        {
            byte[] capture = TestPng.Rgb(32, 20, 1);
            byte[] pinout = TestPng.Rgb(40, 60, 2);
            CollectionComponent component = new(" U7 ", "U7 | VIC-II", "IC | DIP-40", null, null, [new CollectionPin("12", "/RAS", "Output", null, " Pulsing ")]);
            CollectionContentRequest content = Content(Image(ShaOf(capture)) with { PinoutSha256 = ShaOf(pinout) }) with { Components = [component] };

            await this.UploadAsync(Anna, capture);

            CollectionOutcome missing = await this.CreateAsync(Anna, content);

            Assert.Equal(CollectionVerdict.MissingImages, missing.Verdict);
            Assert.Equal([ShaOf(pinout)], missing.Missing);

            await this.UploadAsync(Anna, pinout);

            CollectionOutcome created = await this.CreateAsync(Anna, content);

            Assert.Equal(CollectionVerdict.Created, created.Verdict);

            CollectionRecord kept = this.thisStore.Collections[created.Answer!.Id];

            Assert.Equal(ShaOf(pinout), kept.Images[0].Entry.PinoutSha256);
            Assert.Equal(new CollectionComponent("U7", "U7 | VIC-II", "IC | DIP-40", null, null, null) with { Pins = null }, Assert.Single(kept.Components!) with { Pins = null });
            Assert.Equal([new CollectionPin("12", "/RAS", "Output", null, "Pulsing")], kept.Components![0].Pins);

            ShownCollection? served = await CollectionFlows.FindShownImageToAsync(
                created.Answer.Id, ShaOf(pinout), _ => Task.FromResult<AccountRecord?>(null), this.thisStore, this.thisAccounts, Now, CancellationToken.None);

            Assert.NotNull(served);
        }

        // The hourly sweep lets go of a picture no collection shows - a pinout drawing a collection shows stays.
        [Fact]
        public async Task The_sweep_keeps_a_pinout_drawing_a_collection_shows()
        {
            byte[] capture = TestPng.Rgb(32, 20, 3);
            byte[] pinout = TestPng.Rgb(40, 60, 4);

            await this.UploadAsync(Anna, capture);
            await this.UploadAsync(Anna, pinout);
            await this.CreateAsync(Anna, Content(Image(ShaOf(capture)) with { PinoutSha256 = ShaOf(pinout) }));

            await CollectionFlows.SweepAsync(this.thisStore, this.thisFiles, this.thisAccounts, this.thisMailer, Site, Now.AddDays(3), CancellationToken.None);

            Assert.True(File.Exists(this.thisFiles.PathOf(ShaOf(pinout))));
            Assert.Contains(this.thisStore.Uploads, upload => upload.Sha256 == ShaOf(pinout));
        }

        // ---- "Download all" ---------------------------------------------------------------------------------------------

        // Every capture in the page's order, named as the viewer's "Download" names it.
        [Fact]
        public void Download_all_names_each_capture_by_its_place_and_anchor()
        {
            CollectionRecord collection = new(
                "k7mq3x9vbn2p", Anna.Id, "Commodore 64", "250407", null, null, null, [], Now, Now, Now.AddYears(1), null, null, null,
                [new(Image(new string('a', 64), "U7", "12"), 8, 8), new(Image(new string('b', 64), "U7", "12"), 8, 8), new(Image(new string('c', 64), "R5", "1"), 8, 8)],
                3);

            Assert.Equal(
                ["001-U7-12.png", "002-U7-12-2.png", "003-R5-1.png"],
                CollectionDownload.EntriesOf(collection).Select(entry => entry.Name));
        }

        // Kestrel refuses a synchronous write to a response - and ZipArchive makes one whatever is asked of it (this test caught
        // .NET 10's async API writing each entry's data descriptor synchronously) - so the response is written only
        // asynchronously, here into a stream that throws on any synchronous write; a picture gone meanwhile is left out.
        [Fact]
        public async Task Download_all_writes_its_zip_without_a_synchronous_write()
        {
            byte[] first = TestPng.Rgb(8, 8, 1);
            byte[] second = TestPng.Rgb(8, 8, 2);
            var pictures = new Dictionary<string, byte[]> { ["a"] = first, ["b"] = second };
            using var target = new AsyncOnlyStream();

            await CollectionDownload.WriteAsync(
                target,
                [new("001-U7-12.png", "a"), new("002-U8-1.png", "gone"), new("003-U8-2.png", "b")],
                sha => pictures.TryGetValue(sha, out byte[]? bytes) ? new MemoryStream(bytes) : null,
                CancellationToken.None);

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(target.Written.ToArray()));

            Assert.Equal(["001-U7-12.png", "003-U8-2.png"], zip.Entries.Select(entry => entry.FullName));

            using var read = new MemoryStream();
            using (Stream entry = zip.Entries[1].Open())
                entry.CopyTo(read);

            Assert.Equal(second, read.ToArray());
        }

        // A response body as Kestrel's: write only forward, asynchronously.
        private sealed class AsyncOnlyStream : Stream
        {
            public MemoryStream Written { get; } = new();

            public override bool CanRead => false;

            public override bool CanSeek => false;

            public override bool CanWrite => true;

            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => this.Written.Length;
                set => throw new NotSupportedException();
            }

            public override void Flush() => throw new InvalidOperationException("Synchronous operations are disallowed.");

            public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous operations are disallowed.");

            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                this.Written.Write(buffer.Span);
                return ValueTask.CompletedTask;
            }

            public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
                this.WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                    this.Written.Dispose();

                base.Dispose(disposing);
            }
        }
    }
}
