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
            new(sha, label, pin, "/RAS", "Pulsing", "50us", "2V", "1V", null, null, LooksWrong: false, CapturedUtc: Now);

        private static CollectionContentRequest Content(params CollectionImageEntry[] images) =>
            new("Commodore 64", "250407", "PAL", "No picture", "It went black.", ["https://www.lemon64.com/forum/viewtopic.php?t=1"], CollectionRules.DeleteAt(Now, "1y"), images);

        private CollectionStorage Storage(long maximum = 0, long? free = null, long reserve = 0) => new(maximum, () => free, reserve);

        private Task<CollectionOutcome> UploadAsync(AccountRecord? account, byte[] bytes, CollectionStorage? storage = null, DateTimeOffset? when = null) =>
            CollectionFlows.UploadImageAsync(account, ShaOf(bytes), bytes, this.thisStore, this.thisFiles, storage ?? this.Storage(), when ?? Now, CancellationToken.None);

        private Task<CollectionOutcome> CreateAsync(AccountRecord? account, CollectionContentRequest content, DateTimeOffset? when = null) =>
            CollectionFlows.CreateAsync(account, content, this.thisStore, this.thisFiles, Site, when ?? Now, CancellationToken.None);

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

            CollectionOutcome outcome = await CollectionFlows.ListMineAsync(Anna, this.thisStore, Site, CancellationToken.None);

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

        [Fact]
        public async Task An_owner_whose_account_is_locked_or_gone_is_mailed_nothing()
        {
            CollectionAnswer shared = await this.SharedAsync(Anna, lifetime: "1y");
            this.thisAccounts.Accounts[Anna.Id] = Anna with { IsLocked = true };

            await this.SweepAsync(shared.DeleteAtUtc.AddDays(-29));

            Assert.Empty(this.thisMailer.Sent);
            Assert.NotNull(this.thisStore.Collections[shared.Id].RemindedUtc);
        }
    }
}
