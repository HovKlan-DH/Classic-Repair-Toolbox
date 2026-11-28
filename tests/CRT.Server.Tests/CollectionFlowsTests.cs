using System.Security.Cryptography;
using CRT.Data.Tests;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Sharing an oscilloscope collection (CollectionFlows): owned by its KEY, no account (owner decision, 2026-11-28; server
    // 17.0.0), what a collection may show, the limits per internet connection, and the hourly sweep - on the fake store, the
    // pictures in a folder of the test's own.
    // ###########################################################################################
    public sealed class CollectionFlowsTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private const string Site = "https://classic-repair-toolbox.dk";

        // Two computers' keys, each from an internet connection of its own.
        private static readonly string AnnaKey = new('a', CollectionRules.OwnerKeyLength);

        private static readonly string BoKey = new('b', CollectionRules.OwnerKeyLength);

        private const string AnnaConnection = "198.51.100.7";

        private const string BoConnection = "2001:db8:5:7::/64";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-collections-" + Guid.NewGuid().ToString("N"));

        private readonly FakeCollectionStore thisStore = new();

        private readonly CollectionFiles thisFiles;

        // The website's signed-in readers: the administrator, and a maintainer who is none.
        private static readonly AccountRecord TheAdministrator = Account(9, "admin@example.com", "Dennis@CRT", administrator: true);

        private static readonly AccountRecord AMaintainer = Account(2, "bo@example.com", "Bo");

        public CollectionFlowsTests()
        {
            Directory.CreateDirectory(this.thisRoot);
            this.thisFiles = new CollectionFiles(this.thisRoot);
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

        private static AccountRecord Account(long id, string email, string name, bool administrator = false) =>
            new(id, email, email, "hash", name, IsVerified: true, administrator, IsLocked: false, Now.AddDays(-10), null);

        private static string ShaOf(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

        private static string ConnectionOf(string? key) => key == BoKey ? BoConnection : AnnaConnection;

        private static CollectionImageEntry Image(string sha, string label = "U7", string pin = "12") =>
            new(sha, label, pin, "/RAS", "Pulsing", "50us", "2V", "1V", null, CapturedUtc: Now);

        private static CollectionContentRequest Content(params CollectionImageEntry[] images) =>
            new("Commodore 64", "250407", "PAL", "No picture", "It went black.", ["https://www.lemon64.com/forum/viewtopic.php?t=1"], CollectionRules.DeleteAt(Now, "1y"), images);

        private CollectionStorage Storage(long maximum = 0, long? free = null, long reserve = 0) => new(maximum, () => free, reserve);

        private Task<CollectionOutcome> UploadAsync(string? key, byte[] bytes, CollectionStorage? storage = null, DateTimeOffset? when = null, string? connection = null) =>
            CollectionFlows.UploadImageAsync(key, connection ?? ConnectionOf(key), ShaOf(bytes), bytes, this.thisStore, this.thisFiles, storage ?? this.Storage(), when ?? Now, CancellationToken.None);

        private Task<CollectionOutcome> CreateAsync(string? key, CollectionContentRequest content, DateTimeOffset? when = null, string? connection = null) =>
            CollectionFlows.CreateAsync(key, connection ?? ConnectionOf(key), content, this.thisStore, this.thisFiles, Site, when ?? Now, CancellationToken.None);

        private Task<CollectionOutcome> UpdateAsync(string? key, string id, CollectionContentRequest content, DateTimeOffset? when = null) =>
            CollectionFlows.UpdateAsync(key, id, content, this.thisStore, this.thisFiles, Site, when ?? Now, CancellationToken.None);

        private Task<CollectionOutcome> DeleteAsync(string? key, string id) =>
            CollectionFlows.DeleteAsync(key, id, this.thisStore, Now, CancellationToken.None);

        // A collection of one capture, shared under `key` - uploaded, then sent.
        private async Task<CollectionAnswer> SharedAsync(string key, int seed = 1, DateTimeOffset? when = null, string lifetime = "1y", string? connection = null)
        {
            byte[] png = TestPng.Rgb(32, 20, seed);

            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(key, png, when: when, connection: connection)).Verdict);

            CollectionOutcome created = await this.CreateAsync(
                key, Content(Image(ShaOf(png))) with { DeleteAtUtc = CollectionRules.DeleteAt(when ?? Now, lifetime) }, when, connection);

            Assert.Equal(CollectionVerdict.Created, created.Verdict);
            return created.Answer!;
        }

        // ---- The key ------------------------------------------------------------------------

        // A request from a CRT that shared with an account carries no key - told to update CRT. A key of another shape is none.
        [Fact]
        public async Task A_request_with_no_key_shares_nothing_and_is_told_to_update_CRT()
        {
            foreach (string? key in new[] { null, "", "../../etc/passwd", AnnaKey.ToUpperInvariant(), AnnaKey[..63] })
            {
                CollectionOutcome upload = await this.UploadAsync(key, TestPng.Rgb(8, 8));
                CollectionOutcome create = await this.CreateAsync(key, Content(Image(new string('a', 64))));

                Assert.Equal((CollectionVerdict.BadRequest, CollectionFlows.NoKey), (upload.Verdict, upload.Errors.Single()));
                Assert.Equal((CollectionVerdict.BadRequest, CollectionFlows.NoKey), (create.Verdict, create.Errors.Single()));
            }

            Assert.Empty(this.thisStore.Uploads);
            Assert.Empty(this.thisStore.Collections);
        }

        // Only the key's hash is kept - as a session's: the database read off a backup changes nobody's page.
        [Fact]
        public async Task Only_the_hash_of_the_key_is_kept()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal(CollectionFlows.OwnerHashOf(AnnaKey), kept.OwnerKeyHash);
            Assert.NotEqual(AnnaKey, kept.OwnerKeyHash);
            Assert.All(this.thisStore.Uploads, upload => Assert.Equal(kept.OwnerKeyHash, upload.OwnerKeyHash));
            Assert.Null(CollectionFlows.OwnerHashOf("not a key"));
        }

        // ---- Uploading a capture --------------------------------------------------------

        [Fact]
        public async Task A_capture_uploaded_is_kept_once_and_its_size_answered()
        {
            byte[] png = TestPng.Rgb(800, 480);

            CollectionOutcome outcome = await this.UploadAsync(AnnaKey, png);

            Assert.Equal(CollectionVerdict.Done, outcome.Verdict);
            Assert.Equal(new CollectionImageUploadAnswer(ShaOf(png), 800, 480), outcome.Upload);
            Assert.Equal(png, File.ReadAllBytes(this.thisFiles.PathOf(ShaOf(png))));
            Assert.Equal(AnnaConnection, Assert.Single(this.thisStore.Uploads).Connection);

            // The same again is done, and counts nothing more.
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(AnnaKey, png)).Verdict);
            Assert.Single(this.thisStore.Uploads);
        }

        [Fact]
        public async Task Bytes_that_are_not_their_checksum_are_refused_and_nothing_is_kept()
        {
            byte[] png = TestPng.Rgb(8, 8);

            CollectionOutcome outcome = await CollectionFlows.UploadImageAsync(
                AnnaKey, AnnaConnection, new string('a', 64), png, this.thisStore, this.thisFiles, this.Storage(), Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal([CollectionFlows.ChecksumMismatch], outcome.Errors);
            Assert.Empty(Directory.GetFiles(this.thisRoot, "*", SearchOption.AllDirectories));
        }

        [Fact]
        public async Task A_file_that_is_no_capture_is_refused_and_nothing_is_kept()
        {
            byte[] notPng = "this is a zip file, honestly"u8.ToArray();

            CollectionOutcome outcome = await this.UploadAsync(AnnaKey, notPng);

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal([CollectionPng.NotAPicture], outcome.Errors);
            Assert.Empty(this.thisStore.Uploads);
            Assert.Empty(Directory.GetFiles(this.thisRoot, "*", SearchOption.AllDirectories));
        }

        [Fact]
        public async Task An_address_naming_no_checksum_leads_nowhere()
        {
            CollectionOutcome outcome = await CollectionFlows.UploadImageAsync(
                AnnaKey, AnnaConnection, "../../etc/passwd", TestPng.Rgb(8, 8), this.thisStore, this.thisFiles, this.Storage(), Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
        }

        // *** THE DAY'S LIMIT IS THE INTERNET CONNECTION'S, NOT THE KEY'S *** - a key is free to make, so a new key for every
        // capture would otherwise pass any limit.
        [Fact]
        public async Task An_internet_connection_uploads_its_day_s_captures_and_no_more_whatever_key_it_sends()
        {
            for (int i = 0; i < CollectionRules.MaximumUploadsPerDay; i++)
            {
                this.thisStore.Uploads.Add(new CollectionUploadRecord(
                    CollectionFlows.OwnerHashOf(CollectionRules.NewOwnerKey())!,
                    i.ToString("x64", System.Globalization.CultureInfo.InvariantCulture),
                    10, 8, 8, Now.AddHours(-2), AnnaConnection));
            }

            Assert.Equal(CollectionVerdict.TooMany, (await this.UploadAsync(AnnaKey, TestPng.Rgb(8, 8))).Verdict);
            Assert.Equal(CollectionVerdict.TooMany, (await this.UploadAsync(CollectionRules.NewOwnerKey(), TestPng.Rgb(8, 8), connection: AnnaConnection)).Verdict);

            // Another connection's day is its own, and a day later it is past.
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(BoKey, TestPng.Rgb(8, 8))).Verdict);
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(AnnaKey, TestPng.Rgb(8, 8), when: Now.AddDays(1))).Verdict);
        }

        // A connection not known is counted as ONE - an unknown sender is limited as a whole.
        [Fact]
        public async Task Uploads_from_a_connection_not_known_are_counted_together()
        {
            byte[] png = TestPng.Rgb(8, 8, seed: 1);

            await CollectionFlows.UploadImageAsync(AnnaKey, null, ShaOf(png), png, this.thisStore, this.thisFiles, this.Storage(), Now, CancellationToken.None);

            CollectionUploadRecord kept = Assert.Single(this.thisStore.Uploads);

            Assert.Equal(CollectionFlows.UnknownConnection, kept.Connection);
        }

        [Fact]
        public async Task With_no_room_left_a_capture_is_refused_saying_so()
        {
            Assert.Equal(CollectionVerdict.NoRoom, (await this.UploadAsync(AnnaKey, TestPng.Rgb(8, 8), this.Storage(maximum: 10))).Verdict);
            Assert.Equal(CollectionVerdict.NoRoom, (await this.UploadAsync(AnnaKey, TestPng.Rgb(8, 8), this.Storage(free: 1000, reserve: 5000))).Verdict);

            // A disk that cannot be measured is taken as having room, as the blob store takes it.
            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(AnnaKey, TestPng.Rgb(8, 8), this.Storage(free: null, reserve: 5000))).Verdict);
        }

        // Bytes kept already take no more room - another key sending the same capture is not refused for a full disk.
        [Fact]
        public async Task A_capture_kept_already_needs_no_room()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(AnnaKey, png);

            Assert.Equal(CollectionVerdict.Done, (await this.UploadAsync(BoKey, png, this.Storage(maximum: 1))).Verdict);
            Assert.Equal(2, this.thisStore.Uploads.Count);
        }

        // ---- Sharing a collection ----------------------------------------------------------

        [Fact]
        public async Task A_collection_naming_captures_not_uploaded_is_answered_with_them()
        {
            byte[] first = TestPng.Rgb(8, 8, seed: 1);
            byte[] second = TestPng.Rgb(8, 8, seed: 2);

            await this.UploadAsync(AnnaKey, first);

            CollectionOutcome outcome = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(first)), Image(ShaOf(second), "U8", "3")));

            Assert.Equal(CollectionVerdict.MissingImages, outcome.Verdict);
            Assert.Equal([ShaOf(second)], outcome.Missing);
            Assert.Empty(this.thisStore.Collections);
        }

        [Fact]
        public async Task A_collection_with_its_captures_uploaded_is_shared_under_a_new_id()
        {
            byte[] png = TestPng.Rgb(800, 480);

            await this.UploadAsync(AnnaKey, png);

            CollectionOutcome outcome = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.Created, outcome.Verdict);
            Assert.True(CollectionRules.IsId(outcome.Answer!.Id));
            Assert.Equal($"{Site}/collections/{outcome.Answer.Id}", outcome.Answer.Url);
            Assert.Equal(1, outcome.Answer.ImageCount);

            CollectionRecord kept = this.thisStore.Collections[outcome.Answer.Id];

            Assert.Equal((AnnaConnection, Now), (kept.Connection, kept.CreatedUtc));
            Assert.Equal((800, 480), (kept.Images[0].Width, kept.Images[0].Height));
        }

        // *** A CHECKSUM READ OFF SOMEBODY ELSE'S PAGE IS NO WAY TO SHOW THEIR PICTURE. ***
        [Fact]
        public async Task A_collection_shows_only_captures_its_own_key_uploaded()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(AnnaKey, png);

            CollectionOutcome outcome = await this.CreateAsync(BoKey, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.MissingImages, outcome.Verdict);
            Assert.Equal([ShaOf(png)], outcome.Missing);
        }

        // The bytes gone from disk (moved to another disk and one lost, say): the capture is asked for again, not shown broken.
        [Fact]
        public async Task A_capture_whose_bytes_are_gone_is_asked_for_again()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(AnnaKey, png);
            File.Delete(this.thisFiles.PathOf(ShaOf(png)));

            CollectionOutcome outcome = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.MissingImages, outcome.Verdict);

            // Sent again, the bytes are kept again.
            await this.UploadAsync(AnnaKey, png);
            Assert.Equal(CollectionVerdict.Created, (await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))))).Verdict);
        }

        [Fact]
        public async Task A_collection_breaking_a_rule_is_refused_in_the_rules_words()
        {
            CollectionOutcome outcome = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(TestPng.Rgb(8, 8)))) with { HardwareName = "" });

            Assert.Equal(CollectionVerdict.BadRequest, outcome.Verdict);
            Assert.Equal(["The collection does not say the hardware it is about."], outcome.Errors);
        }

        // A NEW collection a day per connection, each under a key of its own; another connection's, the next day, and a
        // collection shared AGAIN (its page changed) count nothing against it.
        [Fact]
        public async Task An_internet_connection_shares_so_many_new_collections_a_day()
        {
            string firstKey = CollectionRules.NewOwnerKey();

            await this.SharedAsync(firstKey, seed: 0, connection: AnnaConnection);

            for (int i = 1; i < CollectionRules.MaximumCollectionsPerDay; i++)
                await this.SharedAsync(CollectionRules.NewOwnerKey(), seed: i, connection: AnnaConnection);

            string oneMore = CollectionRules.NewOwnerKey();
            byte[] png = TestPng.Rgb(8, 8, seed: 999);

            await this.UploadAsync(oneMore, png, connection: AnnaConnection);

            CollectionOutcome today = await this.CreateAsync(oneMore, Content(Image(ShaOf(png))), connection: AnnaConnection);

            Assert.Equal(CollectionVerdict.TooMany, today.Verdict);
            Assert.Contains("internet connection", today.Errors[0], StringComparison.Ordinal);

            // Its page shared again is no new collection.
            byte[] first = TestPng.Rgb(32, 20, 0);
            Assert.Equal(CollectionVerdict.Done, (await this.CreateAsync(firstKey, Content(Image(ShaOf(first))) with { Title = "Again" }, connection: AnnaConnection)).Verdict);

            Assert.Equal(CollectionVerdict.Created, (await this.CreateAsync(oneMore, Content(Image(ShaOf(png))), connection: BoConnection)).Verdict);
        }

        // ---- Changing and deleting one -----------------------------------------------------

        [Fact]
        public async Task A_collection_changed_keeps_its_id_creation_and_connection_and_takes_its_new_captures()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            byte[] added = TestPng.Rgb(16, 16, seed: 5);

            await this.UploadAsync(AnnaKey, added, connection: BoConnection);

            CollectionImageEntry first = this.thisStore.Collections[shared.Id].Images[0].Entry;
            CollectionOutcome outcome = await this.UpdateAsync(AnnaKey, shared.Id, Content(first, Image(ShaOf(added), "U8", "1")) with { Title = "Found it" }, Now.AddDays(3));

            Assert.Equal(CollectionVerdict.Done, outcome.Verdict);
            Assert.Equal(shared.Id, outcome.Answer!.Id);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal("Found it", kept.Title);
            Assert.Equal((Now, AnnaConnection), (kept.CreatedUtc, kept.Connection));
            Assert.Equal(Now.AddDays(3), kept.UpdatedUtc);
            Assert.Equal(["U7", "U8"], kept.Images.Select(image => image.Entry.ComponentLabel));
        }

        // Whether a collection exists is nobody else's business.
        [Fact]
        public async Task Another_key_finds_no_such_collection_to_change_or_delete()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);

            CollectionOutcome changed = await this.UpdateAsync(BoKey, shared.Id, Content(Image(new string('a', 64))));
            CollectionOutcome deleted = await this.DeleteAsync(BoKey, shared.Id);
            CollectionOutcome noKey = await this.DeleteAsync(null, shared.Id);

            Assert.Equal(CollectionVerdict.NotFound, changed.Verdict);
            Assert.Equal((CollectionVerdict.NotFound, CollectionFlows.NoSuchCollection), (deleted.Verdict, deleted.Errors.Single()));
            Assert.Equal((CollectionVerdict.BadRequest, CollectionFlows.NoKey), (noKey.Verdict, noKey.Errors.Single()));
            Assert.False(this.thisStore.Collections[shared.Id].IsDeleted);
        }

        // *** A COLLECTION SHARED WITH AN ACCOUNT BEFORE SERVER 17.0.0 IS OWNED BY NO KEY. *** Its page stays until its date, or the
        // administrator's delete on the website - no key changes, stops or deletes it.
        [Fact]
        public async Task A_collection_shared_with_an_account_before_is_owned_by_no_key()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            this.thisStore.Collections[shared.Id] = this.thisStore.Collections[shared.Id] with { OwnerKeyHash = null };

            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;

            Assert.Equal(CollectionVerdict.NotFound, (await this.UpdateAsync(AnnaKey, shared.Id, Content(image))).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await this.DeleteAsync(AnnaKey, shared.Id)).Verdict);
            Assert.Equal(CollectionVerdict.NotFound, (await CollectionFlows.SetSharedAsync(AnnaKey, shared.Id, false, this.thisStore, Site, Now, CancellationToken.None)).Verdict);
            Assert.NotNull(await this.ShownToAsync(shared.Id, null));
            Assert.False(CollectionFlows.IsOwnedBy(this.thisStore.Collections[shared.Id], CollectionFlows.OwnerHashOf(AnnaKey)!));
        }

        [Fact]
        public async Task The_owner_deletes_a_collection_and_it_can_no_longer_be_changed()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);

            Assert.Equal(CollectionVerdict.Removed, (await this.DeleteAsync(AnnaKey, shared.Id)).Verdict);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal(CollectionDeletedReasons.Owner, kept.DeletedReason);
            Assert.Null(kept.Title);
            Assert.Empty(kept.Images);

            // Deleting it again is done; changing it is told it went.
            Assert.Equal(CollectionVerdict.Removed, (await this.DeleteAsync(AnnaKey, shared.Id)).Verdict);

            CollectionOutcome changed = await this.UpdateAsync(AnnaKey, shared.Id, Content(Image(new string('a', 64))));

            Assert.Equal(CollectionVerdict.Gone, changed.Verdict);
            Assert.Equal(["That collection was deleted on 2026-November-26."], changed.Errors);
        }

        // The owner deletes their own in CRT, whose window keeps in step with the server - the website's delete is the
        // administrator's alone.
        [Fact]
        public async Task The_website_s_delete_is_the_administrator_s_alone()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);

            CollectionOutcome byMaintainer = await CollectionFlows.DeleteFromWebAsync(AMaintainer, shared.Id, this.thisStore, Now, CancellationToken.None);
            CollectionOutcome signedOut = await CollectionFlows.DeleteFromWebAsync(null, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal((CollectionVerdict.Forbidden, CollectionFlows.OnlyTheAdministratorDeletesHere), (byMaintainer.Verdict, byMaintainer.Errors.Single()));
            Assert.Equal(CollectionVerdict.Unauthorized, signedOut.Verdict);
            Assert.False(this.thisStore.Collections[shared.Id].IsDeleted);

            CollectionOutcome byAdministrator = await CollectionFlows.DeleteFromWebAsync(TheAdministrator, shared.Id, this.thisStore, Now, CancellationToken.None);

            Assert.Equal(CollectionVerdict.Removed, byAdministrator.Verdict);
            Assert.Equal(CollectionDeletedReasons.Administrator, this.thisStore.Collections[shared.Id].DeletedReason);
        }

        // ---- The hourly sweep ----------------------------------------------------------------

        private Task<CollectionSweepResult> SweepAsync(DateTimeOffset when) =>
            CollectionFlows.SweepAsync(this.thisStore, this.thisFiles, when, CancellationToken.None);

        [Fact]
        public async Task A_collection_whose_time_is_up_is_deleted_and_its_capture_goes_a_day_later()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey, lifetime: "1m");
            string sha = this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256;

            Assert.Equal(new CollectionSweepResult(0, 0), await this.SweepAsync(Now.AddDays(20)));

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

            await this.UploadAsync(AnnaKey, held);
            await this.UploadAsync(AnnaKey, free);

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

            await this.UploadAsync(AnnaKey, failing);
            await this.UploadAsync(AnnaKey, fine);
            this.thisStore.IsUploadedFails.Add(ShaOf(failing));

            CollectionSweepResult result = await this.SweepAsync(Now.AddDays(2));

            Assert.Equal([ShaOf(failing)], result.FilesNotDeleted);
            Assert.Equal(1, result.FilesDeleted);
            Assert.False(this.thisFiles.Has(ShaOf(fine)));
        }

        // Another key sending the very same picture as the sweep lets go of it keeps the bytes - its upload wrote none, as they
        // were here (code review, 2026-11-27).
        [Fact]
        public async Task A_picture_sent_again_under_another_key_as_it_is_swept_stays()
        {
            byte[] png = TestPng.Rgb(8, 8, seed: 4);

            await this.UploadAsync(AnnaKey, png);
            this.thisStore.AfterDeleteUnusedUploads = () =>
                this.thisStore.Uploads.Add(new CollectionUploadRecord(CollectionFlows.OwnerHashOf(BoKey)!, ShaOf(png), png.Length, 8, 8, Now.AddDays(2), BoConnection));

            CollectionSweepResult result = await this.SweepAsync(Now.AddDays(2));

            Assert.Equal(0, result.FilesDeleted);
            Assert.True(this.thisFiles.Has(ShaOf(png)));
        }

        // A capture uploaded for a collection about to be sent is not taken from under it.
        [Fact]
        public async Task A_capture_uploaded_today_and_shown_nowhere_yet_stays_a_day()
        {
            byte[] png = TestPng.Rgb(8, 8);

            await this.UploadAsync(AnnaKey, png);

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

            await this.UploadAsync(AnnaKey, png);
            await this.UploadAsync(AnnaKey, png, when: Now.AddDays(10));

            await this.SweepAsync(Now.AddDays(10).AddHours(23));

            Assert.True(this.thisFiles.Has(ShaOf(png)));
            Assert.Equal(Now.AddDays(10), Assert.Single(this.thisStore.Uploads).UploadedUtc);
        }

        // Two keys sent the same picture: one collection gone takes only its own key's upload, never the bytes the other still
        // shows.
        [Fact]
        public async Task A_capture_another_collection_still_shows_stays()
        {
            CollectionAnswer annas = await this.SharedAsync(AnnaKey, seed: 7, lifetime: "1m");
            await this.SharedAsync(BoKey, seed: 7, lifetime: "1y");

            string sha = this.thisStore.Collections[annas.Id].Images[0].Entry.Sha256;

            CollectionSweepResult result = await this.SweepAsync(annas.DeleteAtUtc.AddMinutes(1));

            Assert.Equal(1, result.Expired);
            Assert.Equal(0, result.FilesDeleted);
            Assert.True(this.thisFiles.Has(sha));
            Assert.Single(this.thisStore.Uploads);
        }

        // The deletion date as the database keeps it, whole milliseconds - else the same date shared again read as later than
        // the stored one (code review, 2026-11-27).
        [Fact]
        public async Task The_deletion_date_is_kept_to_the_millisecond_so_the_same_date_again_is_no_later()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey, lifetime: "1y");
            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;
            DateTimeOffset deleteAt = this.thisStore.Collections[shared.Id].DeleteAtUtc.AddTicks(1234);

            await this.UpdateAsync(AnnaKey, shared.Id, Content(image) with { DeleteAtUtc = deleteAt });

            Assert.Equal(0, this.thisStore.Collections[shared.Id].DeleteAtUtc.UtcTicks % TimeSpan.TicksPerMillisecond);
            Assert.Equal(CollectionFlows.WholeMilliseconds(deleteAt), this.thisStore.Collections[shared.Id].DeleteAtUtc);
        }

        // ---- A first share sent again under its key ----------------------------------------------

        // A first share the server kept whose answer was lost is sent again under the same key: the page it made is changed,
        // never a second one made that only the administrator could take down (code review, 2026-11-27).
        [Fact]
        public async Task A_first_share_sent_again_under_its_key_changes_the_page_it_made()
        {
            byte[] png = TestPng.Rgb(32, 20, seed: 3);
            await this.UploadAsync(AnnaKey, png);

            CollectionOutcome first = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));
            CollectionOutcome again = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))) with { Title = "Typed since" });

            Assert.Equal(first.Answer!.Id, again.Answer!.Id);
            Assert.Equal("Typed since", Assert.Single(this.thisStore.Collections.Values).Title);

            // Another key's first share is a page of its own.
            await this.UploadAsync(BoKey, png);
            CollectionOutcome bos = await this.CreateAsync(BoKey, Content(Image(ShaOf(png))));

            Assert.NotEqual(first.Answer.Id, bos.Answer!.Id);
        }

        // A first share retried while the first is still being written waits for it, then finds its page under the key - looked
        // up before the lock, it found none and made a second page (code review, 2026-11-27).
        [Fact]
        public async Task A_first_share_retried_while_the_first_is_written_makes_no_second_page()
        {
            byte[] png = TestPng.Rgb(32, 20, seed: 5);
            await this.UploadAsync(AnnaKey, png);

            Task<CollectionOutcome>? retried = null;
            this.thisStore.WhileCountingCreated = () => retried = this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));

            CollectionOutcome first = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));
            CollectionOutcome again = await retried!;

            Assert.Equal(first.Answer!.Id, again.Answer!.Id);
            Assert.Single(this.thisStore.Collections.Values);
        }

        // A page taken down is no page to change: shared again under its key, a new one is made.
        [Fact]
        public async Task A_key_whose_page_was_taken_down_makes_a_new_page()
        {
            byte[] png = TestPng.Rgb(32, 20, seed: 4);
            await this.UploadAsync(AnnaKey, png);

            CollectionOutcome first = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));
            await this.DeleteAsync(AnnaKey, first.Answer!.Id);

            CollectionOutcome again = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(png))));

            Assert.Equal(CollectionVerdict.Created, again.Verdict);
            Assert.NotEqual(first.Answer.Id, again.Answer!.Id);
        }

        // Its time up, a collection is gone for "Share" too, which then makes a new page rather than bring the old one back - as
        // its page says (code review, 2026-11-27).
        [Fact]
        public async Task A_collection_whose_time_is_up_is_gone_for_share_too()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey, lifetime: "1m");
            DateTimeOffset after = shared.DeleteAtUtc.AddMinutes(1);
            CollectionImageEntry image = this.thisStore.Collections[shared.Id].Images[0].Entry;
            CollectionContentRequest again = Content(image) with { DeleteAtUtc = CollectionRules.DeleteAt(after, "1y") };

            Assert.Equal(CollectionVerdict.Gone, (await this.UpdateAsync(AnnaKey, shared.Id, again, after)).Verdict);

            await this.UploadAsync(AnnaKey, TestPng.Rgb(32, 20, 1), when: after);
            CollectionOutcome created = await this.CreateAsync(AnnaKey, again, after);

            Assert.Equal(CollectionVerdict.Created, created.Verdict);
            Assert.NotEqual(shared.Id, created.Answer!.Id);
        }

        // ---- Shown to a reader (code review, 2026-11-27) --------------------------------------------

        private Task<ShownCollection?> ShownToAsync(string id, AccountRecord? viewer, Action? asked = null, DateTimeOffset? when = null, bool evenIfDeleted = false) =>
            CollectionFlows.FindShownToAsync(
                id,
                _ =>
                {
                    asked?.Invoke();
                    return Task.FromResult(viewer);
                },
                this.thisStore,
                when ?? Now,
                CancellationToken.None,
                evenIfDeleted);

        // Its time up, a collection is deleted from that moment - not only once the hourly sweep makes it so.
        [Fact]
        public async Task A_collection_whose_time_is_up_is_shown_as_deleted_before_the_sweep()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey, lifetime: "1m");
            DateTimeOffset after = shared.DeleteAtUtc.AddMinutes(1);

            Assert.Null(await this.ShownToAsync(shared.Id, null, when: after));

            ShownCollection? asDeleted = await this.ShownToAsync(shared.Id, null, when: after, evenIfDeleted: true);

            Assert.Equal((true, CollectionDeletedReasons.Expired), (asDeleted?.Collection.IsDeleted, asDeleted?.Collection.DeletedReason));
            Assert.NotNull(await this.ShownToAsync(shared.Id, null, when: shared.DeleteAtUtc.AddMinutes(-1)));
        }

        // A picture's every request asks this - a collection shown to all reads no session.
        [Fact]
        public async Task A_collection_shown_to_all_asks_nothing_of_the_reader()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            bool asked = false;

            ShownCollection? shown = await this.ShownToAsync(shared.Id, AMaintainer, () => asked = true);

            Assert.Equal((shared.Id, false), (shown?.Collection.Id, shown?.IsOffline));
            Assert.False(asked);
        }

        // ---- "Stop sharing" (owner decision, 2026-11-27: "Offline, same link back"; migration 0028) ---------------------------

        private Task<CollectionOutcome> SetSharedAsync(string? key, string id, bool shared, DateTimeOffset? when = null) =>
            CollectionFlows.SetSharedAsync(key, id, shared, this.thisStore, Site, when ?? Now, CancellationToken.None);

        // Stopped, its captures stay on the server and its date keeps counting - nothing of it is deleted.
        [Fact]
        public async Task Stopping_sharing_takes_the_page_offline_and_keeps_all_of_it()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);

            CollectionOutcome stopped = await this.SetSharedAsync(AnnaKey, shared.Id, false, Now.AddDays(2));

            Assert.Equal(CollectionVerdict.Done, stopped.Verdict);
            Assert.Equal(shared.Url, stopped.Answer!.Url);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal(Now.AddDays(2), kept.UnsharedUtc);
            Assert.False(kept.IsDeleted);
            Assert.Single(kept.Images);
            Assert.Equal(shared.DeleteAtUtc, kept.DeleteAtUtc);
        }

        // To everybody else only that it is not shared at the moment - to the administrator the page still, said so.
        [Fact]
        public async Task A_collection_whose_sharing_stopped_is_withheld_from_all_but_the_administrator()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            await this.SetSharedAsync(AnnaKey, shared.Id, false);

            ShownCollection? toMaintainer = await this.ShownToAsync(shared.Id, AMaintainer);
            ShownCollection? signedOut = await this.ShownToAsync(shared.Id, null);
            ShownCollection? toAdministrator = await this.ShownToAsync(shared.Id, TheAdministrator);

            Assert.Equal((true, CollectionOffline.SharingStopped), (toMaintainer!.IsWithheld, toMaintainer.Offline));
            Assert.True(signedOut!.IsWithheld);
            Assert.Equal((false, CollectionOffline.SharingStopped), (toAdministrator!.IsWithheld, toAdministrator.Offline));
        }

        // Its pictures go with the page: none for anybody but the administrator.
        [Fact]
        public async Task A_collection_whose_sharing_stopped_shows_nobody_else_its_pictures()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            string sha = this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256;
            await this.SetSharedAsync(AnnaKey, shared.Id, false);

            Task<ShownCollection?> ImageTo(AccountRecord? viewer) =>
                CollectionFlows.FindShownImageToAsync(shared.Id, sha, _ => Task.FromResult(viewer), this.thisStore, Now, CancellationToken.None);

            Assert.Null(await ImageTo(AMaintainer));
            Assert.Null(await ImageTo(null));

            // The administrator sees the page said to be offline - its pictures too, so it can check them.
            Assert.NotNull(await ImageTo(TheAdministrator));
        }

        // "Share" - the whole collection sent again - brings it back under the SAME link; and so does sharing it again alone.
        [Fact]
        public async Task Sharing_again_brings_the_page_back_under_the_same_link()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            CollectionImageEntry first = this.thisStore.Collections[shared.Id].Images[0].Entry;

            await this.SetSharedAsync(AnnaKey, shared.Id, false);

            CollectionOutcome sentAgain = await this.UpdateAsync(AnnaKey, shared.Id, Content(first), Now.AddDays(1));

            Assert.Equal((CollectionVerdict.Done, shared.Url), (sentAgain.Verdict, sentAgain.Answer!.Url));
            Assert.Null(this.thisStore.Collections[shared.Id].UnsharedUtc);

            await this.SetSharedAsync(AnnaKey, shared.Id, false);
            Assert.Equal(CollectionVerdict.Done, (await this.SetSharedAsync(AnnaKey, shared.Id, true)).Verdict);
            Assert.Null(this.thisStore.Collections[shared.Id].UnsharedUtc);
            Assert.False((await this.ShownToAsync(shared.Id, null))!.IsWithheld);
        }

        // The key's alone, as sharing it is; a deleted one is gone, its page saying so.
        [Fact]
        public async Task Only_its_key_stops_sharing_a_collection_and_a_deleted_one_is_gone()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);

            Assert.Equal(CollectionVerdict.NotFound, (await this.SetSharedAsync(BoKey, shared.Id, false)).Verdict);
            Assert.Equal(CollectionVerdict.BadRequest, (await this.SetSharedAsync(null, shared.Id, false)).Verdict);
            Assert.Null(this.thisStore.Collections[shared.Id].UnsharedUtc);

            await this.DeleteAsync(AnnaKey, shared.Id);

            Assert.Equal(CollectionVerdict.Gone, (await this.SetSharedAsync(AnnaKey, shared.Id, false)).Verdict);
        }

        // Stopped twice keeps the first stop's time - the page has been offline since then.
        [Fact]
        public async Task Stopping_sharing_again_keeps_when_it_first_stopped()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);

            await this.SetSharedAsync(AnnaKey, shared.Id, false, Now.AddDays(1));
            await this.SetSharedAsync(AnnaKey, shared.Id, false, Now.AddDays(5));

            Assert.Equal(Now.AddDays(1), this.thisStore.Collections[shared.Id].UnsharedUtc);
        }

        // ---- One capture deleted from the page at once (owner request, 2026-11-28; server 16.1.0) -------------------------------

        private Task<CollectionOutcome> RemoveImageAsync(string? key, string id, string? sha256, DateTimeOffset? when = null) =>
            CollectionFlows.RemoveImageAsync(key, id, sha256, this.thisStore, Site, when ?? Now, CancellationToken.None);

        // A collection of two captures, shared under Anna's key - and the two checksums.
        private async Task<(CollectionAnswer Shared, string First, string Second)> SharedWithTwoAsync()
        {
            byte[] first = TestPng.Rgb(32, 20, 1);
            byte[] second = TestPng.Rgb(32, 20, 2);
            await this.UploadAsync(AnnaKey, first);
            await this.UploadAsync(AnnaKey, second);

            CollectionOutcome created = await this.CreateAsync(AnnaKey, Content(Image(ShaOf(first)), Image(ShaOf(second), pin: "13")));

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

            CollectionOutcome removed = await this.RemoveImageAsync(AnnaKey, shared.Id, first, Now.AddDays(40));

            Assert.Equal(CollectionVerdict.Done, removed.Verdict);
            Assert.Equal(1, removed.Answer!.ImageCount);
            Assert.Equal(shared.DeleteAtUtc, removed.Answer.DeleteAtUtc);

            CollectionRecord kept = this.thisStore.Collections[shared.Id];

            Assert.Equal([second], kept.Images.Select(image => image.Entry.Sha256));
            Assert.Equal(before.DeleteAtUtc, kept.DeleteAtUtc);
            Assert.Equal(before.UnsharedUtc, kept.UnsharedUtc);
            Assert.Equal(before.Title, kept.Title);
            Assert.Equal(Now.AddDays(40), kept.UpdatedUtc);
        }

        // The page reads the collection as it is now - the capture is not on it any more.
        [Fact]
        public async Task A_deleted_capture_is_on_the_page_no_more()
        {
            (CollectionAnswer shared, string first, _) = await this.SharedWithTwoAsync();

            await this.RemoveImageAsync(AnnaKey, shared.Id, first);

            ShownCollection? page = await this.ShownToAsync(shared.Id, null);

            Assert.DoesNotContain(page!.Collection.Images, image => image.Entry.Sha256 == first);
        }

        // The last one may go too - the page then says it has none - and one deleted twice, or never there, is done.
        [Fact]
        public async Task The_last_capture_may_go_and_one_not_there_is_done()
        {
            CollectionAnswer shared = await this.SharedAsync(AnnaKey);
            string only = this.thisStore.Collections[shared.Id].Images[0].Entry.Sha256;

            Assert.Equal(CollectionVerdict.Done, (await this.RemoveImageAsync(AnnaKey, shared.Id, only)).Verdict);
            Assert.Equal(CollectionVerdict.Done, (await this.RemoveImageAsync(AnnaKey, shared.Id, only)).Verdict);
            Assert.Empty(this.thisStore.Collections[shared.Id].Images);
            Assert.False(this.thisStore.Collections[shared.Id].IsDeleted);
        }

        // Its key's alone - any other's "not found" (the administrator deletes a whole collection on the website); a deleted
        // collection is gone; an address naming no checksum is refused.
        [Fact]
        public async Task Only_its_key_deletes_a_capture_and_a_deleted_collection_is_gone()
        {
            (CollectionAnswer shared, string first, _) = await this.SharedWithTwoAsync();

            Assert.Equal(CollectionVerdict.NotFound, (await this.RemoveImageAsync(BoKey, shared.Id, first)).Verdict);
            Assert.Equal(CollectionVerdict.BadRequest, (await this.RemoveImageAsync(null, shared.Id, first)).Verdict);
            Assert.Equal(CollectionVerdict.BadRequest, (await this.RemoveImageAsync(AnnaKey, shared.Id, first.ToUpperInvariant())).Verdict);
            Assert.Equal(2, this.thisStore.Collections[shared.Id].Images.Count);

            await this.DeleteAsync(AnnaKey, shared.Id);

            Assert.Equal(CollectionVerdict.Gone, (await this.RemoveImageAsync(AnnaKey, shared.Id, first)).Verdict);
        }

        // ---- The oscilloscopes the captures were taken with (owner request, 2026-11-28; migration 0030) --------------------------

        [Fact]
        public async Task The_oscilloscopes_sent_are_kept_trimmed_each_once()
        {
            byte[] capture = TestPng.Rgb(32, 20, 1);
            await this.UploadAsync(AnnaKey, capture);

            CollectionOutcome created = await this.CreateAsync(
                AnnaKey,
                Content(Image(ShaOf(capture))) with { Oscilloscopes = [" Rigol DS1054Z ", "rigol ds1054z", "Siglent SDS1104X-E"] });

            Assert.Equal(CollectionVerdict.Created, created.Verdict);
            Assert.Equal(["Rigol DS1054Z", "Siglent SDS1104X-E"], this.thisStore.Collections[created.Answer!.Id].Oscilloscopes);
        }

        // ---- A capture's component and pinout drawing (owner request, 2026-11-27; migration 0029) ------------------------------

        // The drawing is a picture of the key's as a capture is: asked for when not uploaded, then kept with the capture, its
        // component's facts with the collection - and served as one of the collection's pictures.
        [Fact]
        public async Task A_capture_s_pinout_drawing_is_a_picture_of_the_collection_and_its_component_is_kept()
        {
            byte[] capture = TestPng.Rgb(32, 20, 1);
            byte[] pinout = TestPng.Rgb(40, 60, 2);
            CollectionComponent component = new(" U7 ", "U7 | VIC-II", "IC | DIP-40", null, null, [new CollectionPin("12", "/RAS", "Output", null, " Pulsing ")]);
            CollectionContentRequest content = Content(Image(ShaOf(capture)) with { PinoutSha256 = ShaOf(pinout) }) with { Components = [component] };

            await this.UploadAsync(AnnaKey, capture);

            CollectionOutcome missing = await this.CreateAsync(AnnaKey, content);

            Assert.Equal(CollectionVerdict.MissingImages, missing.Verdict);
            Assert.Equal([ShaOf(pinout)], missing.Missing);

            await this.UploadAsync(AnnaKey, pinout);

            CollectionOutcome created = await this.CreateAsync(AnnaKey, content);

            Assert.Equal(CollectionVerdict.Created, created.Verdict);

            CollectionRecord kept = this.thisStore.Collections[created.Answer!.Id];

            Assert.Equal(ShaOf(pinout), kept.Images[0].Entry.PinoutSha256);
            Assert.Equal(new CollectionComponent("U7", "U7 | VIC-II", "IC | DIP-40", null, null, null) with { Pins = null }, Assert.Single(kept.Components!) with { Pins = null });
            Assert.Equal([new CollectionPin("12", "/RAS", "Output", null, "Pulsing")], kept.Components![0].Pins);

            ShownCollection? served = await CollectionFlows.FindShownImageToAsync(
                created.Answer.Id, ShaOf(pinout), _ => Task.FromResult<AccountRecord?>(null), this.thisStore, Now, CancellationToken.None);

            Assert.NotNull(served);
        }

        // The hourly sweep lets go of a picture no collection shows - a pinout drawing a collection shows stays.
        [Fact]
        public async Task The_sweep_keeps_a_pinout_drawing_a_collection_shows()
        {
            byte[] capture = TestPng.Rgb(32, 20, 3);
            byte[] pinout = TestPng.Rgb(40, 60, 4);

            await this.UploadAsync(AnnaKey, capture);
            await this.UploadAsync(AnnaKey, pinout);
            await this.CreateAsync(AnnaKey, Content(Image(ShaOf(capture)) with { PinoutSha256 = ShaOf(pinout) }));

            await this.SweepAsync(Now.AddDays(3));

            Assert.True(File.Exists(this.thisFiles.PathOf(ShaOf(pinout))));
            Assert.Contains(this.thisStore.Uploads, upload => upload.Sha256 == ShaOf(pinout));
        }

        // ---- "Download all" ---------------------------------------------------------------------------------------------

        // Every capture in the page's order, named as the viewer's "Download" names it.
        [Fact]
        public void Download_all_names_each_capture_by_its_place_and_anchor()
        {
            CollectionRecord collection = new(
                "k7mq3x9vbn2p", CollectionFlows.OwnerHashOf(AnnaKey), "Commodore 64", "250407", null, null, null, [], Now, Now, Now.AddYears(1), null, null,
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
