using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the BlobStore behaviour added by the security review (2026-09-25). The resumable
    // upload and hash-on-arrival behaviour it already had is covered through SubmissionFlowTests.
    //
    // A real temp folder, for the reason SubmissionFlowTests gives: the properties worth testing ARE
    // the filesystem behaviour - that a copy which does not verify never replaces the real file,
    // that a stray file is not mistaken for a blob.
    // ###########################################################################################
    public sealed class BlobStoreTests : IDisposable
    {
        private readonly string thisRoot;

        public BlobStoreTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-blobstore-tests", Guid.NewGuid().ToString("N"));
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

        private BlobStore Store() => new(Path.Combine(this.thisRoot, "store"), NullLogger<BlobStore>.Instance);

        private static string HashOf(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

        private static async Task<string> PutAsync(BlobStore store, byte[] bytes)
        {
            string hash = BlobStoreTests.HashOf(bytes);

            using var content = new MemoryStream(bytes);
            await store.AppendChunkAsync(1, hash, 0, content);
            Assert.True(await store.TryCompleteAsync(1, hash));

            return hash;
        }

        private string StoredPathOf(string hash) =>
            Path.Combine(this.thisRoot, "store", BlobStorePaths.BlobFolderName, hash[..2], hash[2..4], hash);

        // -----------------------------------------------------------------------------------
        // One request at a time per partial upload.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE RACE THIS CLOSES. *** Completion hashed the partial file, closed it, and then
        // moved it - so an append landing in between entered the store under a hash its bytes no
        // longer matched, and every later submission citing that hash published the poisoned copy.
        // Append and completion now share a lock; holding it must hold both of them off.
        // ###########################################################################################
        [Fact]
        public async Task An_append_waits_while_the_upload_is_locked()
        {
            BlobStore store = this.Store();
            string hash = BlobStoreTests.HashOf([1, 2, 3]);

            SemaphoreSlim uploadLock = store.LockFor(5, hash);
            await uploadLock.WaitAsync();

            Task<BlobChunkResult> append;

            try
            {
                append = store.AppendChunkAsync(5, hash, 0, new MemoryStream([1, 2, 3]));

                await Task.Delay(200);
                Assert.False(append.IsCompleted, "an append must wait for the upload's lock");
            }
            finally
            {
                uploadLock.Release();
            }

            Assert.True((await append).IsAccepted);
        }

        [Fact]
        public async Task A_completion_waits_while_the_upload_is_locked()
        {
            BlobStore store = this.Store();
            byte[] bytes = [4, 5, 6];
            string hash = BlobStoreTests.HashOf(bytes);

            await store.AppendChunkAsync(5, hash, 0, new MemoryStream(bytes));

            SemaphoreSlim uploadLock = store.LockFor(5, hash);
            await uploadLock.WaitAsync();

            Task<bool> complete;

            try
            {
                complete = store.TryCompleteAsync(5, hash);

                await Task.Delay(200);
                Assert.False(complete.IsCompleted, "a completion must wait for the upload's lock");
            }
            finally
            {
                uploadLock.Release();
            }

            Assert.True(await complete);
        }

        // Two appends at the SAME offset used to both pass the offset check and both write. Now the
        // second sees the first's bytes and is told where to resume.
        [Fact]
        public async Task A_second_append_at_the_same_offset_is_told_where_to_resume()
        {
            BlobStore store = this.Store();
            string hash = BlobStoreTests.HashOf([9, 9, 9, 9]);

            BlobChunkResult first = await store.AppendChunkAsync(5, hash, 0, new MemoryStream([9, 9]));
            BlobChunkResult second = await store.AppendChunkAsync(5, hash, 0, new MemoryStream([9, 9]));

            Assert.True(first.IsAccepted);
            Assert.False(second.IsAccepted);
            Assert.Equal(2, second.ResumeFrom);
        }

        // -----------------------------------------------------------------------------------
        // Verified copies out of the store.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_copy_lands_with_the_blobs_exact_bytes_and_leaves_no_temporary_file()
        {
            BlobStore store = this.Store();
            byte[] bytes = Encoding.UTF8.GetBytes("board scan");
            string hash = await BlobStoreTests.PutAsync(store, bytes);

            string destination = Path.Combine(this.thisRoot, "tree", "Commodore", "C64", "250407", "a.png");

            Assert.Equal(BlobCopyResult.Copied, await store.TryCopyToAsync(hash, destination));
            Assert.Equal(bytes, await File.ReadAllBytesAsync(destination));
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(destination)!));
        }

        // ###########################################################################################
        // *** THE COPY IS VERIFIED, AND A BAD ONE NEVER REPLACES THE REAL FILE. *** A blob that
        // changed on disk after it was accepted - by the race above, a bad sector, a hand edit -
        // used to be copied into the published tree on trust.
        // ###########################################################################################
        [Fact]
        public async Task A_blob_that_no_longer_matches_its_hash_is_NOT_copied_and_the_old_file_survives()
        {
            BlobStore store = this.Store();
            string hash = await BlobStoreTests.PutAsync(store, Encoding.UTF8.GetBytes("genuine"));

            await File.WriteAllTextAsync(this.StoredPathOf(hash), "tampered");

            string destination = Path.Combine(this.thisRoot, "tree", "a.png");
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            await File.WriteAllTextAsync(destination, "the published file");

            Assert.Equal(BlobCopyResult.HashMismatch, await store.TryCopyToAsync(hash, destination));
            Assert.Equal("the published file", await File.ReadAllTextAsync(destination));
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(destination)!));
        }

        [Fact]
        public async Task A_missing_blob_copies_nothing()
        {
            BlobStore store = this.Store();

            Assert.Equal(
                BlobCopyResult.Missing,
                await store.TryCopyToAsync(new string('a', 64), Path.Combine(this.thisRoot, "tree", "a.png")));
        }

        // -----------------------------------------------------------------------------------
        // Importing a file the published tree already holds (2026-09-25).
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** WHAT MAKES A ONE-COMPONENT TEXT FIX UPLOAD NOTHING. *** The published tree was never
        // in the blob store, so the first submission to any board uploaded every file of it
        // (1,212 files, 121 MB, reported). An import takes the server's own copy instead - and it
        // enters the store on the same terms as an upload: hashed as it is written, moved in only
        // on a match.
        // ###########################################################################################
        [Fact]
        public async Task An_import_takes_a_published_file_into_the_store_and_leaves_no_temporary_file()
        {
            BlobStore store = this.Store();
            byte[] bytes = Encoding.UTF8.GetBytes("published board scan");
            string hash = BlobStoreTests.HashOf(bytes);

            string published = Path.Combine(this.thisRoot, "tree", "Commodore", "C64", "250407", "a.png");
            Directory.CreateDirectory(Path.GetDirectoryName(published)!);
            await File.WriteAllBytesAsync(published, bytes);

            Assert.True(await store.TryImportAsync(published, hash));

            Assert.True(store.Contains(hash));
            Assert.Equal(bytes, await File.ReadAllBytesAsync(this.StoredPathOf(hash)));
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(this.StoredPathOf(hash))!));

            // The published file is COPIED, never moved: it is still what every client syncs.
            Assert.Equal(bytes, await File.ReadAllBytesAsync(published));
        }

        // ###########################################################################################
        // *** THE STORE'S ONE PROMISE HOLDS FOR AN IMPORT TOO. *** The caller decided to import
        // because a cached hash said the file was unchanged; a cache can be stale. A blob filed
        // under a hash its bytes do not match would be "already held" for every later submission
        // citing it, and would fail their finalise with no way to resubmit past it.
        // ###########################################################################################
        [Fact]
        public async Task An_import_whose_bytes_do_not_match_the_hash_stores_nothing()
        {
            BlobStore store = this.Store();
            string claimed = BlobStoreTests.HashOf(Encoding.UTF8.GetBytes("what the cache remembered"));

            string published = Path.Combine(this.thisRoot, "tree", "a.png");
            Directory.CreateDirectory(Path.GetDirectoryName(published)!);
            await File.WriteAllTextAsync(published, "what is on disk now");

            Assert.False(await store.TryImportAsync(published, claimed));

            Assert.False(store.Contains(claimed));

            // No temporary file left behind beside where the blob would have gone.
            string blobFolder = Path.GetDirectoryName(this.StoredPathOf(claimed))!;
            Assert.True(!Directory.Exists(blobFolder) || Directory.GetFiles(blobFolder).Length == 0);
        }

        // Identical bytes by definition, so the copy already in the store is kept and the import
        // counts as done - the same rule TryCompleteAsync applies to two uploads of one file.
        [Fact]
        public async Task An_import_of_a_blob_already_held_succeeds_and_keeps_the_stored_copy()
        {
            BlobStore store = this.Store();
            byte[] bytes = Encoding.UTF8.GetBytes("shared image");
            string hash = await BlobStoreTests.PutAsync(store, bytes);

            string published = Path.Combine(this.thisRoot, "tree", "a.png");
            Directory.CreateDirectory(Path.GetDirectoryName(published)!);
            await File.WriteAllBytesAsync(published, bytes);

            Assert.True(await store.TryImportAsync(published, hash));
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(this.StoredPathOf(hash))!));
        }

        [Fact]
        public async Task An_import_of_a_file_that_is_not_there_stores_nothing()
        {
            BlobStore store = this.Store();
            string hash = new('b', 64);

            Assert.False(await store.TryImportAsync(Path.Combine(this.thisRoot, "tree", "gone.png"), hash));
            Assert.False(store.Contains(hash));
        }

        [Fact]
        public async Task Verify_hands_back_the_opening_bytes_of_an_intact_blob_and_flags_a_changed_one()
        {
            BlobStore store = this.Store();
            byte[] bytes = Encoding.UTF8.GetBytes("0123456789");
            string hash = await BlobStoreTests.PutAsync(store, bytes);

            BlobVerification intact = await store.VerifyAsync(hash, headBytes: 4);

            Assert.True(intact.IsVerified);
            Assert.Equal(Encoding.UTF8.GetBytes("0123"), intact.Head);

            await File.WriteAllTextAsync(this.StoredPathOf(hash), "changed!!!");

            BlobVerification changed = await store.VerifyAsync(hash, headBytes: 4);

            Assert.True(changed.Exists);
            Assert.False(changed.HashMatches);
            Assert.False((await store.VerifyAsync(new string('b', 64), 4)).Exists);
        }

        // -----------------------------------------------------------------------------------
        // The disk reserve, and enumeration for the collector.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_reserve_is_honoured_and_an_unmeasurable_disk_counts_as_room()
        {
            var tight = new BlobStore(this.thisRoot, NullLogger<BlobStore>.Instance, () => 1_000, minimumFreeBytes: 600);
            var unknown = new BlobStore(this.thisRoot, NullLogger<BlobStore>.Instance, () => null, minimumFreeBytes: 600);

            Assert.True(tight.HasRoomFor(400));
            Assert.False(tight.HasRoomFor(401));

            // A transient error reading the disk must not refuse every upload.
            Assert.True(unknown.HasRoomFor(long.MaxValue));
        }

        // The collector deletes what this lists, so it must list only what the store itself put
        // there - a stray file, or one filed in the wrong bucket, is not a blob to delete.
        [Fact]
        public async Task Only_real_blobs_are_listed_for_collection()
        {
            BlobStore store = this.Store();
            string hash = await BlobStoreTests.PutAsync(store, [7, 7, 7]);

            string blobs = Path.Combine(this.thisRoot, "store", BlobStorePaths.BlobFolderName);
            await File.WriteAllTextAsync(Path.Combine(blobs, "notes.txt"), "not a blob");

            string misfiled = new('f', 64);
            Directory.CreateDirectory(Path.Combine(blobs, "00", "00"));
            await File.WriteAllTextAsync(Path.Combine(blobs, "00", "00", misfiled), "wrong bucket");

            Assert.Equal([hash], store.EnumerateBlobHashes().ToList());
        }

        [Fact]
        public async Task Deleting_a_blob_removes_it_and_a_second_delete_reports_nothing_done()
        {
            BlobStore store = this.Store();
            string hash = await BlobStoreTests.PutAsync(store, [8, 8]);

            Assert.True(store.TryDeleteBlob(hash));
            Assert.False(store.Contains(hash));
            Assert.False(store.TryDeleteBlob(hash));
        }
    }
}
