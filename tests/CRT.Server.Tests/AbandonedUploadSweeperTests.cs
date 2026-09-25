using System.Reflection;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The abandoned-upload sweep actually RUNNING (code review, 2026-09-25).
    //
    // SubmissionFlows.CollectAbandonedAsync was written and tested (SubmissionFlowTests), and
    // SubmissionEndpoints relies on it to bound the disk an anonymous caller can fill - but no
    // hosted service, timer or endpoint ever called it. These tests pin the three things that were
    // missing: that one pass collects, that a failing pass does not throw (a BackgroundService that
    // throws stops the whole host), and that the sweeper is registered as a hosted service at all.
    //
    // No database and no network: the in-memory store, a temp blob folder, and a hosted service
    // started and stopped in-process - no HTTP pipeline (see the test csproj's header).
    // ###########################################################################################
    public sealed class AbandonedUploadSweeperTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-sweeper-" + Guid.NewGuid().ToString("N"));

        public AbandonedUploadSweeperTests() => Directory.CreateDirectory(this.thisRoot);

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A temp folder left behind is harmless.
            }
        }

        private BlobStore Blobs() => new(this.thisRoot, NullLogger<BlobStore>.Instance);

        // A submission still in its upload state whose window has already closed.
        private static FakeSubmissionStore StoreWithAnExpiredUpload()
        {
            var store = new FakeSubmissionStore();

            store.Submissions[7] = new SubmissionRecord(
                7,
                "Commodore/C64/250407/Data.xlsx",
                AccountId: null,
                ContactEmail: "someone@example.com",
                UploadTokenHash: "hash",
                BaseRevision: "2026-09-01",
                SubmissionState.Uploading,
                Summary: null,
                FormatVersion: 1,
                CreatedUtc: AbandonedUploadSweeperTests.Now - TimeSpan.FromDays(2),
                ExpiresUtc: AbandonedUploadSweeperTests.Now - TimeSpan.FromDays(1),
                DecidedUtc: null);

            return store;
        }

        [Fact]
        public async Task One_pass_collects_an_upload_whose_window_has_closed()
        {
            FakeSubmissionStore store = AbandonedUploadSweeperTests.StoreWithAnExpiredUpload();

            int collected = await AbandonedUploadSweeper.SweepOnceAsync(
                store, this.Blobs(), AbandonedUploadSweeperTests.Now, NullLogger.Instance, CancellationToken.None);

            Assert.Equal(1, collected);
            Assert.Equal(SubmissionState.Abandoned, store.Submissions[7].State);
        }

        // ###########################################################################################
        // *** THE PASS ALSO RUNS THE TWO COLLECTORS NOTHING USED TO RUN (security review,
        // 2026-09-25). *** A collector that exists but is never called is the exact defect this
        // class's own header records for the abandoned-upload sweep; this pins that the new ones
        // are called too. A blob no submission references at all must be gone after one pass.
        // ###########################################################################################
        [Fact]
        public async Task One_pass_also_deletes_a_blob_no_live_submission_needs()
        {
            BlobStore blobs = this.Blobs();
            byte[] bytes = [1, 2, 3, 4];
            string hash = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));

            await blobs.AppendChunkAsync(99, hash, 0, new MemoryStream(bytes));
            Assert.True(await blobs.TryCompleteAsync(99, hash));

            await AbandonedUploadSweeper.SweepOnceAsync(
                new FakeSubmissionStore(), blobs, AbandonedUploadSweeperTests.Now, NullLogger.Instance, CancellationToken.None);

            Assert.False(blobs.Contains(hash));
        }

        // ###########################################################################################
        // *** A FAILING PASS MUST NOT THROW. *** From .NET 8 an exception escaping a
        // BackgroundService stops the host, so one database error during a sweep would take the
        // whole contribution service down. The pass logs and answers 0; the next one retries.
        // ###########################################################################################
        [Fact]
        public async Task A_pass_whose_store_fails_answers_zero_instead_of_throwing()
        {
            ISubmissionStore failing = ThrowingStore.Create();

            int collected = await AbandonedUploadSweeper.SweepOnceAsync(
                failing, this.Blobs(), AbandonedUploadSweeperTests.Now, NullLogger.Instance, CancellationToken.None);

            Assert.Equal(0, collected);
        }

        [Fact]
        public async Task The_started_service_sweeps_straight_away()
        {
            // The first pass runs at start-up, so a restart after an outage collects at once
            // rather than an interval later.
            FakeSubmissionStore store = AbandonedUploadSweeperTests.StoreWithAnExpiredUpload();

            using var sweeper = new AbandonedUploadSweeper(
                store,
                this.Blobs(),
                NullLogger<AbandonedUploadSweeper>.Instance,
                TimeProvider.System,
                TimeSpan.FromHours(1));

            await sweeper.StartAsync(CancellationToken.None);

            try
            {
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);

                while (store.Submissions[7].State != SubmissionState.Abandoned && DateTime.UtcNow < deadline)
                {
                    await Task.Delay(20);
                }

                Assert.Equal(SubmissionState.Abandoned, store.Submissions[7].State);
            }
            finally
            {
                await sweeper.StopAsync(CancellationToken.None);
            }
        }

        [Fact]
        public void The_sweeper_is_registered_as_a_HOSTED_service()
        {
            // The defect was never the sweep itself - it was that nothing ran it. A hosted service
            // is what the host starts with app.Run().
            var services = new ServiceCollection();

            services.AddAbandonedUploadSweeper();

            Assert.Contains(
                services,
                descriptor => descriptor.ServiceType == typeof(IHostedService)
                    && descriptor.ImplementationType == typeof(AbandonedUploadSweeper));
        }

        // An ISubmissionStore whose every call throws - built with DispatchProxy so the test does
        // not have to spell out the whole interface.
        public class ThrowingStore : DispatchProxy
        {
            public static ISubmissionStore Create() => DispatchProxy.Create<ISubmissionStore, ThrowingStore>();

            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args) =>
                throw new InvalidOperationException("The database is unavailable.");
        }
    }
}
