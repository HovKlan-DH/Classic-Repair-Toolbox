namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Runs SubmissionFlows.CollectAbandonedAsync on a timer, for as long as the service runs
    // (code review, 2026-09-25).
    //
    // *** THE SWEEP EXISTED, BUT NOTHING EVER RAN IT. *** CollectAbandonedAsync was written and
    // tested, and SubmissionEndpoints' own header leans on it: an anonymous POST needs no account
    // and no rate limit, so "the 24-hour abandoned-upload sweep is what actually bounds the
    // damage". No hosted service, timer or endpoint called it. Every submission that was created
    // and never finalised - and every partial upload refused for being over the size cap - stayed
    // on disk and in the database for ever.
    //
    // THE RULE STAYS IN SubmissionFlows; this class only decides WHEN. That keeps the rim/flow
    // split the rest of the service follows: the timing is the only thing here, and the one pass
    // it runs (SweepOnceAsync) is tested against the in-memory store.
    //
    // *** A FAILED PASS IS LOGGED AND SWALLOWED. *** From .NET 8 an exception escaping a
    // BackgroundService STOPS THE HOST by default, so one database hiccup at 03:00 would take the
    // whole contribution service down. A missed sweep costs nothing - the next one collects the
    // same rows - so the pass reports and waits for the next tick instead.
    // ###########################################################################################
    public sealed class AbandonedUploadSweeper : BackgroundService
    {
        // Hourly. Uploads expire after SubmissionFlows.UploadWindow (24 hours), so this collects
        // each one within an hour of its window closing - far below anything that could fill a
        // disk - while costing one indexed query an hour.
        public static readonly TimeSpan Interval = TimeSpan.FromHours(1);

        private readonly ISubmissionStore thisStore;
        private readonly BlobStore thisBlobs;
        private readonly ILogger<AbandonedUploadSweeper> thisLogger;
        private readonly TimeProvider thisTime;
        private readonly TimeSpan thisInterval;

        public AbandonedUploadSweeper(
            ISubmissionStore store,
            BlobStore blobs,
            ILogger<AbandonedUploadSweeper> logger)
            : this(store, blobs, logger, TimeProvider.System, AbandonedUploadSweeper.Interval)
        {
        }

        // The seam: tests supply the clock and the interval. Internal, so dependency injection -
        // which only considers public constructors - always takes the one above.
        internal AbandonedUploadSweeper(
            ISubmissionStore store,
            BlobStore blobs,
            ILogger<AbandonedUploadSweeper> logger,
            TimeProvider time,
            TimeSpan interval)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);
            ArgumentNullException.ThrowIfNull(logger);
            ArgumentNullException.ThrowIfNull(time);

            this.thisStore = store;
            this.thisBlobs = blobs;
            this.thisLogger = logger;
            this.thisTime = time;
            this.thisInterval = interval;
        }

        // ###########################################################################################
        // One pass at start-up - so a restart after a long outage collects at once rather than an
        // hour later - then one per interval until the service stops.
        // ###########################################################################################
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            using var timer = new PeriodicTimer(this.thisInterval, this.thisTime);

            try
            {
                do
                {
                    await AbandonedUploadSweeper.SweepOnceAsync(
                        this.thisStore,
                        this.thisBlobs,
                        this.thisTime.GetUtcNow(),
                        this.thisLogger,
                        stoppingToken);
                }
                while (await timer.WaitForNextTickAsync(stoppingToken));
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                // The service is stopping. Nothing to report.
            }
        }

        // ###########################################################################################
        // One sweep. Returns how many submissions were collected; 0 when the pass failed, after
        // logging why - see the class header for why it must never throw.
        // ###########################################################################################
        internal static async Task<int> SweepOnceAsync(
            ISubmissionStore store,
            BlobStore blobs,
            DateTimeOffset now,
            ILogger logger,
            CancellationToken cancellationToken)
        {
            try
            {
                int collected = await SubmissionFlows.CollectAbandonedAsync(store, blobs, now, cancellationToken);

                if (collected > 0)
                {
                    logger.LogInformation(
                        "Collected {Count} abandoned submission upload(s).", collected);
                }

                return collected;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception ex)
            {
                logger.LogWarning(
                    ex,
                    "The abandoned-upload sweep failed; it will run again in {Interval}.",
                    AbandonedUploadSweeper.Interval);

                return 0;
            }
        }
    }

    // ###########################################################################################
    // The registration Program.cs calls, kept beside the class so a test can prove the sweeper
    // is registered as a HOSTED service - the one thing that makes it run at all.
    // ###########################################################################################
    public static class AbandonedUploadSweeperRegistration
    {
        public static IServiceCollection AddAbandonedUploadSweeper(this IServiceCollection services)
        {
            ArgumentNullException.ThrowIfNull(services);

            return services.AddHostedService<AbandonedUploadSweeper>();
        }
    }
}
