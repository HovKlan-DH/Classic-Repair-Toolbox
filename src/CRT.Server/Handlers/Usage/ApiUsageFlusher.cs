namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // Writes the API usage counted in memory (ApiUsageCounter) into crt_api_calls every few minutes,
    // and once more as the service stops - so a restart for a deploy loses nothing.
    //
    // *** A FAILED WRITE IS PUT BACK, LOGGED AND SWALLOWED. *** The tallies go back into the counter
    // for the next write (ApiUsageCounter.PutBack), and nothing escapes: from .NET 8 an exception
    // escaping a BackgroundService stops the host, and statistics must never take the contribution
    // service down - AbandonedUploadSweeper's rule.
    //
    // THE RULE IS FlushAsync; this class only decides WHEN, like the sweeper.
    // ###########################################################################################
    public sealed class ApiUsageFlusher : BackgroundService
    {
        // Every five minutes: at most a few hundred rows a write, and a crash loses five minutes of
        // counting at worst.
        public static readonly TimeSpan Interval = TimeSpan.FromMinutes(5);

        private readonly ApiUsageCounter thisCounter;
        private readonly IApiUsageStore thisStore;
        private readonly ILogger<ApiUsageFlusher> thisLogger;

        public ApiUsageFlusher(ApiUsageCounter counter, IApiUsageStore store, ILogger<ApiUsageFlusher> logger)
        {
            ArgumentNullException.ThrowIfNull(counter);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(logger);

            this.thisCounter = counter;
            this.thisStore = store;
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // One write: everything counted so far, or - when the store refuses - all of it back in the
        // counter. True when it was written (or there was nothing to write).
        //
        // The whole write holds the counter's writes (ApiUsageCounter.HoldWritesAsync), the put-back
        // included, so a reset of the contribution data never lands between taking the tallies and
        // writing them. Given up before its turn, it writes nothing and the tallies stay counted.
        // ###########################################################################################
        public static async Task<bool> FlushAsync(
            ApiUsageCounter counter,
            IApiUsageStore store,
            ILogger logger,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(counter);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(logger);

            IDisposable hold;

            try
            {
                hold = await counter.HoldWritesAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                return false;
            }

            using (hold)
            {
                IReadOnlyList<ApiUsageTally> tallies = counter.TakeAll();

                if (tallies.Count == 0)
                    return true;

                try
                {
                    await store.AddAsync(tallies, cancellationToken);
                    return true;
                }
                catch (Exception ex)
                {
                    counter.PutBack(tallies);

                    if (ex is not OperationCanceledException)
                        logger.LogWarning(ex, "The API usage counts could not be written; they are kept for the next try.");

                    return false;
                }
            }
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            using var timer = new PeriodicTimer(ApiUsageFlusher.Interval);

            try
            {
                while (await timer.WaitForNextTickAsync(stoppingToken))
                    await ApiUsageFlusher.FlushAsync(this.thisCounter, this.thisStore, this.thisLogger, stoppingToken);
            }
            catch (OperationCanceledException)
            {
                // Stopping - StopAsync writes what is left.
            }
        }

        // The last write, as the service stops: with a token of its own, since the host's is about to
        // be cancelled - and bounded, so a database that is away cannot hold up the shutdown.
        public override async Task StopAsync(CancellationToken cancellationToken)
        {
            await base.StopAsync(cancellationToken);

            using var last = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            await ApiUsageFlusher.FlushAsync(this.thisCounter, this.thisStore, this.thisLogger, last.Token);
        }
    }
}
