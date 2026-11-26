using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // THE HOURLY SWEEP OF SHARED COLLECTIONS (2026-11-26): CollectionFlows.SweepAsync, run once at start and then every hour -
    // collections whose time is up deleted, their owners reminded 30 days before, pictures nobody shows any more taken away.
    // Nothing while sharing is not set up. A failed sweep is logged and tried again an hour later; it never stops the service.
    // ###########################################################################################
    public sealed class CollectionSweeper(
        ICollectionStore store,
        IAccountStore accounts,
        IEmailSender mailer,
        ServerOptions options,
        ILogger<CollectionSweeper> logger) : BackgroundService
    {
        public static readonly TimeSpan Interval = TimeSpan.FromHours(1);

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            if (!options.IsCollectionSharingConfigured)
                return;

            using var timer = new PeriodicTimer(CollectionSweeper.Interval);

            try
            {
                do
                {
                    try
                    {
                        CollectionSweepResult result = await CollectionFlows.SweepAsync(
                            store,
                            new CollectionFiles(options.CollectionsRoot!),
                            accounts,
                            mailer,
                            AccountPages.BaseUrl(options.PublicApiBaseUrl),
                            DateTimeOffset.UtcNow,
                            stoppingToken);

                        if (result.Expired + result.Reminded + result.FilesDeleted > 0)
                        {
                            logger.LogInformation(
                                "Collections: {Expired} deleted as their time was up, {Reminded} owner(s) reminded, {Files} capture(s) nobody shows deleted.",
                                result.Expired,
                                result.Reminded,
                                result.FilesDeleted);
                        }
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        logger.LogWarning(ex, "The collections' sweep failed; it runs again in {Interval}.", CollectionSweeper.Interval);
                    }
                }
                while (await timer.WaitForNextTickAsync(stoppingToken));
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
            }
        }
    }
}
