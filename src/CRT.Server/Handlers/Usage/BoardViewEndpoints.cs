using System.Threading.RateLimiting;
using CRT.Server.Configuration;
using Handlers.DataHandling;
using Microsoft.AspNetCore.RateLimiting;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // BOARD VIEWS, over HTTP (owner request, 2026-09-27). A rim over BoardViewFlows and nothing else.
    //
    //   POST /api/usage/board-views   BoardViewReport -> 202 stored (or a batch stored before),
    //                                                    400 a report CRT should drop,
    //                                                    429 too many from this address.
    //
    // ANONYMOUS, like the launch check-in: CRT sends it without an account. The body is at most
    // BoardViewRules.MaxViewsPerReport short views - well inside the 64 KB default limit.
    //
    // *** RATE LIMITED BY ADDRESS IN MEMORY ONLY. *** The submission limits count what the database
    // holds per address; board views must hold no address at all, so this route uses ASP.NET's own
    // limiter, whose partitions live in memory and are dropped when idle. A refused report is kept
    // by CRT and sent again later (BoardViewContract.DeliveryFor).
    // ###########################################################################################
    public static class BoardViewEndpoints
    {
        public const string RateLimitPolicy = "board-views";

        public static void MapBoardViewEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            app.MapPost("/api/" + BoardViewContract.PathUnderApi, BoardViewEndpoints.RecordAsync)
                .RequireRateLimiting(BoardViewEndpoints.RateLimitPolicy);
        }

        // The limiter the route above names - registered with the other services.
        public static void AddBoardViewRateLimit(this IServiceCollection services)
        {
            services.AddRateLimiter(limiter =>
            {
                limiter.RejectionStatusCode = StatusCodes.Status429TooManyRequests;

                limiter.AddPolicy(BoardViewEndpoints.RateLimitPolicy, context =>
                    RateLimitPartition.GetFixedWindowLimiter(
                        context.Connection.RemoteIpAddress?.ToString() ?? "unknown",
                        _ => new FixedWindowRateLimiterOptions
                        {
                            PermitLimit = BoardViewRules.MaxReportsPerAddressPerHour,
                            Window = TimeSpan.FromHours(1),
                            QueueLimit = 0
                        }));
            });
        }

        private static async Task<IResult> RecordAsync(
            BoardViewReport? report,
            HttpContext context,
            BoardViewNameDirectory names,
            ICountryLookup countries,
            IBoardViewStore store,
            ServerOptions options,
            ILoggerFactory loggers,
            CancellationToken cancellationToken)
        {
            BoardViewOutcome outcome = await BoardViewFlows.RecordAsync(
                report,
                context.Connection.RemoteIpAddress,
                names.Find,
                countries,
                store,
                options.CountLocalNetworkBoardViews,
                DateTimeOffset.UtcNow,
                cancellationToken);

            // Said, so the project owner can see a test from home arrive - never with the address.
            if (outcome.IsFromLocalNetwork && !outcome.IsRefused)
            {
                ILogger logger = loggers.CreateLogger("CRT.Server.BoardViews");

                if (outcome.IsRepeat)
                    logger.LogInformation("A batch of board views from a local-network address arrived again and was stored before.");
                else
                    logger.LogInformation("A batch of board views from a local-network address arrived: {Stored} stored, {Ignored} not stored.", outcome.StoredViews, outcome.IgnoredViews);
            }

            return outcome.IsRefused
                ? Results.BadRequest(new { error = outcome.Reason })
                : Results.StatusCode(StatusCodes.Status202Accepted);
        }
    }
}
