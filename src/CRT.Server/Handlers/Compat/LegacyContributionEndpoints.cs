using System.Text;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Compat
{
    // ###########################################################################################
    // CRT 2.x's "SEND CONTRIBUTION", now that the old contribution upload is gone (owner decision,
    // 2026-10-04).
    //
    //   POST /api/legacy/contribution   CRT 2.x's contribution upload -> 426 "OUTDATED_VERSION 3.0.0 - ..."
    //                                                                     429 too many from this address.
    //
    // The old contribution method is not supported alongside the new pipeline (owner decision,
    // 2026-09-23), so a 2.x CRT cannot contribute any more - but it should be TOLD so, not shown
    // "HTTP 404". CRT 2.5.0 and later already turn exactly this answer into "This application
    // version [2.5.0] is too old to contribute data - please update to version [3.0.0] or newer"
    // (their ContributionPackaging.TryParseOutdatedVersionResponse, which reads an OUTDATED_VERSION
    // answer in exactly these words); older 2.x builds show the 426 and log the text. Apache forwards
    // the old address CRT 2.x posts to, https://classic-repair-toolbox.dk/app-contribution/api/,
    // here.
    //
    // *** THE TEXT IS A CONTRACT WITH BUILDS THAT CAN NEVER CHANGE. *** The token, the space and the
    // version straight after it are what 2.5.0 reads; ApiCompatibilityTests runs it through a copy of
    // that parser. Do not reword the start of it.
    //
    // The upload is READ before the answer is written: CRT 2.x reads the answer only after its whole
    // zip is sent, and Apache turns an answer that arrives mid-upload into a 502. Nothing is kept.
    // ###########################################################################################
    public static class LegacyContributionEndpoints
    {
        public const string PathUnderApi = "legacy/contribution";

        public const string RateLimitPolicy = "legacy-contribution";

        // What CRT 2.5.0 looks for, then the first CRT that contributes through the new pipeline.
        public const string OutdatedToken = "OUTDATED_VERSION";
        public const string FirstVersionWithNewPipeline = "3.0.0";

        // A 2.x contributor trying again and again is not an attack, but the body is large.
        public const int MaxPerAddressPerHour = 10;

        public static void MapLegacyContributionEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            app.MapPost("/api/" + LegacyContributionEndpoints.PathUnderApi, LegacyContributionEndpoints.ReceiveAsync)
                .WithBodyLimit(RequestBodyLimits.LegacyContributionBytes)
                .RequireRateLimiting(LegacyContributionEndpoints.RateLimitPolicy);
        }

        // The limiter the route above names - registered with the other services.
        public static void AddLegacyContributionRateLimit(this IServiceCollection services) =>
            services.AddPerAddressRateLimit(LegacyContributionEndpoints.RateLimitPolicy, LegacyContributionEndpoints.MaxPerAddressPerHour);

        // ###########################################################################################
        // The answer, in the words CRT 2.x has always been given. The version CRT 2.x sent is read
        // from its User-Agent ("CRT 2.5.0") - the form carries it too, but the answer is the same
        // either way and the form is not worth parsing for it.
        // ###########################################################################################
        public static string OutdatedAnswer(string? userAgent)
        {
            string version = CrtVersion.FromUserAgent(userAgent)?.ToString() ?? "unknown";
            string first = LegacyContributionEndpoints.FirstVersionWithNewPipeline;

            return $"{LegacyContributionEndpoints.OutdatedToken} {first} - this application version [{version}] " +
                   $"is too old to contribute data - please update to version [{first}] or newer.";
        }

        // ###########################################################################################
        // *** AN UPLOAD OVER THE LIMIT IS STILL TOLD TO UPDATE (code review, 2026-10-04). *** Kestrel
        // stops a body past RequestBodyLimits.LegacyContributionBytes by throwing from the read; left
        // to escape, that was a bare 413 and CRT 2.x never saw the words this route exists to send.
        // The answer is attempted anyway, as ClientVersionGate does - a client that went away simply
        // does not read it.
        // ###########################################################################################
        internal static async Task<IResult> ReceiveAsync(HttpContext context, CancellationToken cancellationToken)
        {
            try
            {
                await context.Request.Body.CopyToAsync(Stream.Null, cancellationToken);
            }
            catch (Exception ex) when (ex is BadHttpRequestException or IOException)
            {
                // Too large, or cut off - the answer is still worth trying.
            }

            return Results.Text(
                LegacyContributionEndpoints.OutdatedAnswer(context.Request.Headers.UserAgent.ToString()),
                "text/plain",
                Encoding.UTF8,
                ClientVersionContract.OutdatedStatus);
        }
    }
}
