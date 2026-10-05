using System.Text;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // THE LAUNCH CHECK-IN, over HTTP (owner request, 2026-10-03). A rim over CheckInFormReader and
    // CheckInFlow and nothing else.
    //
    //   POST /api/usage/check-in   CRT's launch check-in -> 200 "Thanks for checking :-)" (stored,
    //                                                       or from a local network: not stored),
    //                                                       400 not a CRT check-in,
    //                                                       415 not a form,
    //                                                       429 too many from this address.
    //
    // *** OLDER CRTs REACH IT TOO. *** They post the same form to the old check-in address,
    // /app-checkin/, and Apache forwards that here - which is why the answer is the plain text
    // check-ins have always been given.
    //
    // ANONYMOUS, as check-ins have always been, like board views: CRT checks in without an account.
    // Limited per address in memory (there used to be no limit at all); the body to the default
    // 64 KB, stated on the route so RequestBodyLimitsTests sees the decision.
    // ###########################################################################################
    public static class CheckInEndpoints
    {
        public const string RateLimitPolicy = "check-in";

        public static void MapCheckInEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            app.MapPost("/api/" + CheckInContract.PathUnderApi, CheckInEndpoints.RecordAsync)
                .WithBodyLimit(RequestBodyLimits.DefaultBytes)
                .RequireRateLimiting(CheckInEndpoints.RateLimitPolicy);
        }

        // The limiter the route above names - registered with the other services.
        public static void AddCheckInRateLimit(this IServiceCollection services) =>
            services.AddPerAddressRateLimit(CheckInEndpoints.RateLimitPolicy, CheckInFlow.MaxPerAddressPerHour);

        private static async Task<IResult> RecordAsync(
            HttpContext context,
            ICountryLookup countries,
            ICheckInStore store,
            ILoggerFactory loggers,
            CancellationToken cancellationToken)
        {
            CheckInFormRead read = await CheckInFormReader.ReadAsync(context.Request, cancellationToken);

            if (read.IsRefused)
                return Results.Text(read.RefusalReason, "text/plain", Encoding.UTF8, read.RefusalStatus);

            CheckInOutcome outcome = await CheckInFlow.RecordAsync(
                read.Request!,
                context.Connection.RemoteIpAddress,
                countries,
                store,
                cancellationToken);

            if (outcome.IsRefused)
                return Results.Text(outcome.Reason, "text/plain", Encoding.UTF8, StatusCodes.Status400BadRequest);

            // Said, so the project owner can see a test from home arrive - never with the address.
            if (outcome.IsFromLocalNetwork)
            {
                loggers.CreateLogger("CRT.Server.CheckIn").LogInformation(
                    "A check-in from a local-network address arrived - answered, and not stored (a check-in from the server's own network never is).");
            }

            return Results.Text(CheckInContract.ThanksAnswer, "text/plain", Encoding.UTF8, StatusCodes.Status200OK);
        }
    }
}
