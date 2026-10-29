using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // POST /api/submissions/oscilloscope-models - an oscilloscope model sent in from CRT's
    // Oscilloscope tab (CRT 4.0 step 5). A rim over OscilloscopeModelFlows.SubmitAsync, under the
    // contributor routes' per-address request limit. No account needed; an Authorization header is
    // honoured, so a signed-in sender is told at the account's address. The body is one model of a
    // few dozen short strings - at the default body limit.
    //
    //   201 Created  OscilloscopeModelSubmissionAnswer
    //   400          { errors: [...] } - the problems with the model, the shape CRT reads a
    //                board submission's refusal in
    //   429          { errors: [...] } - this address sent its day's models
    // ###########################################################################################
    public static class OscilloscopeModelEndpoints
    {
        public const string Route = "/api/submissions/oscilloscope-models";

        public static void MapOscilloscopeModelEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            app.MapPost(OscilloscopeModelEndpoints.Route, OscilloscopeModelEndpoints.SubmitAsync)
                .RequireRateLimiting(SubmissionEndpoints.RateLimitPolicy);
        }

        private static async Task<IResult> SubmitAsync(
            OscilloscopeModelSubmissionRequest request,
            HttpContext context,
            IAccountStore accounts,
            IOscilloscopeModelStore store,
            IEmailSender mailer,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            AccountRecord? account = await SubmissionEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            OscilloscopeModelSubmitOutcome outcome = await OscilloscopeModelFlows.SubmitAsync(
                request,
                account,
                SubmissionEndpoints.ClientAddress(context),
                context.Request.Headers.UserAgent.ToString(),
                store,
                accounts,
                mailer,
                DateTimeOffset.UtcNow,
                loggerFactory.CreateLogger(typeof(OscilloscopeModelFlows).FullName!),
                cancellationToken);

            if (outcome.Answer is not null)
                return Results.Created($"{OscilloscopeModelEndpoints.Route}/{outcome.Answer.Id}", outcome.Answer);

            return Results.Json(
                new { errors = outcome.Problems },
                statusCode: outcome.IsRateLimited ? StatusCodes.Status429TooManyRequests : StatusCodes.Status400BadRequest);
        }
    }
}
