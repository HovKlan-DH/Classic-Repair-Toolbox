using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The administrator's half of the oscilloscope models sent in from CRT (CRT 4.0 step 5) - a rim
    // over OscilloscopeModelFlows, in AdminEndpoints' group, so every route is refused to anybody but
    // an administrator. The bodies are CRT.Data's records (OscilloscopeModelDecisionRequest), a model
    // id - small, at the default body limit.
    //
    //   GET  /api/admin/oscilloscope-models          200 OscilloscopeModelQueueAnswer
    //   POST /api/admin/oscilloscope-models/publish  200 OscilloscopeModelDecisionAnswer; 404 no such
    //                                                model waiting; 409 decided meanwhile, or a main
    //                                                file that could not be written
    //   POST /api/admin/oscilloscope-models/reject   the same, and 400 with no reason
    // ###########################################################################################
    public static partial class AdminEndpoints
    {
        private static void MapOscilloscopeModelRoutes(RouteGroupBuilder admin)
        {
            admin.MapGet("/oscilloscope-models", AdminEndpoints.ListOscilloscopeModelsAsync);
            admin.MapPost("/oscilloscope-models/publish", AdminEndpoints.PublishOscilloscopeModelAsync);
            admin.MapPost("/oscilloscope-models/reject", AdminEndpoints.RejectOscilloscopeModelAsync);
        }

        private static async Task<IResult> ListOscilloscopeModelsAsync(
            HttpContext context,
            IAccountStore accounts,
            IOscilloscopeModelStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            return Results.Ok(await OscilloscopeModelFlows.ListAsync(store, accounts, options, cancellationToken));
        }

        private static async Task<IResult> PublishOscilloscopeModelAsync(
            OscilloscopeModelDecisionRequest request,
            HttpContext context,
            IAccountStore accounts,
            IOscilloscopeModelStore store,
            IEmailSender mailer,
            ServerOptions options,
            PublishLock publishLock,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            OscilloscopeModelDecisionOutcome outcome = await OscilloscopeModelFlows.PublishAsync(
                access!,
                request,
                store,
                accounts,
                mailer,
                options,
                tree => DataChecksumManifest.Write(tree.Root, tree.PublicBaseUrl, tree.ManifestPath),
                publishLock,
                DateTimeOffset.UtcNow,
                loggerFactory.CreateLogger(typeof(OscilloscopeModelFlows).FullName!),
                cancellationToken);

            return AdminEndpoints.ResultOf(outcome);
        }

        private static async Task<IResult> RejectOscilloscopeModelAsync(
            OscilloscopeModelDecisionRequest request,
            HttpContext context,
            IAccountStore accounts,
            IOscilloscopeModelStore store,
            IEmailSender mailer,
            PublishLock publishLock,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            OscilloscopeModelDecisionOutcome outcome = await OscilloscopeModelFlows.RejectAsync(
                access!,
                request,
                store,
                accounts,
                mailer,
                publishLock,
                DateTimeOffset.UtcNow,
                loggerFactory.CreateLogger(typeof(OscilloscopeModelFlows).FullName!),
                cancellationToken);

            return AdminEndpoints.ResultOf(outcome);
        }

        private static IResult ResultOf(OscilloscopeModelDecisionOutcome outcome) =>
            outcome.Answer is not null
                ? Results.Ok(outcome.Answer)
                : Results.Json(new { error = outcome.Error }, statusCode: outcome.Status);
    }
}
