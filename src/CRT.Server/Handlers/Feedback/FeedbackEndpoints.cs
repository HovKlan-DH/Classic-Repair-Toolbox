using System.Text;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Feedback
{
    // ###########################################################################################
    // FEEDBACK, over HTTP (owner request, 2026-10-03). A rim over FeedbackFormReader and
    // FeedbackFlow and nothing else.
    //
    //   POST /api/feedback   CRT's Feedback tab form -> 200 "Success",
    //                                                   400 nothing to send / an unreadable form,
    //                                                   413 too large, 415 not a form,
    //                                                   429 too many from this address,
    //                                                   502 the mail could not be sent,
    //                                                   507 no room on the disk for it.
    //
    // *** OLDER CRTs REACH IT TOO. *** They post the same form to /app-feedback/, the PHP page's
    // address, and Apache forwards that here (DEPLOYMENT.md) - which is why the answer is the
    // PHP's plain-text "Success" rather than JSON: an installed CRT reads exactly that.
    //
    // ANONYMOUS, like the PHP page: anybody may send feedback. Limited per address in memory, like
    // board views - and the body to FeedbackContract.MaximumRequestBytes (256 MB), on the route.
    // ###########################################################################################
    public static class FeedbackEndpoints
    {
        public const string RateLimitPolicy = "feedback";

        public static void MapFeedbackEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            app.MapPost("/api/" + FeedbackContract.PathUnderApi, FeedbackEndpoints.ReceiveAsync)
                .WithBodyLimit(RequestBodyLimits.FeedbackBytes)
                .RequireRateLimiting(FeedbackEndpoints.RateLimitPolicy);
        }

        // The limiter the route above names - registered with the other services.
        public static void AddFeedbackRateLimit(this IServiceCollection services) =>
            services.AddPerAddressRateLimit(FeedbackEndpoints.RateLimitPolicy, FeedbackFlow.MaxPerAddressPerHour);

        private static async Task<IResult> ReceiveAsync(
            HttpContext context,
            ServerOptions options,
            IEmailSender mailer,
            FeedbackStorage storage,
            ILoggerFactory loggers,
            CancellationToken cancellationToken)
        {
            string root = options.FeedbackRoot!;

            // The zip is written to disk before it is unpacked - so ask for room first.
            if (!FeedbackFlow.HasRoomForUpload(context.Request.ContentLength, Program.FreeBytesAt(root), options.MinimumFreeDiskBytes))
                return Results.Text("The server has no room for this just now.", "text/plain", Encoding.UTF8, StatusCodes.Status507InsufficientStorage);

            // The zip lands beside the feedback folders while it is read, named so it is never
            // taken for one, and is gone again whatever happens.
            string incoming = Path.Combine(root, $".incoming-{Guid.NewGuid():N}.zip");

            try
            {
                FeedbackFormRead read = await FeedbackFormReader.ReadAsync(
                    context.Request.ContentType, context.Request.Body, incoming, cancellationToken);

                if (read.IsRefused)
                    return Results.Text(read.RefusalReason, "text/plain", Encoding.UTF8, read.RefusalStatus);

                // The folder's total and the one-unpack-at-a-time gate are the service's one
                // FeedbackStorage, shared by every feedback.
                var destination = new FeedbackDestination(
                    root,
                    options.FeedbackToAddress!,
                    () => Program.FreeBytesAt(root),
                    options.MinimumFreeDiskBytes,
                    FeedbackFlow.NewReference,
                    options.FeedbackMaxStoredBytes,
                    storage.StoredBytes,
                    UnpackGate: storage.UnpackGate,
                    Saved: storage.Saved,
                    Removed: storage.Removed);

                FeedbackOutcome outcome = await FeedbackFlow.HandleAsync(read.Form!, destination, mailer, cancellationToken);

                // Said without the address or a word of the text - the mail has those.
                ILogger logger = loggers.CreateLogger("CRT.Server.Feedback");

                if (outcome.IsDelivered)
                    logger.LogInformation("Feedback received and mailed: {Reference}, {Files} file(s) saved.", outcome.Reference ?? "no files", outcome.FilesSaved);
                else
                    logger.LogWarning("Feedback NOT delivered ({Status}): {Reference}, {Files} file(s) saved.", outcome.StatusCode, outcome.Reference ?? "no files", outcome.FilesSaved);

                return Results.Text(outcome.Answer, "text/plain", Encoding.UTF8, outcome.StatusCode);
            }
            finally
            {
                try
                {
                    File.Delete(incoming);
                }
                catch (IOException)
                {
                    // Left behind, it is one stray file named so nobody mistakes it.
                }
            }
        }
    }
}
