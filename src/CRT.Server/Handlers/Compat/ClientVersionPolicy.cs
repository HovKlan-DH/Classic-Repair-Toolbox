using Handlers.DataHandling;
using Microsoft.AspNetCore.Http;

namespace CRT.Server.Handlers.Compat
{
    // ###########################################################################################
    // WHICH INSTALLED CRTs EACH PART OF THE API STILL SERVES (owner request, 2026-10-04: "It is
    // important that all older versions will continue to work, including the checkin and feedback").
    //
    // *** THE DEFAULT IS EVERY VERSION, AND IT IS MEANT TO STAY THAT WAY. *** A change to the API is
    // made so installed CRTs keep working - fields added, never renamed or removed; a new meaning on
    // a new route - and the compatibility check (CRT.Server.Tests' ApiCompatibilityTests) fails on
    // anything a released CRT relied on that has gone. A minimum below is the LAST resort, for a
    // change that truly cannot be made that way: the CRTs below it are then told, in words, to
    // update (ClientVersionContract.Outdated, HTTP 426), instead of failing in a way nobody can read.
    //
    // Three areas:
    //   - Forever: the launch check-in, feedback, board views, the health check and the old 2.x
    //     contribution address. NEVER refused, whatever the minimums say - every CRT ever released
    //     sends these, and the owner's promise is that they keep working.
    //   - Submissions: what the Drafts tab's Submit and its status checks use.
    //   - Maintainer: the review, administration and account routes - the Maintainer tab. Its users
    //     are few and known, so this is where a minimum is acceptable sooner (owner, 2026-10-04).
    //
    // Two ways a gated area turns a CRT away, both answered with the same "update CRT" (426):
    //   - THE API REVISION (2026-10-04) - the one in force. A CRT built for an older revision than
    //     this server's (ClientVersionContract.ApiRevision, which a test raises whenever the API
    //     changes in a way such a CRT cannot use) is told to update. Nobody picks a version: the
    //     revision is in the source both ends are built from.
    //   - A MINIMUM CRT VERSION per area - the older, manual lever, for a decision about released
    //     versions. None is set.
    //
    // FAIL OPEN: a request naming no CRT version or no API revision (a browser opening a mailed link,
    // curl, a CRT 2.x, a CRT whose version does not parse) is never refused on that count - the rule
    // the old contribution upload kept too.
    // ###########################################################################################
    public enum ClientArea
    {
        Forever,
        Submissions,
        Maintainer
    }

    public sealed class ClientVersionPolicy
    {
        // ###########################################################################################
        // *** THE MINIMUMS IN FORCE. *** Null is "every CRT". Raising one is a MAJOR server bump
        // (VERSION.md) and is recorded there with the versions it turns away - see CLAUDE.md,
        // "Installed CRTs keep working".
        // ###########################################################################################
        public static ClientVersionPolicy Current { get; } =
            new(submissions: null, maintainer: null, apiRevision: ClientVersionContract.ApiRevision);

        // The path prefixes of the two gated areas; anything else is Forever. A prefix matches the
        // path itself or the path followed by "/", so "/api/reviewer" is not "/api/review".
        private static readonly (string Prefix, ClientArea Area)[] Areas =
        [
            ("/api/submissions", ClientArea.Submissions),
            ("/api/review", ClientArea.Maintainer),
            ("/api/admin", ClientArea.Maintainer),
            ("/api/accounts", ClientArea.Maintainer)
        ];

        public ClientVersionPolicy(CrtVersion? submissions, CrtVersion? maintainer, int? apiRevision = null)
        {
            this.Submissions = submissions;
            this.Maintainer = maintainer;
            this.ApiRevision = apiRevision;
        }

        public CrtVersion? Submissions { get; }

        public CrtVersion? Maintainer { get; }

        // The oldest API revision the gated areas serve - the server's own (Current). Null serves any.
        public int? ApiRevision { get; }

        public static ClientArea AreaOf(string? path)
        {
            string value = path ?? string.Empty;

            foreach ((string prefix, ClientArea area) in ClientVersionPolicy.Areas)
            {
                if (value.Equals(prefix, StringComparison.OrdinalIgnoreCase) ||
                    value.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase))
                {
                    return area;
                }
            }

            return ClientArea.Forever;
        }

        public CrtVersion? MinimumFor(string? path) => ClientVersionPolicy.AreaOf(path) switch
        {
            ClientArea.Submissions => this.Submissions,
            ClientArea.Maintainer => this.Maintainer,
            _ => null
        };

        // ###########################################################################################
        // The refusal for this request, or null to let it through. Only a request to a gated area is
        // refused: one naming an API revision BELOW the server's (apiRevisionHeader, the value of
        // ClientVersionContract.ApiRevisionHeader), or one whose User-Agent names a CRT version below
        // the area's minimum. A NEWER revision is served - the server is deployed before the CRT
        // release that needs it, and the newer CRT is the one being tried out.
        // ###########################################################################################
        public ClientOutdatedAnswer? Refuse(string? path, string? userAgent, string? apiRevisionHeader = null)
        {
            ClientArea area = ClientVersionPolicy.AreaOf(path);

            if (area == ClientArea.Forever)
                return null;

            CrtVersion? client = CrtVersion.FromUserAgent(userAgent);

            if (this.ApiRevision is int served &&
                ClientVersionContract.ApiRevisionFrom(apiRevisionHeader) is int sent &&
                sent < served)
            {
                return ClientVersionContract.OutdatedApi(client);
            }

            CrtVersion? minimum = this.MinimumFor(path);

            return minimum is null || client is null || client >= minimum
                ? null
                : ClientVersionContract.Outdated(client, minimum);
        }
    }

    // ###########################################################################################
    // The policy, in the request pipeline: after routing and the body-size limit (Program.cs), before
    // the rate limiter - so a refused CRT spends none of its address's allowance.
    // ###########################################################################################
    public static class ClientVersionGate
    {
        public static void UseClientVersionGate(this IApplicationBuilder app, ClientVersionPolicy policy)
        {
            ArgumentNullException.ThrowIfNull(app);
            ArgumentNullException.ThrowIfNull(policy);

            ILogger logger = app.ApplicationServices
                .GetRequiredService<ILoggerFactory>()
                .CreateLogger("CRT.Server.ClientVersion");

            app.Use((context, next) => ClientVersionGate.HandleAsync(context, next, policy, logger));
        }

        // ###########################################################################################
        // One request. Separate from the registration so a test can drive it with a DefaultHttpContext
        // - nothing listens and no request is served (CLAUDE.md test rule 6).
        //
        // *** THE BODY IS READ BEFORE THE ANSWER IS WRITTEN. *** A CRT uploading a blob is still
        // sending when the refusal is ready, and Apache in front of the service turns an answer that
        // arrives mid-upload into a 502 - so the client would never see the words this exists for.
        // Reading stops at the route's own body limit (set before this runs), and a body that cannot
        // be read to the end still gets the answer.
        // ###########################################################################################
        internal static async Task HandleAsync(HttpContext context, RequestDelegate next, ClientVersionPolicy policy, ILogger logger)
        {
            ClientOutdatedAnswer? refusal = policy.Refuse(
                context.Request.Path.Value,
                context.Request.Headers.UserAgent.ToString(),
                context.Request.Headers[ClientVersionContract.ApiRevisionHeader].ToString());

            if (refusal is null)
            {
                await next(context);
                return;
            }

            logger.LogInformation(
                "Told a CRT to update: {Method} {Path} from {UserAgent} (API revision {Revision}), which is older than {Minimum}.",
                context.Request.Method,
                context.Request.Path.Value,
                context.Request.Headers.UserAgent.ToString(),
                context.Request.Headers[ClientVersionContract.ApiRevisionHeader].ToString(),
                refusal.MinimumVersion ?? $"API revision {policy.ApiRevision}");

            try
            {
                await context.Request.Body.CopyToAsync(Stream.Null, context.RequestAborted);
            }
            catch (Exception ex) when (ex is BadHttpRequestException or IOException or OperationCanceledException)
            {
                // Too large, or the client went away - the answer is still worth trying.
            }

            context.Response.StatusCode = ClientVersionContract.OutdatedStatus;

            await context.Response.WriteAsJsonAsync(refusal, ReviewApiContract.WireSettings, context.RequestAborted);
        }
    }
}
