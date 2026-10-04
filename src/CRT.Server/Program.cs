using System.Reflection;
using System.Text.Json;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Compat;
using CRT.Server.Handlers.Database;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Feedback;
using CRT.Server.Handlers.Health;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;
using CRT.Server.Handlers;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.HttpOverrides;

namespace CRT.Server
{
    // ###########################################################################################
    // Entry point and wiring for the CRT contribution API.
    //
    // THIS FILE WIRES; IT DOES NOT DECIDE. Every rule lives in a pure class under Handlers/ so it
    // can be unit tested without a server. A rule that ends up expressed here is a rule no test
    // can reach - the same defect CLAUDE.md describes for logic trapped inside a UserControl.
    //
    // Phase 3 step 0 is deliberately just the health endpoint: the systemd -> Apache -> TLS ->
    // SELinux chain has several independent ways to fail, and debugging that chain at the same
    // time as debugging business logic is what turns one evening into one week. Deploy something
    // that does nothing, prove it answers over HTTPS, then add code. See DEPLOYMENT.md.
    // ###########################################################################################
    public static class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // ---------------------------------------------------------------------------------
            // UNDER SYSTEMD, LOG IN SYSTEMD'S FORMAT (2026-09-27). The default console format
            // writes "crit: Category[0]" with the message on a second line and nothing the journal
            // reads as a level, so the journal filed EVERY line as info - and
            // "journalctl -p warning", DEPLOYMENT.md's way to read the reasons without a core dump,
            // showed none of ours. The systemd format writes each entry on ONE line, prefixed with
            // its syslog level, which the journal strips and records: a crit is then a crit.
            // systemd sets JOURNAL_STREAM exactly when the output goes to the journal, so running
            // the service by hand in a terminal keeps the readable default.
            // ---------------------------------------------------------------------------------
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("JOURNAL_STREAM")))
                builder.Logging.AddSystemdConsole();

            // ---------------------------------------------------------------------------------
            // The service listens on 127.0.0.1 ONLY and is reached through the existing Apache
            // site. It must never bind a public interface - that is a hard requirement of the
            // security model, not a preference, and the systemd unit sets ASPNETCORE_URLS to a
            // loopback address to enforce it. This line is the in-process backstop for the case
            // where the unit file is wrong or the service is started by hand: without it,
            // Kestrel's default binding would be reachable from outside.
            // ---------------------------------------------------------------------------------
            if (string.IsNullOrWhiteSpace(builder.Configuration["ASPNETCORE_URLS"]) &&
                string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("ASPNETCORE_URLS")))
            {
                builder.WebHost.UseUrls("http://127.0.0.1:5199");
            }

            // ---------------------------------------------------------------------------------
            // Behind Apache, every request arrives from 127.0.0.1. Without this, the client IP
            // the service sees is the proxy's, which would silently turn every per-IP rate limit
            // into one global bucket - a security control that looks present and is not.
            //
            // KnownProxies is restricted to loopback deliberately: trusting X-Forwarded-For from
            // anywhere would let a caller spoof their own address and defeat the same limits.
            // ---------------------------------------------------------------------------------
            builder.Services.Configure<ForwardedHeadersOptions>(options =>
            {
                options.ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto;
                options.KnownProxies.Clear();
                options.KnownProxies.Add(System.Net.IPAddress.Loopback);
                options.KnownProxies.Add(System.Net.IPAddress.IPv6Loopback);
            });

            var options = new ServerOptions();
            builder.Configuration.GetSection(ServerOptions.SectionName).Bind(options);

            Program.AddServerServices(builder.Services, options);

            var app = builder.Build();

            // ---------------------------------------------------------------------------------
            // Install CRT.Data's logging sink before anything else runs, so a shared class logs
            // into the journal rather than silently discarding its output.
            // ---------------------------------------------------------------------------------
            CrtLog.Sink = new ServerLogAdapter(
                app.Services.GetRequiredService<ILogger<ServerLogAdapter>>());

            // ---------------------------------------------------------------------------------
            // CONFIGURATION IS VALIDATED BEFORE THE SERVICE LISTENS, AND A FAILURE IS FATAL.
            //
            // This is the enforcement half of "no default for the data tree": the values that
            // decide where data goes have no fallback, so a service that was not told explicitly
            // stops here instead of starting and guessing. The process exits with
            // ServerExitCodes.ConfigurationRefused, systemd reports the unit failed and - because
            // the unit names that code in RestartPreventExitStatus - does not retry it, and the
            // journal carries the reasons: an outage, which is noticed at once and harms nobody,
            // rather than an exposure.
            //
            // Health is NOT exempt. Serving health from a misconfigured process would report
            // "ok" for something that must not be running at all.
            //
            // The two probes are passed in rather than called inside the validator so the rule
            // itself stays pure and unit-testable. Writability is answered by attempting a real
            // write: permission bits, group membership, ACLs, a read-only mount and systemd's own
            // ProtectSystem can each make a writable-LOOKING directory refuse.
            // ---------------------------------------------------------------------------------
            IReadOnlyList<string> failures = ServerOptionsValidator.Validate(
                options,
                Directory.Exists,
                TreeWriteAccess.CanWrite);

            if (failures.Count > 0)
            {
                ILogger<ServerOptions> configurationLogger =
                    app.Services.GetRequiredService<ILogger<ServerOptions>>();

                foreach (string failure in failures)
                    configurationLogger.LogCritical("Configuration error: {Failure}", failure);

                // EXIT, not throw - see ServerExitCodes for the core dump every five seconds that
                // throwing produced. Disposing the app first flushes the console logger, which
                // writes on a background thread and would lose the lines above at process exit.
                ((IDisposable)app).Dispose();

                Console.Error.WriteLine(
                    $"The service cannot start: {failures.Count} configuration error(s) in " +
                    $"appsettings.Production.json. See the journal for details " +
                    $"(journalctl -u crt-server -n 20 --no-pager -p warning). " +
                    $"First: {failures[0]}");

                Environment.ExitCode = ServerExitCodes.ConfigurationRefused;
                return;
            }

            // ---------------------------------------------------------------------------------
            // BRING THE SCHEMA UP TO DATE BEFORE SERVING ANYTHING, AND FAIL THE START IF IT
            // CANNOT BE DONE.
            //
            // Migrations run here rather than from a separate command because this service is
            // deployed by hand: a separate step is a step that gets forgotten, and the failure
            // mode of forgetting is a service running against a schema it does not match, which
            // surfaces as confusing runtime errors rather than as a clear refusal.
            //
            // Blocking on the task is correct in Main - there is nothing else for this thread to
            // do, and the service must not listen until the schema is right. Every rule about
            // WHICH migrations run lives in MigrationPlan, unit tested without a database.
            //
            // A FAILURE EXITS, IT DOES NOT THROW (2026-09-27) - the same reason a refused setting
            // exits (ServerExitCodes): a throw out of Main aborts the process, which wrote a core
            // dump into the journal every five seconds while systemd restarted it, and the one line
            // saying what was wrong was never logged at crit at all. ServerExitCodes.ForStartupFailure
            // decides the code: a failed migration is not retried, an unreachable database is.
            // ---------------------------------------------------------------------------------
            string migrationsDirectory = Path.Combine(AppContext.BaseDirectory, "Migrations");

            ILogger migrationLogger = app.Services.GetRequiredService<ILoggerFactory>().CreateLogger("CRT.Server.Migrations");

            try
            {
                MigrationRunner.MigrateAsync(options.ConnectionString!, migrationsDirectory, migrationLogger)
                    .GetAwaiter()
                    .GetResult();
            }
            catch (Exception ex) when (ServerExitCodes.ForStartupFailure(ex) is int exitCode)
            {
                if (exitCode == ServerExitCodes.DatabaseUnreachable)
                    migrationLogger.LogCritical("The database cannot be reached: {Reason} - the service tries again in a few seconds.", ex.Message);
                else
                    migrationLogger.LogCritical("{Reason} Once it is fixed: sudo systemctl restart crt-server", ex.Message);

                // Disposing the app flushes the console logger, which writes on a background
                // thread and would lose the line above at process exit.
                ((IDisposable)app).Dispose();

                Environment.ExitCode = exitCode;
                return;
            }

            app.UseForwardedHeaders();

            // Routing runs HERE, explicitly, so the body-size middleware below sees the endpoint it
            // chose. Left implicit, WebApplication would still route first - but the limit's
            // correctness should not rest on a default nobody wrote down.
            app.UseRouting();

            // ---------------------------------------------------------------------------------
            // Which CRT versions call which route (owner request, 2026-10-04) - counted in memory
            // here, after routing named the route and BEFORE anything can refuse the request: a CRT
            // turned away as outdated or over its rate limit still called the route, and is exactly
            // what retiring one has to know about. ApiUsageFlusher writes the counts every few
            // minutes; a request matching no route is not counted (ApiUsageRules).
            // ---------------------------------------------------------------------------------
            ApiUsageCounter apiUsage = app.Services.GetRequiredService<ApiUsageCounter>();

            app.Use(async (context, next) =>
            {
                ApiUsageRules.Count(apiUsage, context.GetEndpoint(), context.Request.Headers.UserAgent.ToString(), DateTimeOffset.UtcNow);
                await next(context);
            });

            // ---------------------------------------------------------------------------------
            // A body-size limit for every request, taken from the ROUTE routing chose - each route
            // carries its own, written where it is mapped (`.WithBodyLimit(...)`) - BEFORE the
            // endpoint reads a byte (security review, 2026-09-25; on the routes themselves since the
            // code review the same day). A route with none, and a request matching no route, get
            // the small default. Without this every route - including the anonymous submission
            // endpoint, which stores its body in the database - accepted Kestrel's default ~30 MB.
            //
            // IsReadOnly is true once a body has started to be read, and setting the limit then
            // throws; nothing has read it yet at this point, so the guard is belt and braces.
            // ---------------------------------------------------------------------------------
            app.Use(async (context, next) =>
            {
                IHttpMaxRequestBodySizeFeature? limit = context.Features.Get<IHttpMaxRequestBodySizeFeature>();

                if (limit is { IsReadOnly: false })
                    limit.MaxRequestBodySize = RequestBodyLimits.For(context.GetEndpoint());

                await next(context);
            });

            // ---------------------------------------------------------------------------------
            // A CRT older than the part of the API it calls is told to update - HTTP 426 with
            // CRT.Data's ClientOutdatedAnswer - rather than failing in a way nobody can read
            // (2026-10-04). ClientVersionPolicy.Current serves this server's API revision and newer
            // and sets no minimum version; the check-in, feedback and board views are never refused. After the body-size
            // limit, which bounds the body it reads before answering, and before the rate limiter,
            // so a refused CRT spends none of its address's allowance.
            // ---------------------------------------------------------------------------------
            app.UseClientVersionGate(ClientVersionPolicy.Current);

            // Board views', feedback's and the check-in's per-address limits - after routing, which
            // names the route's policy, and before the endpoint runs.
            app.UseRateLimiter();

            Program.MapServerEndpoints(app);

            app.Run();
        }

        // ###########################################################################################
        // Every service the endpoints take, registered - split out of Main so a test can build the
        // server's REAL route table with the real registrations (RequestBodyLimitsTests): routing
        // decides from these which handler parameters come from the container and which from the
        // body. Registering constructs nothing; the stores only open a connection when resolved.
        // ###########################################################################################
        internal static void AddServerServices(IServiceCollection services, ServerOptions options)
        {
            // The JSON both ends of the review API agree on - CRT.Data's ReviewApiContract, which
            // the Maintainer tab serialises its requests with too (code review, 2026-09-25).
            services.ConfigureHttpJsonOptions(json =>
                ReviewApiContract.ApplyWireSettings(json.SerializerOptions));

            services.AddSingleton(options);

            // ---------------------------------------------------------------------------------
            // Account services. Both are stateless and cheap to hold, so singletons.
            //
            // IEmailSender is registered as the INTERFACE deliberately: it is the seam that lets
            // the flows be tested without sending mail (CLAUDE.md test rule 6 forbids a test that
            // needs a network call). The SMTP implementation behind it is an untested I/O
            // boundary, the same as ScopeScpiClient and MiniproProcessRunner in CRT.App.
            // ---------------------------------------------------------------------------------
            services.AddSingleton<Argon2PasswordHasher>();
            services.AddSingleton<IEmailSender, SmtpEmailSender>();
            services.AddSingleton<IAccountStore, MySqlAccountStore>();

            // ---------------------------------------------------------------------------------
            // Submissions (Phase 4). BlobStoreRoot has no default and is validated below, so the
            // "!" here cannot be reached with a null - a misconfigured service never gets this far.
            // ---------------------------------------------------------------------------------
            services.AddSingleton<ISubmissionStore, MySqlSubmissionStore>();
            //
            // The free-space probe and reserve are what stop contributions filling the disk the
            // web site and the database share (security review, 2026-09-25) - see
            // BlobStore.HasRoomFor. The probe answers null when it cannot measure, which the store
            // treats as room rather than refusing every upload over a transient error.
            services.AddSingleton(provider => new BlobStore(
                options.BlobStoreRoot!,
                provider.GetRequiredService<ILogger<BlobStore>>(),
                () => Program.FreeBytesAt(options.BlobStoreRoot!),
                options.MinimumFreeDiskBytes));

            // Collects submissions whose upload window closed without a finalise - their partial
            // blobs and their rows. Anonymous submitting is what makes this necessary: it is what
            // bounds the disk an abandoned or hostile upload can take. It existed as a flow for
            // months with nothing running it; see AbandonedUploadSweeper's header. Hosted services
            // start with app.Run(), so it never touches a schema the migrations below have not
            // brought up to date.
            services.AddAbandonedUploadSweeper();

            // Reads the published board a submission is compared against (Phase 5, task 3).
            // A singleton because it holds no per-request state; BoardDataReader's own cache sits
            // behind it and is shared deliberately, so two maintainers opening submissions for the
            // same board do not each pay for a parse.
            services.AddSingleton<PublishedBoardReader>();

            // Publishing (Phase 5, tasks 5 and 6). Singletons for the same reason: neither holds
            // per-request state, and both take their inputs as arguments rather than as fields.
            //
            // ApprovePublishFlow is one of the two irreversible operations the service exposes - it
            // overwrites a published board with no retained revision behind it - so its own header
            // is worth reading before changing anything it touches. The other is SystemDeletionFlow.
            services.AddSingleton<PublishExecutor>();

            // The one lock every write to a published tree takes - the BETA publish and the
            // production promotion both. See PublishLock.
            services.AddSingleton<PublishLock>();
            services.AddSingleton<ApprovePublishFlow>();

            // BETA to production (2026-09-25). Switched off until the three Production* settings
            // are set; see ServerOptions.
            services.AddSingleton<ProductionPromotionFlow>();

            // Rolling a BETA board back to production's state, returning its submissions to the
            // queue (owner decision, 2026-09-27) - the production promotion's mirror image.
            services.AddSingleton<BetaRollbackFlow>();

            // Deleting a system from both trees and the database (owner request, 2026-10-03) - the
            // administrator's. Takes the publish lock: it writes both trees and their manifests.
            services.AddSingleton<SystemDeletionFlow>();

            // Placing a new system in the drop-down lists (owner request, 2026-09-27). Takes the
            // publish lock too: placing a system already in BETA writes BETA's main Excel data file.
            services.AddSingleton<SystemListingFlow>();

            // Tells the contributor what a maintainer decided. A singleton for the same reason as
            // the two above - it holds only the mailer seam and a logger, and takes everything
            // about a particular submission as arguments.
            services.AddSingleton<SubmissionNotifier>();

            // ---------------------------------------------------------------------------------
            // Board views (owner request, 2026-09-27): which boards CRT users look at, and in which
            // country - see CRT.Data's BoardViewContract. The country lookup is the one network call
            // the service makes on a request (ip-api.com, as the launch check-in uses); it keeps
            // addresses in memory for an hour and writes none of them anywhere.
            // ---------------------------------------------------------------------------------
            services.AddSingleton<IBoardViewStore, MySqlBoardViewStore>();
            services.AddSingleton<ICountryLookup, IpApiCountryLookup>();
            services.AddSingleton<BoardViewNameDirectory>();
            services.AddBoardViewRateLimit();

            // Feedback from CRT's Feedback tab (2026-10-03) - the old PHP page's job. Anonymous,
            // limited per address in memory like board views. FeedbackStorage is what every
            // feedback with files shares: one unpack at a time, and the folder's remembered total.
            services.AddFeedbackRateLimit();
            services.AddSingleton(new FeedbackStorage(() => FeedbackFlow.StoredBytesUnder(options.FeedbackRoot!)));

            // CRT's launch check-in (2026-10-03) - the old app-checkin PHP page's job, into the
            // crt_update table it wrote. Shares the board views' country lookup.
            services.AddSingleton<ICheckInStore, MySqlCheckInStore>();
            services.AddCheckInRateLimit();

            // CRT 2.x's contribution upload (2026-10-04), answered "please update" - anonymous,
            // limited per address in memory like feedback.
            services.AddLegacyContributionRateLimit();

            // ---------------------------------------------------------------------------------
            // Which CRT versions call which route (owner request, 2026-10-04): counted in memory by
            // the pipeline (Main), written into crt_api_calls by the flusher every few minutes and
            // as the service stops, shown under Account > "API usage" with every mapped route
            // (ApiRouteList, handed the route table by MapServerEndpoints).
            // ---------------------------------------------------------------------------------
            services.AddSingleton<ApiUsageCounter>();
            services.AddSingleton<IApiUsageStore, MySqlApiUsageStore>();
            services.AddSingleton<ApiRouteList>();
            services.AddHostedService<ApiUsageFlusher>();

            // Resetting the contribution data for going live (owner request, 2026-10-04) - see
            // DataResetFlow. Its one store empties every table in a single transaction.
            services.AddSingleton<IDataResetStore, MySqlDataResetStore>();
        }

        // ###########################################################################################
        // Every route the service answers - split out of Main for the same test. Each route that
        // takes more than the default body carries its own limit where it is mapped; see
        // RequestBodyLimits.
        // ###########################################################################################
        internal static void MapServerEndpoints(WebApplication app)
        {
            // ---------------------------------------------------------------------------------
            // Health. Unauthenticated on purpose - it is what the project owner and any uptime check
            // call, and it reveals nothing beyond "the process is alive" plus the deployed build.
            // See HealthReport's header for what it deliberately does not report.
            // ---------------------------------------------------------------------------------
            string? informationalVersion =
                HealthReport.ReadInformationalVersion(Assembly.GetExecutingAssembly());

            app.MapGet("/api/health", () =>
                Results.Ok(HealthReport.Build(DateTimeOffset.UtcNow, informationalVersion)));

            // ---------------------------------------------------------------------------------
            // Accounts. Every handler in there is a thin rim over AccountFlows - see
            // AccountEndpoints' header for why the status codes are part of the security design
            // rather than presentation.
            // ---------------------------------------------------------------------------------
            app.MapAccountEndpoints();
            app.MapSubmissionEndpoints();
            app.MapReviewEndpoints();
            app.MapAdminEndpoints();
            app.MapProductionEndpoints();
            app.MapSystemEndpoints();

            // Board views from CRT - anonymous, rate limited per address in memory.
            app.MapBoardViewEndpoints();

            // Feedback from CRT's Feedback tab, and from older CRTs through Apache's forward of
            // /app-feedback/ - anonymous, rate limited per address in memory.
            app.MapFeedbackEndpoints();

            // CRT's launch check-in, and older CRTs' through Apache's forward of /app-checkin/ -
            // anonymous, rate limited per address in memory.
            app.MapCheckInEndpoints();

            // CRT 2.x's contribution upload, through Apache's forward of /app-contribution/api/ once
            // the old PHP page is gone - answered "please update" in the words 2.5.0 understands.
            app.MapLegacyContributionEndpoints();

            // Every route above, for Account > "API usage" - read when it is asked for, so the list is
            // the route table as it is, never a copy kept by hand.
            app.Services.GetRequiredService<ApiRouteList>().Use(((IEndpointRouteBuilder)app).DataSources);
        }

        // ###########################################################################################
        // The free space on the disk holding `directory`, or null when it cannot be measured.
        //
        // DriveInfo on Linux resolves the mount the path lives on, which is the number that
        // matters: the blob store sharing a partition with the site and the database is exactly
        // the case the reserve exists for. Null on failure rather than zero, so a transient error
        // reading the mount table does not refuse every upload - see BlobStore.HasRoomFor.
        // ###########################################################################################
        internal static long? FreeBytesAt(string directory)
        {
            try
            {
                return new DriveInfo(Path.GetFullPath(directory)).AvailableFreeSpace;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
            {
                return null;
            }
        }
    }
}
