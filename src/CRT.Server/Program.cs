using System.Reflection;
using System.Text.Json;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Database;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Health;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
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

            builder.Services.ConfigureHttpJsonOptions(options =>
            {
                options.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
                options.SerializerOptions.DefaultIgnoreCondition =
                    System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull;
            });

            var options = new ServerOptions();
            builder.Configuration.GetSection(ServerOptions.SectionName).Bind(options);
            builder.Services.AddSingleton(options);

            // ---------------------------------------------------------------------------------
            // Account services. Both are stateless and cheap to hold, so singletons.
            //
            // IEmailSender is registered as the INTERFACE deliberately: it is the seam that lets
            // the flows be tested without sending mail (CLAUDE.md test rule 6 forbids a test that
            // needs a network call). The SMTP implementation behind it is an untested I/O
            // boundary, the same as ScopeScpiClient and MiniproProcessRunner in CRT.App.
            // ---------------------------------------------------------------------------------
            builder.Services.AddSingleton<Argon2PasswordHasher>();
            builder.Services.AddSingleton<IEmailSender, SmtpEmailSender>();
            builder.Services.AddSingleton<IAccountStore, MySqlAccountStore>();

            // ---------------------------------------------------------------------------------
            // Submissions (Phase 4). BlobStoreRoot has no default and is validated below, so the
            // "!" here cannot be reached with a null - a misconfigured service never gets this far.
            // ---------------------------------------------------------------------------------
            builder.Services.AddSingleton<ISubmissionStore, MySqlSubmissionStore>();
            builder.Services.AddSingleton(provider => new BlobStore(
                options.BlobStoreRoot!,
                provider.GetRequiredService<ILogger<BlobStore>>()));

            // Collects submissions whose upload window closed without a finalise - their partial
            // blobs and their rows. Anonymous submitting is what makes this necessary: it is what
            // bounds the disk an abandoned or hostile upload can take. It existed as a flow for
            // months with nothing running it; see AbandonedUploadSweeper's header. Hosted services
            // start with app.Run(), so it never touches a schema the migrations below have not
            // brought up to date.
            builder.Services.AddAbandonedUploadSweeper();

            // Reads the published board a submission is compared against (Phase 5, task 3).
            // A singleton because it holds no per-request state; BoardDataReader's own cache sits
            // behind it and is shared deliberately, so two reviewers opening submissions for the
            // same board do not each pay for a parse.
            builder.Services.AddSingleton<PublishedBoardReader>();

            // Publishing (Phase 5, tasks 5 and 6). Singletons for the same reason: neither holds
            // per-request state, and both take their inputs as arguments rather than as fields.
            //
            // ApprovePublishFlow is the ONLY irreversible operation the service exposes - it
            // overwrites a published board with no retained revision behind it - so its own header
            // is worth reading before changing anything it touches.
            builder.Services.AddSingleton<PublishExecutor>();
            builder.Services.AddSingleton<ApprovePublishFlow>();

            // Tells the contributor what a reviewer decided. A singleton for the same reason as
            // the two above - it holds only the mailer seam and a logger, and takes everything
            // about a particular submission as arguments.
            builder.Services.AddSingleton<SubmissionNotifier>();

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
            // stops here instead of starting and guessing. Throwing means the process exits
            // non-zero, systemd reports the unit failed, and the journal carries the reasons -
            // an outage, which is noticed at once and harms nobody, rather than an exposure.
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
                Program.CanWriteToDirectory);

            if (failures.Count > 0)
            {
                ILogger<ServerOptions> configurationLogger =
                    app.Services.GetRequiredService<ILogger<ServerOptions>>();

                foreach (string failure in failures)
                    configurationLogger.LogCritical("Configuration error: {Failure}", failure);

                throw new InvalidOperationException(
                    $"The service cannot start: {failures.Count} configuration error(s) in " +
                    $"appsettings.Production.json. See the journal for details " +
                    $"(journalctl -u crt-server -n 20 --no-pager -p warning). " +
                    $"First: {failures[0]}");
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
            // ---------------------------------------------------------------------------------
            string migrationsDirectory = Path.Combine(AppContext.BaseDirectory, "Migrations");

            MigrationRunner.MigrateAsync(
                options.ConnectionString!,
                migrationsDirectory,
                app.Services.GetRequiredService<ILoggerFactory>().CreateLogger("CRT.Server.Migrations"))
                .GetAwaiter()
                .GetResult();

            app.UseForwardedHeaders();

            // ---------------------------------------------------------------------------------
            // Health. Unauthenticated on purpose - it is what the maintainer and any uptime check
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

            app.Run();
        }

        // ###########################################################################################
        // Answers "can this process actually create a file here?" by trying, then cleaning up.
        //
        // Inspecting permission bits would be the obvious implementation and would be wrong: the
        // service's ability to write depends on its group membership, on ACLs, on whether the mount
        // is read-only, and on systemd's ProtectSystem=strict plus ReadWritePaths - none of which
        // are visible in the mode bits. The only honest answer comes from attempting it, and this
        // runs once at startup so the cost is irrelevant.
        //
        // A unique name is used rather than a fixed one so two services starting at the same moment
        // cannot collide, and the file is removed in a finally so a probe never leaves litter in the
        // data tree.
        // ###########################################################################################
        private static bool CanWriteToDirectory(string directory)
        {
            string probePath = Path.Combine(directory, $".crt-server-write-probe-{Guid.NewGuid():N}");

            try
            {
                using (FileStream stream = File.Create(probePath, 1, FileOptions.DeleteOnClose))
                {
                    stream.WriteByte(0);
                }

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                return false;
            }
            finally
            {
                // DeleteOnClose normally handles this; the explicit delete covers the case where the
                // handle was closed by an exception path before the flag could take effect.
                try
                {
                    if (File.Exists(probePath))
                        File.Delete(probePath);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Nothing useful to do - a stray zero-byte probe file is harmless, and throwing
                    // from a cleanup path would mask the real result.
                }
            }
        }
    }
}
