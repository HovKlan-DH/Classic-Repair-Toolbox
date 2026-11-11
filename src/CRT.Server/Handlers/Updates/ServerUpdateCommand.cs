using System.Reflection;
using System.Text.RegularExpressions;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // `CRT.Server update-server` - ONE LOOK FOR A NEWER SERVER BUILD, AND INSTALLING IT (owner request,
    // 2026-11-09; ServerUpdater says how).
    //
    //   dotnet CRT.Server.dll update-server --channel=releases|main
    //                                       --folder=<the instance's folder, e.g. .../crt-server>
    //                                       --service=<its unit, e.g. crt-server.service>
    //                                       --health=<its health address, http://127.0.0.1:<port>/api/health>
    //                                       [--gh=<the gh tool, /usr/bin/gh>] [--dry-run] [--retry-failed]
    //
    // Run by crt-server-update.timer (and crt-server-test-update.timer for the test instance) as the
    // crt-deploy user, INSTEAD of the service: no web host, no settings, no database. Every value is on
    // the unit's command line, which only root may change; where builds come from and who must have built
    // them is fixed in the code (ServerReleaseFeed). `--dry-run` downloads and proves the newest build and
    // says whether it would be installed; `--retry-failed` tries a build again that did not come up before.
    //
    // Exit codes: 0 nothing wrong (up to date, installed, held for the project owner, a dry run), 1 a build
    // failed, was refused, or GitHub could not be asked - the status file says which - and 2 the command line
    // was wrong.
    // ###########################################################################################
    public static class ServerUpdateCommand
    {
        public const string Name = "update-server";

        public const int Done = 0;
        public const int Failed = 1;
        public const int WrongCommandLine = 2;

        public const string UsageText =
            "Usage: CRT.Server update-server --channel=releases|main --folder=<instance folder> " +
            "--service=<unit>.service --health=http://127.0.0.1:<port>/api/health [--gh=<path to gh>] [--dry-run] [--retry-failed]";

        public const string DefaultGh = "/usr/bin/gh";

        public sealed record Options(
            string Channel,
            string Folder,
            string Service,
            Uri Health,
            string Gh,
            bool DryRun,
            bool RetryFailed);

        private static readonly Regex ServiceName = new("^[a-z0-9][a-z0-9_.@-]{0,60}\\.service$", RegexOptions.CultureInvariant);

        public static bool IsCommand(IReadOnlyList<string> args) =>
            args.Count > 0 && string.Equals(args[0], ServerUpdateCommand.Name, StringComparison.Ordinal);

        // ###########################################################################################
        // The options after the command's name, or null with the reason. Each value option once, as
        // --name=value; anything else refused. The folder and gh absolute with no "..", the service a unit
        // name, the health address the service's own on this machine - never anywhere the network reaches.
        // ###########################################################################################
        public static Options? Parse(IReadOnlyList<string> args, out string why)
        {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            bool dryRun = false;
            bool retryFailed = false;

            foreach (string arg in args)
            {
                if (arg == "--dry-run")
                {
                    dryRun = true;
                    continue;
                }

                if (arg == "--retry-failed")
                {
                    retryFailed = true;
                    continue;
                }

                int equals = arg.IndexOf('=');
                string name = equals < 0 ? arg : arg[..equals];
                string value = equals < 0 ? string.Empty : arg[(equals + 1)..].Trim().Trim('"');

                if (name is not ("--channel" or "--folder" or "--service" or "--health" or "--gh"))
                {
                    why = $"Unknown option [{arg}].";
                    return null;
                }

                if (!values.TryAdd(name, value))
                {
                    why = $"[{name}] is given twice.";
                    return null;
                }

                if (value.Length == 0)
                {
                    why = $"[{name}] has no value.";
                    return null;
                }
            }

            foreach (string required in new[] { "--channel", "--folder", "--service", "--health" })
            {
                if (!values.ContainsKey(required))
                {
                    why = $"[{required}] is missing.";
                    return null;
                }
            }

            string channel = values["--channel"];

            if (channel is not (ServerUpdateOutcomes.ReleasesChannel or ServerUpdateOutcomes.MainChannel))
            {
                why = $"[--channel] is [{channel}] - it is [{ServerUpdateOutcomes.ReleasesChannel}] or [{ServerUpdateOutcomes.MainChannel}].";
                return null;
            }

            string gh = values.TryGetValue("--gh", out string? given) ? given : ServerUpdateCommand.DefaultGh;

            foreach ((string option, string path) in new[] { ("--folder", values["--folder"]), ("--gh", gh) })
            {
                if (!ServerUpdateCommand.IsPlainAbsolutePath(path))
                {
                    why = $"[{option}] is [{path}] - it must be a full path with no \"..\" in it.";
                    return null;
                }
            }

            string service = values["--service"];

            if (!ServerUpdateCommand.ServiceName.IsMatch(service))
            {
                why = $"[--service] is [{service}] - it must be a unit's name, such as crt-server.service.";
                return null;
            }

            if (!Uri.TryCreate(values["--health"], UriKind.Absolute, out Uri? health) ||
                health.Scheme != Uri.UriSchemeHttp ||
                !health.IsLoopback ||
                !string.Equals(health.AbsolutePath, "/api/health", StringComparison.Ordinal) ||
                health.Query.Length > 0)
            {
                why = $"[--health] is [{values["--health"]}] - it must be the service's own health address on this machine, http://127.0.0.1:<port>/api/health.";
                return null;
            }

            why = string.Empty;
            return new Options(channel, Path.TrimEndingDirectorySeparator(values["--folder"]), service, health, gh, dryRun, retryFailed);
        }

        // A full path with no "." or ".." part - on the box "/..." (read as such on Windows too, where the
        // tests read the shipped units); a test's temporary folder on Windows too.
        public static bool IsPlainAbsolutePath(string path) =>
            (path.StartsWith('/') || Path.IsPathFullyQualified(path)) &&
            !path.Replace('\\', '/').Split('/').Any(part => part is "." or "..");

        public static int Run(IReadOnlyList<string> args, TextWriter output)
        {
            if (ServerUpdateCommand.Parse(args.Skip(1).ToList(), out string why) is not Options options)
            {
                output.WriteLine(why);
                output.WriteLine(ServerUpdateCommand.UsageText);
                return ServerUpdateCommand.WrongCommandLine;
            }

            if (OperatingSystem.IsWindows())
            {
                output.WriteLine("update-server runs on the server's Linux box only - it switches links and asks systemd.");
                return ServerUpdateCommand.WrongCommandLine;
            }

            string version = typeof(ServerUpdateCommand).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? "unknown";
            string ghHome = Path.Combine(options.Folder, ServerUpdater.ReleasesFolderName, ServerUpdater.WorkFolderName, "gh");

            using var host = new ServerUpdateHost(options.Gh, ghHome, options.Health, $"CRT.Server-update-server/{version}");

            var updater = new ServerUpdater(
                new ServerUpdater.Options(options.Channel, options.Folder, options.Service, options.DryRun, options.RetryFailed),
                host,
                output);

            try
            {
                return updater.RunAsync(CancellationToken.None).GetAwaiter().GetResult()
                    ? ServerUpdateCommand.Done
                    : ServerUpdateCommand.Failed;
            }
            catch (Exception ex)
            {
                output.WriteLine($"The look for a newer server build failed - {ex}");
                return ServerUpdateCommand.Failed;
            }
        }
    }
}
