using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // *** THE SERVER UPDATES ITSELF FROM GITHUB (owner request, 2026-11-09: "would it be possible to
    // automatically update the server from GitHub? So, when there is a new 'real' version of the server,
    // then its server automatically update to this, and when there is a new alpha/beta server available,
    // then that server will update to that? The goal would be for me to not manually update this, copying
    // files and restarting service and checking if it is alive ... This should of course be as secure as
    // possible"). ***
    //
    // One look, run by a systemd timer as a user of its own (ServerUpdateCommand, INSTALLING.md
    // "Automatic updates from GitHub"). An instance's folder holds its builds side by side:
    //
    //   crt-server/
    //     app -> releases/current          the unit's WorkingDirectory and ExecStart, unchanged
    //     releases/                        the updater's, readable by the service's group
    //       10.2.1/  10.3.0/               one folder per build, named by its server version
    //       current -> 10.3.0              the build the service runs - switched in one step
    //       update-status.json             what the last look did (ServerUpdateStatus)
    //       .updater/                      downloads on their way in, the gh tool's caches, the lock
    //     settings/appsettings.Production.json   linked into every build; never in a download
    //
    // A look: the newest release of the channel (ServerReleaseFeed) -> its build downloaded -> GitHub's
    // attestation checked by `gh attestation verify`, pinned to this repository, its release workflow, main
    // and a GitHub-hosted runner - nothing unproven is even unpacked -> unpacked into a folder of its own
    // (ServerArchive) -> its own crt-server-release.json says its version and whether it is a pre-release
    // -> never older than the build running, never a pre-release on the live service, never a build that
    // failed before -> held while it ships a systemd or polkit file that differs from the one installed
    // (those need root, and the project owner) -> "current" pointed at it, the service restarted, and
    // /api/health asked until it answers with the new version -> otherwise back to the build before.
    //
    // The decisions are all here, against IServerUpdateHost, so the tests drive every path with a fake
    // host; ServerUpdateHost is the thin real one (HTTP, gh, systemctl, links).
    // ###########################################################################################
    internal sealed class ServerUpdater
    {
        public const string ReleasesFolderName = "releases";
        public const string CurrentLinkName = "current";
        public const string WorkFolderName = ".updater";
        public const string SettingsFileName = "appsettings.Production.json";

        // The polkit rule letting the updater restart the services - one file for both instances.
        public const string PolkitRuleFile = "50-crt-server-update.rules";

        // Put in a held build's folder by the project owner when the installed systemd/polkit files are
        // the ones wanted, though they differ from the build's (a unit changed by hand on purpose).
        public const string SystemFilesInstalledMarker = "system-files-installed";

        // How long a restarted service has to answer with its new version - its database migrations run first.
        public static readonly TimeSpan HealthWait = TimeSpan.FromSeconds(120);
        public static readonly TimeSpan HealthPoll = TimeSpan.FromSeconds(3);

        // Builds kept besides the one running: the one it replaced, and one more.
        public const int OlderBuildsKept = 2;

        // ###########################################################################################
        // *** A LOOK EVERY MINUTE, GITHUB'S API SELDOM (owner request, 2026-11-09: "Instead of updating only every 15
        // minutes, can it then update every minute?"). *** The timer runs a look every minute, but GitHub lets a
        // box with no login ask its API 60 times an hour - two instances every minute would ask 120. So a
        // minute's look reads the releases feed (ServerReleaseFeed.FeedUrl, not counted) and asks the API only:
        //   - when the feed changed - a release published, edited or deleted;
        //   - every minute for LookCloselyAfterAChange after it changed - a release is published before its files
        //     are uploaded, and its server build must not be missed for that;
        //   - at least every FullLookEvery anyway - a feed GitHub could not serve, or something it does not show;
        //   - and for a dry run or a retry asked by hand.
        // That is some 4 API calls an hour per instance, and a few more around a release. A build held for the
        // project owner is tried again every minute with no GitHub at all.
        // ###########################################################################################
        public static readonly TimeSpan FullLookEvery = TimeSpan.FromMinutes(15);
        public static readonly TimeSpan LookCloselyAfterAChange = TimeSpan.FromMinutes(10);

        public sealed record Options(
            string Channel,
            string InstanceFolder,
            string Service,
            bool DryRun = false,
            bool RetryFailed = false);

        private readonly Options thisOptions;
        private readonly IServerUpdateHost thisHost;
        private readonly TextWriter thisLog;

        public ServerUpdater(Options options, IServerUpdateHost host, TextWriter log)
        {
            this.thisOptions = options;
            this.thisHost = host;
            this.thisLog = log;
        }

        public string ReleasesFolder => Path.Combine(this.thisOptions.InstanceFolder, ServerUpdater.ReleasesFolderName);

        public string StatusPath => Path.Combine(this.ReleasesFolder, ServerUpdateStatus.FileName);

        public string SettingsPath => Path.Combine(this.thisOptions.InstanceFolder, "settings", ServerUpdater.SettingsFileName);

        private string WorkFolder => Path.Combine(this.ReleasesFolder, ServerUpdater.WorkFolderName);

        // ###########################################################################################
        // One look. True when nothing went wrong (up to date, installed, held, a dry run); false when it
        // failed, was refused or could not look - systemd then marks the run failed, and the status says why.
        // ###########################################################################################
        public async Task<bool> RunAsync(CancellationToken cancellationToken)
        {
            ServerUpdateStatus status = ServerUpdateStatus.Read(this.StatusPath) ?? new ServerUpdateStatus();
            DateTimeOffset now = this.thisHost.UtcNow;
            status.Channel = this.thisOptions.Channel;
            status.CheckedUtc = now;

            string? running = this.thisHost.PointedAt(this.ReleasesFolder);

            if (running == null || !ServerReleaseInfo.IsFolderSafeVersion(running) || !Directory.Exists(Path.Combine(this.ReleasesFolder, running)))
            {
                return this.Finish(status, ServerUpdateOutcomes.Error,
                    $"[{Path.Combine(this.ReleasesFolder, ServerUpdater.CurrentLinkName)}] is not a link to a build in [{this.ReleasesFolder}] - " +
                    "set the folders up as src/CRT.Server/INSTALLING.md's \"Automatic updates from GitHub\" says.");
            }

            Directory.CreateDirectory(this.WorkFolder);

            using FileStream? held = ServerUpdater.TryLock(Path.Combine(this.WorkFolder, "lock"));

            if (held == null)
            {
                this.thisLog.WriteLine("Another look is running - this one does nothing.");
                return true;
            }

            if (!await this.IsFullLookDueAsync(status, now, cancellationToken))
                return await this.QuietLookAsync(status, running, cancellationToken);

            status.FullLookUtc = now;

            string releasesJson;

            try
            {
                releasesJson = await this.thisHost.ReadReleasesAsync(cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                return this.Finish(status, ServerUpdateOutcomes.Error, $"GitHub's list of releases could not be read - {ex.Message}");
            }

            if (ServerReleaseFeed.Newest(releasesJson, this.thisOptions.Channel, out string why, out bool readable) is not { } candidate)
            {
                status.NewestVersion = null;
                status.NewestRelease = null;

                return readable
                    ? this.Finish(status, ServerUpdateOutcomes.UpToDate, $"{why} {running} runs.")
                    : this.Finish(status, ServerUpdateOutcomes.Error, why);
            }

            status.NewestRelease = candidate.Tag;

            // The build looked at last time: nothing is downloaded again for it.
            if (status.ExaminedAssetId == candidate.ArchiveAssetId &&
                await this.AgainAsync(status, candidate, running, cancellationToken) is bool done)
            {
                return done;
            }

            return await this.ExamineAsync(status, candidate, running, cancellationToken);
        }

        // ###########################################################################################
        // Whether this look asks GitHub's API (above, FullLookEvery). Reads the feed, noting when it changed.
        // ###########################################################################################
        private async Task<bool> IsFullLookDueAsync(ServerUpdateStatus status, DateTimeOffset now, CancellationToken cancellationToken)
        {
            string? fingerprint;

            try
            {
                fingerprint = await this.thisHost.ReadFeedFingerprintAsync(cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                fingerprint = null;
            }

            if (fingerprint != null && !string.Equals(fingerprint, status.FeedFingerprint, StringComparison.Ordinal))
            {
                status.FeedFingerprint = fingerprint;
                status.FeedChangedUtc = now;
            }

            return this.thisOptions.DryRun ||
                   this.thisOptions.RetryFailed ||
                   status.FullLookUtc is not DateTimeOffset last ||
                   now - last >= ServerUpdater.FullLookEvery ||
                   (fingerprint != null && status.FeedChangedUtc is DateTimeOffset changed && now - changed < ServerUpdater.LookCloselyAfterAChange);
        }

        // ###########################################################################################
        // A minute with nothing new on GitHub: what the last look found stands, said quietly - nothing in the
        // journal every minute. A build held for the project owner is tried again (its files may be installed now).
        // ###########################################################################################
        private async Task<bool> QuietLookAsync(ServerUpdateStatus status, string running, CancellationToken cancellationToken)
        {
            if (status.Outcome == ServerUpdateOutcomes.Held &&
                status.ExaminedVersion is string waiting &&
                ServerReleaseInfo.IsFolderSafeVersion(waiting) &&
                Directory.Exists(Path.Combine(this.ReleasesFolder, waiting)) &&
                ServerUpdater.Compare(waiting, running) > 0)
            {
                return await this.InstallAsync(status, status.NewestRelease ?? "?", waiting, running, cancellationToken, quietIfStillHeld: true);
            }

            this.Finish(status, status.Outcome, status.Message, quiet: true);
            return true;
        }

        // ###########################################################################################
        // The same build as the last look: what that look came to stands - unless it was held and may go
        // in now, or a retry was asked for. Null: look at it afresh (download it again).
        // ###########################################################################################
        private async Task<bool?> AgainAsync(ServerUpdateStatus status, ServerReleaseFeed.Candidate candidate, string running, CancellationToken cancellationToken)
        {
            string? seen = ServerReleaseInfo.IsFolderSafeVersion(status.ExaminedVersion) ? status.ExaminedVersion : null;

            switch (status.Outcome)
            {
                case ServerUpdateOutcomes.Failed:
                case ServerUpdateOutcomes.Refused:
                    // Said again as it was said - still the newest build, still not installed - and quietly: it went
                    // wrong once, and systemd was told then. A refusal may have been gh not reaching Sigstore, so a
                    // retry looks at that build again too.
                    if (this.thisOptions.RetryFailed)
                        return null;

                    this.Finish(status, status.Outcome, status.Message, quiet: true);
                    return true;

                case ServerUpdateOutcomes.Held when seen != null && Directory.Exists(Path.Combine(this.ReleasesFolder, seen)):
                    status.NewestVersion = seen;
                    return await this.InstallAsync(status, candidate.Tag, seen, running, cancellationToken);
            }

            if (seen != null && ServerUpdater.Compare(seen, running) <= 0)
            {
                status.NewestVersion = seen;
                return this.Finish(status, ServerUpdateOutcomes.UpToDate, ServerUpdater.UpToDateSentence(seen, running, candidate.Tag));
            }

            return null;
        }

        // ###########################################################################################
        // A build not looked at before: downloaded, proven, unpacked, read - and installed when it is newer.
        // ###########################################################################################
        private async Task<bool> ExamineAsync(ServerUpdateStatus status, ServerReleaseFeed.Candidate candidate, string running, CancellationToken cancellationToken)
        {
            string incoming = Path.Combine(this.WorkFolder, "incoming");
            ServerUpdater.DeleteFolder(incoming);
            Directory.CreateDirectory(incoming);

            string archive = Path.Combine(incoming, ServerReleaseFeed.ArchiveName);
            string bundle = Path.Combine(incoming, ServerReleaseFeed.BundleName);
            string unpacked = Path.Combine(incoming, "build");

            try
            {
                this.thisLog.WriteLine($"Downloading the server build of CRT release [{candidate.Tag}]...");

                try
                {
                    await this.thisHost.DownloadAsync(candidate.ArchiveUrl, archive, ServerReleaseFeed.ArchiveMaxBytes, cancellationToken);
                    await this.thisHost.DownloadAsync(candidate.BundleUrl, bundle, ServerReleaseFeed.BundleMaxBytes, cancellationToken);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    return this.Finish(status, ServerUpdateOutcomes.Error, $"The server build of CRT release [{candidate.Tag}] could not be downloaded - {ex.Message}");
                }

                // *** NOTHING UNPROVEN IS UNPACKED. ***
                AttestationCheck proof = await this.thisHost.VerifyAsync(archive, bundle, cancellationToken);

                if (!proof.Genuine)
                {
                    return this.Refuse(status, candidate, null,
                        $"The server build of CRT release [{candidate.Tag}] was NOT installed: GitHub's attestation does not prove " +
                        $"this repository's release workflow built it - {proof.Reason}");
                }

                try
                {
                    ServerArchive.Extract(archive, unpacked);
                }
                catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
                {
                    return this.Refuse(status, candidate, null, $"The server build of CRT release [{candidate.Tag}] could not be unpacked - {ex.Message}");
                }

                string infoPath = Path.Combine(unpacked, ServerReleaseInfo.FileName);
                ServerReleaseInfo? info = ServerReleaseInfo.Read(File.Exists(infoPath) ? File.ReadAllText(infoPath) : null, out string why);

                if (info == null)
                    return this.Refuse(status, candidate, null, $"The server build of CRT release [{candidate.Tag}] was not installed - {why}");

                if (!string.IsNullOrEmpty(proof.Commit) && !string.Equals(proof.Commit, info.Commit, StringComparison.OrdinalIgnoreCase))
                {
                    return this.Refuse(status, candidate, null,
                        $"The server build of CRT release [{candidate.Tag}] was not installed: it says it was built from [{info.Commit}], its attestation [{proof.Commit}].");
                }

                if (info.PreRelease && this.thisOptions.Channel == ServerUpdateOutcomes.ReleasesChannel)
                {
                    return this.Refuse(status, candidate, info.ServerVersion,
                        $"The server build of CRT [{info.CrtVersion}] is a pre-release's, and this service installs releases only.");
                }

                status.NewestVersion = info.ServerVersion;
                status.ExaminedAssetId = candidate.ArchiveAssetId;
                status.ExaminedVersion = info.ServerVersion;

                if (ServerUpdater.Compare(info.ServerVersion, running) <= 0)
                    return this.Finish(status, ServerUpdateOutcomes.UpToDate, ServerUpdater.UpToDateSentence(info.ServerVersion, running, candidate.Tag));

                if (string.Equals(status.FailedVersion, info.ServerVersion, StringComparison.OrdinalIgnoreCase) && !this.thisOptions.RetryFailed)
                {
                    return this.Finish(status, ServerUpdateOutcomes.Failed,
                        $"{info.ServerVersion} did not come up before, so it is not tried again by itself - a newer build is, " +
                        $"or `update-server --retry-failed` tries this one again. {running} runs.");
                }

                // Into a folder of its own, beside the others - never over the one running, which is older.
                string target = Path.Combine(this.ReleasesFolder, info.ServerVersion);
                ServerUpdater.DeleteFolder(target);
                Directory.Move(unpacked, target);
                this.thisHost.LinkSettings(target, this.SettingsPath);

                return await this.InstallAsync(status, candidate.Tag, info.ServerVersion, running, cancellationToken);
            }
            finally
            {
                ServerUpdater.DeleteFolder(incoming);
            }
        }

        // ###########################################################################################
        // The build in releases/`version`, unpacked and proven, goes live: unless its systemd or polkit
        // files wait for the project owner, or this is a dry run.
        // ###########################################################################################
        private async Task<bool> InstallAsync(ServerUpdateStatus status, string tag, string version, string running, CancellationToken cancellationToken, bool quietIfStillHeld = false)
        {
            string target = Path.Combine(this.ReleasesFolder, version);
            IReadOnlyList<string> waiting = this.SystemFilesWaiting(target, Path.Combine(this.ReleasesFolder, running));

            if (waiting.Count > 0)
            {
                return this.Finish(status, ServerUpdateOutcomes.Held,
                    $"{version} (CRT release [{tag}]) waits for you: it ships {string.Join(", ", waiting)} changed, which only root installs. " +
                    $"Copy them from [{target}] to /etc (units to /etc/systemd/system, the polkit rule to /etc/polkit-1/rules.d), and the " +
                    $"next look installs it. {running} runs meanwhile.",
                    quiet: quietIfStillHeld);
            }

            if (this.thisOptions.DryRun)
            {
                // Looked at afresh next time: a dry run installs nothing.
                status.ExaminedAssetId = null;
                return this.Finish(status, ServerUpdateOutcomes.WouldInstall, $"{version} (CRT release [{tag}]) would be installed - a dry run, so {running} runs on.");
            }

            this.thisLog.WriteLine($"Installing {version} in place of {running}...");
            this.thisHost.PointAt(this.ReleasesFolder, version);

            if (await this.thisHost.RestartAsync(this.thisOptions.Service, cancellationToken) is string refused)
            {
                // Not restarted, so the build before still runs - pointed back at it, and tried again at the next look.
                this.thisHost.PointAt(this.ReleasesFolder, running);

                return this.Finish(status, ServerUpdateOutcomes.Held,
                    $"{version} is ready, but [{this.thisOptions.Service}] could not be restarted - {refused}. Is the polkit rule " +
                    $"{ServerUpdater.PolkitRuleFile} installed? {running} runs meanwhile; the next look tries again.",
                    quiet: quietIfStillHeld);
            }

            if (await this.WaitForVersionAsync(version, cancellationToken) is null)
            {
                status.InstalledUtc = this.thisHost.UtcNow;
                status.InstalledVersion = version;
                status.InstalledFromVersion = running;
                status.FailedVersion = null;
                this.KeepFewBuilds(version, running);

                return this.Finish(status, ServerUpdateOutcomes.Installed, $"Installed {version} (CRT release [{tag}]) in place of {running}.");
            }

            // *** IT DID NOT COME UP: BACK TO THE BUILD BEFORE. ***
            string state = await this.thisHost.ServiceStateAsync(this.thisOptions.Service, cancellationToken);
            this.thisLog.WriteLine($"{version} did not answer as itself ({state}) - going back to {running}.");

            this.thisHost.PointAt(this.ReleasesFolder, running);
            status.FailedVersion = version;

            string? back = await this.thisHost.RestartAsync(this.thisOptions.Service, cancellationToken) is string notRestarted
                ? $"it could not be restarted - {notRestarted}"
                : await this.WaitForVersionAsync(running, cancellationToken);

            return this.Finish(status, ServerUpdateOutcomes.Failed, back == null
                ? $"{version} (CRT release [{tag}]) did not come up ({state}), so {running} runs again. " +
                  "Its journal says why (journalctl -u " + this.thisOptions.Service + "). It is not tried again by itself."
                : $"{version} (CRT release [{tag}]) did not come up ({state}), and {running} did not come back either: " +
                  $"{back}. THE SERVICE IS DOWN - see its journal (journalctl -u {this.thisOptions.Service}). A build past a database " +
                  "migration cannot go back: the build before refuses the newer database.");
        }

        // ###########################################################################################
        // Asks /api/health until it answers with `version`: null then, else what it answered last. Gives up
        // after HealthWait, or as soon as systemd says the service failed.
        // ###########################################################################################
        private async Task<string?> WaitForVersionAsync(string version, CancellationToken cancellationToken)
        {
            DateTimeOffset until = this.thisHost.UtcNow + ServerUpdater.HealthWait;
            string? answered = null;

            while (true)
            {
                answered = await this.thisHost.HealthVersionAsync(cancellationToken);

                if (string.Equals(answered?.Trim(), version, StringComparison.OrdinalIgnoreCase))
                    return null;

                if (this.thisHost.UtcNow >= until)
                    return answered == null ? "it did not answer" : $"it answered as {answered}";

                if ((await this.thisHost.ServiceStateAsync(this.thisOptions.Service, cancellationToken)).StartsWith("failed", StringComparison.Ordinal))
                    return "it stopped";

                await this.thisHost.DelayAsync(ServerUpdater.HealthPoll, cancellationToken);
            }
        }

        // ###########################################################################################
        // The systemd and polkit files of this instance the build in `newBuild` ships changed from those of
        // `runningBuild` and that are not installed yet - empty when the build may go in. Their names.
        // ###########################################################################################
        public IReadOnlyList<string> SystemFilesWaiting(string newBuild, string runningBuild)
        {
            if (File.Exists(Path.Combine(newBuild, ServerUpdater.SystemFilesInstalledMarker)))
                return [];

            var waiting = new List<string>();

            foreach (string name in ServerUpdater.SystemFilesOf(this.thisOptions.Service))
            {
                byte[]? shipped = ServerUpdater.ReadOrNull(Path.Combine(newBuild, name));

                if (shipped == null || ServerUpdater.SameBytes(shipped, ServerUpdater.ReadOrNull(Path.Combine(runningBuild, name))))
                    continue;

                if (ServerUpdater.SameBytes(shipped, this.thisHost.ReadInstalledCopy(name)))
                    continue;

                waiting.Add(name);
            }

            return waiting;
        }

        // ###########################################################################################
        // The files root installs for the instance whose service is `service` ("crt-server.service"): its unit,
        // its updater's unit and timer ("crt-server-update.service", ".timer"), and the polkit rule.
        // ###########################################################################################
        public static IReadOnlyList<string> SystemFilesOf(string service)
        {
            string name = service.EndsWith(".service", StringComparison.Ordinal) ? service[..^".service".Length] : service;

            return [$"{name}.service", $"{name}-update.service", $"{name}-update.timer", ServerUpdater.PolkitRuleFile];
        }

        // The builds kept: the one running, the one it replaced, and the newest others up to OlderBuildsKept.
        private void KeepFewBuilds(string running, string replaced)
        {
            var builds = Directory.GetDirectories(this.ReleasesFolder)
                .Select(Path.GetFileName)
                .OfType<string>()
                .Where(ServerReleaseInfo.IsFolderSafeVersion)
                .Where(name => !string.Equals(name, running, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(name => name, Comparer<string>.Create((a, b) => ServerUpdater.Compare(a, b)))
                .ToList();

            var kept = new List<string> { replaced };
            kept.AddRange(builds.Where(name => !string.Equals(name, replaced, StringComparison.OrdinalIgnoreCase)).Take(ServerUpdater.OlderBuildsKept - 1));

            foreach (string old in builds.Where(name => !kept.Contains(name, StringComparer.OrdinalIgnoreCase)))
            {
                try
                {
                    ServerUpdater.DeleteFolder(Path.Combine(this.ReleasesFolder, old));
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    this.thisLog.WriteLine($"The old build {old} could not be deleted - {ex.Message}");
                }
            }
        }

        // Refused for good: the same build is not downloaded again, only said again (AgainAsync).
        private bool Refuse(ServerUpdateStatus status, ServerReleaseFeed.Candidate candidate, string? version, string message)
        {
            status.ExaminedAssetId = candidate.ArchiveAssetId;
            status.ExaminedVersion = version;
            return this.Finish(status, ServerUpdateOutcomes.Refused, message);
        }

        // ###########################################################################################
        // Writes what this look came to. `quiet`: a minute's look saying again what an earlier one said - not
        // written to the journal, which would otherwise gain a line a minute per instance.
        // ###########################################################################################
        private bool Finish(ServerUpdateStatus status, string outcome, string message, bool quiet = false)
        {
            status.Outcome = outcome;
            status.Message = message;

            if (!quiet)
                this.thisLog.WriteLine(message);

            try
            {
                if (Directory.Exists(this.ReleasesFolder))
                    status.Write(this.StatusPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                this.thisLog.WriteLine($"The status could not be written to [{this.StatusPath}] - {ex.Message}");
            }

            return !ServerUpdateOutcomes.NeedsALook(outcome) || outcome == ServerUpdateOutcomes.Held;
        }

        private static string UpToDateSentence(string newest, string running, string tag) =>
            string.Equals(newest, running, StringComparison.OrdinalIgnoreCase)
                ? $"Up to date: {running} is the newest server build (CRT release [{tag}])."
                : $"Up to date: {running} runs, and the newest server build, {newest} (CRT release [{tag}]), is older - never installed backwards.";

        // SemVer order, a version that does not parse lowest.
        public static int Compare(string? a, string? b)
        {
            bool aParses = CrtVersion.TryParse(a, out CrtVersion? first);
            bool bParses = CrtVersion.TryParse(b, out CrtVersion? second);

            return (aParses, bParses) switch
            {
                (true, true) => first!.CompareTo(second!),
                (true, false) => 1,
                (false, true) => -1,
                _ => 0,
            };
        }

        private static FileStream? TryLock(string path)
        {
            try
            {
                return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return null;
            }
        }

        private static byte[]? ReadOrNull(string path) => File.Exists(path) ? File.ReadAllBytes(path) : null;

        // ###########################################################################################
        // Two systemd or polkit files the same, line ends aside: a build published by hand on Windows ships
        // them with "\r\n", every build of the release workflow with "\n" - the same file all the same, and
        // a build held for that alone would wait for nothing.
        // ###########################################################################################
        private static bool SameBytes(byte[]? a, byte[]? b) =>
            a != null && b != null && ServerUpdater.WithoutCarriageReturns(a).SequenceEqual(ServerUpdater.WithoutCarriageReturns(b));

        private static byte[] WithoutCarriageReturns(byte[] bytes) => bytes.Where(b => b != (byte)'\r').ToArray();

        private static void DeleteFolder(string folder)
        {
            if (Directory.Exists(folder))
                Directory.Delete(folder, recursive: true);
        }
    }

    // What `gh attestation verify` said: genuine with the commit it was built from, or why not.
    internal sealed record AttestationCheck(bool Genuine, string? Commit, string Reason);

    // ###########################################################################################
    // Everything the updater does outside its own folder - GitHub, the gh tool, systemd, links. The real
    // one is ServerUpdateHost; the tests' is a fake.
    // ###########################################################################################
    internal interface IServerUpdateHost
    {
        DateTimeOffset UtcNow { get; }

        Task<string> ReadReleasesAsync(CancellationToken cancellationToken);

        // The releases feed's fingerprint (ServerReleaseFeed.FingerprintOf), or null when it could not be read.
        Task<string?> ReadFeedFingerprintAsync(CancellationToken cancellationToken);

        Task DownloadAsync(string url, string toPath, long maxBytes, CancellationToken cancellationToken);

        Task<AttestationCheck> VerifyAsync(string archivePath, string bundlePath, CancellationToken cancellationToken);

        // The build "releases/current" points at (its folder's name), or null when it is no link.
        string? PointedAt(string releasesFolder);

        // Points "releases/current" at `version`'s folder, in one step.
        void PointAt(string releasesFolder, string version);

        // Links the settings file into a build's folder.
        void LinkSettings(string buildFolder, string settingsPath);

        // Restarts the service: null when done, else why not.
        Task<string?> RestartAsync(string service, CancellationToken cancellationToken);

        // The version /api/health answers with, or null when it does not answer.
        Task<string?> HealthVersionAsync(CancellationToken cancellationToken);

        // What systemd says of the service, in words ("active", "failed (exit 78)").
        Task<string> ServiceStateAsync(string service, CancellationToken cancellationToken);

        Task DelayAsync(TimeSpan delay, CancellationToken cancellationToken);

        // The installed copy of a systemd or polkit file, or null when it is not there or cannot be read.
        byte[]? ReadInstalledCopy(string fileName);
    }
}
