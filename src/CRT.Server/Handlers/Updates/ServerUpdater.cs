using System.Security.Cryptography;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // *** THE SERVER UPDATES ITSELF FROM GITHUB (owner request, 2026-11-09: "would it be possible to
    // automatically update the server from GitHub? So, when there is a new 'real' version of the server,
    // then its server automatically update to this, and when there is a new alpha/beta server available,
    // then that server will update to that? The goal would be for me to not manually update this, copying
    // files and restarting service and checking if it is alive ... This should of course be as secure as
    // possible"; and the same day: "the REAL version should work on a release ... But ALPHA/BETA works on
    // commit", "can it then update every minute?"). ***
    //
    // One look a minute, run by a systemd timer as a user of its own (ServerUpdateCommand, INSTALLING.md
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
    // A look: the channel's pointer (ServerReleaseFeed - a few lines on GitHub naming the newest build: a real
    // release's for the live service, main's for the test instance) -> nothing more when it names the build
    // already looked at, or no newer version -> the build downloaded -> GitHub's attestation checked by `gh
    // attestation verify`, pinned to this repository, the channel's workflow, main and a GitHub-hosted runner
    // - nothing unproven is even unpacked -> unpacked into a folder of its own (ServerArchive) -> its own
    // crt-server-release.json says its version and whether it is a pre-release's -> never older than the
    // build running, never a pre-release's on the live service, never a build that failed before -> held
    // while it ships a systemd or polkit file that differs from the one installed (those need root, and the
    // project owner) -> "current" pointed at it, the service restarted, and /api/health asked until it
    // answers with the new version -> otherwise back to the build before.
    //
    // The journal hears of a look only when it has something new to say - a minute saying again what the
    // last said writes nothing there (the status file notes every look's time all the same).
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

        // A build GitHub's attestation check refused is looked at again after this long - the check may have failed only
        // for not reaching Sigstore (code review, 2026-11-11). An hour: a forged build costs a download an hour, no more.
        public static readonly TimeSpan AttestationRetryAfter = TimeSpan.FromHours(1);

        // Builds kept besides the one running: the one it replaced, and one more.
        public const int OlderBuildsKept = 2;

        public sealed record Options(
            string Channel,
            string InstanceFolder,
            string Service,
            bool DryRun = false,
            bool RetryFailed = false);

        private readonly Options thisOptions;
        private readonly IServerUpdateHost thisHost;
        private readonly TextWriter thisLog;

        // What the last look said - a look saying it again says nothing in the journal.
        private (string Outcome, string Message) thisSaidBefore;

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
        // One look. True when nothing new went wrong (up to date, installed, held, a dry run, or a failure said
        // before); false when a build failed or was refused at this look, or GitHub could not be asked -
        // systemd then marks the run failed, and the status says why.
        // ###########################################################################################
        public async Task<bool> RunAsync(CancellationToken cancellationToken)
        {
            if (!this.IsSetUp(out _))
            {
                return this.Finish(this.ReadStatus(), ServerUpdateOutcomes.Error, this.NotSetUpSentence());
            }

            Directory.CreateDirectory(this.WorkFolder);

            using FileStream? held = ServerUpdater.TryLock(Path.Combine(this.WorkFolder, "lock"));

            if (held == null)
            {
                this.thisLog.WriteLine("Another look is running - this one does nothing.");
                return true;
            }

            // The status and the build running are read only once the lock is held (code review, 2026-11-11): read before it,
            // a look taking the lock just as another let go of it went by what that one found before it installed - and
            // unpacked the build just installed again over the folder the service runs from.
            ServerUpdateStatus status = this.ReadStatus();

            if (!this.IsSetUp(out string? running))
            {
                return this.Finish(status, ServerUpdateOutcomes.Error, this.NotSetUpSentence());
            }

            string? pointerJson;

            try
            {
                pointerJson = await this.thisHost.ReadPointerAsync(ServerReleaseFeed.PointerUrlFor(this.thisOptions.Channel), cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                return this.Finish(status, ServerUpdateOutcomes.Error, $"GitHub could not be asked for the newest server build - {ex.Message}");
            }

            if (pointerJson == null)
            {
                status.NewestVersion = null;
                status.NewestRelease = null;

                return this.Finish(status, ServerUpdateOutcomes.UpToDate, this.thisOptions.Channel == ServerUpdateOutcomes.MainChannel
                    ? $"No build of main is published yet. {running} runs."
                    : $"The newest CRT release carries no server build yet. {running} runs.");
            }

            if (ServerBuildPointer.Read(pointerJson, this.thisOptions.Channel, out string why) is not { } pointer)
                return this.Finish(status, ServerUpdateOutcomes.Error, why);

            status.NewestRelease = pointer.Tag;

            // The build looked at last time: nothing is downloaded again for it. A dry run always looks afresh.
            if (!this.thisOptions.DryRun &&
                string.Equals(pointer.Sha256, status.ExaminedSha256, StringComparison.Ordinal) &&
                await this.AgainAsync(status, pointer, running, cancellationToken) is bool done)
            {
                return done;
            }

            // ###########################################################################################
            // A build that says it is no newer is not even downloaded - a new commit that left the server's version
            // as it was, mostly. The pointer is not proven, but all it can do here is spare a download.
            // ###########################################################################################
            if (!this.thisOptions.DryRun && !this.thisOptions.RetryFailed && ServerUpdater.Compare(pointer.ServerVersion, running) <= 0)
            {
                status.ExaminedSha256 = pointer.Sha256;
                status.ExaminedVersion = pointer.ServerVersion;
                status.NewestVersion = pointer.ServerVersion;
                return this.Finish(status, ServerUpdateOutcomes.UpToDate, ServerUpdater.UpToDateSentence(pointer.ServerVersion, running, pointer.Source));
            }

            return await this.ExamineAsync(status, pointer, running, cancellationToken);
        }

        // What the last look wrote, this look's channel and time on it - and what it said kept, so the same said again
        // stays out of the journal.
        private ServerUpdateStatus ReadStatus()
        {
            ServerUpdateStatus status = ServerUpdateStatus.Read(this.StatusPath) ?? new ServerUpdateStatus();
            this.thisSaidBefore = (status.Outcome, status.Message);
            status.Channel = this.thisOptions.Channel;
            status.CheckedUtc = this.thisHost.UtcNow;

            return status;
        }

        // Whether "releases/current" links to a build in the releases folder - `running`, that build's folder name.
        private bool IsSetUp([System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? running)
        {
            running = this.thisHost.PointedAt(this.ReleasesFolder);

            return running != null && ServerReleaseInfo.IsFolderSafeVersion(running) && Directory.Exists(Path.Combine(this.ReleasesFolder, running));
        }

        private string NotSetUpSentence() =>
            $"[{Path.Combine(this.ReleasesFolder, ServerUpdater.CurrentLinkName)}] is not a link to a build in [{this.ReleasesFolder}] - " +
            "set the folders up as src/CRT.Server/INSTALLING.md's \"Automatic updates from GitHub\" says.";

        // ###########################################################################################
        // The same build as the last look: what that look came to stands - unless it was held and may go
        // in now, or a retry was asked for. Null: look at it afresh (download it again).
        // ###########################################################################################
        private async Task<bool?> AgainAsync(ServerUpdateStatus status, ServerBuildPointer pointer, string running, CancellationToken cancellationToken)
        {
            string? seen = ServerReleaseInfo.IsFolderSafeVersion(status.ExaminedVersion) ? status.ExaminedVersion : null;

            switch (status.Outcome)
            {
                case ServerUpdateOutcomes.Failed:
                case ServerUpdateOutcomes.Refused:
                    // Said again as it was said - still the newest build, still not installed - and no new failure:
                    // systemd was told once. A refusal may have been gh not reaching Sigstore, so a retry looks at
                    // that build again too - and so does a look an hour after the attestation check refused it.
                    if (this.thisOptions.RetryFailed)
                        return null;

                    if (status.Outcome == ServerUpdateOutcomes.Refused &&
                        status.AttestationRefusedUtc is { } refused &&
                        this.thisHost.UtcNow - refused >= ServerUpdater.AttestationRetryAfter)
                    {
                        return null;
                    }

                    this.Finish(status, status.Outcome, status.Message);
                    return true;

                case ServerUpdateOutcomes.Held when seen != null && ServerUpdater.Compare(seen, running) > 0 &&
                                                    Directory.Exists(Path.Combine(this.ReleasesFolder, seen)):
                    // Its files may be installed by now - tried every minute, from disk. Never over a build as new or
                    // newer running since (code review, 2026-11-11: one put in by hand went back to the held, older one).
                    status.NewestVersion = seen;
                    return await this.InstallAsync(status, pointer.Source, seen, running, cancellationToken);
            }

            if (seen != null && ServerUpdater.Compare(seen, running) <= 0)
            {
                status.NewestVersion = seen;
                return this.Finish(status, ServerUpdateOutcomes.UpToDate, ServerUpdater.UpToDateSentence(seen, running, pointer.Source));
            }

            return null;
        }

        // ###########################################################################################
        // A build not looked at before: downloaded, proven, unpacked, read - and installed when it is newer.
        // ###########################################################################################
        private async Task<bool> ExamineAsync(ServerUpdateStatus status, ServerBuildPointer pointer, string running, CancellationToken cancellationToken)
        {
            string incoming = Path.Combine(this.WorkFolder, "incoming");
            ServerUpdater.DeleteFolder(incoming);
            Directory.CreateDirectory(incoming);

            string archive = Path.Combine(incoming, ServerReleaseFeed.ArchiveName);
            string bundle = Path.Combine(incoming, ServerReleaseFeed.BundleName);
            string unpacked = Path.Combine(incoming, "build");
            string source = pointer.Source;

            try
            {
                this.thisLog.WriteLine($"Downloading the server build of {source}...");

                try
                {
                    await this.thisHost.DownloadAsync(ServerReleaseFeed.DownloadUrl(this.thisOptions.Channel, pointer.Tag, ServerReleaseFeed.ArchiveName), archive, ServerReleaseFeed.ArchiveMaxBytes, cancellationToken);
                    await this.thisHost.DownloadAsync(ServerReleaseFeed.DownloadUrl(this.thisOptions.Channel, pointer.Tag, ServerReleaseFeed.BundleName), bundle, ServerReleaseFeed.BundleMaxBytes, cancellationToken);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    return this.Finish(status, ServerUpdateOutcomes.Error, $"The server build of {source} could not be downloaded - {ex.Message}. The next minute tries again.");
                }

                // The files of main's release are replaced at every push: one replaced between the pointer and the
                // download is no build yet - the next minute finds them together.
                if (!string.Equals(ServerUpdater.Sha256Of(archive), pointer.Sha256, StringComparison.Ordinal))
                {
                    return this.Finish(status, ServerUpdateOutcomes.Error,
                        $"The server build of {source} changed while it was downloaded - a newer one was being put up. The next minute tries again.");
                }

                // *** NOTHING UNPROVEN IS UNPACKED. *** A gh that cannot be run at all proves nothing either - refused as
                // the check refuses, so looked at again an hour on (code review, 2026-11-11: the exception ended the look
                // unrecorded, and every minute downloaded the build again).
                AttestationCheck proof;

                try
                {
                    proof = await this.thisHost.VerifyAsync(archive, bundle, this.thisOptions.Channel, cancellationToken);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    proof = new AttestationCheck(false, null, $"gh could not be run - {ex.Message}");
                }

                if (!proof.Genuine)
                {
                    return this.Refuse(status, pointer, null,
                        $"The server build of {source} was NOT installed: GitHub's attestation does not prove this repository's " +
                        $"workflow built it - {proof.Reason}", byAttestation: true);
                }

                try
                {
                    ServerArchive.Extract(archive, unpacked);
                }
                catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
                {
                    return this.Refuse(status, pointer, null, $"The server build of {source} could not be unpacked - {ex.Message}");
                }

                string infoPath = Path.Combine(unpacked, ServerReleaseInfo.FileName);
                ServerReleaseInfo? info = ServerReleaseInfo.Read(File.Exists(infoPath) ? File.ReadAllText(infoPath) : null, out string why);

                if (info == null)
                    return this.Refuse(status, pointer, null, $"The server build of {source} was not installed - {why}");

                if (!string.IsNullOrEmpty(proof.Commit) && !string.Equals(proof.Commit, info.Commit, StringComparison.OrdinalIgnoreCase))
                {
                    return this.Refuse(status, pointer, null,
                        $"The server build of {source} was not installed: it says it was built from [{info.Commit}], its attestation [{proof.Commit}].");
                }

                if (info.PreRelease && this.thisOptions.Channel == ServerUpdateOutcomes.ReleasesChannel)
                {
                    return this.Refuse(status, pointer, info.ServerVersion,
                        $"The server build of CRT [{info.CrtVersion}] is a pre-release's, and this service installs real releases only.");
                }

                status.NewestVersion = info.ServerVersion;
                status.ExaminedSha256 = pointer.Sha256;
                status.ExaminedVersion = info.ServerVersion;

                if (ServerUpdater.Compare(info.ServerVersion, running) <= 0)
                    return this.Finish(status, ServerUpdateOutcomes.UpToDate, ServerUpdater.UpToDateSentence(info.ServerVersion, running, source));

                if (string.Equals(status.FailedVersion, info.ServerVersion, StringComparison.OrdinalIgnoreCase) && !this.thisOptions.RetryFailed)
                {
                    return this.Finish(status, ServerUpdateOutcomes.Failed,
                        $"{info.ServerVersion} did not come up before, so it is not tried again by itself - a newer build is, " +
                        $"or `update-server --retry-failed` tries this one again. {running} runs.");
                }

                // Into a folder of its own, beside the others - never over the one running, which is older.
                string target = Path.Combine(this.ReleasesFolder, info.ServerVersion);
                bool markerStays = this.MarkerCarriesOver(target, unpacked);
                ServerUpdater.DeleteFolder(target);
                Directory.Move(unpacked, target);
                this.thisHost.LinkSettings(target, this.SettingsPath);

                if (markerStays)
                    File.WriteAllText(Path.Combine(target, ServerUpdater.SystemFilesInstalledMarker), string.Empty);

                return await this.InstallAsync(status, source, info.ServerVersion, running, cancellationToken);
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
        private async Task<bool> InstallAsync(ServerUpdateStatus status, string source, string version, string running, CancellationToken cancellationToken)
        {
            string target = Path.Combine(this.ReleasesFolder, version);
            IReadOnlyList<string> waiting = this.SystemFilesWaiting(target, Path.Combine(this.ReleasesFolder, running));

            if (waiting.Count > 0)
                return this.Finish(status, ServerUpdateOutcomes.Held, ServerUpdater.HeldForSystemFiles(version, source, running, target, waiting));

            if (this.thisOptions.DryRun)
            {
                // Looked at afresh next time: a dry run installs nothing.
                status.ExaminedSha256 = null;
                return this.Finish(status, ServerUpdateOutcomes.WouldInstall, $"{version} ({source}) would be installed - a dry run, so {running} runs on.");
            }

            this.thisLog.WriteLine($"Installing {version} in place of {running}...");
            this.thisHost.PointAt(this.ReleasesFolder, version);

            if (await this.thisHost.RestartAsync(this.thisOptions.Service, cancellationToken) is string refused)
            {
                // Not restarted, so the build before still runs - pointed back at it, and tried again at the next look.
                this.thisHost.PointAt(this.ReleasesFolder, running);

                return this.Finish(status, ServerUpdateOutcomes.Held,
                    $"{version} is ready, but [{this.thisOptions.Service}] could not be restarted - {refused}. Is the polkit rule " +
                    $"{ServerUpdater.PolkitRuleFile} installed? {running} runs meanwhile; the next minute tries again.");
            }

            if (await this.WaitForVersionAsync(version, cancellationToken) is null)
            {
                status.InstalledUtc = this.thisHost.UtcNow;
                status.InstalledVersion = version;
                status.InstalledFromVersion = running;
                status.FailedVersion = null;
                this.KeepFewBuilds(version, running);

                return this.Finish(status, ServerUpdateOutcomes.Installed, $"Installed {version} ({source}) in place of {running}.");
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
                ? $"{version} ({source}) did not come up ({state}), so {running} runs again. " +
                  "Its journal says why (journalctl -u " + this.thisOptions.Service + "). It is not tried again by itself."
                : $"{version} ({source}) did not come up ({state}), and {running} did not come back either: " +
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
        // Whether the project owner's marker in `heldBuild` - a build of the same version, about to be replaced by the
        // next build of it (main's next push) - still says what it said: true when it is there and the build in
        // `newBuild` ships every one of this instance's systemd and polkit files exactly as the held build did (code
        // review, 2026-11-11: deleted with the old files, the build was held again for the very same files).
        // ###########################################################################################
        private bool MarkerCarriesOver(string heldBuild, string newBuild)
        {
            if (!File.Exists(Path.Combine(heldBuild, ServerUpdater.SystemFilesInstalledMarker)))
                return false;

            return ServerUpdater.SystemFilesOf(this.thisOptions.Service).All(name =>
            {
                byte[]? held = ServerUpdater.ReadOrNull(Path.Combine(heldBuild, name));
                byte[]? shipped = ServerUpdater.ReadOrNull(Path.Combine(newBuild, name));

                return held is null ? shipped is null : ServerUpdater.SameBytes(held, shipped);
            });
        }

        // ###########################################################################################
        // What a build held for its systemd or polkit files says: the very commands that install them. A unit
        // is installed READABLE TO EVERYBODY (owner's box, 2026-11-09: copied with `cp` it came out 0640
        // root:root - the build's own mode - and this updater, which reads it as its own user to see it is in,
        // held 10.3.1 on with the unit installed). The polkit rule's folder is root's alone, so a rule
        // installed there is never seen from here: the marker says it is in.
        // ###########################################################################################
        internal static string HeldForSystemFiles(string version, string source, string running, string target, IReadOnlyList<string> waiting)
        {
            string[] units = waiting.Where(name => name != ServerUpdater.PolkitRuleFile).ToArray();
            string message = $"{version} ({source}) waits for you: it ships {string.Join(", ", waiting)} changed, which only root installs.";

            if (units.Length > 0)
            {
                message += $" Install {(units.Length == 1 ? "it" : "them")} readable to everybody, as this updater reads " +
                    $"{(units.Length == 1 ? "it" : "them")} to see {(units.Length == 1 ? "it is" : "they are")} in: sudo install -m 0644 " +
                    string.Join(" ", units.Select(name => $"{target}/{name}")) + " /etc/systemd/system/ && sudo systemctl daemon-reload.";
            }

            if (waiting.Contains(ServerUpdater.PolkitRuleFile))
            {
                message += $" The polkit rule: sudo install -m 0644 {target}/{ServerUpdater.PolkitRuleFile} /etc/polkit-1/rules.d/ - " +
                    $"its folder is root's, so this updater cannot see it there: then sudo touch {target}/{ServerUpdater.SystemFilesInstalledMarker}.";
            }

            return message + $" The next minute installs it. {running} runs meanwhile.";
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

        // Refused: the same build is not downloaded again, only said again (AgainAsync) - for good, unless it was the
        // attestation check (`byAttestation`), which may have failed only for not reaching Sigstore: looked at again
        // once AttestationRetryAfter has passed.
        private bool Refuse(ServerUpdateStatus status, ServerBuildPointer pointer, string? version, string message, bool byAttestation = false)
        {
            status.ExaminedSha256 = pointer.Sha256;
            status.ExaminedVersion = version;
            status.AttestationRefusedUtc = byAttestation ? this.thisHost.UtcNow : null;
            return this.Finish(status, ServerUpdateOutcomes.Refused, message);
        }

        // ###########################################################################################
        // Writes what this look came to - to the journal only when it says something the last look did not, as
        // a look every minute would otherwise write a line a minute per instance.
        // ###########################################################################################
        private bool Finish(ServerUpdateStatus status, string outcome, string message)
        {
            status.Outcome = outcome;
            status.Message = message;

            if (outcome != ServerUpdateOutcomes.Refused)
                status.AttestationRefusedUtc = null;

            if ((outcome, message) != this.thisSaidBefore)
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

        private static string UpToDateSentence(string newest, string running, string source) =>
            string.Equals(newest, running, StringComparison.OrdinalIgnoreCase)
                ? $"Up to date: {running} is the newest server build ({source})."
                : $"Up to date: {running} runs, and the newest server build, {newest} ({source}), is not newer - never installed backwards.";

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

        // The SHA-256 of a file, lower-case hex - as the workflows write it into the pointer with sha256sum.
        private static string Sha256Of(string path)
        {
            using FileStream file = File.OpenRead(path);
            return Convert.ToHexString(SHA256.HashData(file)).ToLowerInvariant();
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
        // them with "\r\n", every build of a workflow with "\n" - the same file all the same, and a build held
        // for that alone would wait for nothing.
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

        // The channel's pointer (ServerReleaseFeed.PointerUrlFor), or null when GitHub has none (404). Throws
        // when GitHub cannot be asked.
        Task<string?> ReadPointerAsync(string url, CancellationToken cancellationToken);

        Task DownloadAsync(string url, string toPath, long maxBytes, CancellationToken cancellationToken);

        // gh attestation verify, held to `channel`'s workflow (ServerUpdateHost.VerifyArguments).
        Task<AttestationCheck> VerifyAsync(string archivePath, string bundlePath, string channel, CancellationToken cancellationToken);

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
