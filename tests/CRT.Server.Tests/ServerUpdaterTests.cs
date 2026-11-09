using System.Security.Cryptography;
using System.Text.Json;
using CRT.Server.Handlers.Updates;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // THE SERVER UPDATES ITSELF FROM GITHUB (owner requests, 2026-11-09: "The goal would be for me to not
    // manually update this, copying files and restarting service and checking if it is alive ... This
    // should of course be as secure as possible"; "the REAL version should work on a release ... But
    // ALPHA/BETA works on commit"; "can it then update every minute?"; ServerUpdater).
    //
    // Every path of one look, through the real updater on a real instance folder, with a fake host for
    // GitHub, gh, systemd and the links: a newer build proven genuine goes in and the service answers as it;
    // nothing unproven is unpacked; nothing older, nothing a pre-release's on the live service, nothing that
    // failed before goes in by itself; a build that does not come up is rolled back; a build whose systemd
    // or polkit files changed waits for the project owner; the same build, or one of the same version, is not
    // downloaded again; a minute with nothing new writes nothing to the journal.
    // ###########################################################################################
    public sealed class ServerUpdaterTests : IDisposable
    {
        private const string Running = "10.2.1";

        private readonly string thisInstance = Path.Combine(Path.GetTempPath(), "crt-updater-" + Guid.NewGuid().ToString("N"));
        private readonly FakeHost thisHost = new();

        public ServerUpdaterTests()
        {
            // The build running now, set up as INSTALLING.md says.
            this.AddBuild(ServerUpdaterTests.Running);
            this.thisHost.Current = ServerUpdaterTests.Running;
            this.thisHost.Answering = ServerUpdaterTests.Running;
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisInstance, recursive: true);
            }
            catch (IOException)
            {
                // A temp folder left behind is harmless.
            }
        }

        private string Releases => Path.Combine(this.thisInstance, ServerUpdater.ReleasesFolderName);

        private void AddBuild(string version, Dictionary<string, string>? files = null) =>
            ServerBuilds.Unpack(Path.Combine(this.Releases, version), files ?? ServerBuilds.Files(version));

        private ServerUpdater Updater(string channel = ServerUpdateOutcomes.ReleasesChannel, bool dryRun = false, bool retryFailed = false, TextWriter? log = null) =>
            new(new ServerUpdater.Options(channel, this.thisInstance, "crt-server.service", dryRun, retryFailed), this.thisHost, log ?? TextWriter.Null);

        private ServerUpdateStatus Status => ServerUpdateStatus.Read(Path.Combine(this.Releases, ServerUpdateStatus.FileName))!;

        // ###########################################################################################
        // A build put on GitHub as the workflows put it: the archive and its bundle on the release, and the
        // pointer naming it - a real release's (the newest release's pointer) or main's (the standing
        // pre-release's). `pointerVersion` lets a pointer say another version than the build does.
        // ###########################################################################################
        private void Publish(string tag, byte[] build, string serverVersion, string channel = ServerUpdateOutcomes.ReleasesChannel, string? pointerSha = null)
        {
            this.thisHost.Files[ServerReleaseFeed.DownloadUrl(tag, ServerReleaseFeed.ArchiveName)] = build;
            this.thisHost.Files[ServerReleaseFeed.DownloadUrl(tag, ServerReleaseFeed.BundleName)] = "{\"bundle\":true}"u8.ToArray();
            this.thisHost.Pointers[ServerReleaseFeed.PointerUrlFor(channel)] = JsonSerializer.Serialize(new
            {
                tag,
                serverVersion,
                commit = ServerBuilds.Commit,
                sha256 = pointerSha ?? Convert.ToHexString(SHA256.HashData(build)).ToLowerInvariant(),
            });
        }

        private void PublishRelease(string tag, string serverVersion, bool preRelease = false) =>
            this.Publish(tag, ServerBuilds.Build(serverVersion, preRelease), serverVersion);

        private void PublishMain(string serverVersion, string commit = ServerBuilds.Commit) =>
            this.Publish(ServerReleaseFeed.MainTag, ServerBuilds.Build(serverVersion, preRelease: true, commit: commit), serverVersion, ServerUpdateOutcomes.MainChannel);

        [Fact]
        public async Task A_newer_build_proven_genuine_is_installed_and_the_service_answers_as_it()
        {
            this.PublishRelease("4.1.0", "10.3.0");

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal("10.3.0", this.thisHost.Current);
            Assert.Equal(["crt-server.service"], this.thisHost.Restarted);
            Assert.True(File.Exists(Path.Combine(this.Releases, "10.3.0", "CRT.Server.dll")));
            Assert.True(File.Exists(Path.Combine(this.Releases, "10.3.0", "Migrations", "0001_first.sql")));
            Assert.Equal([ServerUpdateOutcomes.ReleasesChannel], this.thisHost.VerifiedFor);

            // Its settings are the box's own, linked in - never in a download.
            Assert.Equal(Path.Combine(this.thisInstance, "settings", ServerUpdater.SettingsFileName), this.thisHost.SettingsLinked[Path.Combine(this.Releases, "10.3.0")]);

            ServerUpdateStatus status = this.Status;
            Assert.Equal(ServerUpdateOutcomes.Installed, status.Outcome);
            Assert.Equal("10.3.0", status.InstalledVersion);
            Assert.Equal(ServerUpdaterTests.Running, status.InstalledFromVersion);
            Assert.Equal(ServerUpdateOutcomes.ReleasesChannel, status.Channel);
            Assert.Contains("CRT release [4.1.0]", status.Message, StringComparison.Ordinal);

            // Nothing is left on its way in.
            Assert.False(Directory.Exists(Path.Combine(this.Releases, ServerUpdater.WorkFolderName, "incoming")));
        }

        // ###########################################################################################
        // *** EVERY COMMIT TO MAIN, ON THE TEST INSTANCE (owner decision: "ALPHA/BETA works on commit"). *** Its
        // builds are a pre-release's and attested by the push workflow; the live service never even reads their
        // pointer, and the test instance installs them, checked against the push workflow.
        // ###########################################################################################
        [Fact]
        public async Task The_test_instance_installs_main_and_the_live_service_never_sees_it()
        {
            this.PublishMain("10.3.0");

            Assert.True(await this.Updater(ServerUpdateOutcomes.ReleasesChannel).RunAsync(CancellationToken.None));
            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(0, this.thisHost.Downloads);
            Assert.Contains("carries no server build", this.Status.Message, StringComparison.Ordinal);

            Assert.True(await this.Updater(ServerUpdateOutcomes.MainChannel).RunAsync(CancellationToken.None));
            Assert.Equal("10.3.0", this.thisHost.Current);
            Assert.Equal([ServerUpdateOutcomes.MainChannel], this.thisHost.VerifiedFor);
            Assert.Contains("main at [7f3a2c1]", this.Status.Message, StringComparison.Ordinal);
        }

        // A commit that left the server's version as it was: nothing downloaded - the minute costs one small file.
        [Fact]
        public async Task A_new_commit_that_left_the_server_version_as_it_was_downloads_nothing()
        {
            this.PublishMain(ServerUpdaterTests.Running, commit: "0123456789abcdef0123456789abcdef01234567");

            Assert.True(await this.Updater(ServerUpdateOutcomes.MainChannel).RunAsync(CancellationToken.None));

            Assert.Equal(0, this.thisHost.Downloads);
            Assert.Empty(this.thisHost.Restarted);
            Assert.Equal(ServerUpdateOutcomes.UpToDate, this.Status.Outcome);
        }

        // ###########################################################################################
        // *** NOTHING UNPROVEN IS EVEN UNPACKED. *** The attestation not proving this repository's workflow built
        // it: no folder, no restart, refused - and not downloaded again every minute.
        // ###########################################################################################
        [Fact]
        public async Task A_build_its_attestation_does_not_prove_is_neither_unpacked_nor_installed()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.Proof = new AttestationCheck(false, null, "verifying with issuer \"sigstore.dev\"");

            Assert.False(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Empty(this.thisHost.Restarted);
            Assert.False(Directory.Exists(Path.Combine(this.Releases, "10.3.0")));
            Assert.Equal(ServerUpdateOutcomes.Refused, this.Status.Outcome);
            Assert.Contains("NOT installed", this.Status.Message, StringComparison.Ordinal);

            // Said again at the next minute - not downloaded again, and no new failure for systemd: it was told once.
            int downloads = this.thisHost.Downloads;
            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal(downloads, this.thisHost.Downloads);
            Assert.Equal(ServerUpdateOutcomes.Refused, this.Status.Outcome);
        }

        // The build says it came from another commit than its attestation does: refused.
        [Fact]
        public async Task A_build_whose_commit_is_not_its_attestations_is_refused()
        {
            this.Publish("4.1.0", ServerBuilds.Build("10.3.0", commit: "0000000000000000000000000000000000000000"), "10.3.0");

            Assert.False(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(ServerUpdateOutcomes.Refused, this.Status.Outcome);
        }

        // ###########################################################################################
        // The files of main's release are replaced at every push: a build that is not the one its pointer names
        // was being replaced - nothing is taken, and the next minute, with pointer and build together, installs.
        // ###########################################################################################
        [Fact]
        public async Task A_build_replaced_while_it_was_downloaded_is_not_taken()
        {
            this.Publish(ServerReleaseFeed.MainTag, ServerBuilds.Build("10.3.0", preRelease: true), "10.3.0", ServerUpdateOutcomes.MainChannel, pointerSha: new string('a', 64));

            Assert.False(await this.Updater(ServerUpdateOutcomes.MainChannel).RunAsync(CancellationToken.None));
            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Empty(this.thisHost.VerifiedFor);
            Assert.Contains("changed while it was downloaded", this.Status.Message, StringComparison.Ordinal);

            this.PublishMain("10.3.0");

            Assert.True(await this.Updater(ServerUpdateOutcomes.MainChannel).RunAsync(CancellationToken.None));
            Assert.Equal("10.3.0", this.thisHost.Current);
        }

        // A pointer naming another channel's release is no pointer: no URL is ever built from it.
        [Theory]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, ServerReleaseFeed.MainTag)]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, "4.1.0-beta.1")]
        [InlineData(ServerUpdateOutcomes.MainChannel, "4.1.0")]
        public async Task A_pointer_naming_another_channels_release_downloads_nothing(string channel, string tag)
        {
            this.Publish(tag, ServerBuilds.Build("10.3.0", preRelease: true), "10.3.0", channel);

            Assert.False(await this.Updater(channel).RunAsync(CancellationToken.None));

            Assert.Equal(0, this.thisHost.Downloads);
            Assert.Equal(ServerUpdateOutcomes.Error, this.Status.Outcome);
        }

        // ###########################################################################################
        // Never backwards, and nothing downloaded for it: a newer CRT release from an older branch carries an
        // older server. A pointer is not proven - but all a lying one could do here is spare a download.
        // ###########################################################################################
        [Fact]
        public async Task An_older_build_is_never_installed_nor_downloaded()
        {
            this.PublishRelease("4.0.2", "10.1.0");

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(0, this.thisHost.Downloads);
            Assert.Equal(ServerUpdateOutcomes.UpToDate, this.Status.Outcome);
            Assert.Contains("never installed backwards", this.Status.Message, StringComparison.Ordinal);
        }

        // A pointer that says "newer" about an older build: downloaded, and its own description decides.
        [Fact]
        public async Task A_pointer_overstating_its_version_is_overruled_by_the_build_itself()
        {
            this.Publish("4.1.0", ServerBuilds.Build("10.1.0"), serverVersion: "10.9.0");

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Empty(this.thisHost.Restarted);
            Assert.Equal("10.1.0", this.Status.NewestVersion);
        }

        // ###########################################################################################
        // The live service installs real releases only - read from the build's own crt-server-release.json,
        // which the attestation covers, never from GitHub's flags. A release carrying a pre-release's build is
        // refused there.
        // ###########################################################################################
        [Fact]
        public async Task A_pre_release_build_is_never_installed_on_the_live_service()
        {
            this.PublishRelease("4.1.0", "10.3.0", preRelease: true);

            Assert.False(await this.Updater(ServerUpdateOutcomes.ReleasesChannel).RunAsync(CancellationToken.None));
            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(ServerUpdateOutcomes.Refused, this.Status.Outcome);
        }

        // ###########################################################################################
        // *** A BUILD THAT DOES NOT COME UP IS ROLLED BACK - AND NOT TRIED AGAIN BY ITSELF. *** Tried at every
        // minute, it would take the service down every minute.
        // ###########################################################################################
        [Fact]
        public async Task A_build_that_does_not_come_up_is_rolled_back_and_not_tried_again_by_itself()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.DoesNotComeUp.Add("10.3.0");

            Assert.False(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(["crt-server.service", "crt-server.service"], this.thisHost.Restarted);
            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Answering);

            ServerUpdateStatus status = this.Status;
            Assert.Equal(ServerUpdateOutcomes.Failed, status.Outcome);
            Assert.Equal("10.3.0", status.FailedVersion);
            Assert.Contains("failed (exit 78)", status.Message, StringComparison.Ordinal);
            Assert.Contains($"{ServerUpdaterTests.Running} runs again", status.Message, StringComparison.Ordinal);

            int downloads = this.thisHost.Downloads;
            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal(downloads, this.thisHost.Downloads);
            Assert.Equal(2, this.thisHost.Restarted.Count);
            Assert.Equal(ServerUpdateOutcomes.Failed, this.Status.Outcome);
        }

        // "update-server --retry-failed": the project owner fixed what stopped it, and tries again.
        [Fact]
        public async Task Retry_failed_tries_that_build_again()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.DoesNotComeUp.Add("10.3.0");
            await this.Updater().RunAsync(CancellationToken.None);

            this.thisHost.DoesNotComeUp.Clear();

            Assert.True(await this.Updater(retryFailed: true).RunAsync(CancellationToken.None));

            Assert.Equal("10.3.0", this.thisHost.Current);
            Assert.Equal(ServerUpdateOutcomes.Installed, this.Status.Outcome);
            Assert.Null(this.Status.FailedVersion);
        }

        // A newer build after a failed one goes in by itself.
        [Fact]
        public async Task A_newer_build_after_a_failed_one_is_installed()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.DoesNotComeUp.Add("10.3.0");
            await this.Updater().RunAsync(CancellationToken.None);

            this.PublishRelease("4.1.1", "10.3.1");

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal("10.3.1", this.thisHost.Current);
        }

        // ###########################################################################################
        // The build before did not come back either - past a database migration it cannot: said as the
        // service being DOWN, never as rolled back.
        // ###########################################################################################
        [Fact]
        public async Task When_the_build_before_does_not_come_back_either_the_service_is_said_to_be_down()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.DoesNotComeUp.Add("10.3.0");
            this.thisHost.DoesNotComeUp.Add(ServerUpdaterTests.Running);

            Assert.False(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdateOutcomes.Failed, this.Status.Outcome);
            Assert.Contains("THE SERVICE IS DOWN", this.Status.Message, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** A BUILD THAT SHIPS A CHANGED SYSTEMD OR POLKIT FILE WAITS FOR THE PROJECT OWNER. *** Those need
        // root; the build is unpacked and kept, tried again every minute from disk, and goes in at the first
        // minute after they are installed - with no download again.
        // ###########################################################################################
        [Fact]
        public async Task A_build_with_a_changed_unit_waits_until_the_unit_is_installed()
        {
            Dictionary<string, string> files = ServerBuilds.Files("10.3.0");
            files["crt-server.service"] += "ReadWritePaths=/new/folder\n";
            this.Publish("4.1.0", ServerBuilds.Pack(files), "10.3.0");

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Empty(this.thisHost.Restarted);
            Assert.Equal(ServerUpdateOutcomes.Held, this.Status.Outcome);
            Assert.Contains("crt-server.service", this.Status.Message, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(this.Releases, "10.3.0", "crt-server.service")));

            // The command that installs it readable to everybody - copied with cp it came out root's alone on the
            // owner's box, and the updater, reading it as its own user, held the build on (2026-11-09).
            string build = Path.Combine(this.Releases, "10.3.0");
            Assert.Contains($"sudo install -m 0644 {build}/crt-server.service /etc/systemd/system/ && sudo systemctl daemon-reload",
                this.Status.Message, StringComparison.Ordinal);
            Assert.DoesNotContain(ServerUpdater.SystemFilesInstalledMarker, this.Status.Message, StringComparison.Ordinal);

            // Still waiting at the next minute, with no download again.
            int downloads = this.thisHost.Downloads;
            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal(ServerUpdateOutcomes.Held, this.Status.Outcome);

            // Installed by the project owner: the next minute installs the build.
            this.thisHost.Installed["crt-server.service"] = files["crt-server.service"];

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal(downloads, this.thisHost.Downloads);
            Assert.Equal("10.3.0", this.thisHost.Current);
            Assert.Equal(ServerUpdateOutcomes.Installed, this.Status.Outcome);
        }

        // A unit changed by hand on purpose: the marker says the installed files are the ones wanted.
        [Fact]
        public async Task The_marker_lets_a_held_build_in()
        {
            Dictionary<string, string> files = ServerBuilds.Files("10.3.0");
            files[ServerUpdater.PolkitRuleFile] += "// changed\n";
            this.Publish("4.1.0", ServerBuilds.Pack(files), "10.3.0");
            await this.Updater().RunAsync(CancellationToken.None);
            Assert.Equal(ServerUpdateOutcomes.Held, this.Status.Outcome);

            // The rule's folder is root's, so the updater never sees it installed: the message says to put the marker.
            string build = Path.Combine(this.Releases, "10.3.0");
            Assert.Contains($"sudo install -m 0644 {build}/{ServerUpdater.PolkitRuleFile} /etc/polkit-1/rules.d/", this.Status.Message, StringComparison.Ordinal);
            Assert.Contains($"sudo touch {build}/{ServerUpdater.SystemFilesInstalledMarker}", this.Status.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("/etc/systemd/system/", this.Status.Message, StringComparison.Ordinal);

            File.WriteAllText(Path.Combine(this.Releases, "10.3.0", ServerUpdater.SystemFilesInstalledMarker), string.Empty);

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal("10.3.0", this.thisHost.Current);
        }

        // Several units held at once: one command installing them all, said of "them".
        [Fact]
        public void Several_held_units_are_installed_by_one_command()
        {
            string message = ServerUpdater.HeldForSystemFiles("10.3.2", "main at [abc1234]", "10.3.1", "/srv/crt-server-test/releases/10.3.2",
                ["crt-server-test.service", "crt-server-test-update.timer"]);

            Assert.Equal(
                "10.3.2 (main at [abc1234]) waits for you: it ships crt-server-test.service, crt-server-test-update.timer changed, which only " +
                "root installs. Install them readable to everybody, as this updater reads them to see they are in: sudo install -m 0644 " +
                "/srv/crt-server-test/releases/10.3.2/crt-server-test.service /srv/crt-server-test/releases/10.3.2/crt-server-test-update.timer " +
                "/etc/systemd/system/ && sudo systemctl daemon-reload. The next minute installs it. 10.3.1 runs meanwhile.",
                message);
        }

        // The build running published by hand on Windows ("\r\n"), the next by a workflow ("\n"): no change, no hold.
        [Fact]
        public async Task Line_ends_alone_hold_no_build()
        {
            Dictionary<string, string> windows = ServerBuilds.Files(ServerUpdaterTests.Running);

            foreach (string name in ServerBuilds.SystemFiles().Keys)
                windows[name] = windows[name].Replace("\n", "\r\n", StringComparison.Ordinal);

            this.AddBuild(ServerUpdaterTests.Running, windows);
            this.PublishRelease("4.1.0", "10.3.0");

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal("10.3.0", this.thisHost.Current);
        }

        // The test instance's files changing does not hold the live service, and the other way round.
        [Fact]
        public void Only_the_instances_own_systemd_files_and_the_polkit_rule_can_hold_a_build()
        {
            Assert.Equal(
                ["crt-server.service", "crt-server-update.service", "crt-server-update.timer", ServerUpdater.PolkitRuleFile],
                ServerUpdater.SystemFilesOf("crt-server.service"));

            Assert.Equal(
                ["crt-server-test.service", "crt-server-test-update.service", "crt-server-test-update.timer", ServerUpdater.PolkitRuleFile],
                ServerUpdater.SystemFilesOf("crt-server-test.service"));
        }

        // The restart refused (no polkit rule): pointed back - nothing was restarted, so the build before runs - and tried again.
        [Fact]
        public async Task A_refused_restart_points_back_and_tries_again_at_the_next_minute()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.RestartRefusal = "Interactive authentication required.";

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(ServerUpdateOutcomes.Held, this.Status.Outcome);
            Assert.Contains(ServerUpdater.PolkitRuleFile, this.Status.Message, StringComparison.Ordinal);

            this.thisHost.RestartRefusal = null;
            int downloads = this.thisHost.Downloads;

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal("10.3.0", this.thisHost.Current);
            Assert.Equal(downloads, this.thisHost.Downloads);
        }

        [Fact]
        public async Task A_dry_run_proves_the_build_and_installs_nothing()
        {
            this.PublishRelease("4.1.0", "10.3.0");

            Assert.True(await this.Updater(dryRun: true).RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Empty(this.thisHost.Restarted);
            Assert.Single(this.thisHost.VerifiedFor);
            Assert.Equal(ServerUpdateOutcomes.WouldInstall, this.Status.Outcome);
        }

        // The setup's check: a dry run over the very build running still downloads and proves it - so it shows
        // gh and the attestation work on the box - and says it is up to date.
        [Fact]
        public async Task A_dry_run_over_the_build_running_proves_it_and_says_up_to_date()
        {
            this.PublishMain(ServerUpdaterTests.Running);

            Assert.True(await this.Updater(ServerUpdateOutcomes.MainChannel, dryRun: true).RunAsync(CancellationToken.None));

            Assert.Single(this.thisHost.VerifiedFor);
            Assert.Equal(ServerUpdateOutcomes.UpToDate, this.Status.Outcome);
            Assert.StartsWith($"Up to date: {ServerUpdaterTests.Running} is the newest server build (main at [", this.Status.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task The_same_build_is_not_downloaded_again_at_every_minute()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            await this.Updater().RunAsync(CancellationToken.None);
            int downloads = this.thisHost.Downloads;

            Assert.True(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(downloads, this.thisHost.Downloads);
            Assert.Equal(ServerUpdateOutcomes.UpToDate, this.Status.Outcome);
            Assert.Equal("10.3.0", this.Status.InstalledVersion);
        }

        // ###########################################################################################
        // A look every minute (owner request: "can it then update every minute?"): one that has nothing new to
        // say writes nothing to the journal - only its time to the status file.
        // ###########################################################################################
        [Fact]
        public async Task A_minute_saying_again_what_the_last_said_writes_nothing_to_the_journal()
        {
            this.PublishRelease("4.0.1", ServerUpdaterTests.Running);

            using var first = new StringWriter();
            await this.Updater(log: first).RunAsync(CancellationToken.None);
            Assert.Contains("Up to date", first.ToString(), StringComparison.Ordinal);
            DateTimeOffset looked = this.Status.CheckedUtc!.Value;

            this.thisHost.Advance(TimeSpan.FromMinutes(1));

            using var second = new StringWriter();
            Assert.True(await this.Updater(log: second).RunAsync(CancellationToken.None));
            Assert.Equal(string.Empty, second.ToString());
            Assert.True(this.Status.CheckedUtc > looked, "A quiet minute is a look all the same - its time is noted");
        }

        // ###########################################################################################
        // The builds kept: the one running, the one it replaced, and one more - older ones go, so the folder
        // does not grow with every release. Only folders named as versions are touched.
        // ###########################################################################################
        [Fact]
        public async Task Old_builds_go_and_the_two_before_the_running_one_stay()
        {
            this.AddBuild("10.0.0");
            this.AddBuild("10.1.0");
            this.AddBuild("10.2.0");
            Directory.CreateDirectory(Path.Combine(this.Releases, "not-a-build"));
            this.PublishRelease("4.1.0", "10.3.0");

            await this.Updater().RunAsync(CancellationToken.None);

            Assert.Equal(
                ["10.2.0", "10.2.1", "10.3.0", "not-a-build"],
                Directory.GetDirectories(this.Releases).Select(Path.GetFileName).Where(name => !name!.StartsWith('.')).Order(StringComparer.Ordinal));
        }

        // Not set up as INSTALLING.md says ("releases/current" no link to a build): nothing is done.
        [Fact]
        public async Task Without_its_folders_set_up_the_updater_does_nothing()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            this.thisHost.Current = null;

            Assert.False(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(0, this.thisHost.PointerReads);
            Assert.Equal(ServerUpdateOutcomes.Error, this.Status.Outcome);
            Assert.Contains("INSTALLING.md", this.Status.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task GitHub_not_answering_changes_nothing_and_is_said()
        {
            this.thisHost.GitHubDown = true;

            Assert.False(await this.Updater().RunAsync(CancellationToken.None));

            Assert.Equal(ServerUpdaterTests.Running, this.thisHost.Current);
            Assert.Equal(ServerUpdateOutcomes.Error, this.Status.Outcome);
        }

        [Fact]
        public async Task No_server_build_yet_is_up_to_date()
        {
            Assert.True(await this.Updater().RunAsync(CancellationToken.None));
            Assert.Equal(ServerUpdateOutcomes.UpToDate, this.Status.Outcome);
            Assert.Contains("carries no server build", this.Status.Message, StringComparison.Ordinal);

            Assert.True(await this.Updater(ServerUpdateOutcomes.MainChannel).RunAsync(CancellationToken.None));
            Assert.Contains("No build of main", this.Status.Message, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // What the service then says: it runs from releases/10.3.0, and finds the status beside it - the
        // answer Account > "Server version" shows.
        // ###########################################################################################
        [Fact]
        public async Task The_service_running_the_new_build_finds_what_the_updater_did()
        {
            this.PublishRelease("4.1.0", "10.3.0");
            await this.Updater().RunAsync(CancellationToken.None);

            ServerUpdatesAnswer answer = ServerUpdateStatus.AnswerFrom(ServerUpdateStatus.Locate(Path.Combine(this.Releases, "10.3.0")));

            Assert.True(answer.SetUp);
            Assert.Equal(ServerUpdateOutcomes.ReleasesChannel, answer.Channel);
            Assert.Equal(ServerUpdateOutcomes.Installed, answer.Outcome);
            Assert.Equal("10.3.0", answer.InstalledVersion);
            Assert.Equal(ServerUpdaterTests.Running, answer.InstalledFromVersion);
            Assert.NotNull(answer.CheckedUtc);
        }

        // ###########################################################################################
        // GitHub, gh, systemd and the links, faked: a service answers as the build "current" points at once
        // restarted - unless that build is one that does not come up. Time is virtual: a wait moves it on.
        // ###########################################################################################
        private sealed class FakeHost : IServerUpdateHost
        {
            public Dictionary<string, string> Pointers { get; } = new(StringComparer.Ordinal);

            public Dictionary<string, byte[]> Files { get; } = new(StringComparer.Ordinal);

            public Dictionary<string, string> Installed { get; } = new(StringComparer.Ordinal);

            public Dictionary<string, string> SettingsLinked { get; } = new(StringComparer.Ordinal);

            public HashSet<string> DoesNotComeUp { get; } = new(StringComparer.Ordinal);

            public List<string> Restarted { get; } = [];

            // The channel each attestation was checked for.
            public List<string> VerifiedFor { get; } = [];

            public AttestationCheck Proof { get; set; } = new(true, ServerBuilds.Commit, string.Empty);

            public string? Current { get; set; }

            public string? Answering { get; set; }

            public string? RestartRefusal { get; set; }

            public bool GitHubDown { get; set; }

            public int Downloads { get; private set; }

            public int PointerReads { get; private set; }

            public DateTimeOffset UtcNow { get; private set; } = new(2026, 11, 9, 12, 0, 0, TimeSpan.Zero);

            public void Advance(TimeSpan time) => this.UtcNow += time;

            public Task<string?> ReadPointerAsync(string url, CancellationToken cancellationToken)
            {
                this.PointerReads++;

                return this.GitHubDown
                    ? throw new HttpRequestException("No such host is known.")
                    : Task.FromResult(this.Pointers.TryGetValue(url, out string? json) ? json : null);
            }

            public Task DownloadAsync(string url, string toPath, long maxBytes, CancellationToken cancellationToken)
            {
                this.Downloads++;
                File.WriteAllBytes(toPath, this.Files[url]);
                return Task.CompletedTask;
            }

            public Task<AttestationCheck> VerifyAsync(string archivePath, string bundlePath, string channel, CancellationToken cancellationToken)
            {
                this.VerifiedFor.Add(channel);
                Assert.True(File.Exists(archivePath));
                Assert.True(File.Exists(bundlePath));
                return Task.FromResult(this.Proof);
            }

            public string? PointedAt(string releasesFolder) => this.Current;

            public void PointAt(string releasesFolder, string version)
            {
                Assert.True(Directory.Exists(Path.Combine(releasesFolder, version)), $"Pointed at {version}, which is no build there");
                this.Current = version;
            }

            public void LinkSettings(string buildFolder, string settingsPath) => this.SettingsLinked[buildFolder] = settingsPath;

            public Task<string?> RestartAsync(string service, CancellationToken cancellationToken)
            {
                if (this.RestartRefusal != null)
                    return Task.FromResult<string?>(this.RestartRefusal);

                this.Restarted.Add(service);
                this.Answering = this.Current != null && this.DoesNotComeUp.Contains(this.Current) ? null : this.Current;
                return Task.FromResult<string?>(null);
            }

            public Task<string?> HealthVersionAsync(CancellationToken cancellationToken) => Task.FromResult(this.Answering);

            // A build that does not come up stops as a refused setting does: exit 78, not restarted (crt-server.service).
            public Task<string> ServiceStateAsync(string service, CancellationToken cancellationToken) =>
                Task.FromResult(this.Answering == null ? "failed (exit 78)" : "active");

            public Task DelayAsync(TimeSpan delay, CancellationToken cancellationToken)
            {
                this.UtcNow += delay;
                return Task.CompletedTask;
            }

            public byte[]? ReadInstalledCopy(string fileName) =>
                this.Installed.TryGetValue(fileName, out string? content) ? System.Text.Encoding.UTF8.GetBytes(content) : null;
        }
    }
}
