using CRT.Server.Handlers.Updates;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The files the box installs once for the updater (2026-11-09; INSTALLING.md, "Automatic updates from
    // GitHub"): a unit and a timer per instance, and the polkit rule letting it restart the services. Held
    // to the code and to each other:
    //
    //   - each unit's command line is one ServerUpdateCommand accepts, for ITS instance - its folder, its
    //     service, its port, its channel (the live service releases only);
    //   - the two units are the same but for which instance they are, as the service units are
    //     (TestInstanceUnitTests) - a hardening line forgotten in one shows here;
    //   - the updater runs as its own user and writes its instance's builds and nothing else;
    //   - the polkit rule lets that user restart the two services and do nothing else.
    // ###########################################################################################
    public sealed class UpdaterUnitTests
    {
        private const string Site = "/mydir/http/classic-repair-toolbox.dk";

        private static readonly string[] InstanceKeys = ["Description=", "After=", "ExecStart=", "ReadWritePaths="];

        private static string ShippedPath(string name)
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);
            return Path.Combine(folder!, "src", "CRT.Server", name);
        }

        private static List<string> Settings(string name) =>
            File.ReadAllLines(UpdaterUnitTests.ShippedPath(name))
                .Select(line => line.Trim())
                .Where(line => line.Length > 0 && !line.StartsWith('#'))
                .ToList();

        private static string Line(string name, string key) =>
            Assert.Single(UpdaterUnitTests.Settings(name), line => line.StartsWith(key, StringComparison.Ordinal));

        private static ServerUpdateCommand.Options CommandOf(string unit)
        {
            string[] words = UpdaterUnitTests.Line(unit, "ExecStart=")["ExecStart=".Length..].Split(' ', StringSplitOptions.RemoveEmptyEntries);

            Assert.Equal("/usr/bin/dotnet", words[0]);
            Assert.Equal(ServerUpdateCommand.Name, words[2]);

            ServerUpdateCommand.Options? options = ServerUpdateCommand.Parse(words.Skip(3).ToList(), out string why);
            Assert.True(options != null, why);
            return options!;
        }

        [Theory]
        [InlineData("crt-server-update.service", "crt-server", "crt-server.service", 5199, ServerUpdateOutcomes.ReleasesChannel)]
        [InlineData("crt-server-test-update.service", "crt-server-test", "crt-server-test.service", 5198, ServerUpdateOutcomes.PreReleasesChannel)]
        public void Each_updater_looks_after_its_own_instance(string unit, string folder, string service, int port, string channel)
        {
            ServerUpdateCommand.Options options = UpdaterUnitTests.CommandOf(unit);

            Assert.Equal(channel, options.Channel);
            Assert.Equal($"{UpdaterUnitTests.Site}/{folder}", options.Folder.Replace('\\', '/'));
            Assert.Equal(service, options.Service);
            Assert.Equal(new Uri($"http://127.0.0.1:{port}/api/health"), options.Health);
            Assert.False(options.DryRun);
            Assert.False(options.RetryFailed);

            // It runs the program the service runs, and its health port is the service's own.
            Assert.StartsWith($"ExecStart=/usr/bin/dotnet {UpdaterUnitTests.Site}/{folder}/app/CRT.Server.dll ", UpdaterUnitTests.Line(unit, "ExecStart="), StringComparison.Ordinal);
            Assert.Equal($"Environment=ASPNETCORE_URLS=http://127.0.0.1:{port}", UpdaterUnitTests.Line(service, "Environment=ASPNETCORE_URLS="));
            Assert.Equal($"ReadWritePaths={UpdaterUnitTests.Site}/{folder}/{ServerUpdater.ReleasesFolderName}", UpdaterUnitTests.Line(unit, "ReadWritePaths="));
        }

        [Fact]
        public void Every_setting_but_which_instance_it_is_equals_between_the_two_updaters()
        {
            static bool IsInstanceLine(string line) => UpdaterUnitTests.InstanceKeys.Any(key => line.StartsWith(key, StringComparison.Ordinal));

            Assert.Equal(
                UpdaterUnitTests.Settings("crt-server-update.service").Where(line => !IsInstanceLine(line)),
                UpdaterUnitTests.Settings("crt-server-test-update.service").Where(line => !IsInstanceLine(line)));

            Assert.Equal(
                UpdaterUnitTests.Settings("crt-server-update.timer").Where(line => !line.StartsWith("Description=", StringComparison.Ordinal) && !line.StartsWith("Unit=", StringComparison.Ordinal)),
                UpdaterUnitTests.Settings("crt-server-test-update.timer").Where(line => !line.StartsWith("Description=", StringComparison.Ordinal) && !line.StartsWith("Unit=", StringComparison.Ordinal)));
        }

        // ###########################################################################################
        // Its own user - never the service's, which must not rewrite its own program - and its own group:
        // NOT the service's crt-data, which reads the settings with the database password in them. What it
        // unpacks reaches the service's group through the releases folder's setgid bit (ServerArchive).
        // ###########################################################################################
        [Theory]
        [InlineData("crt-server-update.service")]
        [InlineData("crt-server-test-update.service")]
        public void The_updater_runs_as_its_own_user_and_writes_nothing_but_its_builds(string unit)
        {
            List<string> settings = UpdaterUnitTests.Settings(unit);

            Assert.Contains("Type=oneshot", settings);
            Assert.Contains("User=crt-deploy", settings);
            Assert.DoesNotContain(settings, line => line.StartsWith("Group=", StringComparison.Ordinal));
            Assert.DoesNotContain(settings, line => line.StartsWith("SupplementaryGroups=", StringComparison.Ordinal));
            Assert.Contains("NoNewPrivileges=true", settings);
            Assert.Contains("ProtectSystem=strict", settings);
            Assert.Contains("ProtectHome=true", settings);
            Assert.Contains("PrivateTmp=true", settings);
            Assert.Single(settings, line => line.StartsWith("ReadWritePaths=", StringComparison.Ordinal));
            Assert.DoesNotContain("User=crt-server", settings);
            Assert.DoesNotContain("User=root", settings);
        }

        // ###########################################################################################
        // Every minute (owner request, 2026-11-09: "can it then update every minute?"), counted from the END of
        // the look before - so a look that installs (a download, a restart, two minutes of health checks) never
        // runs into the next one - and within seconds, not systemd's default minute of slack.
        // ###########################################################################################
        [Theory]
        [InlineData("crt-server-update.timer", "crt-server-update.service")]
        [InlineData("crt-server-test-update.timer", "crt-server-test-update.service")]
        public void Each_timer_starts_its_own_updater_a_minute_after_its_last_look(string timer, string unit)
        {
            Assert.Equal($"Unit={unit}", UpdaterUnitTests.Line(timer, "Unit="));
            Assert.Equal("OnActiveSec=1min", UpdaterUnitTests.Line(timer, "OnActiveSec="));
            Assert.Equal("OnUnitInactiveSec=1min", UpdaterUnitTests.Line(timer, "OnUnitInactiveSec="));
            Assert.Equal("AccuracySec=5s", UpdaterUnitTests.Line(timer, "AccuracySec="));
            Assert.DoesNotContain(UpdaterUnitTests.Settings(timer), line => line.StartsWith("OnCalendar=", StringComparison.Ordinal));
            Assert.Equal("WantedBy=timers.target", UpdaterUnitTests.Line(timer, "WantedBy="));
        }

        // ###########################################################################################
        // *** THE POLKIT RULE: RESTART, THESE TWO SERVICES, THIS USER - NOTHING ELSE. *** It is the one thing
        // the updater may do as more than itself.
        // ###########################################################################################
        [Fact]
        public void The_polkit_rule_lets_the_updater_restart_the_two_services_and_nothing_else()
        {
            string rule = File.ReadAllText(UpdaterUnitTests.ShippedPath(ServerUpdater.PolkitRuleFile));

            Assert.Contains("action.id === \"org.freedesktop.systemd1.manage-units\"", rule, StringComparison.Ordinal);
            Assert.Contains("subject.user === \"crt-deploy\"", rule, StringComparison.Ordinal);
            Assert.Contains("action.lookup(\"verb\") === \"restart\"", rule, StringComparison.Ordinal);
            Assert.Contains("action.lookup(\"unit\") === \"crt-server.service\"", rule, StringComparison.Ordinal);
            Assert.Contains("action.lookup(\"unit\") === \"crt-server-test.service\"", rule, StringComparison.Ordinal);

            // One rule, one YES - and no unit, verb or user beyond those.
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(rule, "polkit\\.Result\\.YES"));
            Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(rule, "action\\.lookup\\(\"unit\"\\)").Count);
            Assert.DoesNotContain("||  subject", rule, StringComparison.Ordinal);
        }

        // The files a build is held for are files every build ships - so the hold looks at something real.
        [Theory]
        [InlineData("crt-server.service")]
        [InlineData("crt-server-test.service")]
        public void Every_file_that_can_hold_a_build_is_shipped(string service)
        {
            foreach (string name in ServerUpdater.SystemFilesOf(service))
                Assert.True(File.Exists(UpdaterUnitTests.ShippedPath(name)), $"{name} is not shipped");
        }
    }
}
