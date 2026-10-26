namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The systemd unit of the contribution service's TEST instance, crt-server-test.service - CRT
    // 4.0's pre-releases submit to it, beside the live service (owner decision, 2026-10-25).
    //
    // Two units drift: a hardening line added to the live one and forgotten in the test one, or a
    // path copied across that lets the test instance write the live service's uploads. So the test
    // unit is held to the live unit line by line: only the five lines that say WHICH instance it is
    // may differ, and each of those is checked for what it must say.
    // ###########################################################################################
    public sealed class TestInstanceUnitTests
    {
        private const string Site = "/mydir/http/classic-repair-toolbox.dk";

        private static readonly string[] InstanceKeys =
        [
            "Description=",
            "WorkingDirectory=",
            "ExecStart=",
            "Environment=ASPNETCORE_URLS=",
            "ReadWritePaths=",
        ];

        [Fact]
        public void Every_setting_but_which_instance_it_is_equals_the_live_units()
        {
            List<string> live = TestInstanceUnitTests.Settings("crt-server.service");
            List<string> test = TestInstanceUnitTests.Settings("crt-server-test.service");

            Assert.Equal(
                live.Where(line => !TestInstanceUnitTests.IsInstanceLine(line)),
                test.Where(line => !TestInstanceUnitTests.IsInstanceLine(line)));

            // And it has each instance line, once.
            Assert.All(TestInstanceUnitTests.InstanceKeys, key => Assert.Single(test, line => line.StartsWith(key, StringComparison.Ordinal)));
        }

        [Fact]
        public void It_listens_on_its_own_loopback_port()
        {
            Assert.Equal("Environment=ASPNETCORE_URLS=http://127.0.0.1:5198", TestInstanceUnitTests.Line("crt-server-test.service", "Environment=ASPNETCORE_URLS="));
            Assert.Equal("Environment=ASPNETCORE_URLS=http://127.0.0.1:5199", TestInstanceUnitTests.Line("crt-server.service", "Environment=ASPNETCORE_URLS="));
        }

        [Fact]
        public void It_runs_its_own_program_with_its_own_settings()
        {
            string app = TestInstanceUnitTests.Site + "/crt-server-test/app";

            Assert.Equal($"WorkingDirectory={app}", TestInstanceUnitTests.Line("crt-server-test.service", "WorkingDirectory="));
            Assert.Equal($"ExecStart=/usr/bin/dotnet {app}/CRT.Server.dll", TestInstanceUnitTests.Line("crt-server-test.service", "ExecStart="));
        }

        // The 4.0 trees it publishes into and its own uploads and feedback - never the live service's
        // folders, never the Excel trees CRT 3.x reads.
        [Fact]
        public void It_writes_the_CRT_4_0_trees_and_its_own_folders_only()
        {
            string[] paths = TestInstanceUnitTests.Line("crt-server-test.service", "ReadWritePaths=")["ReadWritePaths=".Length..]
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);

            Assert.Equal(
                [
                    TestInstanceUnitTests.Site + "/public_html/app-data-v4-BETA",
                    TestInstanceUnitTests.Site + "/public_html/app-data-v4",
                    TestInstanceUnitTests.Site + "/crt-server-test/blobs",
                    TestInstanceUnitTests.Site + "/crt-server-test/feedback",
                ],
                paths);
        }

        // Every line that sets something - comments, blank lines and section headings left out.
        private static List<string> Settings(string unit) =>
            File.ReadAllLines(TestInstanceUnitTests.UnitPath(unit))
                .Select(line => line.Trim())
                .Where(line => line.Length > 0 && !line.StartsWith('#') && !line.StartsWith('['))
                .ToList();

        private static string Line(string unit, string key) =>
            Assert.Single(TestInstanceUnitTests.Settings(unit), line => line.StartsWith(key, StringComparison.Ordinal));

        private static bool IsInstanceLine(string line) =>
            TestInstanceUnitTests.InstanceKeys.Any(key => line.StartsWith(key, StringComparison.Ordinal));

        private static string UnitPath(string unit)
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "src", "CRT.Server", unit);
        }
    }
}
