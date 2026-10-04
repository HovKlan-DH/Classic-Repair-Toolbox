using System.Reflection;
using System.Text.RegularExpressions;
using CRT.Server.Handlers.Health;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Holds CRT.Server's declared version and its written history together.
    //
    // WHY THIS EXISTS. The project owner bumps CRT by hand (CRT Maintainer was merged into it on 2026-09-29) but asked
    // (2026-09-26) for THIS service's version to be handled for them - it only has to be visibly
    // versioned so a deployment can be identified. The policy therefore lives in
    // src/CRT.Server/VERSION.md and says two edits go together: InformationalVersion in the csproj,
    // and a row in that file's history table.
    //
    // "Two files must agree" is precisely the kind of rule this repository keeps learning cannot be
    // left to memory - the same reasoning as TestPathLiteralTests and WikiHelpPageNamesTests. A
    // version bumped in the csproj with no history row is a deployment nobody can explain later,
    // and it fails silently: the build is green and /api/health answers perfectly.
    //
    // The Stop hook (.claude/hooks/server-version-bump.sh) is the other half - it notices server
    // code changing while the version stands still. It can only WARN, because which bump a change
    // deserves is a judgement call. This test covers the part that IS objective.
    //
    // No server, no database, no network: it reads two files and the assembly's own attribute.
    // ###########################################################################################
    public class ServerVersionTests
    {
        // The version as the BUILD produced it, with the SDK's "+<commit>" metadata stripped the
        // same way the health endpoint strips it.
        private static string DeclaredVersion()
        {
            string? raw = HealthReport.ReadInformationalVersion(typeof(HealthReport).Assembly);
            Assert.False(string.IsNullOrWhiteSpace(raw), "CRT.Server carries no InformationalVersion.");

            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, raw);
            return status.Version;
        }

        private static string RepoRoot()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);
            return folder!;
        }

        private static string VersionHistoryPath() =>
            Path.Combine(ServerVersionTests.RepoRoot(), "src", "CRT.Server", "VERSION.md");

        // -----------------------------------------------------------------------------------

        // The whole point of the exercise: a version that means something to a human reading
        // /api/health. "unknown" would mean the build lost the attribute entirely.
        [Fact]
        public void The_service_reports_a_real_version_rather_than_unknown()
        {
            string version = ServerVersionTests.DeclaredVersion();

            Assert.NotEqual(HealthReport.UnknownVersion, version);
            Assert.DoesNotContain("$(", version, StringComparison.Ordinal);
        }

        // SemVer 2.0.0 is what VERSION.md commits to, so the number has to BE one: three numeric
        // parts, with an optional pre-release suffix. A version like "alpha3" or "1.0" would make
        // the MAJOR/MINOR/PATCH policy meaningless, and nothing else would object to it.
        [Fact]
        public void The_declared_version_is_valid_semver()
        {
            string version = ServerVersionTests.DeclaredVersion();

            // MAJOR.MINOR.PATCH, then an optional -prerelease of dot-separated alphanumerics.
            Assert.Matches(
                @"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(-[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*)?$",
                version);
        }

        // The history table is the record of WHY a deployment's version is what it is. A bump that
        // never reached it leaves the top row describing an older build than the one running, which
        // is worse than no record at all - it is a confident wrong answer.
        [Fact]
        public void The_declared_version_is_the_newest_row_in_the_version_history()
        {
            string version = ServerVersionTests.DeclaredVersion();
            string historyPath = ServerVersionTests.VersionHistoryPath();

            Assert.True(File.Exists(historyPath), $"The version history is missing: {historyPath}");

            string history = File.ReadAllText(historyPath);

            // Rows look like: | `1.0.0-alpha.3` | 2026-09-26 | - | ... |
            MatchCollection rows = Regex.Matches(history, @"^\|\s*`([^`]+)`\s*\|", RegexOptions.Multiline);

            Assert.True(rows.Count > 0, "The version history table has no version rows.");

            string newest = rows[0].Groups[1].Value.Trim();

            Assert.Equal(version, newest);
        }

        // Every version is listed once. A duplicated row means two different changes are both
        // claiming the same number, so the history can no longer say which one is deployed.
        [Fact]
        public void No_version_is_listed_twice_in_the_history()
        {
            string history = File.ReadAllText(ServerVersionTests.VersionHistoryPath());

            List<string> versions = Regex
                .Matches(history, @"^\|\s*`([^`]+)`\s*\|", RegexOptions.Multiline)
                .Select(match => match.Groups[1].Value.Trim())
                .ToList();

            Assert.Equal(versions.Distinct(StringComparer.Ordinal).Count(), versions.Count);
        }
    }
}
