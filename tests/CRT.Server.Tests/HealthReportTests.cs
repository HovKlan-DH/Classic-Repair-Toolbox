using System.Reflection;
using CRT.Server.Handlers.Compat;
using CRT.Server.Handlers.Health;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers what GET /api/health answers with.
    //
    // The endpoint itself is one line in Program.cs and is not tested - what matters is that the
    // payload is well-formed, that a missing version is visible rather than blank, and that the
    // response stays free of anything an unauthenticated caller should not learn. That last one is
    // the point of the "reveals nothing else" test below: this endpoint is reachable from the
    // public internet through the Apache proxy, so a future change adding a database probe or a
    // path to it should have to delete an assertion that says not to.
    // ###########################################################################################
    public class HealthReportTests
    {
        [Fact]
        public void A_health_report_says_ok_and_carries_the_build_version()
        {
            var now = new DateTimeOffset(2026, 9, 20, 12, 0, 0, TimeSpan.Zero);

            HealthStatus status = HealthReport.Build(now, "2.6.0-alpha.2");

            Assert.Equal("ok", status.Status);
            Assert.Equal("2.6.0-alpha.2", status.Version);
            Assert.Equal(now, status.Utc);
        }

        // The API revision it serves (2026-10-04), so the Maintainer tab can say whether CRT or the
        // server is behind - the same number the server's gate refuses older CRTs by.
        [Fact]
        public void A_health_report_carries_the_api_revision_the_server_serves()
        {
            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, "4.6.0");

            Assert.Equal(ClientVersionContract.ApiRevision, status.ApiRevision);
            Assert.Equal(ClientVersionContract.ApiRevision, ClientVersionPolicy.Current.ApiRevision);
        }

        // A blank version must not silently become an empty string in the response: a deployment
        // built without the version attribute should be obvious to whoever is reading the health
        // output, not indistinguishable from a formatting bug.
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void A_missing_version_is_reported_as_unknown_rather_than_blank(string? version)
        {
            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, version);

            Assert.Equal(HealthReport.UnknownVersion, status.Version);
            Assert.NotEqual(string.Empty, status.Version);
        }

        // The SDK appends SourceRevisionId to InformationalVersion, so the raw attribute reads
        // "0.1.0-phase3+7150bc943a81fe...". Publishing that from an unauthenticated, publicly
        // reachable endpoint hands out the exact commit the deployed binary was built from.
        // This was real: the first run of the live endpoint returned the full SHA.
        [Fact]
        public void The_git_commit_is_never_published_in_the_version()
        {
            HealthStatus status = HealthReport.Build(
                DateTimeOffset.UtcNow,
                "0.1.0-phase3+7150bc943a81fe1b74fe50c282ec2cc249ab10d4");

            Assert.Equal("0.1.0-phase3", status.Version);
            Assert.DoesNotContain("+", status.Version, StringComparison.Ordinal);
            Assert.DoesNotContain("7150bc94", status.Version, StringComparison.Ordinal);
        }

        // The pre-release part before the "+" is kept: it distinguishes deployments and carries no
        // information that is not already in the public repository's release list.
        [Fact]
        public void The_prerelease_part_of_the_version_survives()
        {
            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, "2.6.0-alpha.2+abcdef");

            Assert.Equal("2.6.0-alpha.2", status.Version);
        }

        // A version that is nothing BUT build metadata must not collapse to an empty string.
        [Fact]
        public void A_version_that_is_only_build_metadata_reports_unknown()
        {
            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, "+abcdef");

            Assert.Equal(HealthReport.UnknownVersion, status.Version);
        }

        // The real assembly must not carry a commit through to the response either - this is the
        // end-to-end version of the test above, run against what actually ships.
        [Fact]
        public void The_real_assembly_version_carries_no_build_metadata()
        {
            string? raw = HealthReport.ReadInformationalVersion(typeof(HealthReport).Assembly);

            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, raw);

            Assert.DoesNotContain("+", status.Version, StringComparison.Ordinal);
        }

        [Fact]
        public void A_version_with_surrounding_whitespace_is_trimmed()
        {
            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, "  1.2.3  ");

            Assert.Equal("1.2.3", status.Version);
        }

        // The timestamp is normalised to UTC so two deployments in different timezones cannot
        // report times that look like they disagree.
        [Fact]
        public void The_timestamp_is_normalised_to_utc()
        {
            var localTime = new DateTimeOffset(2026, 9, 20, 14, 0, 0, TimeSpan.FromHours(2));

            HealthStatus status = HealthReport.Build(localTime, "1.0.0");

            Assert.Equal(TimeSpan.Zero, status.Utc.Offset);
            Assert.Equal(12, status.Utc.Hour);
        }

        // Health is unauthenticated and publicly reachable. It must report liveness and the build,
        // and nothing that helps someone map the box: no database state, no configuration, no
        // paths, no runtime or OS version. Adding any of those means deliberately changing this
        // test, which is the point - see HealthReport's header for the reasoning.
        //
        // The API revision (2026-10-04) was added deliberately, as part of "the build": it says which
        // shape of the API is deployed, the number every published CRT build carries in its requests
        // anyway, and nothing about the box.
        [Fact]
        public void A_health_report_reveals_nothing_beyond_liveness_and_the_build()
        {
            HealthStatus status = HealthReport.Build(DateTimeOffset.UtcNow, "1.0.0");

            PropertyInfo[] properties = typeof(HealthStatus).GetProperties();

            Assert.Equal(4, properties.Length);
            Assert.Contains(properties, p => p.Name == nameof(HealthStatus.Status));
            Assert.Contains(properties, p => p.Name == nameof(HealthStatus.Version));
            Assert.Contains(properties, p => p.Name == nameof(HealthStatus.Utc));
            Assert.Contains(properties, p => p.Name == nameof(HealthStatus.ApiRevision));
            Assert.NotNull(status);
        }

        [Fact]
        public void The_informational_version_is_read_from_the_assembly()
        {
            // The server assembly sets InformationalVersion in its csproj, so reading it back from
            // the assembly under test proves the wiring the endpoint depends on.
            string? version = HealthReport.ReadInformationalVersion(typeof(HealthReport).Assembly);

            Assert.False(string.IsNullOrWhiteSpace(version));
        }

        [Fact]
        public void Reading_the_version_from_a_null_assembly_is_refused()
        {
            Assert.Throws<ArgumentNullException>(() => HealthReport.ReadInformationalVersion(null!));
        }
    }
}
