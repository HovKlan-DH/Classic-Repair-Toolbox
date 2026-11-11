using System.Text.Json;
using CRT.Server.Handlers.Updates;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Where a server build comes from and who must have built it (ServerReleaseFeed, ServerBuildPointer;
    // owner decisions, 2026-11-09: "the REAL version should work on a release ... But ALPHA/BETA works on
    // commit"). The live service reads GitHub's newest release, the test instance the standing pre-release of
    // main; each trusts only its own workflow; a pointer naming anything else is no pointer, so no address is
    // ever built from it.
    // ###########################################################################################
    public sealed class ServerReleaseFeedTests
    {
        private const string Sha = "3b8ac6b30336802fc1a858d7c084e11cdf24ac1a761ca90b68022d7d729208de";

        private static string Pointer(string tag, string serverVersion = "10.3.0", string sha256 = ServerReleaseFeedTests.Sha) =>
            JsonSerializer.Serialize(new { tag, serverVersion, commit = ServerBuilds.Commit, sha256 });

        [Fact]
        public void The_live_service_reads_the_newest_release_and_the_test_instance_mains_standing_pre_release()
        {
            Assert.Equal(
                "https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/latest/download/crt-server-linux-x64.json",
                ServerReleaseFeed.PointerUrlFor(ServerUpdateOutcomes.ReleasesChannel));

            Assert.Equal(
                "https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/crt-server-main/crt-server-linux-x64.json",
                ServerReleaseFeed.PointerUrlFor(ServerUpdateOutcomes.MainChannel));

            Assert.Equal(
                "https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/4.1.0/crt-server-linux-x64.tar.gz",
                ServerReleaseFeed.DownloadUrl("4.1.0", ServerReleaseFeed.ArchiveName));
        }

        // ###########################################################################################
        // *** EACH CHANNEL TRUSTS ITS OWN WORKFLOW ONLY. *** A build of main - attested by the push workflow - is no
        // build the live service takes, whatever release it is put on; a release's is the release workflow's.
        // ###########################################################################################
        [Fact]
        public void Each_channel_trusts_only_its_own_workflow_on_main()
        {
            Assert.Equal("HovKlan-DH/Classic-Repair-Toolbox/.github/workflows/build-and-release.yml", ServerReleaseFeed.SignerWorkflowFor(ServerUpdateOutcomes.ReleasesChannel));
            Assert.Equal("HovKlan-DH/Classic-Repair-Toolbox/.github/workflows/build-and-unittest.yml", ServerReleaseFeed.SignerWorkflowFor(ServerUpdateOutcomes.MainChannel));
            Assert.Equal("refs/heads/main", ServerReleaseFeed.SourceRef);
            Assert.Equal("crt-server-main", ServerReleaseFeed.MainTag);
        }

        [Fact]
        public void A_pointer_says_its_release_version_commit_and_checksum()
        {
            ServerBuildPointer? pointer = ServerBuildPointer.Read(ServerReleaseFeedTests.Pointer("4.1.0"), ServerUpdateOutcomes.ReleasesChannel, out string why);

            Assert.NotNull(pointer);
            Assert.Equal(string.Empty, why);
            Assert.Equal("4.1.0", pointer.Tag);
            Assert.Equal("10.3.0", pointer.ServerVersion);
            Assert.Equal(ServerReleaseFeedTests.Sha, pointer.Sha256);
            Assert.Equal("CRT release [4.1.0]", pointer.Source);

            Assert.Equal("main at [7f3a2c1]", ServerBuildPointer.Read(ServerReleaseFeedTests.Pointer("crt-server-main"), ServerUpdateOutcomes.MainChannel, out _)!.Source);
        }

        // ###########################################################################################
        // A pointer is never proven - but no address is built from one that names another channel's release, a
        // tag that could climb out of the downloads, no version or no checksum.
        // ###########################################################################################
        [Theory]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, "crt-server-main", "10.3.0", ServerReleaseFeedTests.Sha)]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, "4.1.0-beta.1", "10.3.0", ServerReleaseFeedTests.Sha)]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, "../4.1.0", "10.3.0", ServerReleaseFeedTests.Sha)]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, "4.1.0/x", "10.3.0", ServerReleaseFeedTests.Sha)]
        [InlineData(ServerUpdateOutcomes.MainChannel, "4.1.0", "10.3.0", ServerReleaseFeedTests.Sha)]
        [InlineData(ServerUpdateOutcomes.MainChannel, "crt-server-main", "../10.3.0", ServerReleaseFeedTests.Sha)]
        [InlineData(ServerUpdateOutcomes.MainChannel, "crt-server-main", "10.3.0", "not-a-checksum")]
        [InlineData(ServerUpdateOutcomes.MainChannel, "crt-server-main", "10.3.0", "3B8AC6B30336802FC1A858D7C084E11CDF24AC1A761CA90B68022D7D729208DE")]
        public void A_pointer_naming_anything_but_its_channels_release_and_a_build_is_refused(string channel, string tag, string serverVersion, string sha256)
        {
            Assert.Null(ServerBuildPointer.Read(ServerReleaseFeedTests.Pointer(tag, serverVersion, sha256), channel, out string why));
            Assert.NotEqual(string.Empty, why);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        [InlineData("[]")]
        public void A_pointer_that_cannot_be_read_is_refused(string? json) =>
            Assert.Null(ServerBuildPointer.Read(json, ServerUpdateOutcomes.ReleasesChannel, out _));
    }
}
