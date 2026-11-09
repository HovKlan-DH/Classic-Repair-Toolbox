using CRT.Server.Handlers.Updates;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // What a server build says about itself (crt-server-release.json, 2026-11-09) - its version names the
    // folder it goes into, so a version that is no plain SemVer version, or could name another folder, is
    // refused.
    // ###########################################################################################
    public sealed class ServerReleaseInfoTests
    {
        [Fact]
        public void A_build_says_its_server_version_its_CRT_release_its_commit_and_whether_it_is_a_pre_release()
        {
            ServerReleaseInfo? info = ServerReleaseInfo.Read(ServerBuilds.ReleaseInfo("10.3.0", preRelease: true, crtVersion: "4.1.0-beta.2"), out string why);

            Assert.NotNull(info);
            Assert.Equal(string.Empty, why);
            Assert.Equal("10.3.0", info.ServerVersion);
            Assert.Equal("4.1.0-beta.2", info.CrtVersion);
            Assert.Equal(ServerBuilds.Commit, info.Commit);
            Assert.True(info.PreRelease);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not json")]
        [InlineData("{}")]
        [InlineData("{\"serverVersion\":\"../10.3.0\",\"crtVersion\":\"4.1.0\"}")]
        [InlineData("{\"serverVersion\":\"10.3\",\"crtVersion\":\"4.1.0\"}")]
        [InlineData("{\"serverVersion\":\"10.3.0\",\"crtVersion\":\"latest\"}")]
        public void A_build_whose_info_is_missing_or_wrong_is_refused(string? json)
        {
            Assert.Null(ServerReleaseInfo.Read(json, out string why));
            Assert.NotEqual(string.Empty, why);
        }

        [Theory]
        [InlineData("10.3.0", true)]
        [InlineData("10.3.0-beta.1", true)]
        [InlineData("10.3.0+abc", false)]
        [InlineData("../10.3.0", false)]
        [InlineData("10.3.0/x", false)]
        [InlineData("10.3.0-..", false)]
        [InlineData("current", false)]
        [InlineData("", false)]
        [InlineData(null, false)]
        public void Only_a_plain_version_names_a_build_folder(string? version, bool safe) =>
            Assert.Equal(safe, ServerReleaseInfo.IsFolderSafeVersion(version));
    }
}
