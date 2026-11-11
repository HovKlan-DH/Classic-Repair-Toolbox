using System.Text.Json;
using System.Text.Json.Serialization;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // WHERE A SERVER BUILD COMES FROM, AND WHO MUST HAVE BUILT IT (owner decisions, 2026-11-09: "GitHub
    // attestations"; and then "the REAL version should work on a release ... But ALPHA/BETA works on commit").
    //
    //   - THE LIVE SERVICE ("releases"): the server build a real CRT release carries - built, attested and
    //     attached by the release workflow, build-and-release.yml. Found through GitHub's newest release.
    //   - THE TEST INSTANCE ("main"): the server build of every commit pushed to main whose tests passed -
    //     built, attested and put by the push workflow, build-and-unittest.yml, on ONE standing pre-release,
    //     MainTag, its files replaced at every push. No CRT release is needed for it.
    //
    // Each carries three files: ArchiveName, the build; BundleName, GitHub's attestation that the named
    // workflow built exactly those bytes from main; PointerName, a few lines saying which build that is
    // (ServerBuildPointer) - read every minute, so a minute costs one small download and no GitHub API call
    // (60 an hour without a login). The pointer is not proven: it only decides whether to download. What is
    // installed is decided by the build's own crt-server-release.json, inside the attested bytes.
    //
    // Everything here is fixed in the code, never read from a setting: the box trusts builds of THIS
    // repository's two workflows on main, and nothing a setting could point elsewhere.
    // ###########################################################################################
    internal static class ServerReleaseFeed
    {
        public const string Repository = "HovKlan-DH/Classic-Repair-Toolbox";

        public const string SourceRef = "refs/heads/main";

        // The standing pre-release the push workflow keeps the newest build of main on.
        public const string MainTag = "crt-server-main";

        public const string ArchiveName = "crt-server-linux-x64.tar.gz";
        public const string BundleName = ServerReleaseFeed.ArchiveName + ".sigstore.json";
        public const string PointerName = "crt-server-linux-x64.json";

        // Where a release's files are downloaded from - nothing else is ever fetched.
        public const string DownloadPrefix = "https://github.com/" + ServerReleaseFeed.Repository + "/releases/download/";

        public const long ArchiveMaxBytes = 200L * 1024 * 1024;
        public const long BundleMaxBytes = 1L * 1024 * 1024;
        public const long PointerMaxBytes = 64L * 1024;

        // ###########################################################################################
        // The workflow whose builds a channel installs - the attestation must name it.
        // ###########################################################################################
        public static string SignerWorkflowFor(string channel) =>
            ServerReleaseFeed.Repository + "/.github/workflows/" +
            (channel == ServerUpdateOutcomes.MainChannel ? "build-and-unittest.yml" : "build-and-release.yml");

        // ###########################################################################################
        // The pointer a channel reads every minute: GitHub's newest release's (a pre-release is never the
        // newest, so never the live service's), or the standing pre-release's.
        // ###########################################################################################
        public static string PointerUrlFor(string channel) =>
            channel == ServerUpdateOutcomes.MainChannel
                ? ServerReleaseFeed.DownloadPrefix + ServerReleaseFeed.MainTag + "/" + ServerReleaseFeed.PointerName
                : "https://github.com/" + ServerReleaseFeed.Repository + "/releases/latest/download/" + ServerReleaseFeed.PointerName;

        // A file of release `tag` - a tag the pointer named, held to its channel's shape first (ServerBuildPointer).
        public static string DownloadUrl(string tag, string fileName) =>
            ServerReleaseFeed.DownloadPrefix + Uri.EscapeDataString(tag) + "/" + fileName;
    }

    // ###########################################################################################
    // WHICH BUILD A RELEASE CARRIES - "crt-server-linux-x64.json", written by the workflows beside the build
    // (2026-11-09):
    //
    //   tag            the release it is on - a CRT release's version, or ServerReleaseFeed.MainTag
    //   serverVersion  the server version inside it
    //   commit         the commit it was built from
    //   sha256         the build's checksum - a new build is a new checksum, even of the same version
    //
    // NOT PROVEN, and so never trusted with more than whether to download: a pointer naming an older or the
    // same version only spares a download; one naming a newer leads to a download, whose attestation and own
    // description decide everything (ServerUpdater).
    // ###########################################################################################
    internal sealed class ServerBuildPointer
    {
        [JsonPropertyName("tag")] public string Tag { get; set; } = string.Empty;

        [JsonPropertyName("serverVersion")] public string ServerVersion { get; set; } = string.Empty;

        [JsonPropertyName("commit")] public string Commit { get; set; } = string.Empty;

        [JsonPropertyName("sha256")] public string Sha256 { get; set; } = string.Empty;

        // ###########################################################################################
        // The pointer in `json` for `channel`, or null with the reason. Its tag must be its channel's - a real
        // release's version for the live service, the standing pre-release for the test instance - so no URL is
        // ever built from anything else; its checksum a SHA-256; its version one that can name a folder.
        // ###########################################################################################
        public static ServerBuildPointer? Read(string? json, string channel, out string why)
        {
            ServerBuildPointer? pointer;

            try
            {
                pointer = string.IsNullOrWhiteSpace(json) ? null : JsonSerializer.Deserialize<ServerBuildPointer>(json);
            }
            catch (JsonException ex)
            {
                why = $"The release's {ServerReleaseFeed.PointerName} could not be read - {ex.Message}";
                return null;
            }

            if (pointer == null)
            {
                why = $"The release's {ServerReleaseFeed.PointerName} is empty.";
                return null;
            }

            bool tagFits = channel == ServerUpdateOutcomes.MainChannel
                ? pointer.Tag == ServerReleaseFeed.MainTag
                : ServerReleaseInfo.IsFolderSafeVersion(pointer.Tag) && !pointer.Tag.Contains('-');

            if (!tagFits)
            {
                why = $"The release's {ServerReleaseFeed.PointerName} names [{pointer.Tag}], which is no release of this channel.";
                return null;
            }

            if (!ServerReleaseInfo.IsFolderSafeVersion(pointer.ServerVersion))
            {
                why = $"The release's {ServerReleaseFeed.PointerName} names the server version [{pointer.ServerVersion}], which is no version.";
                return null;
            }

            if (!IsSha256(pointer.Sha256))
            {
                why = $"The release's {ServerReleaseFeed.PointerName} names no checksum.";
                return null;
            }

            why = string.Empty;
            return pointer;
        }

        // A lower-case hex SHA-256, as the workflows write it with sha256sum.
        public static bool IsSha256(string? value) =>
            value is { Length: 64 } && value.All(c => c is (>= '0' and <= '9') or (>= 'a' and <= 'f'));

        // Where it came from, in the status's words: "CRT release [4.1.0]" or "main at [63431ff]".
        public string Source =>
            this.Tag == ServerReleaseFeed.MainTag
                ? $"main at [{(this.Commit.Length > 7 ? this.Commit[..7] : this.Commit)}]"
                : $"CRT release [{this.Tag}]";
    }
}
