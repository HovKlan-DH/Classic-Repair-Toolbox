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
    //     built and attested by the push workflow, build-and-unittest.yml, and put on a release of ITS SERVER
    //     VERSION in a repository of its own, MainBuildsRepository (owner request, 2026-11-11: a standing
    //     pre-release here "will polute the list, as this is not something anyone would need to download"),
    //     its files replaced at every push while the version stands. Found through that repository's newest
    //     release. No CRT release is needed for it.
    //
    // Where a build is DOWNLOADED from proves nothing: the attestation names this repository's workflow
    // whichever repository's release carries the bytes (Repository, SignerWorkflowFor).
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
        // The repository whose workflows build the server - the one every attestation must name.
        public const string Repository = "HovKlan-DH/Classic-Repair-Toolbox";

        // The repository the push workflow puts main's builds in, one release per server version, tagged with it.
        public const string MainBuildsRepository = "HovKlan-DH/CRT-Server";

        public const string SourceRef = "refs/heads/main";

        public const string ArchiveName = "crt-server-linux-x64.tar.gz";
        public const string BundleName = ServerReleaseFeed.ArchiveName + ".sigstore.json";
        public const string PointerName = "crt-server-linux-x64.json";

        // The repository a channel's builds are downloaded from - nothing else is ever fetched.
        public static string DownloadRepositoryFor(string channel) =>
            channel == ServerUpdateOutcomes.MainChannel ? ServerReleaseFeed.MainBuildsRepository : ServerReleaseFeed.Repository;

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
        // The pointer a channel reads every minute: the newest release's of its repository - CRT's newest
        // release (a pre-release is never the newest, so never the live service's), or the newest server
        // version's build of main.
        // ###########################################################################################
        public static string PointerUrlFor(string channel) =>
            "https://github.com/" + ServerReleaseFeed.DownloadRepositoryFor(channel) + "/releases/latest/download/" + ServerReleaseFeed.PointerName;

        // A file of `channel`'s release `tag` - a tag the pointer named, held to its channel's shape first
        // (ServerBuildPointer).
        public static string DownloadUrl(string channel, string tag, string fileName) =>
            "https://github.com/" + ServerReleaseFeed.DownloadRepositoryFor(channel) + "/releases/download/" + Uri.EscapeDataString(tag) + "/" + fileName;
    }

    // ###########################################################################################
    // WHICH BUILD A RELEASE CARRIES - "crt-server-linux-x64.json", written by the workflows beside the build
    // (2026-11-09):
    //
    //   tag            the release it is on - a CRT release's version, or for main's builds the server version
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

        // Read for the test instance's channel - a build of main, not a CRT release.
        [JsonIgnore] private bool thisOfMain;

        // ###########################################################################################
        // The pointer in `json` for `channel`, or null with the reason. Its tag must be its channel's - a real
        // release's version for the live service, the server version it carries for the test instance - so no
        // URL is ever built from anything else; its checksum a SHA-256; its version one that can name a folder.
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

            pointer.thisOfMain = channel == ServerUpdateOutcomes.MainChannel;

            bool tagFits = pointer.thisOfMain
                ? ServerReleaseInfo.IsFolderSafeVersion(pointer.Tag) && string.Equals(pointer.Tag, pointer.ServerVersion, StringComparison.Ordinal)
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
            this.thisOfMain
                ? $"main at [{(this.Commit.Length > 7 ? this.Commit[..7] : this.Commit)}]"
                : $"CRT release [{this.Tag}]";
    }
}
