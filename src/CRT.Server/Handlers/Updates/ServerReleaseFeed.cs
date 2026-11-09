using System.Globalization;
using System.Text.Json;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // WHERE A SERVER BUILD COMES FROM, AND WHICH ONE IS NEXT (owner decisions, 2026-11-09: the server
    // builds come "With CRT releases", checked by "GitHub attestations").
    //
    // The release workflow builds the server beside CRT and attaches it to the same GitHub Release:
    // ArchiveName, the build, and BundleName, GitHub's attestation that this repository's release
    // workflow built exactly those bytes. The live service installs from releases, the test instance
    // from pre-releases too - so each CRT talks to a server built from the same commit as itself.
    //
    // Everything here is fixed in the code, never read from a setting: the box trusts builds of THIS
    // repository's release workflow on main, and nothing a setting could point elsewhere.
    // ###########################################################################################
    internal static class ServerReleaseFeed
    {
        public const string Repository = "HovKlan-DH/Classic-Repair-Toolbox";

        // The workflow whose builds are trusted - the one Dennis runs from the Actions tab - on main.
        public const string SignerWorkflow = ServerReleaseFeed.Repository + "/.github/workflows/build-and-release.yml";
        public const string SourceRef = "refs/heads/main";

        public const string ArchiveName = "crt-server-linux-x64.tar.gz";
        public const string BundleName = ServerReleaseFeed.ArchiveName + ".sigstore.json";

        // The newest releases first; a pre-release run deletes the pre-releases before it, so a page holds them all.
        public const string ReleasesUrl = "https://api.github.com/repos/" + ServerReleaseFeed.Repository + "/releases?per_page=30";

        // ###########################################################################################
        // *** THE LOOK EVERY MINUTE (owner request, 2026-11-09: "Instead of updating only every 15 minutes, can it
        // then update every minute?"). *** GitHub lets a box with no login ask its API 60 times an hour, and the
        // two instances asking every minute would ask 120 - so a minute's look reads the releases' Atom feed, a
        // page of github.com the API's limit does not count, and only a change there (or ServerUpdater's
        // safety net) makes the updater ask the API. The feed lists every published release, pre-releases too.
        // ###########################################################################################
        public const string FeedUrl = "https://github.com/" + ServerReleaseFeed.Repository + "/releases.atom";

        // ###########################################################################################
        // What the feed says, as one string that changes whenever a release is published, edited or deleted:
        // every entry's id and time, hashed. Null when `atomXml` is no Atom feed (an error page, a DTD).
        // ###########################################################################################
        public static string? FingerprintOf(string? atomXml)
        {
            if (string.IsNullOrWhiteSpace(atomXml))
                return null;

            try
            {
                // A DTD refused outright, nothing resolved - no entity is ever expanded. (XDocument.Parse alone
                // PROCESSES a DTD - ServerReleaseFeedTests caught that.)
                var settings = new System.Xml.XmlReaderSettings
                {
                    DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = 4_000_000,
                };

                using var reader = System.Xml.XmlReader.Create(new StringReader(atomXml), settings);
                System.Xml.Linq.XDocument document = System.Xml.Linq.XDocument.Load(reader);
                System.Xml.Linq.XNamespace atom = "http://www.w3.org/2005/Atom";

                if (document.Root is not { } feed || feed.Name != atom + "feed")
                    return null;

                string entries = string.Join(
                    "\n",
                    feed.Elements(atom + "entry").Select(entry => $"{entry.Element(atom + "id")?.Value}\t{entry.Element(atom + "updated")?.Value}"));

                return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(entries))).ToLowerInvariant();
            }
            catch (System.Xml.XmlException)
            {
                return null;
            }
        }

        // Where a release's files are downloaded from - nothing else is ever fetched.
        public const string DownloadPrefix = "https://github.com/" + ServerReleaseFeed.Repository + "/releases/download/";

        public const long ArchiveMaxBytes = 200L * 1024 * 1024;
        public const long BundleMaxBytes = 1L * 1024 * 1024;

        // One release that carries a server build: its CRT tag, whether it is a pre-release, and its two files.
        public sealed record Candidate(
            string Tag,
            bool IsPreRelease,
            DateTimeOffset PublishedUtc,
            string ArchiveUrl,
            long ArchiveAssetId,
            long ArchiveBytes,
            string BundleUrl);

        // ###########################################################################################
        // The newest release with a server build for `channel` (ServerUpdateOutcomes' channel words) in
        // GitHub's answer to ReleasesUrl, or null with the reason - `readable` false when the answer was
        // no list of releases at all. A draft is never one; a pre-release only for the pre-releases
        // channel; a release whose files lie anywhere but this repository's downloads, or whose build is
        // larger than ArchiveMaxBytes, is passed over.
        // ###########################################################################################
        public static Candidate? Newest(string releasesJson, string channel, out string why, out bool readable)
        {
            bool takesPreReleases = string.Equals(channel, ServerUpdateOutcomes.PreReleasesChannel, StringComparison.Ordinal);
            var candidates = new List<Candidate>();
            readable = false;

            try
            {
                using JsonDocument document = JsonDocument.Parse(releasesJson);

                if (document.RootElement.ValueKind != JsonValueKind.Array)
                {
                    why = "GitHub's list of releases was not a list.";
                    return null;
                }

                readable = true;

                foreach (JsonElement release in document.RootElement.EnumerateArray())
                {
                    if (ServerReleaseFeed.CandidateOf(release) is { } candidate &&
                        (takesPreReleases || !candidate.IsPreRelease))
                    {
                        candidates.Add(candidate);
                    }
                }
            }
            catch (JsonException ex)
            {
                why = $"GitHub's list of releases could not be read - {ex.Message}";
                return null;
            }

            Candidate? newest = candidates.OrderByDescending(candidate => candidate.PublishedUtc).FirstOrDefault();

            why = newest == null
                ? $"No {(takesPreReleases ? "release or pre-release" : "release")} on GitHub carries a server build yet."
                : string.Empty;

            return newest;
        }

        private static Candidate? CandidateOf(JsonElement release)
        {
            if (release.ValueKind != JsonValueKind.Object ||
                ServerReleaseFeed.Bool(release, "draft") ||
                ServerReleaseFeed.String(release, "tag_name") is not { Length: > 0 } tag ||
                !release.TryGetProperty("assets", out JsonElement assets) ||
                assets.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            JsonElement? archive = null;
            JsonElement? bundle = null;

            foreach (JsonElement asset in assets.EnumerateArray())
            {
                string? name = asset.ValueKind == JsonValueKind.Object ? ServerReleaseFeed.String(asset, "name") : null;

                if (string.Equals(name, ServerReleaseFeed.ArchiveName, StringComparison.Ordinal))
                    archive = asset;
                else if (string.Equals(name, ServerReleaseFeed.BundleName, StringComparison.Ordinal))
                    bundle = asset;
            }

            if (archive is not { } archiveAsset || bundle is not { } bundleAsset)
                return null;

            string? archiveUrl = ServerReleaseFeed.String(archiveAsset, "browser_download_url");
            string? bundleUrl = ServerReleaseFeed.String(bundleAsset, "browser_download_url");
            long archiveBytes = ServerReleaseFeed.Long(archiveAsset, "size") ?? -1;

            if (!ServerReleaseFeed.IsOurDownload(archiveUrl, tag, ServerReleaseFeed.ArchiveName) ||
                !ServerReleaseFeed.IsOurDownload(bundleUrl, tag, ServerReleaseFeed.BundleName) ||
                archiveBytes <= 0 || archiveBytes > ServerReleaseFeed.ArchiveMaxBytes ||
                ServerReleaseFeed.Long(archiveAsset, "id") is not long assetId)
            {
                return null;
            }

            DateTimeOffset published = DateTimeOffset.TryParse(
                ServerReleaseFeed.String(release, "published_at") ?? ServerReleaseFeed.String(release, "created_at"),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal,
                out DateTimeOffset at)
                ? at
                : DateTimeOffset.MinValue;

            return new Candidate(tag, ServerReleaseFeed.Bool(release, "prerelease"), published, archiveUrl!, assetId, archiveBytes, bundleUrl!);
        }

        // ###########################################################################################
        // A download address of this repository's release `tag`, for `fileName` - exactly, so a list
        // somebody tampered with cannot send the updater to another place. HTTPS, no "..", no query.
        // ###########################################################################################
        public static bool IsOurDownload(string? url, string tag, string fileName) =>
            url != null &&
            string.Equals(url, ServerReleaseFeed.DownloadPrefix + Uri.EscapeDataString(tag) + "/" + fileName, StringComparison.Ordinal) &&
            !tag.Contains("..", StringComparison.Ordinal) &&
            !tag.Contains('/') &&
            !tag.Contains('\\');

        private static string? String(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;

        private static bool Bool(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.True;

        private static long? Long(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out long number)
                ? number
                : null;
    }
}
