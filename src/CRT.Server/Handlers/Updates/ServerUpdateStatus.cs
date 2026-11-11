using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // WHAT THE UPDATER DID, KEPT FOR THE SERVICE TO SAY (2026-11-09) - "update-status.json" in the
    // instance's releases folder, beside the builds ("crt-server/releases/update-status.json").
    //
    // The updater (ServerUpdater) writes it after every look, replacing it whole. The service reads it
    // when a maintainer opens Account > "Server version" (ReviewEndpoints' server-status, as CRT.Data's
    // ServerUpdatesAnswer). It finds it beside the folder it runs from: the build it runs is
    // "releases/<version>", so the file is in that folder's parent (Locate). A service whose program
    // folder is not a release - one installed by hand - finds none: "not set up".
    //
    // Besides what CRT shows, it keeps what the next look needs: the build already looked at (so the same
    // build is not downloaded again every minute) and a build that failed (never tried again by itself - it
    // would take the service down at every look).
    // ###########################################################################################
    internal sealed class ServerUpdateStatus
    {
        public const string FileName = "update-status.json";

        // The most the service reads - the file is a few hundred bytes.
        public const int MaxBytes = 64 * 1024;

        [JsonPropertyName("channel")] public string Channel { get; set; } = string.Empty;

        [JsonPropertyName("checkedUtc")] public DateTimeOffset? CheckedUtc { get; set; }

        [JsonPropertyName("outcome")] public string Outcome { get; set; } = string.Empty;

        [JsonPropertyName("message")] public string Message { get; set; } = string.Empty;

        [JsonPropertyName("newestVersion")] public string? NewestVersion { get; set; }

        [JsonPropertyName("newestRelease")] public string? NewestRelease { get; set; }

        [JsonPropertyName("installedUtc")] public DateTimeOffset? InstalledUtc { get; set; }

        [JsonPropertyName("installedVersion")] public string? InstalledVersion { get; set; }

        [JsonPropertyName("installedFromVersion")] public string? InstalledFromVersion { get; set; }

        // A build that did not come up: never installed again by itself (ServerUpdater, --retry-failed).
        [JsonPropertyName("failedVersion")] public string? FailedVersion { get; set; }

        // The build looked at last - by its checksum, as its pointer names it - and its version: the next look
        // skips the download for it.
        [JsonPropertyName("examinedSha256")] public string? ExaminedSha256 { get; set; }

        [JsonPropertyName("examinedVersion")] public string? ExaminedVersion { get; set; }

        // When GitHub's attestation check refused the build looked at last - null for any other outcome. Such a refusal
        // may have been gh not reaching Sigstore, so that build is looked at again once ServerUpdater.AttestationRetryAfter
        // has passed (code review, 2026-11-11: refused for good, the live service waited for a hand-run --retry-failed).
        [JsonPropertyName("attestationRefusedUtc")] public DateTimeOffset? AttestationRefusedUtc { get; set; }

        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }

        private static readonly JsonSerializerOptions WriteOptions = new() { WriteIndented = true };

        // ###########################################################################################
        // The status in `path`, or null when there is none or it cannot be read - the updater then starts
        // afresh, which costs one download.
        // ###########################################################################################
        public static ServerUpdateStatus? Read(string path)
        {
            try
            {
                var file = new FileInfo(path);

                if (!file.Exists || file.Length > ServerUpdateStatus.MaxBytes)
                    return null;

                return JsonSerializer.Deserialize<ServerUpdateStatus>(File.ReadAllText(path, Encoding.UTF8));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                return null;
            }
        }

        // ###########################################################################################
        // Writes the status to `path` whole: a file beside it first, then renamed over it, so the service
        // never reads half of one.
        // ###########################################################################################
        public void Write(string path)
        {
            string temporary = path + ".writing";
            File.WriteAllText(temporary, JsonSerializer.Serialize(this, ServerUpdateStatus.WriteOptions), new UTF8Encoding(false));

            if (!OperatingSystem.IsWindows())
                File.SetUnixFileMode(temporary, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead);

            File.Move(temporary, path, overwrite: true);
        }

        // ###########################################################################################
        // Where the service finds the status, given the folder it runs from (its content root): in that
        // folder's parent when it runs a release ("releases/<version>" - the parent is "releases"), or in
        // "releases" beside it when it runs through a link of another name ("crt-server/app"). The first
        // that exists, or null.
        // ###########################################################################################
        public static string? Locate(string contentRoot)
        {
            string folder = Path.TrimEndingDirectorySeparator(Path.GetFullPath(contentRoot));
            string? parent = Path.GetDirectoryName(folder);

            if (parent == null)
                return null;

            string[] candidates =
            [
                Path.Combine(parent, ServerUpdateStatus.FileName),
                Path.Combine(parent, ServerUpdater.ReleasesFolderName, ServerUpdateStatus.FileName),
            ];

            return candidates.FirstOrDefault(File.Exists);
        }

        // ###########################################################################################
        // What the service answers about its updates (CRT.Data's ServerUpdatesAnswer), from the status at
        // `path`: not set up without one, said as unreadable when it cannot be read.
        // ###########################################################################################
        public static ServerUpdatesAnswer AnswerFrom(string? path)
        {
            if (path == null)
                return new ServerUpdatesAnswer(SetUp: false);

            if (ServerUpdateStatus.Read(path) is not ServerUpdateStatus status)
            {
                return new ServerUpdatesAnswer(
                    SetUp: true,
                    Outcome: ServerUpdateOutcomes.Error,
                    Message: "The updater's status file could not be read.");
            }

            return new ServerUpdatesAnswer(
                SetUp: true,
                Channel: status.Channel,
                CheckedUtc: status.CheckedUtc,
                Outcome: status.Outcome,
                Message: status.Message,
                NewestVersion: status.NewestVersion,
                InstalledUtc: status.InstalledUtc,
                InstalledVersion: status.InstalledVersion,
                InstalledFromVersion: status.InstalledFromVersion);
        }
    }
}
