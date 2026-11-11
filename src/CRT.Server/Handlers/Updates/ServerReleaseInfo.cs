using System.Text.Json;
using System.Text.Json.Serialization;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // WHAT A SERVER BUILD SAYS ABOUT ITSELF - "crt-server-release.json" at the top of the build the
    // release workflow attaches (2026-11-09). Written by the workflow, and so covered by GitHub's
    // attestation with the rest of the build: the version installed and the channel are read from HERE,
    // never from the GitHub release around it, whose name and pre-release flag anybody with write access
    // to the repository could change without building anything.
    //
    //   serverVersion  CRT.Server.csproj's InformationalVersion - the folder it goes into, and what
    //                  /api/health must answer once it runs.
    //   crtVersion     the CRT release it came with.
    //   commit         the commit it was built from.
    //   preRelease     whether that CRT release is a pre-release - never installed on the live service.
    // ###########################################################################################
    internal sealed class ServerReleaseInfo
    {
        public const string FileName = "crt-server-release.json";

        [JsonPropertyName("serverVersion")] public string ServerVersion { get; set; } = string.Empty;

        [JsonPropertyName("crtVersion")] public string CrtVersion { get; set; } = string.Empty;

        [JsonPropertyName("commit")] public string Commit { get; set; } = string.Empty;

        [JsonPropertyName("preRelease")] public bool PreRelease { get; set; }

        // ###########################################################################################
        // The info in `json`, or null with the reason: the server version a three-part SemVer version
        // that can name a folder, the CRT version one too.
        // ###########################################################################################
        public static ServerReleaseInfo? Read(string? json, out string why)
        {
            ServerReleaseInfo? info;

            try
            {
                info = string.IsNullOrWhiteSpace(json) ? null : JsonSerializer.Deserialize<ServerReleaseInfo>(json);
            }
            catch (JsonException ex)
            {
                why = $"{ServerReleaseInfo.FileName} could not be read - {ex.Message}";
                return null;
            }

            if (info == null)
            {
                why = $"The build carries no {ServerReleaseInfo.FileName}.";
                return null;
            }

            if (!ServerReleaseInfo.IsFolderSafeVersion(info.ServerVersion))
            {
                why = $"The build's server version [{info.ServerVersion}] is not a version.";
                return null;
            }

            if (!global::Handlers.DataHandling.CrtVersion.TryParse(info.CrtVersion, out _))
            {
                why = $"The build's CRT version [{info.CrtVersion}] is not a version.";
                return null;
            }

            why = string.Empty;
            return info;
        }

        // A version that is a SemVer version AND safe as a folder name: letters, digits, "." and "-" only.
        public static bool IsFolderSafeVersion(string? version) =>
            !string.IsNullOrEmpty(version) &&
            version.Length <= 64 &&
            version.All(c => char.IsAsciiLetterOrDigit(c) || c is '.' or '-') &&
            global::Handlers.DataHandling.CrtVersion.TryParse(version, out _);
    }
}
