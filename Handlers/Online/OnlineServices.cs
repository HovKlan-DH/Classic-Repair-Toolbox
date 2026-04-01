using CRT;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Handlers.OnlineHandling
{
    internal sealed class DataFileEntry
    {
        [JsonPropertyName("file")] public string File { get; init; } = string.Empty;
        [JsonPropertyName("checksum")] public string Checksum { get; init; } = string.Empty;
        [JsonPropertyName("url")] public string Url { get; init; } = string.Empty;
    }

    public static class OnlineServices
    {
        private static readonly string UserAgent;
        private static readonly Uri TrustedManifestUri;
        private static readonly string TrustedDownloadAuthority;

        static OnlineServices()
        {
            OnlineServices.UserAgent = $"{AppConfig.AppShortName} {AppConfig.AppVersionString}";
            OnlineServices.TrustedManifestUri = new Uri(AppConfig.ChecksumsUrl, UriKind.Absolute);
            OnlineServices.TrustedDownloadAuthority = OnlineServices.TrustedManifestUri.Authority;
        }

        // ###########################################################################################
        // Asks server for newest version.
        // Reports the app version and OS details. Runs silently - failures are only logged.
        // ###########################################################################################
        public static async Task CheckInVersionAsync()
        {
            try
            {
                var osHighLevel = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "Windows"
                    : RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "macOS"
                    : RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ? "Linux"
                    : RuntimeInformation.IsOSPlatform(OSPlatform.FreeBSD) ? "FreeBSD"
                    : "Unknown";

                var osVersion = RuntimeInformation.OSDescription;

                var cpu = RuntimeInformation.ProcessArchitecture switch
                {
                    Architecture.X64 => "64-bit",
                    Architecture.X86 => "32-bit",
                    Architecture.Arm64 => "ARM 64-bit",
                    Architecture.Arm => "ARM 32-bit",
                    var a => a.ToString()
                };

                using var http = new HttpClient { Timeout = AppConfig.ApiTimeout };
                http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", OnlineServices.UserAgent);

                var payload = new List<KeyValuePair<string, string>>
                {
                    new("control", AppConfig.AppShortName),
                    new("osHighlevel", osHighLevel),
                    new("osVersion", osVersion),
                    new("cpu", cpu),
                };

                using var response = await http.PostAsync(AppConfig.CheckVersionUrl, new FormUrlEncodedContent(payload));
                var responseBody = await response.Content.ReadAsStringAsync();
                Logger.Info($"Online check-in completed:");
                Logger.Info($"    HTTP:[{(int)response.StatusCode}]");
                Logger.Info($"    Sent:[{osHighLevel}]");
                Logger.Info($"    Sent:[{osVersion}]");
                Logger.Info($"    Sent:[{cpu}]");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Online check-in failed: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Fetches and parses the online checksum manifest. Returns null on failure.
        // onStatus: optional callback for failure messages.
        // ###########################################################################################
        internal static async Task<List<DataFileEntry>?> FetchManifestAsync(Action<string>? onStatus = null)
        {
            using var http = new HttpClient { Timeout = AppConfig.ApiTimeout };
            http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", OnlineServices.UserAgent);

            string json;
            try
            {
                if (!string.Equals(OnlineServices.TrustedManifestUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Critical($"Manifest URL must use HTTPS: [{OnlineServices.TrustedManifestUri}]");
                    onStatus?.Invoke("Sync failed - see log");
                    return null;
                }

                json = await http.GetStringAsync(OnlineServices.TrustedManifestUri);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Failed to fetch checksum manifest: {ex.Message}");
                onStatus?.Invoke("Sync failed - see log");
                return null;
            }

            try
            {
                var entries = JsonSerializer.Deserialize<List<DataFileEntry>>(json);

                if (entries is null || entries.Count == 0)
                {
                    Logger.Warning("Checksum manifest is empty");
                    onStatus?.Invoke("No files in manifest");
                    return null;
                }

                Logger.Info($"Online source checksum manifest fetched:");
                Logger.Info($"    [{entries.Count}] files available online");
                return entries;
            }
            catch (Exception ex)
            {
                Logger.Critical($"Failed to parse checksum manifest: {ex.Message}");
                onStatus?.Invoke("Sync failed - see log");
                return null;
            }
        }

        // ###########################################################################################
        // Compares entries from a pre-fetched manifest against local files and downloads anything
        // that is missing or has changed. Runs in two phases: verify checksums, then download.
        // filter:   optional predicate on the file path — when null, all entries are processed.
        // onStatus: optional callback for general progress messages.
        // onFile:   optional callback fired with each file path as it is being downloaded.
        // label:    optional context label used to improve log messages (e.g. "board data files").
        // Returns the number of files that were successfully new or updated.
        // ###########################################################################################
        internal static async Task<int> SyncFilesAsync(
            List<DataFileEntry> manifest,
            string dataRoot,
            Func<string, bool>? filter = null,
            Action<string>? onStatus = null,
            Action<string>? onFile = null,
            string? label = null)
        {
            using var http = new HttpClient { Timeout = AppConfig.DownloadTimeout };
            http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", OnlineServices.UserAgent);

            var entries = filter == null ? manifest : manifest.FindAll(e => filter(e.File));

            if (entries.Count == 0)
                return 0;

            // Phase 1: Validate manifest entries and compare local checksums against the manifest —
            // no file callback here, the splash only shows filenames during actual downloads in Phase 2
            var validEntries = new List<DataFileEntry>();
            var toDownload = new List<(DataFileEntry Entry, string LocalPath, Uri DownloadUri, string ExpectedChecksum, bool IsNew)>();
            int invalidCount = 0;

            foreach (var entry in entries)
            {
                if (!OnlineServices.TryValidateManifestEntry(
                    entry,
                    dataRoot,
                    out string localPath,
                    out Uri? downloadUri,
                    out string expectedChecksum,
                    out string failureReason))
                {
                    invalidCount++;
                    Logger.Warning($"[{entry.File}] [Rejected] [{failureReason}]");
                    continue;
                }

                validEntries.Add(entry);

                if (!File.Exists(localPath))
                {
                    toDownload.Add((entry, localPath, downloadUri!, expectedChecksum, true));
                }
                else
                {
                    var localChecksum = await OnlineServices.ComputeChecksumAsync(localPath);
                    if (!string.Equals(localChecksum, expectedChecksum, StringComparison.Ordinal))
                        toDownload.Add((entry, localPath, downloadUri!, expectedChecksum, false));
                }
            }

            // Phase 2: Download files that are new or changed
            if (toDownload.Count == 0)
            {
                if (validEntries.Count == 0)
                {
                    Logger.Warning($"No valid {(label ?? "files")} were accepted from the manifest");
                    onStatus?.Invoke("No valid files in manifest");
                    onFile?.Invoke(string.Empty);
                    return 0;
                }

                if (validEntries.Count == 1 && label != null)
                    Logger.Info($"{label} [{validEntries[0].File}] is up to date");
                else
                    Logger.Info($"All local [{validEntries.Count}] valid {label ?? "files"} are up to date");

                onStatus?.Invoke("All local files are up to date");
                onFile?.Invoke(string.Empty);
                return 0;
            }

            // Only log the detail header when syncing multiple files
            if (validEntries.Count > 1)
                Logger.Info("Individual file sync status:");

            int newCount = 0, updatedCount = 0, failedCount = 0;
            int downloadIndex = 0;

            foreach (var (entry, localPath, downloadUri, expectedChecksum, isNew) in toDownload)
            {
                downloadIndex++;
                onStatus?.Invoke($"Downloading file [{downloadIndex}] of [{toDownload.Count}] from online source");
                onFile?.Invoke(entry.File);

                if (await OnlineServices.DownloadFileAsync(http, entry, localPath, downloadUri, expectedChecksum, isNew))
                {
                    if (isNew)
                        newCount++;
                    else
                        updatedCount++;
                }
                else
                {
                    failedCount++;
                }
            }

            if (label != null && validEntries.Count == 1)
            {
                if (newCount + updatedCount == 1)
                    Logger.Info($"{label} [{validEntries[0].File}] has been updated");
                // else: failure already logged individually by DownloadFileAsync
            }
            else
            {
                int upToDateCount = validEntries.Count - toDownload.Count;
                Logger.Info($"Sync completed - [{newCount}] new, [{updatedCount}] updated, [{failedCount}] failed, [{invalidCount}] invalid, [{upToDateCount}] up-to-date");
            }

            onStatus?.Invoke($"Sync complete ({newCount} new, {updatedCount} updated, {failedCount} failed, {invalidCount} invalid)");
            onFile?.Invoke(string.Empty);
            return newCount + updatedCount;
        }

        // ###########################################################################################
        // Validates one manifest entry before any local file access or download is attempted.
        // Ensures path containment, checksum shape, and trusted HTTPS URL requirements.
        // ###########################################################################################
        private static bool TryValidateManifestEntry(
            DataFileEntry entry,
            string dataRoot,
            out string localPath,
            out Uri? downloadUri,
            out string expectedChecksum,
            out string failureReason)
        {
            localPath = string.Empty;
            downloadUri = null;
            expectedChecksum = string.Empty;
            failureReason = string.Empty;

            if (!OnlineServices.TryResolveValidatedLocalPath(dataRoot, entry.File, out localPath, out failureReason))
                return false;

            if (!OnlineServices.TryNormalizeManifestChecksum(entry.Checksum, out expectedChecksum))
            {
                failureReason = "manifest checksum is missing or invalid";
                return false;
            }

            if (!OnlineServices.TryCreateTrustedDownloadUri(entry.Url, out downloadUri, out failureReason))
                return false;

            return true;
        }

        // ###########################################################################################
        // Resolves one manifest file path and rejects rooted or traversing paths outside dataRoot.
        // ###########################################################################################
        private static bool TryResolveValidatedLocalPath(
            string dataRoot,
            string manifestFile,
            out string localPath,
            out string failureReason)
        {
            localPath = string.Empty;
            failureReason = string.Empty;

            if (string.IsNullOrWhiteSpace(dataRoot))
            {
                failureReason = "data root is empty";
                return false;
            }

            if (string.IsNullOrWhiteSpace(manifestFile))
            {
                failureReason = "manifest file path is empty";
                return false;
            }

            try
            {
                string normalizedDataRoot = Path.GetFullPath(dataRoot);
                string normalizedRelativePath = manifestFile.Replace('/', Path.DirectorySeparatorChar).Trim();

                if (Path.IsPathRooted(normalizedRelativePath))
                {
                    failureReason = "manifest file path must be relative";
                    return false;
                }

                string candidateLocalPath = Path.GetFullPath(Path.Combine(normalizedDataRoot, normalizedRelativePath));
                string normalizedDataRootWithSeparator = OnlineServices.AppendDirectorySeparator(normalizedDataRoot);
                StringComparison pathComparison = OnlineServices.GetPathComparison();

                if (string.Equals(candidateLocalPath, normalizedDataRoot, pathComparison))
                {
                    failureReason = "manifest file path resolves to the data root";
                    return false;
                }

                if (!candidateLocalPath.StartsWith(normalizedDataRootWithSeparator, pathComparison))
                {
                    failureReason = "manifest file path escapes data root";
                    return false;
                }

                localPath = candidateLocalPath;
                return true;
            }
            catch (Exception ex)
            {
                failureReason = $"manifest file path is invalid: {ex.Message}";
                return false;
            }
        }

        // ###########################################################################################
        // Normalizes a manifest SHA-256 checksum and rejects malformed values.
        // ###########################################################################################
        private static bool TryNormalizeManifestChecksum(string checksum, out string normalizedChecksum)
        {
            normalizedChecksum = string.Empty;

            if (string.IsNullOrWhiteSpace(checksum))
                return false;

            string trimmedChecksum = checksum.Trim();

            if (trimmedChecksum.Length != 64)
                return false;

            for (int i = 0; i < trimmedChecksum.Length; i++)
            {
                char ch = trimmedChecksum[i];
                bool isHex =
                    (ch >= '0' && ch <= '9') ||
                    (ch >= 'a' && ch <= 'f') ||
                    (ch >= 'A' && ch <= 'F');

                if (!isHex)
                    return false;
            }

            normalizedChecksum = trimmedChecksum.ToLowerInvariant();
            return true;
        }

        // ###########################################################################################
        // Validates that a manifest download URL is absolute HTTPS and stays on the trusted authority.
        // ###########################################################################################
        private static bool TryCreateTrustedDownloadUri(string url, out Uri? downloadUri, out string failureReason)
        {
            downloadUri = null;
            failureReason = string.Empty;

            if (string.IsNullOrWhiteSpace(url))
            {
                failureReason = "download URL is empty";
                return false;
            }

            if (!Uri.TryCreate(url.Trim(), UriKind.Absolute, out Uri? candidateUri))
            {
                failureReason = "download URL is not a valid absolute URI";
                return false;
            }

            if (!string.Equals(candidateUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                failureReason = "download URL must use HTTPS";
                return false;
            }

            if (!string.Equals(candidateUri.Authority, OnlineServices.TrustedDownloadAuthority, StringComparison.OrdinalIgnoreCase))
            {
                failureReason = $"download URL authority is not trusted: [{candidateUri.Authority}]";
                return false;
            }

            downloadUri = candidateUri;
            return true;
        }

        // ###########################################################################################
        // Computes the SHA-256 checksum of a local file and returns it as a lowercase hex string.
        // ###########################################################################################
        private static async Task<string> ComputeChecksumAsync(string filePath)
        {
            await using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            /*
            var hash = await SHA256.HashDataAsync(stream);
            return Convert.ToHexStringLower(hash);
            */
            // .NET6 compliant
            using var sha256 = SHA256.Create();
            var hash = await Task.Run(() => sha256.ComputeHash(stream));
            return Convert.ToHexString(hash).ToLowerInvariant();
        }

        // ###########################################################################################
        // Computes the SHA-256 checksum of downloaded bytes and returns it as a lowercase hex string.
        // ###########################################################################################
        private static string ComputeChecksum(byte[] data)
        {
            using var sha256 = SHA256.Create();
            var hash = sha256.ComputeHash(data);
            return Convert.ToHexString(hash).ToLowerInvariant();
        }

        // ###########################################################################################
        // Appends a trailing directory separator when missing so StartsWith path checks are safe.
        // ###########################################################################################
        private static string AppendDirectorySeparator(string path)
        {
            if (string.IsNullOrEmpty(path))
                return Path.DirectorySeparatorChar.ToString();

            if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        // ###########################################################################################
        // Returns the correct filesystem path comparison for the current operating system.
        // ###########################################################################################
        private static StringComparison GetPathComparison()
        {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
        }

        // ###########################################################################################
        // Downloads a single manifest entry, verifies the downloaded checksum, then atomically swaps
        // the local file into place only when all validation checks succeed.
        // ###########################################################################################
        private static async Task<bool> DownloadFileAsync(
            HttpClient http,
            DataFileEntry entry,
            string localPath,
            Uri downloadUri,
            string expectedChecksum,
            bool isNew)
        {
            string? directory = Path.GetDirectoryName(localPath);

            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            string tempPath = localPath + ".tmp";

            try
            {
                using var response = await http.GetAsync(downloadUri);
                var statusCode = (int)response.StatusCode;

                if (response.IsSuccessStatusCode)
                {
                    var data = await response.Content.ReadAsByteArrayAsync();
                    string actualChecksum = OnlineServices.ComputeChecksum(data);

                    if (!string.Equals(actualChecksum, expectedChecksum, StringComparison.Ordinal))
                    {
                        Logger.Warning($"[{entry.File}] [{statusCode}] [Checksum mismatch] [Expected:{expectedChecksum}] [Actual:{actualChecksum}]");
                        try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                        return false;
                    }

                    await File.WriteAllBytesAsync(tempPath, data);
                    File.Move(tempPath, localPath, overwrite: true);
                    Logger.Info($"[{entry.File}] [{statusCode}] [{(isNew ? "New" : "Updated")}]");
                    return true;
                }

                Logger.Warning($"[{entry.File}] [{statusCode}]");
                return false;
            }
            catch (Exception ex)
            {
                Logger.Warning($"[{entry.File}] [Exception] [{ex.Message}]");

                // Clean up temp file if it was left behind
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }

                return false;
            }
        }
    }
}