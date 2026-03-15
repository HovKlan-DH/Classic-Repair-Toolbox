using CRT;
using DataHandling;
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

namespace OnlineHandling
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

        static OnlineServices()
        {
            UserAgent = $"{AppConfig.AppShortName} {AppConfig.AppVersionString}";
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
                http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", UserAgent);

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
            http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", UserAgent);

            string json;
            try
            {
                json = await http.GetStringAsync(AppConfig.ChecksumsUrl);
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
            http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", UserAgent);

            var entries = filter == null ? manifest : manifest.FindAll(e => filter(e.File));

            if (entries.Count == 0)
                return 0;

            // Phase 1: Compare local checksums against the manifest — no file callback here,
            // the splash only shows filenames during actual downloads in Phase 2
            var toDownload = new List<(DataFileEntry Entry, bool IsNew)>();

            foreach (var entry in entries)
            {
                var relativePath = entry.File.Replace('/', Path.DirectorySeparatorChar);
                var localPath = Path.Combine(dataRoot, relativePath);

                if (!File.Exists(localPath))
                {
                    toDownload.Add((entry, true));
                }
                else
                {
                    var localChecksum = await ComputeChecksumAsync(localPath);
                    if (localChecksum != entry.Checksum.ToLowerInvariant())
                        toDownload.Add((entry, false));
                }
            }

            // Phase 2: Download files that are new or changed
            if (toDownload.Count == 0)
            {
                if (entries.Count == 1 && label != null)
                    Logger.Info($"{label} [{entries[0].File}] is up to date");
                else
                    Logger.Info($"All local [{entries.Count}] {label ?? "files"} are up to date");
                onStatus?.Invoke("All local files are up to date");
                onFile?.Invoke(string.Empty);
                return 0;
            }

            // Only log the detail header when syncing multiple files
            if (entries.Count > 1)
                Logger.Info("Individual file sync status:");

            int newCount = 0, updatedCount = 0, failedCount = 0;
            int downloadIndex = 0;

            foreach (var (entry, isNew) in toDownload)
            {
                downloadIndex++;
                onStatus?.Invoke($"Downloading file [{downloadIndex}] of [{toDownload.Count}] from online source");
                onFile?.Invoke(entry.File);
                if (await DownloadFileAsync(http, entry, dataRoot, isNew))
                {
                    if (isNew) newCount++;
                    else updatedCount++;
                }
                else
                {
                    failedCount++;
                }
            }

            if (label != null && entries.Count == 1)
            {
                if (newCount + updatedCount == 1)
                    Logger.Info($"{label} [{entries[0].File}] has been updated");
                // else: failure already logged individually by DownloadFileAsync
            }
            else
            {
                int upToDateCount = entries.Count - toDownload.Count;
                Logger.Info($"Sync completed - [{newCount}] new, [{updatedCount}] updated, [{failedCount}] failed, [{upToDateCount}] up-to-date");
            }

            onStatus?.Invoke($"Sync complete ({newCount} new, {updatedCount} updated, {failedCount} failed)");
            onFile?.Invoke(string.Empty);
            return newCount + updatedCount;
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
        // Downloads a single manifest entry, logs the HTTP status code with New/Updated context,
        // and saves it to the correct local path. Returns true on success, false otherwise.
        // Uses atomic swapping of files.
        // ###########################################################################################
        private static async Task<bool> DownloadFileAsync(HttpClient http, DataFileEntry entry, string dataRoot, bool isNew)
        {
            var relativePath = entry.File.Replace('/', Path.DirectorySeparatorChar);
            var localPath = Path.Combine(dataRoot, relativePath);
            var directory = Path.GetDirectoryName(localPath);

            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            var decodedUrl = Uri.UnescapeDataString(entry.Url);
            var tempPath = localPath + ".tmp";

            try
            {
                using var response = await http.GetAsync(decodedUrl);
                var statusCode = (int)response.StatusCode;

                if (response.IsSuccessStatusCode)
                {
                    var data = await response.Content.ReadAsByteArrayAsync();
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