using Handlers.DataHandling;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace CRT
{
    public static class ExternalTargetLauncher
    {
        // ###########################################################################################
        // Opens a validated external target. Only HTTP/HTTPS URLs are allowed, and local files must
        // resolve inside the configured data-root boundary.
        // ###########################################################################################
        public static bool TryOpen(string target, string? dataRootOverride = null)
        {
            if (string.IsNullOrWhiteSpace(target))
                return false;

            if (ExternalTargetLauncher.TryCreateAllowedWebUri(target, out Uri? webUri))
            {
                return ExternalTargetLauncher.TryStart(webUri.AbsoluteUri, $"URL [{webUri.AbsoluteUri}]");
            }

            string dataRoot = !string.IsNullOrWhiteSpace(dataRootOverride)
                ? dataRootOverride
                : DataManager.DataRoot;

            if (ExternalTargetLauncher.TryResolveDataRootScopedFilePath(target, dataRoot, out string? localPath))
            {
                return ExternalTargetLauncher.TryStart(localPath, $"local file [{localPath}]");
            }

            Debug.WriteLine($"Rejected external target outside allowed scope: {target}");
            return false;
        }

        // ###########################################################################################
        // Validates that a target string is an allowed HTTP/HTTPS URL.
        // ###########################################################################################
        private static bool TryCreateAllowedWebUri(string target, out Uri? uri)
        {
            uri = null;

            if (!Uri.TryCreate(target.Trim(), UriKind.Absolute, out Uri? candidateUri))
                return false;

            if (!string.Equals(candidateUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(candidateUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            uri = candidateUri;
            return true;
        }

        // ###########################################################################################
        // Resolves a local file path and rejects anything outside the configured data-root.
        // Relative paths are resolved against data-root; absolute paths must still stay inside it.
        // ###########################################################################################
        private static bool TryResolveDataRootScopedFilePath(string target, string dataRoot, out string? localPath)
        {
            localPath = null;

            if (string.IsNullOrWhiteSpace(dataRoot) || string.IsNullOrWhiteSpace(target))
                return false;

            try
            {
                string normalizedDataRoot = Path.GetFullPath(dataRoot);
                string normalizedTargetInput = target.Trim().Replace('/', Path.DirectorySeparatorChar);

                string normalizedTarget = Path.IsPathRooted(normalizedTargetInput)
                    ? Path.GetFullPath(normalizedTargetInput)
                    : Path.GetFullPath(Path.Combine(normalizedDataRoot, normalizedTargetInput));

                string normalizedDataRootWithSeparator = ExternalTargetLauncher.AppendDirectorySeparator(normalizedDataRoot);
                StringComparison pathComparison = ExternalTargetLauncher.GetPathComparison();

                if (string.Equals(normalizedTarget, normalizedDataRoot, pathComparison))
                    return false;

                if (!normalizedTarget.StartsWith(normalizedDataRootWithSeparator, pathComparison))
                    return false;

                if (!File.Exists(normalizedTarget))
                    return false;

                localPath = normalizedTarget;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ###########################################################################################
        // Starts a validated target through the operating system shell.
        // ###########################################################################################
        private static bool TryStart(string fileName, string description)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = fileName,
                    UseShellExecute = true
                });

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open {description}: {ex.Message}");
                return false;
            }
        }

        // ###########################################################################################
        // Appends a trailing directory separator when missing so StartsWith path checks stay safe.
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
    }
}