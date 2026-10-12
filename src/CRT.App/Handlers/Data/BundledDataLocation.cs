using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Where an installed CRT keeps the data pack it seeds a new data root from (DataManager's first
    // launch). Windows and Linux carry it as "Data" next to the executable; the macOS release puts
    // it in the app bundle's Contents/Resources (build-and-release.yml), because codesign signs
    // every file under Contents/MacOS as code, keeping a non-code file's signature in extended
    // attributes that an update package does not carry - and the data pack was 11,007 of the
    // 11,255 files there in 3.0.0.
    // ###########################################################################################
    internal static class BundledDataLocation
    {
        public const string FolderName = "Data";

        // ###########################################################################################
        // The bundled data folder for an executable in baseDirectory, or null when there is none
        // (a build from source carries no data pack). Next to the executable wins, so an older
        // macOS layout - and every other platform - reads as before.
        // ###########################################################################################
        public static string? Find(string baseDirectory, Func<string, bool> directoryExists)
        {
            string besideExecutable = Path.Combine(baseDirectory, FolderName);
            if (directoryExists(besideExecutable))
            {
                return besideExecutable;
            }

            string? bundleResources = AppBundleResources(baseDirectory);
            if (bundleResources != null)
            {
                string inResources = Path.Combine(bundleResources, FolderName);
                if (directoryExists(inResources))
                {
                    return inResources;
                }
            }

            return null;
        }

        // ###########################################################################################
        // The bundle's Contents/Resources when baseDirectory is a macOS app bundle's Contents/MacOS,
        // otherwise null - so no other folder's "../Resources" is ever read.
        // ###########################################################################################
        private static string? AppBundleResources(string baseDirectory)
        {
            string trimmed = baseDirectory.TrimEnd('/', '\\');
            DirectoryInfo macos = new DirectoryInfo(trimmed);
            DirectoryInfo? contents = macos.Parent;

            if (!string.Equals(macos.Name, "MacOS", StringComparison.Ordinal) ||
                contents == null ||
                !string.Equals(contents.Name, "Contents", StringComparison.Ordinal))
            {
                return null;
            }

            return Path.Combine(contents.FullName, "Resources");
        }
    }
}
