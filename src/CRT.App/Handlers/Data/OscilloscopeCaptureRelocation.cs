using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // OSCILLOSCOPE CAPTURES AND THE DATA FOLDER (code review, 2026-10-13).
    //
    // Until then a capture was saved, unless the user picked a folder, straight into the synced
    // "Data" folder - the one "Allow deletion of orphan and non-used files" cleans, and every
    // capture is a file no manifest lists. With that box ticked, the next launch deleted them.
    //
    // Three things put that right, and the rules for all three are here:
    //   - the default capture folder is beside Data now (UserSettings.OscilloscopeImageFolder);
    //   - the cleanup never deletes anything inside the capture folder, wherever it is
    //     (IsSameOrInside, used by DataManager's cleanup);
    //   - captures already sitting in the data root's TOP level - where the old default put them -
    //     are moved to the capture folder before any cleanup can run (MoveCapturesOutOfDataRoot).
    //     A file it is not sure of is left where it is, AND kept away from the cleanup.
    // ###########################################################################################
    internal static class OscilloscopeCaptureRelocation
    {
        // What a move found: the files it moved (where from, where to) and the files it left in the
        // data root that the cleanup must not delete.
        internal sealed record Outcome(
            IReadOnlyList<(string From, string To)> Moved,
            IReadOnlyList<string> LeftInPlace);

        // ###########################################################################################
        // What the data folder's cleanup needs to know about captures: the folder they are saved in
        // now (nothing inside it is ever deleted), and where captures found in the data root's top
        // level go - null when the capture folder is itself inside the data folder (a folder the
        // user chose there): then they are where the user wants them, and stay.
        // ###########################################################################################
        internal sealed record Plan(string CaptureFolder, string? MoveTarget)
        {
            internal static Plan ForCurrentSettings(string dataRoot) =>
                Plan.ForFolder(UserSettings.OscilloscopeImageFolder?.Trim() ?? string.Empty, dataRoot);

            // Captures found in the data root go to the capture folder - the new default for nearly
            // everybody, and for a user who chose a folder elsewhere the one the Oscilloscope tab
            // opens - unless that folder is inside the data folder itself.
            internal static Plan ForFolder(string captureFolder, string dataRoot) =>
                new(
                    captureFolder,
                    string.IsNullOrWhiteSpace(captureFolder) || OscilloscopeCaptureRelocation.IsSameOrInside(captureFolder, dataRoot)
                        ? null
                        : captureFolder);
        }

        private static readonly string[] ImageExtensions =
            [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"];

        // ###########################################################################################
        // Whether a file name is one a capture is saved under: "<board label>_<pin or name>.png" or
        // "<board label>_<pin or name>_<region>.png" (TabOscilloscope.Capture.cs builds it).
        // ###########################################################################################
        internal static bool LooksLikeCapture(string fileName) =>
            !string.IsNullOrWhiteSpace(fileName) &&
            string.Equals(Path.GetExtension(fileName), ".png", StringComparison.OrdinalIgnoreCase) &&
            Path.GetFileNameWithoutExtension(fileName).IndexOf('_', StringComparison.Ordinal) > 0;

        // An image that is not shaped like a capture - a photo, a screenshot saved by hand. Not
        // ours to move, and not ours to delete either: it could be the user's.
        internal static bool IsImageInDoubt(string fileName) =>
            !OscilloscopeCaptureRelocation.LooksLikeCapture(fileName) &&
            ImageExtensions.Contains(Path.GetExtension(fileName), StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // Whether `path` is `folder` itself or anything inside it. Both are made full first; the
        // comparison ignores case where the file system does (Windows, macOS).
        // ###########################################################################################
        internal static bool IsSameOrInside(string path, string folder)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(folder))
            {
                return false;
            }

            StringComparison comparison = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;

            string fullPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
            string fullFolder = Path.TrimEndingDirectorySeparator(Path.GetFullPath(folder));

            return string.Equals(fullPath, fullFolder, comparison) ||
                   fullPath.StartsWith(fullFolder + Path.DirectorySeparatorChar, comparison);
        }

        // ###########################################################################################
        // Moves every file in `dataRoot`'s TOP level that looks like a capture and that the manifest
        // does not list (`isInManifest` takes the file NAME, which is its path relative to the data
        // root) into `targetFolder`. Never overwrites: a name already taken there is left where it
        // is, as is a file that could not be moved. Images that do not look like captures are left
        // too. Everything left is in LeftInPlace, for the cleanup to keep away from.
        //
        // Never throws: a folder that cannot be listed moves nothing and leaves nothing to report.
        // ###########################################################################################
        internal static Outcome MoveCapturesOutOfDataRoot(
            string dataRoot,
            string targetFolder,
            Func<string, bool> isInManifest)
        {
            var moved = new List<(string From, string To)>();
            var leftInPlace = new List<string>();

            if (string.IsNullOrWhiteSpace(dataRoot) || !Directory.Exists(dataRoot) ||
                string.IsNullOrWhiteSpace(targetFolder) ||
                OscilloscopeCaptureRelocation.IsSameOrInside(targetFolder, dataRoot))
            {
                return new Outcome(moved, leftInPlace);
            }

            List<string> topLevelFiles;

            try
            {
                topLevelFiles = Directory.EnumerateFiles(dataRoot).ToList();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Oscilloscope captures in the data folder could not be looked for - [{ex.Message}]");
                return new Outcome(moved, leftInPlace);
            }

            foreach (string file in topLevelFiles)
            {
                string name = Path.GetFileName(file);

                if (isInManifest(name))
                {
                    continue;
                }

                if (OscilloscopeCaptureRelocation.IsImageInDoubt(name))
                {
                    leftInPlace.Add(file);
                    continue;
                }

                if (!OscilloscopeCaptureRelocation.LooksLikeCapture(name))
                {
                    continue;
                }

                string target = Path.Combine(targetFolder, name);

                try
                {
                    if (File.Exists(target) || Directory.Exists(target))
                    {
                        Logger.Warning($"Oscilloscope capture left in the data folder - [{name}] is already taken in [{targetFolder}]");
                        leftInPlace.Add(file);
                        continue;
                    }

                    Directory.CreateDirectory(targetFolder);
                    File.Move(file, target);
                    moved.Add((file, target));
                    Logger.Info($"Moved oscilloscope capture out of the data folder: [{file}] -> [{target}]");
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Oscilloscope capture could not be moved out of the data folder - [{file}] [{ex.Message}]");
                    leftInPlace.Add(file);
                }
            }

            return new Outcome(moved, leftInPlace);
        }
    }
}
