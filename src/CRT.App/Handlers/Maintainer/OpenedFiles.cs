using System;
using CRT;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // OPENING A FILE FROM THE SERVER in the program the computer uses for it - the table's hover
    // card (a PDF), and the file tree (owner request, 2026-09-28: "I need to check the Excel file -
    // does its format look correct, and JSON etc."). The bytes are saved into a folder of their
    // own under the temp folder and handed to the operating system.
    //
    // *** ONLY KNOWN TYPES, AND A WEB PAGE AS TEXT. *** See TryGetOpenName. The table opens only
    // what a submission may carry (ReviewTableFiles.TryGetOpenName); the tree also opens a board's
    // own files - the workbook, the highlight file and KiCad data - which the server writes or
    // takes from the data itself, never from an upload.
    // ###########################################################################################
    public static class OpenedFiles
    {
        // What a board holds beyond what a submission may upload: its workbooks, its highlight file
        // and its KiCad data (the types CRT reads - ComponentListBuilder.IsSupportedKiCadRawFile).
        private static readonly IReadOnlySet<string> BoardFileExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx", ".json", ".kicad_pcb", ".kicad_sch", ".kicad_pro"
        };

        // ###########################################################################################
        // The name a tree file is saved as before it is opened: only a type a submission may carry
        // or a board's own file type, a web page as .txt (opened as .html it would run whatever
        // script is in it, in the maintainer's browser), only the file's own name - never its
        // folders - and no character a file system refuses.
        // ###########################################################################################
        public static bool TryGetOpenName(string? path, out string fileName)
        {
            if (ReviewTableFiles.TryGetOpenName(path, out fileName))
                return true;

            string name = Path.GetFileName((path ?? string.Empty).Replace('\\', '/').Split('/')[^1]).Trim();

            if (name.Length == 0 || !OpenedFiles.BoardFileExtensions.Contains(Path.GetExtension(name)))
            {
                fileName = string.Empty;
                return false;
            }

            foreach (char invalid in Path.GetInvalidFileNameChars())
                name = name.Replace(invalid, '_');

            fileName = name;
            return true;
        }

        // ###########################################################################################
        // Where opened files are written: the app's own folder in the temp folder, then
        // "Maintainer". It was "%TEMP%\CRT Maintainer" while the Maintainer tab was a separate
        // application (until 2026-09-29); CRT's name is used now, so everything CRT leaves in the
        // temp folder sits under one folder of its own name.
        // ###########################################################################################
        internal static string TempRoot =>
            Path.Combine(Path.GetTempPath(), AppConfig.AppFolderName, "Maintainer");

        // ###########################################################################################
        // Saves `bytes` as `fileName` in a new folder under TempRoot and
        // hands it to `launch`. Null when it opened, otherwise the sentence to show. Earlier opens'
        // folders are swept first - see Sweep.
        // ###########################################################################################
        public static async Task<string?> SaveAndLaunchAsync(byte[] bytes, string fileName, Func<string, Task<bool>> launch)
        {
            ArgumentNullException.ThrowIfNull(bytes);
            ArgumentNullException.ThrowIfNull(launch);

            try
            {
                string root = OpenedFiles.TempRoot;

                OpenedFiles.Sweep(root, DateTimeOffset.UtcNow);

                string folder = Path.Combine(root, Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(folder);

                string fullPath = Path.Combine(folder, fileName);
                await File.WriteAllBytesAsync(fullPath, bytes);

                return await launch(fullPath) ? null : "This file could not be opened.";
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return $"This file could not be saved to open it ({ex.GetType().Name}).";
            }
        }

        // ###########################################################################################
        // Deletes the folders earlier opens left in the temp folder once they are old enough
        // (ReviewTableFiles.IsStaleOpenedFolder) - code review, 2026-09-26.
        //
        // *** NEVER FAILS THE OPEN IT PRECEDES. *** A folder still held by an open viewer, a
        // permission refusal, or a root that does not exist yet must not stop the maintainer
        // opening the file they asked for, so every failure is passed over and the next sweep
        // tries again. One folder that cannot be deleted does not stop the others.
        // ###########################################################################################
        private static void Sweep(string root, DateTimeOffset nowUtc)
        {
            try
            {
                if (!Directory.Exists(root))
                    return;

                foreach (string folder in Directory.EnumerateDirectories(root))
                {
                    try
                    {
                        if (ReviewTableFiles.IsStaleOpenedFolder(Directory.GetLastWriteTimeUtc(folder), nowUtc))
                            Directory.Delete(folder, recursive: true);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        // Still open in a viewer, or not ours to delete. The next sweep will try again.
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // The temp folder itself could not be listed; opening the file matters more.
            }
        }
    }
}
