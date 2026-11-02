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
    // own under Root and handed to the operating system.
    //
    // *** ONLY KNOWN TYPES, AND A WEB PAGE AS TEXT. *** See TryGetOpenName. The table opens only
    // what a submission may carry (ReviewTableFiles.TryGetOpenName); the tree also opens a board's
    // own files - the workbook, the highlight file and KiCad data - which the server writes or
    // takes from the data itself, never from an upload.
    //
    // *** EVERY TYPE HERE MUST BE ONE ExternalTargetLauncher OPENS (code review, 2026-11-02). *** That
    // is what opens them (MaintainerFileLauncher), with its own list: a type this lets through and it
    // refuses would be fetched and then "could not be opened".
    // Every_type_the_tab_opens_is_one_the_launcher_opens holds the two together.
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
        //
        // A KICAD PROJECT FILE AS .json (code review, 2026-11-02): that is what it is, and what looking
        // at it needs. It lands alone in a folder of its own, so KiCad would open an empty project, and
        // the launcher has no .kicad_pro on its list - which there is no reason to widen for this.
        // ###########################################################################################
        public static bool TryGetOpenName(string? path, out string fileName)
        {
            if (ReviewTableFiles.TryGetOpenName(path, out fileName))
                return true;

            string name = Path.GetFileName((path ?? string.Empty).Replace('\\', '/').Split('/')[^1]).Trim();
            string extension = Path.GetExtension(name);

            if (name.Length == 0 || !OpenedFiles.BoardFileExtensions.Contains(extension))
            {
                fileName = string.Empty;
                return false;
            }

            foreach (char invalid in Path.GetInvalidFileNameChars())
                name = name.Replace(invalid, '_');

            fileName = string.Equals(extension, ".kicad_pro", StringComparison.OrdinalIgnoreCase) ? name + ".json" : name;
            return true;
        }

        // ###########################################################################################
        // Where opened files are written: a folder in the user's OWN app data folder (AppConfig
        // .MaintainerOpenedFilesFolder), never the shared temp one. It was "<temp>/CRT/Maintainer" until
        // 2026-11-02 (code review, Security).
        // ###########################################################################################
        internal static string Root => AppConfig.MaintainerOpenedFilesFolder;

        private static string EarlierRoot => Path.Combine(Path.GetTempPath(), AppConfig.AppFolderName, "Maintainer");

        // ###########################################################################################
        // Saves `bytes` as `fileName` in a new folder under Root and hands it to `launch`. Null when it
        // opened, otherwise the sentence to show. Earlier opens' folders are swept first - see Sweep -
        // and those of a CRT from before the folder moved - see SweepEarlierRoot.
        // ###########################################################################################
        public static Task<string?> SaveAndLaunchAsync(byte[] bytes, string fileName, Func<string, Task<bool>> launch)
        {
            OpenedFiles.SweepEarlierRoot(Path.GetTempPath(), OpenedFiles.EarlierRoot, DateTimeOffset.UtcNow);

            return OpenedFiles.SaveAndLaunchInAsync(OpenedFiles.Root, bytes, fileName, launch);
        }

        // The same under another folder - a test's, never the user's real one.
        internal static async Task<string?> SaveAndLaunchInAsync(string root, byte[] bytes, string fileName, Func<string, Task<bool>> launch)
        {
            ArgumentNullException.ThrowIfNull(bytes);
            ArgumentNullException.ThrowIfNull(launch);

            try
            {
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
        // Deletes the folders earlier opens left under Root once they are old enough
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
                // The folder itself could not be listed; opening the file matters more.
            }
        }

        // ###########################################################################################
        // *** WHAT EARLIER CRTs LEFT IN THE SHARED TEMP FOLDER (code review, 2026-11-02 - Security). *** Moving
        // Root left contributors' files under "<temp>/CRT/Maintainer" for good, readable on a shared Linux
        // machine, and nothing swept there any more. Asked at each open: once it has gone, that is one
        // Directory.Exists.
        //
        // *** THAT FOLDER MAY NOT BE OURS (the very reason Root moved), and this deletes. *** Another account
        // can make "/tmp/CRT/Maintainer" a link to the user's home, so nothing is done through a folder that
        // links elsewhere (FolderLinks), and only a folder an open would have named - a GUID, no link, old
        // enough for no viewer to hold it - goes. Fails as Sweep fails: never the open.
        // ###########################################################################################
        internal static void SweepEarlierRoot(string temp, string earlierRoot, DateTimeOffset nowUtc)
        {
            try
            {
                if (!Directory.Exists(earlierRoot) || FolderLinks.AnyBetween(temp, Path.Combine(earlierRoot, "x")))
                    return;

                foreach (string folder in Directory.EnumerateDirectories(earlierRoot))
                {
                    try
                    {
                        if (Guid.TryParseExact(Path.GetFileName(folder), "N", out _) &&
                            !FolderLinks.AnyBetween(earlierRoot, Path.Combine(folder, "x")) &&
                            ReviewTableFiles.IsStaleOpenedFolder(Directory.GetLastWriteTimeUtc(folder), nowUtc))
                        {
                            Directory.Delete(folder, recursive: true);
                        }
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        // Still open in a viewer, or not ours to delete. The next open will try again.
                    }
                }

                // Gone once empty, and "<temp>/CRT" with it - neither is deleted while it holds anything.
                Directory.Delete(earlierRoot);
                Directory.Delete(Path.GetDirectoryName(earlierRoot)!);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Something is still there; opening the file matters more.
            }
        }
    }
}
