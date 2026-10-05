namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // MAY THE SERVICE WRITE WHERE IT IS ABOUT TO? - asked before the first byte of a publish, a
    // production promotion or a push-back lands (owner report, 2026-09-28).
    //
    // *** WHY: DATA COPIED INTO THE TREES BY HAND. *** The trees are group-writable with the
    // group inherited (INSTALLING.md ("Folders and permissions")), but whatever is copied in as root arrives owned by
    // root with no group write. For a FILE that no longer matters - every served file is
    // replaced by a rename (FileReplacer, VerifiedFileCopy), which needs only its folder. A
    // FOLDER copied in that way is another matter: nothing can be written into it or removed
    // from it, and finding that out at the third file of a board leaves the board half-replaced.
    // It happened: production's files copied into BETA by hand, then an approval that wrote the
    // new workbook and was refused at the highlight file beside it.
    //
    // The service cannot repair it - only the owner of a folder, or root, can change it - so the
    // answer is to REFUSE BEFORE WRITING, name the folders, and put the command that fixes them
    // in the log for the administrator.
    //
    // *** ASKED BY TRYING, not by reading permission bits *** - group membership, ACLs, a
    // read-only mount and systemd's ProtectSystem each make a writable-looking folder refuse. A
    // zero-byte dot-named probe, deleted on close; the startup check uses the same one.
    //
    // *** AFTER THE LINK CHECK, ALWAYS. *** A probe is a real file: probing a folder reached
    // through a symbolic link would create it outside the tree. Every caller checks links first.
    // ###########################################################################################
    public static class TreeWriteAccess
    {
        // The group INSTALLING.md ("Folders and permissions") gives the data trees.
        public const string DataGroup = "crt-data";

        // ###########################################################################################
        // The folders this process may NOT write, among those `paths` would be written into or
        // removed from - each path's own folder, or for a folder that does not exist yet the
        // nearest one above it that does, since that is where it would be created. Never above
        // `treeRoot`: a missing root is reported as the root, and nothing outside it is probed.
        // Each folder once, as full paths, in the order first met. `canWrite` is for tests.
        // ###########################################################################################
        public static IReadOnlyList<string> FoldersRefusing(
            string treeRoot,
            IEnumerable<string> paths,
            Func<string, bool>? canWrite = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(treeRoot);
            ArgumentNullException.ThrowIfNull(paths);

            Func<string, bool> probe = canWrite ?? TreeWriteAccess.CanWrite;
            string root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(treeRoot));

            var asked = new HashSet<string>(StringComparer.Ordinal);
            var refusing = new List<string>();

            foreach (string path in paths)
            {
                string? folder = TreeWriteAccess.ExistingFolderFor(root, path);

                if (folder is null || !asked.Add(folder))
                    continue;

                // A tree root that is not there cannot be written, and is not probed.
                if (!Directory.Exists(folder) || !probe(folder))
                    refusing.Add(folder);
            }

            return refusing;
        }

        // The folder a write of `path` needs: its own, or the nearest existing one above it, never
        // above the root. Null for a path outside the root - callers resolved every path inside
        // it already, so this is only a guard.
        private static string? ExistingFolderFor(string root, string path)
        {
            string? folder = Path.GetDirectoryName(Path.GetFullPath(path));

            if (folder is null || !TreeWriteAccess.IsAtOrInside(root, folder))
                return null;

            while (!Directory.Exists(folder) && !string.Equals(folder, root, StringComparison.Ordinal))
            {
                folder = Path.GetDirectoryName(folder);

                if (folder is null || !TreeWriteAccess.IsAtOrInside(root, folder))
                    return root;
            }

            return folder;
        }

        private static bool IsAtOrInside(string root, string folder)
        {
            string trimmed = Path.TrimEndingDirectorySeparator(folder);

            return string.Equals(trimmed, root, StringComparison.Ordinal) ||
                trimmed.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The sentence a maintainer reads. Names the folders the way the data does
        // ("Commodore/C128/310378"), says nothing was changed, and sends them to whoever runs the
        // server - a maintainer cannot fix a folder, and the command belongs in the log.
        // ###########################################################################################
        public static string RefusalMessage(string treeName, string treeRoot, IReadOnlyList<string> folders)
        {
            ArgumentNullException.ThrowIfNull(folders);

            const int Named = 5;
            string root = Path.GetFullPath(treeRoot);

            IEnumerable<string> names = folders
                .Take(Named)
                .Select(folder => TreeWriteAccess.DisplayName(root, folder));

            string list = string.Join(", ", names.Select(name => $"[{name}]"));

            if (folders.Count > Named)
                list += $" and {folders.Count - Named} more";

            return $"The server is not allowed to write into {list} in the {treeName} data, so nothing was changed. " +
                "This happens when files are copied into the data by hand as another user. Whoever runs the server " +
                "has to give it access again - the server's log gives the command - and then this can be tried again.";
        }

        // ###########################################################################################
        // The command that gives the service its folders back - INSTALLING.md ("Folders and permissions"), for exactly
        // these folders. Logged for the administrator, never shown to a maintainer. Each path is
        // single-quoted for the shell (a board folder or "Shared files" has spaces in it).
        // ###########################################################################################
        public static string FixCommand(IEnumerable<string> folders)
        {
            ArgumentNullException.ThrowIfNull(folders);

            return string.Join(" ; ", folders.Select(folder =>
            {
                string quoted = TreeWriteAccess.ShellQuote(folder);

                return $"sudo chgrp -R {TreeWriteAccess.DataGroup} {quoted} && sudo chmod -R g+rwX {quoted} && " +
                    $"sudo find {quoted} -type d -exec chmod g+s {{}} +";
            }));
        }

        internal static string ShellQuote(string text) => "'" + text.Replace("'", "'\\''") + "'";

        private static string DisplayName(string root, string folder)
        {
            string relative = Path.GetRelativePath(root, folder).Replace(Path.DirectorySeparatorChar, '/');

            return relative == "." ? "the top folder" : relative;
        }

        // ###########################################################################################
        // Can this process create a file in `directory`? Answered by trying - see the header - and
        // cleaned up in a finally, so a probe never leaves litter in a data tree.
        // ###########################################################################################
        public static bool CanWrite(string directory)
        {
            string probePath = Path.Combine(directory, $".crt-server-write-probe-{Guid.NewGuid():N}");

            try
            {
                using (FileStream stream = File.Create(probePath, 1, FileOptions.DeleteOnClose))
                {
                    stream.WriteByte(0);
                }

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                return false;
            }
            finally
            {
                // DeleteOnClose normally handles this; the explicit delete covers the case where the
                // handle was closed by an exception path before the flag could take effect.
                try
                {
                    if (File.Exists(probePath))
                        File.Delete(probePath);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Nothing useful to do - a stray zero-byte probe file is harmless, and throwing
                    // from a cleanup path would mask the real result.
                }
            }
        }
    }
}
