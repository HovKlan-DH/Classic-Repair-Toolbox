using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHETHER A PATH IS REACHED THROUGH A FOLDER THAT LINKS ELSEWHERE (code review, 2026-10-25).
    //
    // A containment check on the TEXT of a path ("is it under the draft folder?") is fooled by a
    // folder on the way that is a symbolic link or a junction: the path reads as inside, and a
    // write or a DELETE lands in the folder linked to. The two places CRT deletes files it decided
    // are its own - a draft taking in a new file (DraftFiles), the KiCad copy tidying itself
    // (BoardKiCadFolder) - ask this first.
    // ###########################################################################################
    public static class FolderLinks
    {
        // ###########################################################################################
        // True when a folder strictly below `root`, on the way to `path`, is a link - or when that
        // cannot be told (a folder that cannot be read counts as one: the answer decides whether a
        // file may be deleted). The root itself is the caller's choice and is not looked at; folders
        // not made yet are not links.
        // ###########################################################################################
        public static bool AnyBetween(string root, string path)
        {
            string top = Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));

            try
            {
                for (string? folder = Path.GetDirectoryName(Path.GetFullPath(path));
                     folder is not null && folder.Length > top.Length;
                     folder = Path.GetDirectoryName(folder))
                {
                    if (Directory.Exists(folder) && new DirectoryInfo(folder).Attributes.HasFlag(FileAttributes.ReparsePoint))
                        return true;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return true;
            }

            return false;
        }
    }
}
