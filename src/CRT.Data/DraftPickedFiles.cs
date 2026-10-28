using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // FILES AN EDITOR PICKED, COPIED INTO A DRAFT FOLDER UNDER THEIR NAMES - and taken back when the
    // change citing them is not saved (code review, 2026-10-26).
    //
    // The component editor copies a picked file to where the draft's next write takes it in
    // (DraftFiles: "<draft>/Component images/U8.png") BEFORE the change citing it is saved, so the
    // name a row stores already names bytes. When that save is not made the change is taken back out
    // (DraftEdits) - and the copies must go with it: left lying there, the next write by ANY editor
    // took them in, so a picture replaced under a name a row already cites entered the draft, and
    // the next submission, after the editor had said "could not be saved".
    //
    // A file a copy replaced - one picked earlier and not taken in yet - is set aside in CRT's own
    // bookkeeping folder (SetAsideFolder, never the contributor's work - DraftFolderLayout) and put
    // back by TakeBack; Keep lets go of it. Any thread; one editor's save at a time.
    // ###########################################################################################
    public sealed class DraftPickedFiles
    {
        public const string SetAsideFolder = DraftFolderLayout.BookkeepingPrefix + "set-aside";

        private readonly string thisDraftFolder;
        private readonly List<(string Copy, string? SetAside)> thisCopies = new();
        private readonly List<string> thisCopied = new();

        // ###########################################################################################
        // The RelativeFile of every file copied, in the order asked - the ones skipped left out, so a
        // caller adds rows only for what is there (code review, 2026-10-28: three pictures picked, one
        // locked by another program, and three rows went in, one naming a file the draft lacked).
        // ###########################################################################################
        public IReadOnlyList<string> Copied => this.thisCopied;

        private DraftPickedFiles(string draftFolder)
        {
            this.thisDraftFolder = draftFolder;
        }

        // ###########################################################################################
        // Copies each picked file (`Source`) to where the draft takes in `RelativeFile`
        // (DraftFileResolver.BuildDraftFileDestination). A picked file that is gone, or that lies
        // where it would be copied to, is skipped; so is one whose copy or set-aside fails - logged,
        // as a copy that failed always was, and with nothing replaced that could not be put back.
        // ###########################################################################################
        public static DraftPickedFiles CopyInto(string draftFolder, IReadOnlyList<(string Source, string RelativeFile)> files)
        {
            ArgumentNullException.ThrowIfNull(files);

            var copies = new DraftPickedFiles(draftFolder);
            var done = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach ((string source, string relativeFile) in files)
            {
                string destination = Path.GetFullPath(DraftFileResolver.BuildDraftFileDestination(draftFolder, relativeFile));

                if (!File.Exists(source) ||
                    string.Equals(Path.GetFullPath(source), destination, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);

                    // A second file to the same place in one save: the first already set aside what was there.
                    string? setAside = !done.Contains(destination) && File.Exists(destination) ? copies.SetAside(destination) : null;

                    File.Copy(source, destination, overwrite: true);

                    copies.thisCopies.Add((destination, setAside));
                    copies.thisCopied.Add(relativeFile);
                    done.Add(destination);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Warning($"Failed to copy attached file [{source}] into draft - [{ex.Message}]");
                }
            }

            return copies;
        }

        // Whether a copy still lies where it was put - not taken in by a write yet.
        public bool AnyLying => this.thisCopies.Any(copy => File.Exists(copy.Copy));

        // The save was made: what the copies replaced is let go of.
        public void Keep()
        {
            foreach ((_, string? setAside) in this.thisCopies)
            {
                if (setAside is not null)
                    DraftPickedFiles.TryDelete(setAside);
            }

            this.thisCopies.Clear();
            this.RemoveSetAsideFolderIfEmpty();
        }

        // ###########################################################################################
        // The save was made, but a write takes in only the names its rows cite (DraftFiles.ListFor):
        // the copies it took in - gone from where they arrived - are kept, what they replaced let go
        // of; each one still lying there goes as the copy of a save not made does, what it replaced
        // coming back (code review, 2026-10-26 - left lying, the next write by any editor took it in
        // unasked). True when every copy was taken in.
        // ###########################################################################################
        public bool KeepTakenIn()
        {
            List<(string Copy, string? SetAside)> lying = this.thisCopies.Where(copy => File.Exists(copy.Copy)).ToList();

            foreach ((string copy, string? setAside) in this.thisCopies.Where(copy => !File.Exists(copy.Copy)))
            {
                if (setAside is not null)
                    DraftPickedFiles.TryDelete(setAside);
            }

            this.thisCopies.Clear();
            this.thisCopies.AddRange(lying);
            this.TakeBack();

            return lying.Count == 0;
        }

        // The save was not made: every copy goes, and what it replaced comes back - last copy first,
        // so a place copied to twice gets back what was there before either.
        public void TakeBack()
        {
            for (int index = this.thisCopies.Count - 1; index >= 0; index--)
            {
                (string copy, string? setAside) = this.thisCopies[index];

                try
                {
                    if (setAside is not null)
                    {
                        File.Move(setAside, copy, overwrite: true);
                    }
                    else if (File.Exists(copy))
                    {
                        File.Delete(copy);
                        DraftPickedFiles.RemoveEmptyFolders(this.thisDraftFolder, Path.GetDirectoryName(copy)!);
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Warning($"Could not take back the copy [{copy}] of a change that was not saved - [{ex.Message}]");
                }
            }

            this.thisCopies.Clear();
            this.RemoveSetAsideFolderIfEmpty();
        }

        // The file at `destination` moved into the set-aside folder; where it went.
        private string SetAside(string destination)
        {
            string folder = Path.Combine(this.thisDraftFolder, DraftPickedFiles.SetAsideFolder);
            Directory.CreateDirectory(folder);

            string path = Path.Combine(folder, Path.GetRandomFileName());
            File.Move(destination, path);

            return path;
        }

        private void RemoveSetAsideFolderIfEmpty()
        {
            string folder = Path.Combine(this.thisDraftFolder, DraftPickedFiles.SetAsideFolder);

            try
            {
                if (Directory.Exists(folder) && !Directory.EnumerateFileSystemEntries(folder).Any())
                    Directory.Delete(folder);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not remove the empty folder [{folder}] - [{ex.Message}]");
            }
        }

        // The folders a removed copy leaves empty, up to (not including) the draft folder.
        private static void RemoveEmptyFolders(string draftFolder, string from)
        {
            string root = Path.GetFullPath(draftFolder);
            string? folder = Path.GetFullPath(from);

            while (folder is not null &&
                   folder.Length > root.Length &&
                   folder.StartsWith(root, StringComparison.OrdinalIgnoreCase) &&
                   Directory.Exists(folder) &&
                   !Directory.EnumerateFileSystemEntries(folder).Any())
            {
                Directory.Delete(folder);
                folder = Path.GetDirectoryName(folder);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not remove the set-aside file [{path}] - [{ex.Message}]");
            }
        }
    }
}
