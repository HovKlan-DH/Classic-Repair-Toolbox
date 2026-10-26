using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // Which list a changed file is in - what it is called where it is offered.
    public enum ChangedFileKind
    {
        ComponentFile,
        ComponentPicture,
        BoardFile,
    }

    // ###########################################################################################
    // One file a draft replaced: a row that cited `PublishedName` (bytes `PublishedSha256`) on the
    // published board cites `DraftName` (bytes `DraftSha256`) in the draft. `RowName` is the row's
    // name ("Datasheet"), the same in both.
    // ###########################################################################################
    public sealed record ChangedFile(ChangedFileKind Kind, string RowName, string PublishedName, string PublishedSha256, string DraftName, string DraftSha256)
    {
        // "Datasheet (6510.pdf)" - the row, and the file it replaced.
        public string DisplayText => $"{this.RowName} ({Path.GetFileName(this.PublishedName)})";
    }

    // The bytes `File` replaced, on another board - under `NameThere`, that board's own name for them.
    public sealed record ChangedFileElsewhere(ChangedFile File, string BoardName, string BoardIdentity, string NameThere)
    {
        // The board - and that board's name for the file when it calls it something else.
        public string DisplayText =>
            string.Equals(BoardFileNames.Normalise(this.NameThere), BoardFileNames.Normalise(this.File.PublishedName), StringComparison.OrdinalIgnoreCase)
                ? this.BoardName
                : $"{this.BoardName} - as {Path.GetFileName(this.NameThere)}";
    }

    // What a board's draft makes of the new file put in place of the old one.
    public enum ChangedFilePut
    {
        // The draft has the published bytes there - the new file goes in.
        Put,

        // The draft has the new bytes already - under that name, or under another its rows cite.
        AlreadySo,

        // No row of the draft cites the name any more, or the draft lists nothing under it.
        NoLongerThere,

        // The draft has bytes of its own there - the contributor changed that board's file.
        OwnVersion,
    }

    // ###########################################################################################
    // "ALSO ON OTHER BOARDS" FOR FILES (CRT 4.0 step 2, 2026-10-26 - the plan: "This datasheet is also
    // used by 11 other boards - change it there too? Each board gets the change in its own draft ... It
    // works for files, text, pins and replacements"; Dennis, 2026-10-24: every board owns its files,
    // with CRT offering to apply a changed file to the other boards that use it).
    //
    // In the CRT 4.0 tree a board lists each of its files by name and SHA-256, and bytes several boards
    // list are stored once. So "the same file on another board" is a board listing the SAME BYTES - in
    // the live data 2,333 files are, nearly all under the same name. When a draft replaces one of a
    // component's files or pictures, or one of the board's own files, the other boards holding the
    // bytes it replaced can be given the new file, each in its own draft.
    //
    //   - THE NEW FILE GOES IN AS A CHANGE OF THAT BOARD'S ROWS (code review, 2026-10-26): under a
    //     name of that board's own (NameFor - the folder of its name for the old file, the new file's
    //     own name, never one the board has), every row there citing the old file pointed at it
    //     (Repointed) - ONE step of that draft's history, which its Undo takes back. Written over the
    //     old file's bytes instead, it was no step at all - a board file had no way back but discarding
    //     the draft - and a file of another type sat under the old one's name ("6510.txt" as
    //     "6510.pdf").
    //
    //   - A REPLACED FILE is a row of the published board and a row of the draft with the same NAME
    //     (trimmed, capitals ignored, in order), citing different bytes. A row renamed and given a new
    //     file at once is not paired - nothing says it is the same entry. Bytes that went two different
    //     ways (two rows, one old file, two new ones) are left out: which one is meant is not known.
    //   - ONLY WHERE ANOTHER BOARD SHOWS IT THE SAME WAY - a component's file or picture, or a board
    //     file. A scope baseline is measured on its board and a schematic image is the board itself
    //     (ComponentCopy's rule: what the board decides itself never comes from another), so a name
    //     cited only by those is never offered.
    //   - THE OTHER BOARD'S DRAFT DECIDES (CanPut): the new file goes in only where that draft still
    //     has the published bytes - never over a version of its own.
    //   - ONLY A NAME THAT IS A BOARD'S FILE NAME (SubmissionPathRules.IsSafelyShaped) - never one of a
    //     tampered tree climbing out of the board's folder.
    //
    // PURE: boards and file lists in, the changes out.
    // ###########################################################################################
    public static class ChangedFiles
    {
        // ###########################################################################################
        // The component `label`'s files and pictures (in `region`) the draft replaced - against the
        // published board and its file list. Nothing without the published board or either list.
        // ###########################################################################################
        public static List<ChangedFile> OfComponent(
            BoardData draft,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draftFiles,
            BoardData? published,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? publishedFiles,
            string? label,
            string? region)
        {
            ArgumentNullException.ThrowIfNull(draft);

            if (published is null || draftFiles is null || publishedFiles is null)
                return [];

            IEnumerable<ChangedFile> files = ChangedFiles.Pair(
                ChangedFileKind.ComponentFile,
                ComponentPaneEdits.FilesOf(published, label).Select(row => (row.Name, row.File)),
                ComponentPaneEdits.FilesOf(draft, label).Select(row => (row.Name, row.File)),
                publishedFiles,
                draftFiles);

            IEnumerable<ChangedFile> pictures = ChangedFiles.Pair(
                ChangedFileKind.ComponentPicture,
                ComponentPaneEdits.PicturesOf(published, label, region).Select(row => (row.Name, row.File)),
                ComponentPaneEdits.PicturesOf(draft, label, region).Select(row => (row.Name, row.File)),
                publishedFiles,
                draftFiles);

            return ChangedFiles.Unambiguous(files.Concat(pictures));
        }

        // The board's own files (Files and links) the draft replaced.
        public static List<ChangedFile> OfBoardFiles(
            BoardData draft,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draftFiles,
            BoardData? published,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? publishedFiles)
        {
            ArgumentNullException.ThrowIfNull(draft);

            if (published is null || draftFiles is null || publishedFiles is null)
                return [];

            return ChangedFiles.Unambiguous(ChangedFiles.Pair(
                ChangedFileKind.BoardFile,
                published.BoardLocalFiles.Select(row => (row.Name, row.File)),
                draft.BoardLocalFiles.Select(row => (row.Name, row.File)),
                publishedFiles,
                draftFiles));
        }

        // ###########################################################################################
        // Where else the bytes each of `files` replaced are - every other board (named, with its board
        // identity, its published data and file list) listing them under a name one of its component
        // files, component pictures or board files cites. In the order of `files`, then of `others`.
        // ###########################################################################################
        public static List<ChangedFileElsewhere> Elsewhere(
            IReadOnlyList<ChangedFile> files,
            IEnumerable<(string BoardName, string BoardIdentity, BoardData Board, IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? Files)> others)
        {
            ArgumentNullException.ThrowIfNull(files);
            ArgumentNullException.ThrowIfNull(others);

            var boards = others
                .Where(other => other.Files is not null)
                .Select(other => (other.BoardName, other.BoardIdentity, Cited: ChangedFiles.SharedNames(other.Board), Files: other.Files!))
                .ToList();

            var found = new List<ChangedFileElsewhere>();

            foreach (ChangedFile file in files)
            {
                foreach ((string boardName, string identity, HashSet<string> cited, IReadOnlyDictionary<string, BoardDocument.StoredFileRow> listed) in boards)
                {
                    // Never a name that is no board's file name - one climbing out of its folder (code
                    // review, 2026-10-26): the new file is copied to wherever the name points.
                    foreach (BoardDocument.StoredFileRow entry in listed.Values
                        .Where(entry => string.Equals(entry.Sha256, file.PublishedSha256, StringComparison.OrdinalIgnoreCase) &&
                                        cited.Contains(BoardFileNames.Normalise(entry.Name)) &&
                                        SubmissionPathRules.IsSafelyShaped(BoardFileNames.Normalise(entry.Name), out _))
                        .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        found.Add(new ChangedFileElsewhere(file, boardName, identity, entry.Name));
                    }
                }
            }

            return found;
        }

        // ###########################################################################################
        // What the draft `board` (its file list `files`) of the board `target` names makes of the new
        // file going in under `target.NameThere`.
        // ###########################################################################################
        public static ChangedFilePut CanPut(BoardData board, IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? files, ChangedFileElsewhere target)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(target);

            string name = BoardFileNames.Normalise(target.NameThere);
            HashSet<string> shared = ChangedFiles.SharedNames(board);

            if (!shared.Contains(name) || ChangedFiles.Listed(files, target.NameThere) is not BoardDocument.StoredFileRow listed)
            {
                // Put in before, its rows cite the new file under a name of its own now.
                bool hasNewBytes = shared.Any(cited => ChangedFiles.Listed(files, cited) is { } row &&
                                                       string.Equals(row.Sha256, target.File.DraftSha256, StringComparison.OrdinalIgnoreCase));

                return hasNewBytes ? ChangedFilePut.AlreadySo : ChangedFilePut.NoLongerThere;
            }

            if (string.Equals(listed.Sha256, target.File.DraftSha256, StringComparison.OrdinalIgnoreCase))
                return ChangedFilePut.AlreadySo;

            return string.Equals(listed.Sha256, target.File.PublishedSha256, StringComparison.OrdinalIgnoreCase)
                ? ChangedFilePut.Put
                : ChangedFilePut.OwnVersion;
        }

        // ###########################################################################################
        // The name the new file gets on the board `target` names: in that board's folder for the old
        // file, under the new file's own name - its own type with it - and never one of `taken`, the
        // names that board has (BoardFileNames.FreeName: "6510 rev B (2).pdf").
        // ###########################################################################################
        public static string NameFor(ChangedFileElsewhere target, IEnumerable<string> taken)
        {
            ArgumentNullException.ThrowIfNull(target);

            string thereFolder = Path.GetDirectoryName(BoardFileNames.Normalise(target.NameThere))?.Replace('\\', '/') ?? string.Empty;
            string fileName = Path.GetFileName(BoardFileNames.Normalise(target.File.DraftName));
            string preferred = thereFolder.Length == 0 ? fileName : $"{thereFolder}/{fileName}";

            return BoardFileNames.FreeName(preferred, taken ?? []);
        }

        // ###########################################################################################
        // `board` with every component file, component picture and board file citing `oldName` (as
        // BoardFileNames spells it, capitals ignored) citing `newName` instead - every other field as it
        // was, every other row the same object. The same board when no row cites it. Scope baselines
        // and schematic images keep the old file: they are the board's own (ComponentCopy's rule).
        // ###########################################################################################
        public static BoardData Repointed(BoardData board, string oldName, string newName)
        {
            ArgumentNullException.ThrowIfNull(board);

            string from = BoardFileNames.Normalise(oldName);

            bool Cites(string? file) => from.Length > 0 && string.Equals(BoardFileNames.Normalise(file), from, StringComparison.OrdinalIgnoreCase);

            if (!board.ComponentLocalFiles.Any(row => Cites(row.File)) &&
                !board.ComponentImages.Any(row => Cites(row.File) && !ComponentImageKinds.IsScopeBaseline(row)) &&
                !board.BoardLocalFiles.Any(row => Cites(row.File)))
            {
                return board;
            }

            List<ComponentLocalFileEntry> componentFiles = board.ComponentLocalFiles
                .Select(row => Cites(row.File) ? new ComponentLocalFileEntry { BoardLabel = row.BoardLabel, Name = row.Name, File = newName } : row)
                .ToList();

            List<ComponentImageEntry> pictures = board.ComponentImages
                .Select(row => Cites(row.File) && !ComponentImageKinds.IsScopeBaseline(row)
                    ? new ComponentImageEntry
                    {
                        BoardLabel = row.BoardLabel,
                        Region = row.Region,
                        Pin = row.Pin,
                        Name = row.Name,
                        ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                        File = newName,
                        Note = row.Note,
                        TimeDiv = row.TimeDiv,
                        VoltsDiv = row.VoltsDiv,
                        TriggerLevelVolts = row.TriggerLevelVolts,
                        MachineState = row.MachineState,
                        Source = row.Source,
                    }
                    : row)
                .ToList();

            List<BoardLocalFileEntry> boardFiles = board.BoardLocalFiles
                .Select(row => Cites(row.File) ? new BoardLocalFileEntry { Category = row.Category, Name = row.Name, File = newName } : row)
                .ToList();

            return board
                .WithComponentSections(componentImages: pictures, componentLocalFiles: componentFiles)
                .WithBoardSections(boardLocalFiles: boardFiles);
        }

        // ---- the rules ----------------------------------------------------------------------------

        // Each row of `before` paired with the first row of `after` of the same name, where their
        // bytes differ.
        private static IEnumerable<ChangedFile> Pair(
            ChangedFileKind kind,
            IEnumerable<(string Name, string File)> before,
            IEnumerable<(string Name, string File)> after,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow> publishedFiles,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow> draftFiles)
        {
            var unpaired = after.ToList();

            foreach ((string name, string file) in before)
            {
                string rowName = name?.Trim() ?? string.Empty;

                if (rowName.Length == 0)
                    continue;

                int index = unpaired.FindIndex(row => string.Equals(row.Name?.Trim(), rowName, StringComparison.OrdinalIgnoreCase));

                if (index < 0)
                    continue;

                string draftFile = unpaired[index].File;
                unpaired.RemoveAt(index);

                if (ChangedFiles.Listed(publishedFiles, file) is not BoardDocument.StoredFileRow was ||
                    ChangedFiles.Listed(draftFiles, draftFile) is not BoardDocument.StoredFileRow now ||
                    !DataTreeLayout.IsSha256(was.Sha256) ||
                    !DataTreeLayout.IsSha256(now.Sha256) ||
                    string.Equals(was.Sha256, now.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                yield return new ChangedFile(kind, rowName, was.Name, was.Sha256, now.Name, now.Sha256);
            }
        }

        // One change per old and new bytes - and none for old bytes that went two different ways.
        private static List<ChangedFile> Unambiguous(IEnumerable<ChangedFile> changes)
        {
            List<ChangedFile> all = changes.ToList();

            var ambiguous = new HashSet<string>(
                all.GroupBy(change => change.PublishedSha256, StringComparer.OrdinalIgnoreCase)
                    .Where(group => group.Select(change => change.DraftSha256).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
                    .Select(group => group.Key),
                StringComparer.OrdinalIgnoreCase);

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            return all
                .Where(change => !ambiguous.Contains(change.PublishedSha256) && seen.Add(change.PublishedSha256))
                .ToList();
        }

        // A listed file by the name a row cites - as it is spelt, else as BoardFileNames spells it.
        private static BoardDocument.StoredFileRow? Listed(IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? files, string? name)
        {
            string trimmed = name?.Trim() ?? string.Empty;

            if (files is null || trimmed.Length == 0)
                return null;

            return files.TryGetValue(trimmed, out BoardDocument.StoredFileRow? listed) ||
                   files.TryGetValue(BoardFileNames.Normalise(trimmed), out listed)
                ? listed
                : null;
        }

        // The names a board's component files, component pictures and board files cite.
        private static HashSet<string> SharedNames(BoardData board) =>
            new(
                board.ComponentLocalFiles.Select(row => row.File)
                    .Concat(board.ComponentImages.Where(row => !ComponentImageKinds.IsScopeBaseline(row)).Select(row => row.File))
                    .Concat(board.BoardLocalFiles.Select(row => row.File))
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Select(name => BoardFileNames.Normalise(name)),
                StringComparer.OrdinalIgnoreCase);
    }
}
