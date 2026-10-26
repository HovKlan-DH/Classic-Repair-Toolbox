using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // What converting one draft did: whether it is now a CRT 4.0 draft (and whether it already
    // was), what the contributor should know about it, and why it was left alone when it was.
    public sealed record DraftConversionResult(bool Converted, bool AlreadyConverted, IReadOnlyList<string> Notes, string? Failure)
    {
        public static DraftConversionResult Failed(string failure) => new(false, false, [], failure);
    }

    // ###########################################################################################
    // CONVERTS A USER'S EXCEL-ERA DRAFT INTO A CRT 4.0 DRAFT, IN PLACE (owner decision, 2026-10-24:
    // "I do expect you also will migrate all existing data").
    //
    // A CRT 4.0 draft is its board's folder under the drafts root, laid out like the 4.0 tree:
    //
    //     <drafts>/Commodore/C64/250407/
    //         board.json          the whole board as the draft has it (BoardDocument) - every file
    //                             named, the published board's KiCad files included
    //         Files/<ab>/<sha>    the bytes of the files the published tree does not hold: what the
    //                             contributor added or replaced
    //         .crt-draft.json     the marker, as before, naming the board by its board file
    //                             ("<id>/board.json")
    //         .crt-base.json      the published board as the draft started from it (a BoardDocument)
    //         .crt-pre-4.0/       the Excel-era draft, every file of it, set aside - nothing is lost
    //
    // *** NAMES AGREE WITH THE PUBLISHED BOARD'S. *** A file the draft has as the published board
    // has it (the same bytes) gets the published name; a file the draft REPLACED keeps the name the
    // published board knows it by (found through the draft's base: what was at that path when the
    // draft started); only a file the draft added is named afresh, by BoardFileNames, never taking
    // a published name. So a converted draft shows the contributor's own changes and nothing else.
    //
    // *** A DRAFT IS CONVERTED EVEN WHEN SOMETHING IS ODD ABOUT IT. *** Unlike the tree's conversion,
    // which a person can fix before it runs, a user cannot fix a draft before 4.0 starts - and a
    // draft left unconverted would be out of reach. So it is converted with what CRT reads of it
    // (as 3.x showed it), the Excel-era files are kept beside it, and Notes say what to look at: a
    // file it cites that exists nowhere, say.
    //
    // *** SAFE TO STOP AT ANY POINT. *** The new draft is built in a working folder first; only then
    // are the old files set aside and the new ones moved into place. A conversion interrupted
    // half-way starts again from the set-aside originals.
    // ###########################################################################################
    public static class DraftConversion
    {
        public const string OriginalsFolder = ".crt-pre-4.0";

        private const string WorkingFolder = ".crt-converting";

        // Written into the set-aside folder once every old file is in it.
        private const string SetAsideMarker = ".crt-set-aside";

        // ###########################################################################################
        // Converts the Excel-era draft in `draftFolder` - the board's folder under the drafts root,
        // board `boardId` ("Commodore/C64/250407"). `oldDataRoot` is the Excel data folder the
        // draft resolved its published files against; `treeContentRoot` is the 4.0 tree's content
        // folder (its "Data"), holding the published board files and the stored bytes; `boardIds`
        // is every board the tree lists, for naming files a row cites from another board's folder.
        // ###########################################################################################
        public static DraftConversionResult Convert(
            string draftFolder,
            string boardId,
            string oldDataRoot,
            string treeContentRoot,
            IReadOnlyCollection<string> boardIds)
        {
            if (string.IsNullOrWhiteSpace(draftFolder) || !Directory.Exists(draftFolder))
                return DraftConversionResult.Failed($"The draft folder [{draftFolder}] is not there.");

            string working = Path.Combine(draftFolder, DraftConversion.WorkingFolder);
            string originals = Path.Combine(draftFolder, DraftConversion.OriginalsFolder);
            bool setAside = File.Exists(Path.Combine(originals, DraftConversion.SetAsideMarker));

            if (File.Exists(Path.Combine(draftFolder, DataTreeLayout.BoardFileName)) &&
                !Directory.Exists(working) &&
                (!Directory.Exists(originals) || setAside))
            {
                return new DraftConversionResult(false, true, [], null);
            }

            try
            {
                // A conversion stopped while setting the old draft aside finishes that first: what
                // is still beside the set-aside folder is the rest of the old draft.
                if (Directory.Exists(originals) && !setAside)
                    DraftConversion.SetAside(draftFolder, originals);

                // Then it starts again from what it set aside.
                string source = Directory.Exists(originals) ? originals : draftFolder;

                if (Directory.Exists(working))
                    Directory.Delete(working, recursive: true);

                Built? built = DraftConversion.Build(source, boardId, oldDataRoot, treeContentRoot, boardIds, working, out string? failure);

                if (built is null)
                {
                    if (Directory.Exists(working))
                        Directory.Delete(working, recursive: true);

                    return DraftConversionResult.Failed(failure!);
                }

                if (!Directory.Exists(originals))
                    DraftConversion.SetAside(draftFolder, originals);

                DraftConversion.MoveIn(working, draftFolder);

                return new DraftConversionResult(true, false, built.Notes, null);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException)
            {
                return DraftConversionResult.Failed($"The draft could not be converted: {ex.Message}");
            }
        }

        private sealed record Built(IReadOnlyList<string> Notes);

        // ###########################################################################################
        // Reads the Excel-era draft at `source` and writes the 4.0 draft into `working`: board.json,
        // the base, the marker and the stored bytes it needs. Null, with the reason, when `source`
        // is no Excel-era draft CRT can read.
        // ###########################################################################################
        private static Built? Build(
            string source,
            string boardId,
            string oldDataRoot,
            string treeContentRoot,
            IReadOnlyCollection<string> boardIds,
            string working,
            out string? failure)
        {
            failure = null;
            var notes = new List<string>();

            DraftMarker? marker = DraftMarkerStore.Load(Path.Combine(source, DraftFolderLayout.DraftMarkerFileName));

            if (marker is null)
            {
                failure = "It has no draft marker (.crt-draft.json) CRT can read.";
                return null;
            }

            List<string> workbooks = Directory.EnumerateFiles(source, "*" + DataGenerationRules.WorkbookExtension)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal) && !DraftFolderLayout.IsDraftOnlyFile(path))
                .ToList();

            if (workbooks.Count != 1)
            {
                failure = workbooks.Count == 0 ? "It holds no board workbook." : $"It holds {workbooks.Count} workbooks, and a draft has one.";
                return null;
            }

            string workbook = workbooks[0];
            BoardData? data = BoardDataReader.ReadWorkbookUncached(workbook);

            if (data is null)
            {
                failure = $"Its workbook [{Path.GetFileName(workbook)}] could not be read.";
                return null;
            }

            // The board's identity as the Excel-era draft knew it - its workbook path from the data
            // root - which every path its rows cite is relative to.
            string excelDataFile = $"{boardId}/{Path.GetFileName(workbook)}";

            var content = new BoardContent(
                data,
                BoardComponentHighlightStorage.LoadAllKiCadCalibrations(workbook),
                BoardComponentHighlightStorage.LoadSchematicRenames(workbook));

            BoardDocument? published = DraftConversion.ReadPublished(treeContentRoot, boardId, notes);
            DraftBase? draftBase = DraftBaseStore.LoadFile(Path.Combine(source, DraftFolderLayout.DraftBaseFileName));

            var namer = new Namer(boardId, boardIds, published, draftBase, excelDataFile);
            var files = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);
            var hashes = new Dictionary<string, (string Sha, long Size)>(StringComparer.OrdinalIgnoreCase);

            void Cite(string? cited)
            {
                if (string.IsNullOrWhiteSpace(cited))
                    return;

                string? bytes = DraftConversion.FindBytes(source, oldDataRoot, excelDataFile, cited);

                if (bytes is null)
                {
                    notes.Add($"A row cites [{cited}], which is neither in the draft nor in the data - the row keeps its name for it.");
                    return;
                }

                (string sha, long size) = DraftConversion.HashOf(bytes, hashes);
                string name = namer.NameOf(cited, sha);

                files[name] = new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = size };
                DraftConversion.StoreUnlessPublished(bytes, sha, name, treeContentRoot, working);
            }

            foreach (BoardSchematicEntry schematic in data.Schematics)
                Cite(schematic.SchematicImageFile);

            foreach (ComponentImageEntry image in data.ComponentImages)
                Cite(image.File);

            foreach (ComponentLocalFileEntry file in data.ComponentLocalFiles)
                Cite(file.File);

            foreach (BoardLocalFileEntry file in data.BoardLocalFiles)
                Cite(file.File);

            // The KiCad files: the published board's, as the draft showed them beside its own, and
            // the draft's own import, which wins for a file of the same name.
            foreach (BoardDocument.StoredFileRow kiCad in (published?.Files ?? []).Where(DraftFiles.IsKiCad))
                files.TryAdd(kiCad.Name, kiCad);

            string kiCadFolder = Path.Combine(source, DataTreeUsage.KiCadFolderName);

            if (Directory.Exists(kiCadFolder))
            {
                foreach (string path in Directory.EnumerateFiles(kiCadFolder, "*", SearchOption.AllDirectories).Order(StringComparer.Ordinal))
                {
                    string relative = Path.GetRelativePath(source, path).Replace('\\', '/');

                    if (!DataChecksumManifest.IsSyncable(relative))
                        continue;

                    (string sha, long size) = DraftConversion.HashOf(path, hashes);

                    files[relative] = new BoardDocument.StoredFileRow { Name = relative, Sha256 = sha, Size = size };
                    DraftConversion.StoreUnlessPublished(path, sha, relative, treeContentRoot, working);
                }
            }

            BoardDocument document = BoardDocumentMapping.ToDocument(content, namer.NameFor, DraftConversion.Sorted(files.Values));

            Directory.CreateDirectory(working);
            DraftConversion.WriteText(Path.Combine(working, DataTreeLayout.BoardFileName), DataTreeJson.Write(document));

            if (DraftConversion.BaseDocument(draftBase, excelDataFile, oldDataRoot, source, namer, hashes) is BoardDocument baseDocument)
                DraftConversion.WriteText(Path.Combine(working, DraftFolderLayout.DraftBaseFileName), DataTreeJson.Write(baseDocument));

            // The marker names the board by its board file now ("<id>/board.json", the identity the
            // main file gives it); the content hashes it kept were of the Excel files, which a 4.0
            // draft is never compared with.
            DraftMarkerStore.Save(Path.Combine(working, DraftFolderLayout.DraftMarkerFileName), marker with
            {
                BoardKey = DataTreeLayout.BoardFilePath(boardId),
                BaseContentHash = string.Empty,
                DriftSeenContentHash = string.Empty,
            });

            return new Built(notes);
        }

        // ###########################################################################################
        // The base - the published board as the Excel-era draft started from it - as a board file,
        // its files named as the draft's are. Null without a base (a draft started before bases
        // existed, or a new board): the 4.0 draft then has none either, as before.
        // ###########################################################################################
        private static BoardDocument? BaseDocument(
            DraftBase? draftBase,
            string excelDataFile,
            string oldDataRoot,
            string source,
            Namer namer,
            Dictionary<string, (string Sha, long Size)> hashes)
        {
            if (draftBase is null)
                return null;

            var files = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);
            var shaByBoardPath = draftBase.Files.ToDictionary(file => file.Path, file => file, StringComparer.OrdinalIgnoreCase);

            foreach (string cited in SubmissionManifestBuilder.CollectReferencedFiles(draftBase.Board))
            {
                string? sha = null;
                long size = 0;

                // A file of the board's own folder has its hash in the base; a shared one is read from
                // the data the draft resolved it against.
                if (DraftFolderLayout.RelativeToBoardFolder(excelDataFile, cited) is string withinBoard &&
                    shaByBoardPath.TryGetValue(withinBoard, out DraftBaseFile? recorded))
                {
                    (sha, size) = (recorded.Sha256.ToLowerInvariant(), recorded.Length);
                }
                else if (DraftConversion.FindBytes(source, oldDataRoot, excelDataFile, cited, draftCopy: false) is string bytes)
                {
                    (sha, size) = DraftConversion.HashOf(bytes, hashes);
                }

                if (sha is null)
                    continue;

                string name = namer.BaseNameOf(cited, sha);
                files.TryAdd(name, new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = size });
            }

            var content = new BoardContent(draftBase.Board, draftBase.Calibrations, []);

            BoardDocument document = BoardDocumentMapping.ToDocument(content, namer.BaseNameFor, DraftConversion.Sorted(files.Values));
            document.RevisionDate = string.IsNullOrWhiteSpace(document.RevisionDate) ? draftBase.Revision : document.RevisionDate;

            return document;
        }

        // ###########################################################################################
        // NAMES FOR A DRAFT'S FILES, agreeing with the published board's (see the class header).
        // ###########################################################################################
        private sealed class Namer
        {
            private readonly BoardFileNames thisNames;
            private readonly BoardFileNames thisBaseNames;
            private readonly Dictionary<string, string> thisPublishedNameBySha;
            private readonly Dictionary<string, string> thisBaseShaByPath;

            public Namer(string boardId, IReadOnlyCollection<string> boardIds, BoardDocument? published, DraftBase? draftBase, string excelDataFile)
            {
                this.thisNames = new BoardFileNames(boardId, boardIds);
                this.thisBaseNames = new BoardFileNames(boardId, boardIds);
                this.thisPublishedNameBySha = new Dictionary<string, string>(StringComparer.Ordinal);

                foreach (BoardDocument.StoredFileRow file in published?.Files ?? [])
                {
                    this.thisNames.Reserve(file.Name, file.Sha256);
                    this.thisBaseNames.Reserve(file.Name, file.Sha256);
                    this.thisPublishedNameBySha.TryAdd(file.Sha256, file.Name);
                }

                // What was at each of the board's own paths when the draft started. The base keeps
                // its checksums in upper-case hex (DraftFingerprint's form); the 4.0 files, lower.
                this.thisBaseShaByPath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (DraftBaseFile file in draftBase?.Files ?? [])
                    this.thisBaseShaByPath.TryAdd(BoardFileNames.Normalise($"{DraftConversion.FolderOf(excelDataFile)}/{file.Path}"), file.Sha256.ToLowerInvariant());
            }

            public string NameOf(string cited, string sha)
            {
                string path = BoardFileNames.Normalise(cited);

                // The published bytes: the published name.
                if (this.thisPublishedNameBySha.TryGetValue(sha, out string? published) && this.IsFreeInDraft(published, sha))
                {
                    this.thisNames.Assign(path, published, sha);
                    return this.InDraft(published, sha);
                }

                // Replaced: the published name of what was at this path when the draft started -
                // unless the draft also keeps those published bytes under that name elsewhere.
                if (this.thisBaseShaByPath.TryGetValue(path, out string? baseSha) &&
                    this.thisPublishedNameBySha.TryGetValue(baseSha, out string? replaced) &&
                    this.IsFreeInDraft(replaced, sha))
                {
                    this.thisNames.Assign(path, replaced, sha);
                    return this.InDraft(replaced, sha);
                }

                return this.InDraft(this.thisNames.Add(path, sha), sha);
            }

            // One name, one file, in the draft: a name the draft already gave other bytes is not
            // given these.
            private readonly Dictionary<string, string> thisDraftShaByName = new(StringComparer.OrdinalIgnoreCase);

            private bool IsFreeInDraft(string name, string sha) =>
                !this.thisDraftShaByName.TryGetValue(name, out string? taken) || string.Equals(taken, sha, StringComparison.Ordinal);

            private string InDraft(string name, string sha)
            {
                this.thisDraftShaByName[name] = sha;
                return name;
            }

            public string NameFor(string cited) => this.thisNames.NameFor(cited);

            public string BaseNameOf(string cited, string sha)
            {
                string path = BoardFileNames.Normalise(cited);

                if (this.thisPublishedNameBySha.TryGetValue(sha, out string? published))
                {
                    this.thisBaseNames.Assign(path, published, sha);
                    return published;
                }

                return this.thisBaseNames.Add(path, sha);
            }

            public string BaseNameFor(string cited) => this.thisBaseNames.NameFor(cited);
        }

        // ---- the bytes -------------------------------------------------------------------------

        // ###########################################################################################
        // Where the Excel-era draft found a cited file: the draft's own copy first - the board's
        // files at their place in the board folder, a new shared file under its whole path
        // (DraftFileResolver's rule) - then the data folder. Null when neither holds it.
        // ###########################################################################################
        private static string? FindBytes(string source, string oldDataRoot, string excelDataFile, string cited, bool draftCopy = true)
        {
            string relative = BoardFileNames.Normalise(cited);

            if (draftCopy)
            {
                string inDraft = DraftFolderLayout.RelativeToBoardFolder(excelDataFile, relative) ?? relative;
                string draftPath = Path.Combine(source, inDraft.Replace('/', Path.DirectorySeparatorChar));

                if (File.Exists(draftPath))
                    return draftPath;
            }

            if (!string.IsNullOrWhiteSpace(oldDataRoot))
            {
                string dataPath = Path.Combine(oldDataRoot, relative.Replace('/', Path.DirectorySeparatorChar));

                if (File.Exists(dataPath))
                    return dataPath;
            }

            return null;
        }

        // Copies bytes into the draft's own store unless the 4.0 tree already stores them.
        private static void StoreUnlessPublished(string bytes, string sha, string name, string treeContentRoot, string working)
        {
            string stored = DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar);

            if (!string.IsNullOrWhiteSpace(treeContentRoot) && File.Exists(Path.Combine(treeContentRoot, stored)))
                return;

            string target = Path.Combine(working, stored);

            if (File.Exists(target))
                return;

            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(bytes, target);

            if (!string.Equals(DataTreeLayout.Sha256Of(target), sha, StringComparison.Ordinal))
                throw new IOException($"[{bytes}] changed while the draft was converted.");
        }

        private static (string Sha, long Size) HashOf(string path, Dictionary<string, (string Sha, long Size)> hashes)
        {
            string full = Path.GetFullPath(path);

            if (!hashes.TryGetValue(full, out (string Sha, long Size) known))
            {
                known = (DataTreeLayout.Sha256Of(full), new FileInfo(full).Length);
                hashes[full] = known;
            }

            return known;
        }

        // ---- reading and writing ---------------------------------------------------------------

        private static BoardDocument? ReadPublished(string treeContentRoot, string boardId, List<string> notes)
        {
            if (string.IsNullOrWhiteSpace(treeContentRoot))
                return null;

            string path = Path.Combine(treeContentRoot, DataTreeLayout.BoardFilePath(boardId).Replace('/', Path.DirectorySeparatorChar));

            if (!File.Exists(path))
                return null;

            try
            {
                return DataTreeJson.ReadBoard(File.ReadAllText(path));
            }
            catch (InvalidDataException ex)
            {
                notes.Add($"The published board could not be read ({ex.Message}), so the draft's files are named afresh.");
                return null;
            }
        }

        // ###########################################################################################
        // Moves every entry of the Excel-era draft into the set-aside folder, then marks that folder
        // complete - only then does anything new go beside it, so an interrupted run always knows
        // which of the two it was doing. Each step is a move inside the draft's own folder.
        // ###########################################################################################
        private static void SetAside(string draftFolder, string originals)
        {
            Directory.CreateDirectory(originals);

            foreach (string entry in Directory.EnumerateFileSystemEntries(draftFolder).ToList())
            {
                string name = Path.GetFileName(entry);

                if (name is DraftConversion.WorkingFolder or DraftConversion.OriginalsFolder)
                    continue;

                string target = Path.Combine(originals, name);

                if (Directory.Exists(entry))
                    Directory.Move(entry, target);
                else
                    File.Move(entry, target);
            }

            File.WriteAllText(Path.Combine(originals, DraftConversion.SetAsideMarker), "The Excel-era draft, kept by CRT 4.0's conversion.\n");
        }

        // The new draft into place, over whatever an interrupted move had already put there.
        private static void MoveIn(string working, string draftFolder)
        {
            foreach (string entry in Directory.EnumerateFileSystemEntries(working).ToList())
            {
                string target = Path.Combine(draftFolder, Path.GetFileName(entry));

                if (Directory.Exists(target))
                    Directory.Delete(target, recursive: true);

                if (Directory.Exists(entry))
                    Directory.Move(entry, target);
                else
                    File.Move(entry, target, overwrite: true);
            }

            Directory.Delete(working, recursive: true);
        }

        private static void WriteText(string path, string text) =>
            File.WriteAllBytes(path, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(text));

        private static List<BoardDocument.StoredFileRow> Sorted(IEnumerable<BoardDocument.StoredFileRow> files) =>
            files
                .OrderBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(file => file.Name, StringComparer.Ordinal)
                .ToList();

        private static string FolderOf(string path)
        {
            int slash = path.LastIndexOf('/');
            return slash < 0 ? string.Empty : path[..slash];
        }
    }
}
