using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeOpenXml;

namespace Handlers.DataHandling
{
    // One file of the Excel-era tree that the CRT 4.0 tree carries.
    public sealed record DataTreeSourceFile(string RelativePath, string FullPath, string Sha256, long Size);

    // One board, converted: its id, the workbook it came from (data-root relative), its board file,
    // and the name it gave each path its rows cite (for anything that later has to follow a path
    // to its new name - a user's own settings, say).
    public sealed record ConvertedBoard(
        string BoardId,
        string SourceWorkbook,
        BoardDocument Document,
        IReadOnlyDictionary<string, string> NamesByCitedPath);

    // What Write did: files written, files already there with the right bytes, stale files removed,
    // and the manifest's entries (-1 when no manifest was asked for).
    public sealed record DataTreeWriteResult(int Written, int Kept, int Removed, int ManifestEntries);

    // ###########################################################################################
    // EVERYTHING A CONVERSION WILL WRITE, worked out before anything is written.
    //
    // `StoredFiles` and `CopiedFiles` are content-folder relative ("Files/3f/3f...pdf",
    // "MiniPro/IC tests/..."), each with the source file it is copied from. `LeftOut` lists the
    // source tree's files the 4.0 tree does not carry - the older generations' workbooks, what
    // only they cite, and "!" notes no board cites - so nobody has to wonder what was dropped.
    //
    // *** A PLAN WITH A PROBLEM IS NOT WRITTEN. *** Every problem is something the conversion
    // would otherwise lose or get wrong; Write refuses such a plan.
    // ###########################################################################################
    public sealed class DataTreeConversionPlan
    {
        public required string SourceDataRoot { get; init; }

        public Version? Generation { get; init; }

        public required MainDocument Main { get; init; }

        public required IReadOnlyList<ConvertedBoard> Boards { get; init; }

        public required IReadOnlyDictionary<string, DataTreeSourceFile> StoredFiles { get; init; }

        public required IReadOnlyDictionary<string, DataTreeSourceFile> CopiedFiles { get; init; }

        // Every source file the conversion carries, by its path in the source tree - two paths to
        // the same bytes are two entries here and one stored file.
        public required IReadOnlyDictionary<string, DataTreeSourceFile> SourceFiles { get; init; }

        public required IReadOnlyList<string> LeftOut { get; init; }

        public required IReadOnlyList<string> Problems { get; init; }

        public required IReadOnlyList<string> Notes { get; init; }

        public bool CanWrite => this.Problems.Count == 0;

        // What the boards' files add up to as each board names them, and what is actually stored.
        public long BoardFileBytes => this.Boards.SelectMany(board => board.Document.Files).Sum(file => file.Size);

        public long StoredBytes => this.StoredFiles.Values.Sum(file => file.Size);
    }

    // ###########################################################################################
    // CONVERTS AN EXCEL-ERA DATA TREE INTO THE CRT 4.0 TREE (owner decisions, 2026-10-24: board data
    // as JSON only, every board owning its files, each file's bytes stored once, all existing data
    // migrated). DataTreeLayout says where everything goes.
    //
    // *** IT CONVERTS THE NEWEST GENERATION, AS CRT READS IT. *** The newest versioned main
    // workbook, the boards it lists, and each board through BoardDataReader - the app's own reader,
    // so the converted board is exactly the board a CRT shows. The older generations serve older
    // CRTs and stay where they are, frozen.
    //
    // *** FAIL CLOSED: ANYTHING IT COULD NOT CARRY IS A PROBLEM, AND A PLAN WITH ONE IS NOT
    // WRITTEN. *** CRT's readers are forgiving - a column they do not know is skipped, a sidecar
    // entry they cannot read is left out - which is right for showing a board and wrong for a
    // one-way conversion. So the converter checks what the readers would quietly drop: a sheet or
    // a column the schema does not know, a value under no known heading, a header row it cannot
    // find, a repeated heading, a sidecar root it does not know or an entry the readers skip, a file
    // a row cites that is not there, a board the main workbook lists that is missing, a board
    // folder of the newest generation the main workbook does not list, an oscilloscope row with
    // values but no model name.
    //
    // *** IT ONLY READS THE SOURCE, AND WRITES ONLY INTO ITS OWN KIND OF FOLDER. *** Write refuses
    // a target inside the source (or the source inside it), and a non-empty target that is not a
    // tree this converter wrote. Into its own tree it writes exactly the plan: files already there
    // with the right bytes are kept, everything else under Data/ is removed - so a second run on
    // release day leaves the same tree a first run on fresh data would.
    // ###########################################################################################
    public static class DataTreeConversion
    {
        private const string GenericSharedMiniProFolder = DataTreeUsage.MiniProTestsFolder;

        // ###########################################################################################
        // Works out the conversion of the tree at `sourceDataRoot` (its "Data" folder). Reads the
        // source and hashes every file it carries; writes nothing.
        // ###########################################################################################
        public static DataTreeConversionPlan Plan(string sourceDataRoot)
        {
            var problems = new List<string>();
            var notes = new List<string>();
            var boards = new List<ConvertedBoard>();
            var stored = new Dictionary<string, DataTreeSourceFile>(StringComparer.Ordinal);
            var copied = new Dictionary<string, DataTreeSourceFile>(StringComparer.Ordinal);
            var main = new MainDocument();

            string root = string.IsNullOrWhiteSpace(sourceDataRoot) ? string.Empty : Path.GetFullPath(sourceDataRoot);

            if (root.Length == 0 || !Directory.Exists(root))
            {
                problems.Add($"The data folder [{sourceDataRoot}] could not be found.");
                return DataTreeConversion.Finish(root, null, main, boards, stored, copied, new Dictionary<string, DataTreeSourceFile>(), [], problems, notes);
            }

            EpplusLicense.Ensure();

            IReadOnlyList<string> files = DataTreeConversion.ListFiles(root, problems);
            ILookup<string, string> onDisk = files.ToLookup(path => path, StringComparer.OrdinalIgnoreCase);
            var hashes = new Dictionary<string, DataTreeSourceFile>(StringComparer.Ordinal);
            var carried = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string? masterPath = MasterListing.NewestMasterPath(root);

            if (masterPath is null)
            {
                problems.Add("The data folder holds no versioned main workbook (Classic-Repair-Toolbox.v<version>.xlsx).");
                return DataTreeConversion.Finish(root, null, main, boards, stored, copied, hashes, files, problems, notes);
            }

            Version? generation = DataGenerationRules.TryReadGeneration(Path.GetFileName(masterPath));
            carried.Add(Path.GetFileName(masterPath));

            IReadOnlyList<MasterListingRow> listed = DataTreeConversion.ReadMain(masterPath, main, problems);

            // Every board the tree holds, for naming files cited from another board's folder.
            List<string> boardIds = listed
                .Select(row => row.BoardId)
                .Concat(files.Where(DataTreeUsage.IsBoardFolderWorkbook).Select(DataTreeConversion.FolderOf))
                .Where(id => id.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            DataTreeConversion.CheckEveryNewestBoardIsListed(files, generation, listed, problems);

            // A board's folder sits beside the tree's own folders, so it may not be one of them.
            foreach (MasterListingRow row in listed.Where(row => DataTreeLayout.ReservedTopFolders.Contains(row.BoardId.Split('/')[0], StringComparer.OrdinalIgnoreCase)))
                problems.Add($"Board [{row.BoardId}] is filed under [{row.BoardId.Split('/')[0]}], a folder name the 4.0 tree keeps for itself.");

            // Read every board first, then hash everything they cite at once - on all cores, as
            // hashing a gigabyte is most of a conversion's time - and only then name the files,
            // in each board's own order.
            List<BoardRead> reads = listed
                .Select(row => DataTreeConversion.ReadBoard(root, row, files, onDisk, carried, problems, notes))
                .OfType<BoardRead>()
                .ToList();

            List<string> miniPro = DataTreeConversion.MiniProFiles(files, problems);

            DataTreeConversion.HashAll(
                root,
                reads.SelectMany(read => read.Cited.Select(cited => cited.Relative)).Concat(miniPro),
                hashes);

            foreach (BoardRead read in reads)
            {
                if (DataTreeConversion.NameFiles(read, boardIds, hashes, stored, carried, problems) is ConvertedBoard board)
                    boards.Add(board);
            }

            foreach (string test in miniPro)
            {
                copied[$"{DataTreeLayout.MiniProTestsFolder}/{test[(DataTreeConversion.GenericSharedMiniProFolder.Length + 1)..]}"] = hashes[test];
                carried.Add(test);
            }

            return DataTreeConversion.Finish(
                root, generation, main, boards, stored, copied, hashes, files.Where(path => !carried.Contains(path)).ToList(), problems, notes);
        }

        // ###########################################################################################
        // Writes the plan into the tree at `targetTreeRoot` - its "Data" folder and, given the
        // tree's public address (".../app-data-v4/Data"), the manifest beside it. Throws
        // InvalidOperationException, writing nothing, for a plan with problems or a target it must
        // not write; IOException when a source file changed since the plan was made.
        // ###########################################################################################
        public static DataTreeWriteResult Write(DataTreeConversionPlan plan, string targetTreeRoot, string? publicBaseUrl)
        {
            ArgumentNullException.ThrowIfNull(plan);

            if (!plan.CanWrite)
                throw new InvalidOperationException(
                    $"The conversion has {plan.Problems.Count} problem(s) and is not written: " + string.Join(" | ", plan.Problems));

            if (string.IsNullOrWhiteSpace(targetTreeRoot))
                throw new InvalidOperationException("No folder was given to write the converted data into.");

            string target = Path.GetFullPath(targetTreeRoot);
            string content = Path.Combine(target, DataTreeLayout.ContentFolder);
            string manifestPath = Path.Combine(target, DataTreeLayout.ManifestFileName);

            DataTreeConversion.GuardTarget(target, plan.SourceDataRoot);

            if (string.IsNullOrWhiteSpace(publicBaseUrl) && File.Exists(manifestPath))
                throw new InvalidOperationException(
                    $"[{target}] has a manifest, which this conversion would leave describing the old files. " +
                    "Give the tree's public address so it is rebuilt.");

            var planned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int written = 0;
            int kept = 0;

            foreach ((string relative, DataTreeSourceFile source) in plan.StoredFiles.Concat(plan.CopiedFiles).OrderBy(pair => pair.Key, StringComparer.Ordinal))
            {
                planned.Add(relative);

                if (DataTreeConversion.CopyVerified(source, Path.Combine(content, DataTreeConversion.Native(relative))))
                    written++;
                else
                    kept++;
            }

            foreach (ConvertedBoard board in plan.Boards)
            {
                string relative = DataTreeLayout.BoardFilePath(board.BoardId);
                planned.Add(relative);

                if (DataTreeConversion.WriteText(Path.Combine(content, DataTreeConversion.Native(relative)), DataTreeJson.Write(board.Document)))
                    written++;
                else
                    kept++;
            }

            planned.Add(DataTreeLayout.MainFileName);

            if (DataTreeConversion.WriteText(Path.Combine(content, DataTreeLayout.MainFileName), DataTreeJson.Write(plan.Main)))
                written++;
            else
                kept++;

            int removed = DataTreeConversion.RemoveUnplanned(content, planned);

            int manifestEntries = -1;

            if (!string.IsNullOrWhiteSpace(publicBaseUrl))
            {
                manifestEntries = DataChecksumManifest.Write(content, publicBaseUrl, manifestPath, hashEveryFile: true);

                if (manifestEntries < 0)
                    throw new IOException($"The manifest [{manifestPath}] could not be written - see the log.");
            }

            return new DataTreeWriteResult(written, kept, removed, manifestEntries);
        }

        // ---- the main workbook -----------------------------------------------------------------

        private static IReadOnlyList<MasterListingRow> ReadMain(string masterPath, MainDocument main, List<string> problems)
        {
            string masterName = Path.GetFileName(masterPath);

            if (!MasterListing.TryRead(masterPath, out IReadOnlyList<MasterListingRow> rows, out string why))
            {
                problems.Add($"The main workbook [{masterName}] could not be read: {why}");
                return [];
            }

            foreach (IGrouping<string, MasterListingRow> twice in rows.GroupBy(row => row.BoardId, StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1))
                problems.Add($"The main workbook [{masterName}] lists board [{twice.Key}] {twice.Count()} times.");

            foreach (MasterListingRow row in rows.Where(row => row.BoardId.Length == 0))
                problems.Add($"The main workbook [{masterName}] lists [{row.ExcelDataFile}], which names no board folder.");

            main.Boards = rows.Select(row => new MainDocument.BoardListingRow
            {
                BoardId = row.BoardId,
                HardwareName = row.HardwareName,
                BoardName = row.BoardName,
                HardwareNotes = row.Notes,
            }).ToList();

            try
            {
                using var stream = new FileStream(masterPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using ExcelPackage package = EpplusLicense.OpenPackage(stream);

                foreach (ExcelWorksheet sheet in package.Workbook.Worksheets)
                {
                    if (string.Equals(sheet.Name, MasterWorkbookSchema.SheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        SheetText? listing = DataTreeConversion.ReadSheet(
                            sheet,
                            [MasterWorkbookSchema.ColHardwareName, MasterWorkbookSchema.ColBoardName, MasterWorkbookSchema.ColExcelDataFile],
                            [MasterWorkbookSchema.ColHardwareName, MasterWorkbookSchema.ColBoardName, MasterWorkbookSchema.ColExcelDataFile, MasterWorkbookSchema.ColHardwareNotes],
                            $"The main workbook [{masterName}]",
                            problems);

                        main.BoardsRevisionDate = listing?.RevisionDate ?? string.Empty;
                    }
                    else if (string.Equals(sheet.Name, MasterWorkbookSchema.OscilloscopeSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        SheetText? scopes = DataTreeConversion.ReadSheet(
                            sheet,
                            [MasterWorkbookSchema.ColBrand, MasterWorkbookSchema.ColSeriesOrModel, MasterWorkbookSchema.ColPort],
                            MasterWorkbookSchema.OscilloscopeColumns,
                            $"The main workbook [{masterName}]",
                            problems);

                        main.OscilloscopesRevisionDate = scopes?.RevisionDate ?? string.Empty;
                        main.Oscilloscopes = scopes is null ? [] : DataTreeConversion.Oscilloscopes(scopes, masterName, problems);
                    }
                    else if (DataTreeConversion.HasAnyText(sheet))
                    {
                        problems.Add($"The main workbook [{masterName}] has a sheet [{sheet.Name}] that CRT does not read.");
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException or InvalidOperationException)
            {
                problems.Add($"The main workbook [{masterName}] could not be read: {ex.Message}");
            }

            return rows;
        }

        // ###########################################################################################
        // The oscilloscopes of the main workbook at `masterPath`, as the conversion reads them - for
        // a CRT 4.0 first launch to find the models a user added to their OWN copy of the workbook
        // (OwnOscilloscopeModels). Problems are added to `problems`; what could be read is returned.
        // `opened` says whether the workbook could be opened at all - false for one not there, held
        // by another program or broken, where nothing says what the user's models are.
        // ###########################################################################################
        public static IReadOnlyList<MainDocument.OscilloscopeRow> ReadOscilloscopes(string masterPath, List<string> problems) =>
            DataTreeConversion.ReadOscilloscopes(masterPath, problems, out _);

        public static IReadOnlyList<MainDocument.OscilloscopeRow> ReadOscilloscopes(string masterPath, List<string> problems, out bool opened)
        {
            ArgumentNullException.ThrowIfNull(problems);

            string masterName = Path.GetFileName(masterPath ?? string.Empty);
            opened = false;

            if (string.IsNullOrWhiteSpace(masterPath) || !File.Exists(masterPath))
            {
                problems.Add($"The main workbook [{masterPath}] is not there.");
                return [];
            }

            EpplusLicense.Ensure();

            try
            {
                using var stream = new FileStream(masterPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using ExcelPackage package = EpplusLicense.OpenPackage(stream);

                ExcelWorksheet? sheet = package.Workbook.Worksheets[MasterWorkbookSchema.OscilloscopeSheetName];
                opened = true;

                if (sheet is null)
                    return [];

                SheetText? scopes = DataTreeConversion.ReadSheet(
                    sheet,
                    [MasterWorkbookSchema.ColBrand, MasterWorkbookSchema.ColSeriesOrModel, MasterWorkbookSchema.ColPort],
                    MasterWorkbookSchema.OscilloscopeColumns,
                    $"The main workbook [{masterName}]",
                    problems);

                return scopes is null ? [] : DataTreeConversion.Oscilloscopes(scopes, masterName, problems);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException or InvalidOperationException)
            {
                problems.Add($"The main workbook [{masterName}] could not be read: {ex.Message}");
                opened = false;
                return [];
            }
        }

        // The Oscilloscope sheet's rows as CRT reads them: a blank brand is the brand above, a row
        // with no model name is skipped - unless it holds something, which would be lost.
        private static List<MainDocument.OscilloscopeRow> Oscilloscopes(SheetText sheet, string masterName, List<string> problems)
        {
            var models = new List<MainDocument.OscilloscopeRow>();
            string brand = string.Empty;

            foreach ((int sheetRow, IReadOnlyDictionary<string, string> cells) in sheet.Rows)
            {
                string Cell(string column) => cells.TryGetValue(column, out string? value) ? value : string.Empty;

                if (Cell(MasterWorkbookSchema.ColBrand).Length > 0)
                    brand = Cell(MasterWorkbookSchema.ColBrand);

                if (Cell(MasterWorkbookSchema.ColSeriesOrModel).Length == 0)
                {
                    if (cells.Any(pair => pair.Value.Length > 0 && pair.Key != MasterWorkbookSchema.ColBrand))
                        problems.Add($"The main workbook [{masterName}] has an oscilloscope row {sheetRow} with values but no series or model.");

                    continue;
                }

                models.Add(new MainDocument.OscilloscopeRow
                {
                    Brand = brand,
                    SeriesOrModel = Cell(MasterWorkbookSchema.ColSeriesOrModel),
                    Port = Cell(MasterWorkbookSchema.ColPort),
                    DebounceTime = Cell(MasterWorkbookSchema.ColDebounceTime),
                    Identify = Cell(MasterWorkbookSchema.ColIdentify),
                    DrainErrorQueue = Cell(MasterWorkbookSchema.ColDrainErrorQueue),
                    OperationComplete = Cell(MasterWorkbookSchema.ColOperationComplete),
                    ClearStatistics = Cell(MasterWorkbookSchema.ColClearStatistics),
                    QueryActiveTrigger = Cell(MasterWorkbookSchema.ColQueryActiveTrigger),
                    Stop = Cell(MasterWorkbookSchema.ColStop),
                    Single = Cell(MasterWorkbookSchema.ColSingle),
                    Run = Cell(MasterWorkbookSchema.ColRun),
                    QueryTriggerMode = Cell(MasterWorkbookSchema.ColQueryTriggerMode),
                    QueryTriggerLevel = Cell(MasterWorkbookSchema.ColQueryTriggerLevel),
                    SetTriggerLevel = Cell(MasterWorkbookSchema.ColSetTriggerLevel),
                    QueryTimeDiv = Cell(MasterWorkbookSchema.ColQueryTimeDiv),
                    SetTimeDiv = Cell(MasterWorkbookSchema.ColSetTimeDiv),
                    QueryVoltsDiv = Cell(MasterWorkbookSchema.ColQueryVoltsDiv),
                    SetVoltsDiv = Cell(MasterWorkbookSchema.ColSetVoltsDiv),
                    DumpImage = Cell(MasterWorkbookSchema.ColDumpImage),
                    TimeDivList = Cell(MasterWorkbookSchema.ColTimeDivList),
                    VoltsDivList = Cell(MasterWorkbookSchema.ColVoltsDivList),
                });
            }

            return models;
        }

        // A board folder holding a workbook of the generation being converted, which the main
        // workbook does not list, would vanish from the 4.0 tree.
        private static void CheckEveryNewestBoardIsListed(
            IReadOnlyList<string> files,
            Version? generation,
            IReadOnlyList<MasterListingRow> listed,
            List<string> problems)
        {
            var listedIds = listed.Select(row => row.BoardId).ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (string workbook in files.Where(DataTreeUsage.IsBoardFolderWorkbook))
            {
                if (DataTreeConversion.FolderOf(workbook) is { Length: > 0 } folder &&
                    !listedIds.Contains(folder) &&
                    Equals(DataGenerationRules.TryReadGeneration(Path.GetFileName(workbook)), generation))
                {
                    problems.Add($"The board workbook [{workbook}] is not listed in the main workbook.");
                }
            }
        }

        // ---- one board -------------------------------------------------------------------------

        // A board read, before its files are named: its content, and every file it cites in the
        // order it shows them (the path as cited, and the file it found in the tree).
        private sealed record BoardRead(
            string BoardId,
            string Workbook,
            BoardContent Content,
            IReadOnlyList<(string Cited, string Relative)> Cited);

        // ###########################################################################################
        // Reads one listed board: its workbook (through BoardDataReader, after checking it holds
        // nothing the reader would skip), its sidecar (likewise), and the files it cites - those
        // its rows name, and everything in its "KiCad data" folder. Null, with the problems said,
        // when any of it could not be carried.
        // ###########################################################################################
        private static BoardRead? ReadBoard(
            string root,
            MasterListingRow row,
            IReadOnlyList<string> files,
            ILookup<string, string> onDisk,
            HashSet<string> carried,
            List<string> problems,
            List<string> notes)
        {
            string boardId = row.BoardId;

            if (boardId.Length == 0)
                return null;

            if (DataTreeConversion.Resolve(onDisk, row.ExcelDataFile, out string? workbook, out string? why) is false)
            {
                problems.Add($"Board [{boardId}]: the main workbook lists [{row.ExcelDataFile}], {why}");
                return null;
            }

            string workbookPath = Path.Combine(root, DataTreeConversion.Native(workbook!));
            int problemsBefore = problems.Count;

            carried.Add(workbook!);
            DataTreeConversion.CheckWorkbookShape(workbookPath, boardId, problems);

            string sidecar = Path.ChangeExtension(workbook!, ".json");

            if (onDisk.Contains(sidecar))
            {
                if (!onDisk[sidecar].Contains(sidecar, StringComparer.Ordinal))
                    problems.Add($"Board [{boardId}]: its JSON file is spelled differently from its workbook [{workbook}], so CRT reads no highlights for it.");

                carried.Add(sidecar);
            }

            BoardData? data = BoardDataReader.ReadWorkbookUncached(workbookPath);

            if (data is null)
            {
                problems.Add($"Board [{boardId}]: the workbook [{workbook}] could not be read.");
                return null;
            }

            var content = new BoardContent(
                data,
                BoardComponentHighlightStorage.LoadAllKiCadCalibrations(workbookPath),
                BoardComponentHighlightStorage.LoadSchematicRenames(workbookPath));

            DataTreeConversion.CheckSidecar(BoardComponentHighlightStorage.GetJsonPath(workbookPath), content, boardId, problems);

            var cited = new List<(string Cited, string Relative)>();

            void Cite(string? path)
            {
                if (string.IsNullOrWhiteSpace(path))
                    return;

                if (DataTreeConversion.Resolve(onDisk, path, out string? relative, out string? missing) is false)
                {
                    problems.Add($"Board [{boardId}]: a row cites [{path}], {missing}");
                    return;
                }

                if (!string.Equals(relative, BoardFileNames.Normalise(path), StringComparison.Ordinal))
                    notes.Add($"Board [{boardId}]: [{path}] is stored as [{relative}] - converted to the stored file.");

                cited.Add((path, relative!));
            }

            // In the order the board shows its files, so the first to ask keeps a name both want.
            foreach (BoardSchematicEntry schematic in data.Schematics)
                Cite(schematic.SchematicImageFile);

            foreach (ComponentImageEntry image in data.ComponentImages)
                Cite(image.File);

            foreach (ComponentLocalFileEntry file in data.ComponentLocalFiles)
                Cite(file.File);

            foreach (BoardLocalFileEntry file in data.BoardLocalFiles)
                Cite(file.File);

            string kiCadFolder = $"{DataTreeConversion.FolderOf(workbook!)}/{DataTreeUsage.KiCadFolderName}/";

            foreach (string kiCadFile in files.Where(path => path.StartsWith(kiCadFolder, StringComparison.OrdinalIgnoreCase)))
                Cite(kiCadFile);

            return problems.Count > problemsBefore ? null : new BoardRead(boardId, workbook!, content, cited);
        }

        // ###########################################################################################
        // Names a read board's files (BoardFileNames, in the order it cites them), records each
        // file's bytes as stored under its checksum, and builds the board file.
        // ###########################################################################################
        private static ConvertedBoard? NameFiles(
            BoardRead read,
            IReadOnlyList<string> boardIds,
            Dictionary<string, DataTreeSourceFile> hashes,
            Dictionary<string, DataTreeSourceFile> stored,
            HashSet<string> carried,
            List<string> problems)
        {
            var names = new BoardFileNames(read.BoardId, boardIds);
            var boardFiles = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach ((string cited, string relative) in read.Cited)
            {
                DataTreeSourceFile source = hashes[relative];
                string name = names.Add(cited, source.Sha256);

                carried.Add(relative);
                boardFiles.TryAdd(name, new BoardDocument.StoredFileRow { Name = name, Sha256 = source.Sha256, Size = source.Size });
                stored.TryAdd(DataTreeLayout.StoredFilePath(source.Sha256, name), source);
            }

            BoardDocument document;

            try
            {
                document = BoardDocumentMapping.ToDocument(
                    read.Content,
                    names.NameFor,
                    boardFiles.Values
                        .OrderBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(file => file.Name, StringComparer.Ordinal)
                        .ToList());
            }
            catch (FormatException ex)
            {
                problems.Add($"Board [{read.BoardId}]: {ex.Message}");
                return null;
            }

            return new ConvertedBoard(read.BoardId, read.Workbook, document, new Dictionary<string, string>(names.NamesByPath, StringComparer.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Everything in a board workbook that BoardDataReader would not read: a sheet the schema
        // does not know, a heading it does not know or that repeats, a value under no known
        // heading, a sheet with values and no header row it can find.
        // ###########################################################################################
        private static void CheckWorkbookShape(string workbookPath, string boardId, List<string> problems)
        {
            try
            {
                using var stream = new FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using ExcelPackage package = EpplusLicense.OpenPackage(stream);

                foreach (ExcelWorksheet sheet in package.Workbook.Worksheets)
                {
                    BoardWorkbookSchema.SheetDefinition? definition = BoardWorkbookSchema.AllSheets
                        .FirstOrDefault(known => string.Equals(known.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase));

                    if (definition is null)
                    {
                        if (DataTreeConversion.HasAnyText(sheet))
                            problems.Add($"Board [{boardId}]: the workbook has a sheet [{sheet.Name}] that CRT does not read.");

                        continue;
                    }

                    DataTreeConversion.ReadSheet(sheet, definition.RequiredHeaders, definition.ColumnOrder, $"Board [{boardId}]", problems);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException or InvalidOperationException)
            {
                problems.Add($"Board [{boardId}]: the workbook could not be read: {ex.Message}");
            }
        }

        // ###########################################################################################
        // The sidecar as it is on disk, against what CRT's readers made of it: a root they do not
        // know, a file that does not parse, or fewer highlights, calibrations or renames than the
        // file holds (the readers skip what they cannot read) - each would be lost.
        // ###########################################################################################
        private static void CheckSidecar(string sidecarPath, BoardContent content, string boardId, List<string> problems)
        {
            if (!File.Exists(sidecarPath))
                return;

            JsonDocument json;

            try
            {
                string text = File.ReadAllText(sidecarPath);

                if (string.IsNullOrWhiteSpace(text))
                    return;

                json = JsonDocument.Parse(text);
            }
            catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
            {
                problems.Add($"Board [{boardId}]: its JSON file could not be read: {ex.Message}");
                return;
            }

            using (json)
            {
                if (json.RootElement.ValueKind != JsonValueKind.Object)
                {
                    problems.Add($"Board [{boardId}]: its JSON file holds no object.");
                    return;
                }

                int highlights = 0;
                int calibrations = 0;
                int renames = 0;

                foreach (JsonProperty part in json.RootElement.EnumerateObject())
                {
                    if (part.NameEquals(BoardSidecarWriter.HighlightsRoot) && part.Value.ValueKind == JsonValueKind.Object)
                    {
                        highlights = part.Value.EnumerateObject()
                            .Where(schematic => schematic.Value.ValueKind == JsonValueKind.Object)
                            .SelectMany(schematic => schematic.Value.EnumerateObject())
                            .Sum(label => label.Value.ValueKind == JsonValueKind.Array ? label.Value.GetArrayLength() : 1);
                    }
                    else if (part.NameEquals(BoardSidecarWriter.CalibrationRoot) && part.Value.ValueKind == JsonValueKind.Object)
                    {
                        calibrations = part.Value.EnumerateObject().Count();
                    }
                    else if (part.NameEquals(SchematicRenameLog.Root) && part.Value.ValueKind == JsonValueKind.Array)
                    {
                        renames = part.Value.GetArrayLength();
                    }
                    else
                    {
                        problems.Add($"Board [{boardId}]: its JSON file has a part [{part.Name}] that CRT does not read.");
                    }
                }

                if (highlights != content.Board.ComponentHighlights.Count)
                    problems.Add($"Board [{boardId}]: its JSON file holds {highlights} highlight(s), and CRT reads {content.Board.ComponentHighlights.Count} of them.");

                if (calibrations != content.Calibrations.Count)
                    problems.Add($"Board [{boardId}]: its JSON file holds {calibrations} KiCad calibration(s), and CRT reads {content.Calibrations.Count} of them.");

                if (renames != content.Renames.Count)
                    problems.Add($"Board [{boardId}]: its JSON file holds {renames} schematic rename(s), and CRT reads {content.Renames.Count} of them.");
            }
        }

        // ---- the MiniPro IC tests --------------------------------------------------------------

        // CRT reads them by folder, so every file in it goes, under DataTreeLayout's own folder.
        private static List<string> MiniProFiles(IReadOnlyList<string> files, List<string> problems)
        {
            string folder = DataTreeConversion.GenericSharedMiniProFolder + "/";
            List<string> tests = files.Where(path => path.StartsWith(folder, StringComparison.OrdinalIgnoreCase)).ToList();

            if (tests.Count == 0)
                problems.Add($"The data folder has no MiniPro IC tests ([{DataTreeConversion.GenericSharedMiniProFolder}]).");

            return tests;
        }

        // ---- reading a sheet -------------------------------------------------------------------

        // A sheet's rows under its header row, keyed by the known headings, and the date on its
        // "# Revision date:" line above the header.
        private sealed record SheetText(string RevisionDate, IReadOnlyList<(int SheetRow, IReadOnlyDictionary<string, string> Cells)> Rows);

        // ###########################################################################################
        // Finds the header row as CRT's readers do (the first row carrying every required heading),
        // and reports what they would skip: a heading not in `knownHeaders`, a repeated one, a value
        // in a column with no known heading, or values with no header row at all. Null without a
        // header row.
        // ###########################################################################################
        private static SheetText? ReadSheet(
            ExcelWorksheet sheet,
            IReadOnlyList<string> requiredHeaders,
            IReadOnlyList<string> knownHeaders,
            string who,
            List<string> problems)
        {
            int maxRow = sheet.Dimension?.End.Row ?? 0;
            int maxColumn = sheet.Dimension?.End.Column ?? 0;

            for (int row = 1; row <= maxRow; row++)
            {
                var headings = new Dictionary<int, string>();

                for (int column = 1; column <= maxColumn; column++)
                {
                    string text = DataTreeConversion.Heading(WorkbookCellText.Trimmed(sheet, row, column));

                    if (text.Length > 0)
                        headings[column] = text;
                }

                if (!requiredHeaders.All(required => headings.Values.Contains(required, StringComparer.OrdinalIgnoreCase)))
                    continue;

                var columns = new Dictionary<int, string>();

                foreach ((int column, string heading) in headings)
                {
                    string? known = knownHeaders.FirstOrDefault(name => string.Equals(name, heading, StringComparison.OrdinalIgnoreCase));

                    if (known is null)
                        problems.Add($"{who}: sheet [{sheet.Name}] has a column [{heading}] that CRT does not read.");
                    else if (columns.ContainsValue(known))
                        problems.Add($"{who}: sheet [{sheet.Name}] has the column [{heading}] twice - CRT reads only the first.");
                    else
                        columns[column] = known;
                }

                var rows = new List<(int, IReadOnlyDictionary<string, string>)>();

                for (int dataRow = row + 1; dataRow <= maxRow; dataRow++)
                {
                    var cells = new Dictionary<string, string>(StringComparer.Ordinal);

                    for (int column = 1; column <= maxColumn; column++)
                    {
                        string value = WorkbookCellText.Trimmed(sheet, dataRow, column);

                        if (columns.TryGetValue(column, out string? heading))
                            cells[heading] = value;
                        else if (value.Length > 0)
                            problems.Add($"{who}: sheet [{sheet.Name}] row {dataRow} has [{value}] under no column CRT reads.");
                    }

                    if (cells.Values.Any(value => value.Length > 0))
                        rows.Add((dataRow, cells));
                }

                return new SheetText(DataTreeConversion.RevisionDateAbove(sheet, row, maxColumn), rows);
            }

            if (DataTreeConversion.HasAnyText(sheet, skipMarkers: true))
                problems.Add($"{who}: sheet [{sheet.Name}] has values but no header row CRT can find.");

            return null;
        }

        // "# Revision date: 2026-October-4" in the rows above the header, as the readers find it.
        private static string RevisionDateAbove(ExcelWorksheet sheet, int headerRow, int maxColumn)
        {
            for (int row = 1; row < headerRow; row++)
            {
                for (int column = 1; column <= maxColumn; column++)
                {
                    string text = WorkbookCellText.Trimmed(sheet, row, column);

                    if (text.StartsWith(MasterWorkbookSchema.RevisionDatePrefix, StringComparison.OrdinalIgnoreCase))
                        return text[MasterWorkbookSchema.RevisionDatePrefix.Length..].Trim();
                }
            }

            return string.Empty;
        }

        // A heading as the readers compare it: Alt+Enter line breaks folded into one space.
        private static string Heading(string text) =>
            string.Join(" ", text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)).Trim();

        // Whether a sheet holds any text - its "#" preamble lines aside when `skipMarkers`. Every
        // workbook CRT writes gives each sheet a header row, so a sheet with text and none is odd
        // enough to look at, its title band included.
        private static bool HasAnyText(ExcelWorksheet sheet, bool skipMarkers = false)
        {
            int maxRow = sheet.Dimension?.End.Row ?? 0;
            int maxColumn = sheet.Dimension?.End.Column ?? 0;

            for (int row = 1; row <= maxRow; row++)
            {
                for (int column = 1; column <= maxColumn; column++)
                {
                    string text = WorkbookCellText.Trimmed(sheet, row, column);

                    if (text.Length == 0 || (skipMarkers && text.StartsWith('#')))
                        continue;

                    return true;
                }
            }

            return false;
        }

        // ---- files -----------------------------------------------------------------------------

        // Every file under the root the manifest would list, root relative with "/" separators.
        private static IReadOnlyList<string> ListFiles(string root, List<string> problems)
        {
            try
            {
                return Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
                    .Select(path => Path.GetRelativePath(root, path).Replace('\\', '/'))
                    .Where(DataChecksumManifest.IsSyncable)
                    .OrderBy(path => path, StringComparer.Ordinal)
                    .ToList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                problems.Add($"The data folder could not be read completely: {ex.Message}");
                return [];
            }
        }

        // ###########################################################################################
        // The file a path names in the tree: the exact spelling when it is there, otherwise the one
        // spelling that differs only in capitals (which every Windows and macOS CRT reads). Several
        // such spellings and no exact one is ambiguous; none is missing.
        // ###########################################################################################
        private static bool Resolve(ILookup<string, string> onDisk, string? path, out string? relative, out string? why)
        {
            string wanted = BoardFileNames.Normalise(path);
            List<string> spellings = onDisk[wanted].ToList();

            relative = spellings.FirstOrDefault(spelling => string.Equals(spelling, wanted, StringComparison.Ordinal))
                       ?? (spellings.Count == 1 ? spellings[0] : null);

            why = relative is not null
                ? null
                : spellings.Count == 0
                    ? "which is not in the data."
                    : $"which is in the data in {spellings.Count} spellings: {string.Join(", ", spellings)}.";

            return relative is not null;
        }

        // Hashes every one of these files not hashed yet, on all cores - a gigabyte of pictures and
        // datasheets is most of a conversion's time.
        private static void HashAll(string root, IEnumerable<string> relativePaths, Dictionary<string, DataTreeSourceFile> hashes)
        {
            List<string> wanted = relativePaths
                .Distinct(StringComparer.Ordinal)
                .Where(relative => !hashes.ContainsKey(relative))
                .ToList();

            var found = new System.Collections.Concurrent.ConcurrentDictionary<string, DataTreeSourceFile>(StringComparer.Ordinal);

            System.Threading.Tasks.Parallel.ForEach(wanted, relative =>
            {
                string full = Path.Combine(root, DataTreeConversion.Native(relative));
                found[relative] = new DataTreeSourceFile(relative, full, DataTreeLayout.Sha256Of(full), new FileInfo(full).Length);
            });

            foreach ((string relative, DataTreeSourceFile file) in found)
                hashes[relative] = file;
        }

        // ---- writing ---------------------------------------------------------------------------

        // ###########################################################################################
        // The target may be an empty or missing folder, or a tree this converter wrote (it holds the
        // 4.0 main file and no workbook) - never the source, inside it, or around it.
        // ###########################################################################################
        private static void GuardTarget(string target, string source)
        {
            string sourceFull = Path.GetFullPath(source);

            if (DataTreeConversion.IsSameOrInside(target, sourceFull) || DataTreeConversion.IsSameOrInside(sourceFull, target))
                throw new InvalidOperationException($"[{target}] and the data being converted [{sourceFull}] overlap - choose a folder of its own.");

            if (!Directory.Exists(target) || !Directory.EnumerateFileSystemEntries(target).Any())
                return;

            bool isConverted = File.Exists(Path.Combine(target, DataTreeLayout.ContentFolder, DataTreeLayout.MainFileName));
            bool holdsWorkbooks = Directory
                .EnumerateFiles(target, "*" + DataGenerationRules.WorkbookExtension, SearchOption.AllDirectories)
                .Any();

            if (!isConverted || holdsWorkbooks)
                throw new InvalidOperationException(
                    $"[{target}] holds other data. Write into an empty folder, or into one this conversion wrote before.");
        }

        private static bool IsSameOrInside(string path, string folder)
        {
            string a = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
            string b = Path.TrimEndingDirectorySeparator(Path.GetFullPath(folder));

            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase) ||
                   a.StartsWith(b + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }

        // Copies a source file into place through a temporary beside the target, checking the copy's
        // bytes are the planned ones. False when the target already held exactly those bytes.
        private static bool CopyVerified(DataTreeSourceFile source, string target)
        {
            if (File.Exists(target) &&
                new FileInfo(target).Length == source.Size &&
                string.Equals(DataTreeLayout.Sha256Of(target), source.Sha256, StringComparison.Ordinal))
            {
                return false;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(target)!);

            FileReplacer.Replace(target, temporary =>
            {
                File.Copy(source.FullPath, temporary, overwrite: true);

                if (!string.Equals(DataTreeLayout.Sha256Of(temporary), source.Sha256, StringComparison.Ordinal))
                    throw new IOException($"[{source.RelativePath}] changed since the conversion was planned - plan it again.");
            });

            return true;
        }

        // Writes the text when the file does not already hold exactly it. False when it did.
        private static bool WriteText(string target, string text)
        {
            byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(text);

            if (File.Exists(target) && File.ReadAllBytes(target).AsSpan().SequenceEqual(bytes))
                return false;

            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            FileReplacer.Replace(target, temporary => File.WriteAllBytes(temporary, bytes));

            return true;
        }

        // Every file under the content folder the plan did not write - a board, a file or a test
        // that is no longer there - and the folders that leaves empty.
        private static int RemoveUnplanned(string content, HashSet<string> planned)
        {
            if (!Directory.Exists(content))
                return 0;

            int removed = 0;

            foreach (string path in Directory.EnumerateFiles(content, "*", SearchOption.AllDirectories).ToList())
            {
                if (planned.Contains(Path.GetRelativePath(content, path).Replace('\\', '/')))
                    continue;

                File.Delete(path);
                removed++;
            }

            foreach (string folder in Directory.EnumerateDirectories(content, "*", SearchOption.AllDirectories).OrderByDescending(path => path.Length).ToList())
            {
                if (!Directory.EnumerateFileSystemEntries(folder).Any())
                    Directory.Delete(folder);
            }

            return removed;
        }

        // ---- small helpers ---------------------------------------------------------------------

        private static string FolderOf(string path)
        {
            int slash = path.LastIndexOf('/');
            return slash < 0 ? string.Empty : path[..slash];
        }

        private static string Native(string relative) => relative.Replace('/', Path.DirectorySeparatorChar);

        private static DataTreeConversionPlan Finish(
            string root,
            Version? generation,
            MainDocument main,
            List<ConvertedBoard> boards,
            Dictionary<string, DataTreeSourceFile> stored,
            Dictionary<string, DataTreeSourceFile> copied,
            IReadOnlyDictionary<string, DataTreeSourceFile> sources,
            IReadOnlyList<string> leftOut,
            List<string> problems,
            List<string> notes) => new()
            {
                SourceDataRoot = root,
                Generation = generation,
                Main = main,
                Boards = boards,
                StoredFiles = stored,
                CopiedFiles = copied,
                SourceFiles = sources,
                LeftOut = leftOut,
                Problems = problems,
                Notes = notes,
            };
    }
}
