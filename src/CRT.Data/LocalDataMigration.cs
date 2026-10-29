using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // TAKING OVER WHAT CRT 3.x LEFT ON THIS MACHINE, AT EVERY 4.0 LAUNCH (owner decision,
    // 2026-10-24: "I do expect you also will migrate all existing data").
    //
    // Three things of a user's own lived where CRT 4.0 no longer reads them:
    //
    //   - DRAFTS made by 3.x - Excel-era board folders in Drafts/ - are converted in place into 4.0
    //     drafts (DraftConversion). Tried at every launch until each is converted: a converted
    //     draft is passed over at once, and one that could not be converted yet (its workbook open
    //     somewhere, say) is tried again next time.
    //   - BOARDS a 3.x "_UserContribution" workbook listed - boards made the old way, kept in the
    //     3.x data folder - are copied into Drafts/ as a board folder put there by hand, ONCE each,
    //     for the import to take in as new boards (DraftFolderImport, which converts them). A board
    //     the 4.0 main file lists is published and is left; a board that has a draft folder already
    //     is left; and a board carried once is never carried again, even after its draft is
    //     discarded (the record, CarriedRecordFileName, in the drafts folder).
    //   - OSCILLOSCOPE MODELS added to the 3.x copy of the main workbook (OwnOscilloscopeModels),
    //     kept in a file of the user's own (OwnOscilloscopeFile).
    //
    // The 3.x data folder is only READ - it stays as it was, for a 3.x CRT and for this.
    // ###########################################################################################
    public static class LocalDataMigration
    {
        // In the drafts folder: the boards already carried in from "_UserContribution" workbooks.
        public const string CarriedRecordFileName = DraftFolderLayout.BookkeepingPrefix + "carried-from-3x.json";

        private const string UserContributionSuffix = "_UserContribution" + DataGenerationRules.WorkbookExtension;

        private static readonly JsonSerializerOptions RecordJson = new() { WriteIndented = true };

        // ###########################################################################################
        // Converts every Excel-era draft in `draftsRoot` (DraftConversion.Convert - the rest as
        // there). Answers each draft it did something with, or tried to: one already converted is
        // left out.
        // ###########################################################################################
        public static IReadOnlyList<(string BoardFolder, DraftConversionResult Result)> ConvertExcelDrafts(
            string draftsRoot,
            string oldDataRoot,
            string treeContentRoot,
            IReadOnlyCollection<string> boardIds)
        {
            var results = new List<(string, DraftConversionResult)>();

            if (string.IsNullOrWhiteSpace(draftsRoot) || !Directory.Exists(draftsRoot))
                return results;

            foreach (string folder in DraftFolderImport.EnumerateBoardFolders(draftsRoot))
            {
                if (!File.Exists(Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName)))
                    continue;

                string boardId = string.Join('/', Path.GetRelativePath(draftsRoot, folder)
                    .Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries));

                DraftConversionResult result = DraftConversion.Convert(folder, boardId, oldDataRoot, treeContentRoot, boardIds);

                if (!result.AlreadyConverted)
                    results.Add((folder, result));
            }

            return results;
        }

        // ###########################################################################################
        // Copies the boards the "_UserContribution" workbooks in `oldDataRoot` list into
        // `draftsRoot` - see the header for which. Answers the board ids carried; `notes` says what
        // could not be. Each board is copied through a ".crt-" folder and moved into place whole,
        // so the import never sees half of one; its workbook, that workbook's JSON sidecar and every
        // other file of the board's folder travel - not another generation's workbook beside it,
        // which would make the folder two boards.
        // ###########################################################################################
        public static IReadOnlyList<string> CarryUserContributionBoards(
            string draftsRoot,
            string oldDataRoot,
            IReadOnlyCollection<string> listedBoardIds,
            List<string> notes)
        {
            ArgumentNullException.ThrowIfNull(notes);

            var carried = new List<string>();

            if (string.IsNullOrWhiteSpace(draftsRoot) || string.IsNullOrWhiteSpace(oldDataRoot) || !Directory.Exists(oldDataRoot))
                return carried;

            var listed = new HashSet<string>(listedBoardIds ?? [], StringComparer.OrdinalIgnoreCase);
            string recordPath = Path.Combine(draftsRoot, LocalDataMigration.CarriedRecordFileName);
            HashSet<string> record = LocalDataMigration.ReadRecord(recordPath);

            foreach (string workbook in Directory.EnumerateFiles(oldDataRoot, "*" + LocalDataMigration.UserContributionSuffix)
                         .Where(path => path.EndsWith(LocalDataMigration.UserContributionSuffix, StringComparison.OrdinalIgnoreCase))
                         .Order(StringComparer.Ordinal))
            {
                if (!MasterListing.TryRead(workbook, out IReadOnlyList<MasterListingRow> rows, out string why))
                {
                    notes.Add($"[{Path.GetFileName(workbook)}] could not be read ({why}), so its boards were not carried over.");
                    continue;
                }

                foreach (MasterListingRow row in rows)
                {
                    string boardId = BoardDescriptorRules.BoardIdFromExcelDataFile(row.ExcelDataFile);

                    if (!BoardDescriptorRules.IsValidBoardId(boardId))
                    {
                        notes.Add($"[{row.ExcelDataFile}] in [{Path.GetFileName(workbook)}] names no board folder CRT can use, so it was not carried over.");
                        continue;
                    }

                    if (listed.Contains(boardId) || record.Contains(boardId))
                        continue;

                    string target = Path.Combine(draftsRoot, boardId.Replace('/', Path.DirectorySeparatorChar));

                    if (Directory.Exists(target))
                    {
                        notes.Add($"[{boardId}] already has a draft folder, so the copy in the CRT 3 data folder was left where it is.");
                        continue;
                    }

                    string source = Path.Combine(oldDataRoot, boardId.Replace('/', Path.DirectorySeparatorChar));
                    string workbookName = Path.GetFileName(row.ExcelDataFile.Trim());

                    if (!File.Exists(Path.Combine(source, workbookName)))
                    {
                        notes.Add($"[{row.ExcelDataFile}] is not in the CRT 3 data folder, so [{boardId}] was not carried over.");
                        continue;
                    }

                    try
                    {
                        LocalDataMigration.CopyBoardFolder(draftsRoot, source, target, workbookName);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        notes.Add($"[{boardId}] could not be copied into the drafts folder ({ex.Message}) - tried again at the next start.");
                        continue;
                    }

                    record.Add(boardId);
                    carried.Add(boardId);

                    // After each, so a stop keeps what was done.
                    LocalDataMigration.WriteRecord(recordPath, record);
                }
            }

            return carried;
        }

        // ###########################################################################################
        // The oscilloscope models the user added to their copy of 3.x's main workbook (the newest
        // generation's, in `oldDataRoot`), measured against the published ones - none when there is
        // no such workbook. A sheet it cannot fully read is noted.
        // ###########################################################################################
        public static IReadOnlyList<MainDocument.OscilloscopeRow> OwnScopeModels(
            string oldDataRoot,
            IReadOnlyList<MainDocument.OscilloscopeRow> published,
            List<string> notes) =>
            LocalDataMigration.OwnScopeModels(oldDataRoot, published, notes, out _);

        // The same, saying whether the search was MADE (code review, 2026-10-25): false when 3.x's
        // workbook is there but could not be opened - held by another program, or broken - so
        // nothing says what the user's models are, and the search is to be made again next time.
        public static IReadOnlyList<MainDocument.OscilloscopeRow> OwnScopeModels(
            string oldDataRoot,
            IReadOnlyList<MainDocument.OscilloscopeRow> published,
            List<string> notes,
            out bool searched)
        {
            ArgumentNullException.ThrowIfNull(notes);

            searched = true;

            if (string.IsNullOrWhiteSpace(oldDataRoot) || !Directory.Exists(oldDataRoot) ||
                MasterListing.NewestMasterPath(oldDataRoot) is not string master)
            {
                return [];
            }

            var problems = new List<string>();
            IReadOnlyList<MainDocument.OscilloscopeRow> local = DataTreeConversion.ReadOscilloscopes(master, problems, out bool opened);

            notes.AddRange(problems.Select(problem => $"The oscilloscope sheet of [{Path.GetFileName(master)}]: {problem}"));

            searched = opened;

            return OwnOscilloscopeModels.Find(local, published ?? []);
        }

        private static void CopyBoardFolder(string draftsRoot, string source, string target, string workbookName)
        {
            string workbookSidecar = Path.ChangeExtension(workbookName, ".json");
            string staging = Path.Combine(draftsRoot, $"{DraftFolderLayout.BookkeepingPrefix}carrying-{Guid.NewGuid():N}");

            try
            {
                foreach (string path in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
                {
                    string relative = Path.GetRelativePath(source, path);
                    string name = Path.GetFileName(path);
                    bool atTop = !relative.Contains(Path.DirectorySeparatorChar) && !relative.Contains(Path.AltDirectorySeparatorChar);

                    if (name.StartsWith("~$", StringComparison.Ordinal))
                        continue;

                    // Another workbook of the folder - an older generation's - and its sidecar.
                    if (atTop &&
                        !string.Equals(name, workbookName, StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(name, workbookSidecar, StringComparison.OrdinalIgnoreCase) &&
                        (name.EndsWith(DataGenerationRules.WorkbookExtension, StringComparison.OrdinalIgnoreCase) ||
                         File.Exists(Path.Combine(source, Path.ChangeExtension(name, DataGenerationRules.WorkbookExtension)))))
                    {
                        continue;
                    }

                    string destination = Path.Combine(staging, relative);
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                    File.Copy(path, destination);
                }

                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                Directory.Move(staging, target);
            }
            finally
            {
                if (Directory.Exists(staging))
                {
                    try
                    {
                        Directory.Delete(staging, recursive: true);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        CrtLog.Warning($"Could not tidy away [{staging}] - [{ex.Message}]");
                    }
                }
            }
        }

        private static HashSet<string> ReadRecord(string path)
        {
            try
            {
                if (File.Exists(path) && JsonSerializer.Deserialize<List<string>>(File.ReadAllText(path)) is { } ids)
                    return new HashSet<string>(ids, StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read [{path}] - [{ex.Message}]; the boards carried from CRT 3 are judged by their draft folders alone");
            }

            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private static void WriteRecord(string path, HashSet<string> ids)
        {
            try
            {
                FileReplacer.Replace(path, temporary => File.WriteAllText(
                    temporary,
                    JsonSerializer.Serialize(ids.Order(StringComparer.OrdinalIgnoreCase).ToList(), LocalDataMigration.RecordJson)));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not write [{path}] - [{ex.Message}]");
            }
        }
    }

    // ###########################################################################################
    // THE USER'S OWN OSCILLOSCOPE MODELS - a file of their own beside the settings
    // (Classic-Repair-Toolbox.own-oscilloscopes.json), a main file holding only oscilloscopes
    // (DataTreeJson). CRT lists them after the published models.
    //
    // At the first 4.0 launch it is made from what the user added to 3.x's main workbook
    // (LocalDataMigration.OwnScopeModels) - EVEN WHEN THAT IS NOTHING: the file being there is what
    // says the search was made, so a model the user deletes later does not come back.
    //
    // *** BUT NOT WHEN THE SEARCH COULD NOT BE MADE (code review, 2026-10-25). *** 3.x's workbook held
    // open by another program, or broken, read as "no models", and the empty file then stopped every
    // later launch from looking - the user's own SCPI command sets were never carried over, and
    // nothing said so. Then nothing is written, and the search is made again at the next start.
    // ###########################################################################################
    public static class OwnOscilloscopeFile
    {
        public const string FileName = "Classic-Repair-Toolbox.own-oscilloscopes.json";

        // The models in the file at `path`, made first from 3.x's workbook in `oldDataRoot` when there
        // is no file yet. Empty for no path, and for a file that cannot be read (noted).
        public static IReadOnlyList<MainDocument.OscilloscopeRow> LoadOrFind(
            string path,
            string oldDataRoot,
            IReadOnlyList<MainDocument.OscilloscopeRow> published,
            List<string> notes) =>
            OwnOscilloscopeFile.LoadOrFind(path, oldDataRoot, published, notes, out _);

        // ###########################################################################################
        // The same, saying when the file must NOT be written this run, and why (code review,
        // 2026-10-29): it could not be read - written, it would replace models nobody could read with
        // the list held, which is none - or the search of 3.x's workbook could not be made, which a
        // file would end for good. Null when it may be written. The Oscilloscope tab's model editor
        // writes nothing while it is not null.
        // ###########################################################################################
        public const string NotReadText =
            "Your own oscilloscope models could not be read when CRT started, so CRT changes none of them until it can - nothing is lost. ";

        public static IReadOnlyList<MainDocument.OscilloscopeRow> LoadOrFind(
            string path,
            string oldDataRoot,
            IReadOnlyList<MainDocument.OscilloscopeRow> published,
            List<string> notes,
            out string? notWritable)
        {
            ArgumentNullException.ThrowIfNull(notes);

            notWritable = null;

            if (string.IsNullOrWhiteSpace(path))
                return [];

            if (!File.Exists(path))
            {
                IReadOnlyList<MainDocument.OscilloscopeRow> found = LocalDataMigration.OwnScopeModels(oldDataRoot, published, notes, out bool searched);

                if (!searched)
                {
                    notes.Add("Your own oscilloscope models from CRT 3 could not be read - looked for again at the next start.");
                    notWritable = OwnOscilloscopeFile.NotReadText +
                        "CRT 3's main data file, where they are, is open in another program or broken. Close it, and start CRT again.";
                    return found;
                }

                try
                {
                    OwnOscilloscopeFile.Write(path, found);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    notes.Add($"Your own oscilloscope models could not be kept in [{path}] ({ex.Message}) - looked for again at the next start.");
                }

                return found;
            }

            try
            {
                return DataTreeJson.ReadMain(File.ReadAllText(path)).Oscilloscopes;
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                notes.Add($"Your own oscilloscope models in [{path}] could not be read ({ex.Message}).");
                notWritable = OwnOscilloscopeFile.NotReadText +
                    $"The file [{path}] could not be read ({ex.Message}). Put it right or move it away, and start CRT again.";
                return [];
            }
        }

        public static void Write(string path, IReadOnlyList<MainDocument.OscilloscopeRow> models)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);

            string text = DataTreeJson.Write(new MainDocument { Oscilloscopes = (models ?? []).ToList() });
            FileReplacer.Replace(path, temporary => File.WriteAllText(temporary, text, new System.Text.UTF8Encoding(false)));
        }
    }
}
