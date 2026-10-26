using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A board the application already knows about, as the import needs to see it: its identity,
    // and whether the main file lists it (published) or only a local source does.
    // ###########################################################################################
    public sealed record KnownDraftBoard(string ExcelDataFile, bool IsPublished);

    public enum DraftFolderImportKind
    {
        NotImported,

        // A copy of a board the main file lists - an ordinary draft over the published board.
        DraftOfPublishedBoard,

        // Anything else - a board with nothing published, exactly as "Add a new board" makes one.
        NewBoard,
    }

    // ###########################################################################################
    // What the import did with one board folder, so the caller can log it. Reason says why a
    // folder was NOT imported; ConvertedFromWorkbook names the Excel workbook an Excel-era copy
    // was converted from.
    // ###########################################################################################
    public sealed class DraftFolderImportOutcome
    {
        public string BoardFolder { get; init; } = string.Empty;

        public string ExcelDataFile { get; init; } = string.Empty;

        public DraftFolderImportKind Kind { get; init; }

        public bool Imported => this.Kind != DraftFolderImportKind.NotImported;

        public string ConvertedFromWorkbook { get; init; } = string.Empty;

        public string Reason { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // TURNS A BOARD FOLDER PUT INTO Drafts/ BY HAND INTO A DRAFT (owner request, 2026-09-27).
    //
    // *** THIS REVERSES "A DRAFT IS ONLY EVER SOMETHING THIS APPLICATION CREATED", on purpose. ***
    // A contributor may already have work on a board - made the old way, in Excel, or on another
    // machine - and the owner's words were that CRT should "import that into its Draft so it can be
    // submitted". Without this such a folder is invisible: it has no marker, so it is not a draft,
    // the Drafts tab never lists it and it can never be submitted.
    //
    // *** THE MARKER STILL MAKES A FOLDER A DRAFT - this only writes one. *** DraftBoardSource,
    // DraftStatusReader, retirement and discard all go on asking the marker and nothing else, so
    // an imported folder is from then on indistinguishable from one the application seeded, and
    // there is no second notion of "draft" to keep in step.
    //
    // WHAT IS A BOARD FOLDER: exactly three levels down (Manufacturer/Hardware/Board, the layout
    // the data tree uses and DraftManager.EnumerateDraftOnlyBoards walks), holding a board file
    // (board.json) or a board workbook (.xlsx) directly. The two SHARED folders are skipped by name
    // at either upper level - a "Shared files/Board local files" folder can hold an .xlsx
    // datasheet, and taking that for a board would put a board called "Shared files" in the lists.
    //
    // *** IN CRT 4.0 THE FOLDER IS TAKEN IN AS A 4.0 DRAFT (2026-10-24). *** Two shapes arrive:
    //   - a 4.0 board folder (a board.json, copied from another machine's drafts, say): it gets its
    //     marker and nothing else changes;
    //   - an Excel-era board folder (ONE workbook, under whatever name - a copy from the old data
    //     folder or from a 3.x draft): it gets an Excel-era marker and is then converted in place
    //     like every 3.x draft at the first 4.0 launch (DraftConversion) - its files named as the
    //     published board names them, its Excel files kept in ".crt-pre-4.0". A folder the
    //     conversion refuses is left as it was, marker and all taken away again, and says why.
    //
    // WHICH KIND OF DRAFT, decided by the known boards and nothing else:
    //   - the main file lists this folder: a draft of that PUBLISHED board. Its base revision is
    //     the one written inside the copy itself, which is the honest answer to "what were these
    //     edits made on top of" - the published board may have moved on since the copy was taken,
    //     and the drift warning then says so;
    //   - anything else: a NEW board, registered under its FOLDER names. Those have to be the
    //     folder names and never a display name from a workbook: a submission sends the
    //     registration's names as the board's parts and the server rebuilds the board id from
    //     them (the identity.board_id_mismatch of 2026-09-23).
    //
    // Only UNMARKED folders are touched, so running this on every load is harmless: a folder is
    // imported once and is an ordinary draft from then on. Nothing is ever deleted or overwritten.
    // ###########################################################################################
    public static class DraftFolderImport
    {
        // ###########################################################################################
        // `oldDataRoot`: the Excel data folder an Excel-era copy's rows may cite shared files from
        // (CRT 3.x's, when it is still there); `treeContentRoot`: the 4.0 tree, whose published
        // board names the files. Either may be empty - the files are then named afresh.
        // ###########################################################################################
        public static IReadOnlyList<DraftFolderImportOutcome> ImportUnmarkedFolders(
            string draftsRoot,
            IEnumerable<KnownDraftBoard>? knownBoards,
            DateTimeOffset nowUtc,
            string oldDataRoot = "",
            string treeContentRoot = "")
        {
            var outcomes = new List<DraftFolderImportOutcome>();

            if (string.IsNullOrWhiteSpace(draftsRoot) || !Directory.Exists(draftsRoot))
            {
                return outcomes;
            }

            List<KnownDraftBoard> known = (knownBoards ?? []).ToList();
            Dictionary<string, KnownDraftBoard> knownById = DraftFolderImport.IndexByBoardId(known);
            IReadOnlyCollection<string> boardIds = knownById.Where(pair => pair.Value.IsPublished).Select(pair => pair.Key).ToList();

            foreach (string boardFolder in DraftFolderImport.EnumerateBoardFolders(draftsRoot))
            {
                if (File.Exists(Path.Combine(boardFolder, DraftFolderLayout.DraftMarkerFileName)))
                {
                    // Already a draft - the ordinary case, and not worth a line in the log.
                    continue;
                }

                outcomes.Add(DraftFolderImport.ImportOne(boardFolder, knownById, boardIds, nowUtc, oldDataRoot, treeContentRoot));
            }

            return outcomes;
        }

        private static DraftFolderImportOutcome ImportOne(
            string boardFolder,
            Dictionary<string, KnownDraftBoard> knownById,
            IReadOnlyCollection<string> boardIds,
            DateTimeOffset nowUtc,
            string oldDataRoot,
            string treeContentRoot)
        {
            string board = Path.GetFileName(boardFolder);
            string hardwareFolder = Path.GetDirectoryName(boardFolder) ?? string.Empty;
            string hardware = Path.GetFileName(hardwareFolder);
            string manufacturer = Path.GetFileName(Path.GetDirectoryName(hardwareFolder) ?? string.Empty);

            string boardId = $"{manufacturer}/{hardware}/{board}";
            string markerPath = Path.Combine(boardFolder, DraftFolderLayout.DraftMarkerFileName);

            try
            {
                string boardFile = Path.Combine(boardFolder, DataTreeLayout.BoardFileName);
                bool isBoardFile = File.Exists(boardFile);
                List<string> workbooks = isBoardFile ? [] : DraftFolderImport.BoardWorkbooksIn(boardFolder);

                if (!isBoardFile && workbooks.Count == 0)
                {
                    return DraftFolderImport.NotImported(boardFolder, "it holds no board file (board.json) and no board workbook (.xlsx)");
                }

                knownById.TryGetValue(boardId, out KnownDraftBoard? known);

                // ###########################################################################################
                // *** THE CAPITALS MUST MATCH, because the published tree's do. *** The lookup is
                // case-insensitive so this case is SEEN rather than silently imported as a second,
                // new board beside the real one. On Linux "commodore/c64" is simply a different
                // folder from "Commodore/C64", so importing it under either name would be wrong.
                // ###########################################################################################
                if (known is not null &&
                    !string.Equals(BoardDescriptorRules.BoardIdFromExcelDataFile(known.ExcelDataFile), boardId, StringComparison.Ordinal))
                {
                    return DraftFolderImport.NotImported(
                        boardFolder,
                        $"its folder names differ only in capitals from [{BoardDescriptorRules.BoardIdFromExcelDataFile(known.ExcelDataFile)}] - " +
                        "rename the folders to match exactly");
                }

                if (known is null && !DraftFolderImport.IsUsableNewBoardIdentity(manufacturer, hardware, board, out string identityReason))
                {
                    return DraftFolderImport.NotImported(boardFolder, identityReason);
                }

                if (workbooks.Count > 1)
                {
                    return DraftFolderImport.NotImported(
                        boardFolder,
                        $"it holds {workbooks.Count} workbooks and it cannot tell which is the board - keep only the board's own");
                }

                bool isPublished = known?.IsPublished == true;
                string excelDataFile = known?.ExcelDataFile ?? DataTreeLayout.BoardFilePath(boardId);
                string createdUtc = nowUtc.ToUniversalTime().ToString("o");

                // An Excel-era copy is marked under its workbook's own path - the identity its rows'
                // paths are relative to - and converted from there.
                string markedAs = isBoardFile ? excelDataFile : $"{boardId}/{Path.GetFileName(workbooks[0])}";
                string baseRevision = !isPublished
                    ? string.Empty
                    : isBoardFile
                        ? BoardDataReader.ReadRevisionDateOnly(boardFile).Trim()
                        : BoardDataReader.ReadRevisionDateOnly(workbooks[0]).Trim();

                DraftMarkerStore.Save(
                    markerPath,
                    new DraftMarker
                    {
                        BoardKey = markedAs,
                        BaseRevision = baseRevision,
                        CreatedUtc = createdUtc,
                        NewBoard = isPublished
                            ? null
                            : new NewBoardRegistration
                            {
                                HardwareName = hardware,
                                BoardName = board,
                                ExcelDataFile = excelDataFile,
                                CreatedUtc = createdUtc,
                            },
                    });

                if (!isBoardFile)
                {
                    DraftConversionResult converted = DraftConversion.Convert(boardFolder, boardId, oldDataRoot, treeContentRoot, boardIds);

                    if (!converted.Converted)
                    {
                        File.Delete(markerPath);
                        return DraftFolderImport.NotImported(boardFolder, converted.Failure ?? "it could not be converted into a CRT 4.0 draft");
                    }

                    foreach (string note in converted.Notes)
                    {
                        CrtLog.Info($"Taking in [{boardFolder}] as a draft - {note}");
                    }
                }

                return new DraftFolderImportOutcome
                {
                    BoardFolder = boardFolder,
                    ExcelDataFile = excelDataFile,
                    Kind = isPublished ? DraftFolderImportKind.DraftOfPublishedBoard : DraftFolderImportKind.NewBoard,
                    ConvertedFromWorkbook = isBoardFile ? string.Empty : Path.GetFileName(workbooks[0]),
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Most likely a file of it is open in another program. Nothing stays written, so
                // the next load simply tries again.
                DraftFolderImport.TryDelete(markerPath);
                return DraftFolderImport.NotImported(boardFolder, ex.Message);
            }
            catch (Exception ex)
            {
                // ###########################################################################################
                // *** ANYTHING ELSE IS THIS FOLDER'S, NOT THE LOAD'S (code review, 2026-09-27). *** The
                // import runs inside DataManager's loading of the main file, before the board list
                // is set - so an exception escaping here left the application with NO boards at all,
                // for an optional convenience. It costs this one folder, and says why.
                // ###########################################################################################
                CrtLog.Warning($"Taking in [{boardFolder}] as a draft failed - [{ex}]");
                DraftFolderImport.TryDelete(markerPath);
                return DraftFolderImport.NotImported(boardFolder, ex.Message);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not remove [{path}] after a failed import - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Every Manufacturer/Hardware/Board folder, skipping the two shared folders by name - and
        // CRT's own ".crt-" folders at every level (a board being carried in by LocalDataMigration
        // is copied through one, and must not be taken in half-copied).
        // ###########################################################################################
        internal static IEnumerable<string> EnumerateBoardFolders(string draftsRoot)
        {
            foreach (string manufacturerFolder in DraftFolderImport.SafeDirectories(draftsRoot))
            {
                if (SubmissionFileScopes.IsSharedFolderName(Path.GetFileName(manufacturerFolder)))
                {
                    continue;
                }

                foreach (string hardwareFolder in DraftFolderImport.SafeDirectories(manufacturerFolder))
                {
                    if (SubmissionFileScopes.IsSharedFolderName(Path.GetFileName(hardwareFolder)))
                    {
                        continue;
                    }

                    foreach (string boardFolder in DraftFolderImport.SafeDirectories(hardwareFolder))
                    {
                        yield return boardFolder;
                    }
                }
            }
        }

        // An unreadable folder is skipped rather than ending the whole walk.
        private static IEnumerable<string> SafeDirectories(string folder)
        {
            try
            {
                return Directory.GetDirectories(folder)
                    .Where(path => !Path.GetFileName(path).StartsWith(DraftFolderLayout.BookkeepingPrefix, StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not look inside [{folder}] for board folders - [{ex.Message}]");
                return [];
            }
        }

        // ###########################################################################################
        // The .xlsx files directly in a board folder, less Excel's "~$" owner file - it appears
        // beside every workbook that is open, and would otherwise make one workbook look like two.
        // Filtered by hand rather than by a "*.xlsx" pattern, whose matching is case-insensitive on
        // Windows and case-sensitive elsewhere.
        // ###########################################################################################
        private static List<string> BoardWorkbooksIn(string boardFolder)
        {
            return Directory.GetFiles(boardFolder)
                .Where(path => path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();
        }

        // ###########################################################################################
        // A new board's folder names become its identity, so they must be names "Add a new board"
        // would have accepted - including the server's rule that they carry no extra spaces
        // (SubmissionValidator's identity.parts_not_canonical).
        // ###########################################################################################
        private static bool IsUsableNewBoardIdentity(string manufacturer, string hardware, string board, out string reason)
        {
            foreach ((string part, string label) in new[] { (manufacturer, "manufacturer"), (hardware, "hardware"), (board, "board") })
            {
                if (!NewBoardIdentity.IsValidPathSegment(part, out string segmentReason))
                {
                    reason = $"its {label} folder name {segmentReason}";
                    return false;
                }

                if (!string.Equals(NewBoardIdentity.SanitizePathSegment(part), part, StringComparison.Ordinal))
                {
                    reason = $"its {label} folder name [{part}] has extra spaces";
                    return false;
                }
            }

            reason = string.Empty;
            return true;
        }

        // One entry per board id, and a PUBLISHED listing wins over a local one for the same
        // folder - it is the one the server knows.
        private static Dictionary<string, KnownDraftBoard> IndexByBoardId(IEnumerable<KnownDraftBoard>? knownBoards)
        {
            var byId = new Dictionary<string, KnownDraftBoard>(StringComparer.OrdinalIgnoreCase);

            foreach (KnownDraftBoard known in knownBoards ?? [])
            {
                string id = BoardDescriptorRules.BoardIdFromExcelDataFile(known.ExcelDataFile);
                if (id.Length == 0)
                {
                    continue;
                }

                if (!byId.TryGetValue(id, out KnownDraftBoard? existing) || (known.IsPublished && !existing.IsPublished))
                {
                    byId[id] = known;
                }
            }

            return byId;
        }

        private static DraftFolderImportOutcome NotImported(string boardFolder, string reason) =>
            new()
            {
                BoardFolder = boardFolder,
                Kind = DraftFolderImportKind.NotImported,
                Reason = reason,
            };
    }
}
