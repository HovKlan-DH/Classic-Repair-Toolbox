using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A system the application already knows about, as the import needs to see it: its identity,
    // and whether the MAIN workbook lists it (published) or only a local source does - a legacy
    // "_UserContribution" workbook, say, whose board lives in Data/ but has never been published.
    // ###########################################################################################
    public sealed record KnownDraftSystem(string ExcelDataFile, bool IsPublished);

    public enum DraftFolderImportKind
    {
        NotImported,

        // A copy of a board the main workbook lists - an ordinary draft over the published board.
        DraftOfPublishedBoard,

        // Anything else - a system with nothing published, exactly as "Add a new system" makes one.
        NewSystem,
    }

    // ###########################################################################################
    // What the import did with one board folder, so the caller can log it. Reason says why a
    // folder was NOT imported; RenamedFrom names the workbook's original file when it had to be
    // renamed to the name the known system uses.
    // ###########################################################################################
    public sealed class DraftFolderImportOutcome
    {
        public string SystemFolder { get; init; } = string.Empty;

        public string ExcelDataFile { get; init; } = string.Empty;

        public DraftFolderImportKind Kind { get; init; }

        public bool Imported => this.Kind != DraftFolderImportKind.NotImported;

        public string RenamedFrom { get; init; } = string.Empty;

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
    // Data/ uses and DraftManager.EnumerateDraftOnlySystems walks), holding an .xlsx directly. The
    // two SHARED folders are skipped by name at either upper level - a "Shared files/Board local
    // files" folder can hold an .xlsx datasheet, and taking that for a board would put a system
    // called "Shared files" in the lists.
    //
    // WHICH KIND OF DRAFT, decided by the known systems and nothing else:
    //   - the main workbook lists this folder: a draft of that PUBLISHED board. Its base revision is
    //     the one written inside the copy itself, which is the honest answer to "what were these
    //     edits made on top of" - the published board may have moved on since the copy was taken,
    //     and the drift warning then says so;
    //   - anything else: a NEW system, registered under its FOLDER names. Those have to be the
    //     folder names and never a display name from a workbook: a submission sends the
    //     registration's names as the system's parts and the server rebuilds the system id from
    //     them (the identity.system_id_mismatch of 2026-09-23). A folder a legacy
    //     "_UserContribution" workbook lists is one of these - it has never been published.
    //
    // A KNOWN system's workbook must carry the name that system uses, because every draft lookup
    // builds the workbook path from it. A folder holding ONE workbook under another name (a copy
    // from an older data generation, say) has it renamed, with its JSON sidecar; with several and
    // none of them right, nothing is guessed.
    //
    // Only UNMARKED folders are touched, so running this on every load is harmless: a folder is
    // imported once and is an ordinary draft from then on. Nothing is ever deleted or overwritten.
    // ###########################################################################################
    public static class DraftFolderImport
    {
        public static IReadOnlyList<DraftFolderImportOutcome> ImportUnmarkedFolders(
            string draftsRoot,
            IEnumerable<KnownDraftSystem>? knownSystems,
            DateTimeOffset nowUtc)
        {
            var outcomes = new List<DraftFolderImportOutcome>();

            if (string.IsNullOrWhiteSpace(draftsRoot) || !Directory.Exists(draftsRoot))
            {
                return outcomes;
            }

            Dictionary<string, KnownDraftSystem> knownById = DraftFolderImport.IndexBySystemId(knownSystems);

            foreach (string boardFolder in DraftFolderImport.EnumerateBoardFolders(draftsRoot))
            {
                if (File.Exists(Path.Combine(boardFolder, DraftFolderLayout.DraftMarkerFileName)))
                {
                    // Already a draft - the ordinary case, and not worth a line in the log.
                    continue;
                }

                outcomes.Add(DraftFolderImport.ImportOne(boardFolder, knownById, nowUtc));
            }

            return outcomes;
        }

        private static DraftFolderImportOutcome ImportOne(
            string boardFolder,
            Dictionary<string, KnownDraftSystem> knownById,
            DateTimeOffset nowUtc)
        {
            string board = Path.GetFileName(boardFolder);
            string hardwareFolder = Path.GetDirectoryName(boardFolder) ?? string.Empty;
            string hardware = Path.GetFileName(hardwareFolder);
            string manufacturer = Path.GetFileName(Path.GetDirectoryName(hardwareFolder) ?? string.Empty);

            string systemId = $"{manufacturer}/{hardware}/{board}";

            try
            {
                List<string> workbooks = DraftFolderImport.BoardWorkbooksIn(boardFolder);
                if (workbooks.Count == 0)
                {
                    return DraftFolderImport.NotImported(boardFolder, "it holds no board workbook (.xlsx)");
                }

                knownById.TryGetValue(systemId, out KnownDraftSystem? known);

                // ###########################################################################################
                // *** THE CAPITALS MUST MATCH, because the published tree's do. *** The lookup is
                // case-insensitive so this case is SEEN rather than silently imported as a second,
                // new system beside the real one. On Linux "commodore/c64" is simply a different
                // folder from "Commodore/C64", so importing it under either name would be wrong.
                // ###########################################################################################
                if (known is not null &&
                    !string.Equals(SystemDescriptorRules.SystemIdFromExcelDataFile(known.ExcelDataFile), systemId, StringComparison.Ordinal))
                {
                    return DraftFolderImport.NotImported(
                        boardFolder,
                        $"its folder names differ only in capitals from [{SystemDescriptorRules.SystemIdFromExcelDataFile(known.ExcelDataFile)}] - " +
                        "rename the folders to match exactly");
                }

                if (known is null && !DraftFolderImport.IsUsableNewSystemIdentity(manufacturer, hardware, board, out string identityReason))
                {
                    return DraftFolderImport.NotImported(boardFolder, identityReason);
                }

                string? wantedName = known is not null
                    ? known.ExcelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries)[^1]
                    : workbooks.Count == 1 ? Path.GetFileName(workbooks[0]) : null;

                if (wantedName is null)
                {
                    return DraftFolderImport.NotImported(
                        boardFolder,
                        $"it holds {workbooks.Count} workbooks and it cannot tell which is the board - keep only the board's own");
                }

                string workbookPath = Path.Combine(boardFolder, wantedName);
                string renamedFrom = string.Empty;

                if (!workbooks.Any(path => string.Equals(Path.GetFileName(path), wantedName, StringComparison.Ordinal)))
                {
                    if (workbooks.Count != 1)
                    {
                        return DraftFolderImport.NotImported(
                            boardFolder,
                            $"none of its {workbooks.Count} workbooks is named [{wantedName}], the name this board uses");
                    }

                    renamedFrom = Path.GetFileName(workbooks[0]);
                    DraftFolderImport.RenameWorkbook(workbooks[0], workbookPath);
                }

                string excelDataFile = known?.ExcelDataFile ?? $"{systemId}/{wantedName}";
                string createdUtc = nowUtc.ToUniversalTime().ToString("o");

                bool isPublished = known?.IsPublished == true;

                DraftMarkerStore.Save(
                    Path.Combine(boardFolder, DraftFolderLayout.DraftMarkerFileName),
                    isPublished
                        ? new DraftMarker
                        {
                            SystemKey = excelDataFile,
                            BaseRevision = BoardDataReader.ReadRevisionDateOnly(workbookPath).Trim(),
                            CreatedUtc = createdUtc,
                        }
                        : new DraftMarker
                        {
                            SystemKey = excelDataFile,
                            BaseRevision = string.Empty,
                            CreatedUtc = createdUtc,
                            NewSystem = new NewSystemRegistration
                            {
                                HardwareName = hardware,
                                BoardName = board,
                                ExcelDataFile = excelDataFile,
                                CreatedUtc = createdUtc,
                            },
                        });

                return new DraftFolderImportOutcome
                {
                    SystemFolder = boardFolder,
                    ExcelDataFile = excelDataFile,
                    Kind = isPublished ? DraftFolderImportKind.DraftOfPublishedBoard : DraftFolderImportKind.NewSystem,
                    RenamedFrom = renamedFrom,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Most likely the workbook is open in Excel and could not be renamed. Nothing was
                // written, so the next load simply tries again.
                return DraftFolderImport.NotImported(boardFolder, ex.Message);
            }
            catch (Exception ex)
            {
                // ###########################################################################################
                // *** ANYTHING ELSE IS THIS FOLDER'S, NOT THE LOAD'S (code review, 2026-09-27). *** The
                // import runs inside DataManager's loading of the main workbook, before the board list
                // is set - so an exception escaping here left the application with NO boards at all,
                // for an optional convenience. It costs this one folder, and says why.
                // ###########################################################################################
                CrtLog.Warning($"Taking in [{boardFolder}] as a draft failed - [{ex}]");
                return DraftFolderImport.NotImported(boardFolder, ex.Message);
            }
        }

        // Every Manufacturer/Hardware/Board folder, skipping the two shared folders by name.
        private static IEnumerable<string> EnumerateBoardFolders(string draftsRoot)
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
                return Directory.GetDirectories(folder);
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
        // A new system's folder names become its identity, so they must be names "Add a new system"
        // would have accepted - including the server's rule that they carry no extra spaces
        // (SubmissionValidator's identity.parts_not_canonical).
        // ###########################################################################################
        private static bool IsUsableNewSystemIdentity(string manufacturer, string hardware, string board, out string reason)
        {
            foreach ((string part, string label) in new[] { (manufacturer, "manufacturer"), (hardware, "hardware"), (board, "board") })
            {
                if (!NewSystemIdentity.IsValidPathSegment(part, out string segmentReason))
                {
                    reason = $"its {label} folder name {segmentReason}";
                    return false;
                }

                if (!string.Equals(NewSystemIdentity.SanitizePathSegment(part), part, StringComparison.Ordinal))
                {
                    reason = $"its {label} folder name [{part}] has extra spaces";
                    return false;
                }
            }

            reason = string.Empty;
            return true;
        }

        // ###########################################################################################
        // Renames the workbook and its JSON sidecar together. The sidecar goes FIRST and is put
        // back if the workbook then cannot move (open in Excel), so a failure never leaves the
        // highlights beside a workbook name that does not exist.
        // ###########################################################################################
        private static void RenameWorkbook(string from, string to)
        {
            string fromSidecar = BoardComponentHighlightStorage.GetJsonPath(from);
            string toSidecar = BoardComponentHighlightStorage.GetJsonPath(to);

            bool sidecarMoved = false;

            if (File.Exists(fromSidecar) && !File.Exists(toSidecar))
            {
                File.Move(fromSidecar, toSidecar);
                sidecarMoved = true;
            }

            try
            {
                File.Move(from, to);
            }
            catch (Exception) when (sidecarMoved)
            {
                File.Move(toSidecar, fromSidecar);
                throw;
            }
        }

        // One entry per system id, and a PUBLISHED listing wins over a local one for the same
        // folder - it is the one the server knows.
        private static Dictionary<string, KnownDraftSystem> IndexBySystemId(IEnumerable<KnownDraftSystem>? knownSystems)
        {
            var byId = new Dictionary<string, KnownDraftSystem>(StringComparer.OrdinalIgnoreCase);

            foreach (KnownDraftSystem known in knownSystems ?? [])
            {
                string id = SystemDescriptorRules.SystemIdFromExcelDataFile(known.ExcelDataFile);
                if (id.Length == 0)
                {
                    continue;
                }

                if (!byId.TryGetValue(id, out KnownDraftSystem? existing) || (known.IsPublished && !existing.IsPublished))
                {
                    byId[id] = known;
                }
            }

            return byId;
        }

        private static DraftFolderImportOutcome NotImported(string boardFolder, string reason) =>
            new()
            {
                SystemFolder = boardFolder,
                Kind = DraftFolderImportKind.NotImported,
                Reason = reason,
            };
    }
}
