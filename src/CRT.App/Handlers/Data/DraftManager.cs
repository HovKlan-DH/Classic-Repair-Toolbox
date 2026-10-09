using CRT;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Resolves and owns the "Drafts/" root - a second tree that sits beside "Data/" in AppData and
    // mirrors its folder structure exactly (NewContributeStrategy.md Phase 2). A contributor's
    // edits are written ONLY here; DataManager's sync continues to overwrite "Data/" freely and
    // never reads from or writes to this root, so a sync can never destroy a draft.
    //
    // This class only resolves the root and maps a board to its folder under it - the same split
    // DataManager (root resolution, sync) keeps from BoardDataReader (board file parsing).
    // Reading/writing one board's draft.json is DraftDataStore, in CRT.Data, since that logic is
    // pure and needed by the Maintainer tab too; only "where is Drafts/" is an app concern.
    //
    // Mirrors WorklogManager's own root-resolution pattern (its own "--workbooks-root=" beside
    // DataManager's "--data-root="): a "--drafts-root=" switch, parsed the same way (case-
    // insensitive, surrounding quotes stripped, first match wins), defaulting to an AppData folder
    // that survives Velopack updates. Call Load() once at startup before any other member is used.
    // ###########################################################################################
    public static class DraftManager
    {
        private const string DraftsRootArg = "--drafts-root=";

        private static string _draftsRoot = string.Empty;

        // The folder drafts are actually being read from and written to - the "--drafts-root="
        // value when one was given, otherwise the AppData default. Mirrors DataManager.DataRoot
        // and WorklogManager.WorkbookRoot, and exposed for the same reason a future "Open drafts
        // folder" Configuration button would need it. Empty when Load has not run or failed.
        public static string DraftsRoot => _draftsRoot;

        // ###########################################################################################
        // Resolves the folder to store drafts in - "--drafts-root=" if given, otherwise a "Drafts"
        // folder beside the default "Data" folder in the user's AppData folder - and points the
        // manager at it. Falls back to an unusable (empty) root silently on any failure, the same
        // fail-soft behaviour WorklogManager.Load uses: a draft is a local convenience, and losing
        // it to a folder-creation failure must never stop the rest of the app from starting.
        // ###########################################################################################
        public static void Load(string[]? args = null)
        {
            try
            {
                string root = ResolveDraftsRoot(args);
                LoadFrom(root);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load drafts: [{ex.Message}] - using defaults");
                _draftsRoot = string.Empty;
            }
        }

        // ###########################################################################################
        // Parses "--drafts-root=" out of the command line, following the same conventions
        // DataManager.ResolveDataRoot and WorklogManager.ResolveExplicitWorkbookRoot use. Unlike
        // those two, this always returns a usable path (never null) because there is no separate
        // "explicit vs default" caller-visible distinction needed here yet - Load() is the only
        // caller. internal, so tests can point it at a temp folder without going through Load().
        // ###########################################################################################
        internal static string ResolveDraftsRoot(string[]? args)
        {
            foreach (var arg in args ?? [])
            {
                if (arg.StartsWith(DraftsRootArg, StringComparison.OrdinalIgnoreCase))
                    return arg[DraftsRootArg.Length..].Trim('"', '\'');
            }

            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(appData, AppConfig.AppFolderName, AppConfig.DraftsFolderName);
        }

        // ###########################################################################################
        // Points the manager at an explicit drafts root folder, creating it if missing. Load()
        // resolves the real AppData location and calls this; splitting the two lets the test suite
        // point at a temporary folder instead of the user's real one - never call Load() from a
        // test, the same rule CLAUDE.md states for WorklogManager.Load/DataManager.InitializeAsync.
        //
        // A blank path resets to the unloaded state (no folder is created for an empty string, and
        // GetBoardFolder/LoadDraftFor then fail soft) - used by the test suite to restore a known
        // "not loaded" state between tests, since this is static singleton state that otherwise
        // persists across every test in the collection.
        // ###########################################################################################
        internal static void LoadFrom(string draftsRootPath)
        {
            if (string.IsNullOrWhiteSpace(draftsRootPath))
            {
                _draftsRoot = string.Empty;
                return;
            }

            _draftsRoot = draftsRootPath;
            Directory.CreateDirectory(_draftsRoot);
            Logger.Info($"Drafts root is [{_draftsRoot}]");
        }

        // ###########################################################################################
        // Maps a board's ExcelDataFile ("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx") to its
        // folder under Drafts/ ("<DraftsRoot>/Commodore/C64/250407") - the manufacturer/hardware/
        // board segments only, dropping the file name, so the draft folder mirrors the same
        // Manufacturer/Hardware/Board layout "Data/" itself uses (see CLAUDE.md's "Content"
        // section) and is easy to find by hand. Uses "/" as the separator to split on, matching
        // ExcelDataFile's own convention (the sync manifest's separator - see DataManager), not
        // Path.DirectorySeparatorChar, which would not match on Windows.
        //
        // Returns empty when the root has not loaded or the path has no folder segments to take -
        // callers must treat that as "no draft available" rather than resolving a bogus path.
        // ###########################################################################################
        public static string GetBoardFolder(string excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(_draftsRoot) || string.IsNullOrWhiteSpace(excelDataFile))
            {
                return string.Empty;
            }

            string[] segments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length < 2)
            {
                return string.Empty;
            }

            // All but the last segment (the file name itself) become path segments under the root.
            string[] folderSegments = new string[segments.Length - 1];
            Array.Copy(segments, folderSegments, folderSegments.Length);

            return Path.Combine(_draftsRoot, Path.Combine(folderSegments));
        }

        // ###########################################################################################
        // The draft folder whose files the SCREEN should show for this board - empty when the
        // board on screen is not the draft, which includes "View boards as officially published".
        //
        // Use this, not GetBoardFolder, wherever a board's images, local files or KiCad data are
        // looked up for DISPLAY. GetBoardFolder is where a draft lives and is still right for
        // writing into one; this is whether the board being shown IS that draft. The rule itself
        // is DraftBoardSource.ViewedDraftFolder (pure, unit tested) - this only supplies the roots
        // and the user's toggle.
        // ###########################################################################################
        public static string GetViewedBoardFolder(string excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(_draftsRoot) || string.IsNullOrWhiteSpace(excelDataFile))
            {
                return string.Empty;
            }

            return DraftBoardSource.ViewedDraftFolder(
                DataManager.DataRoot,
                _draftsRoot,
                excelDataFile,
                UserSettings.ViewOfficialPublishedOnly);
        }

        // ###########################################################################################
        // Every board in hardwareBoards that has a non-empty local draft - "list every board with
        // local changes" (NewContributeStrategy.md Phase 2, session 2b, task 7).
        //
        // Takes the candidate list rather than reading DataManager.HardwareBoards itself, the same
        // split DraftDataStore/DraftManager already keep from DataManager: this class resolves
        // WHERE a board's draft folder is, never which boards exist. It also sidesteps a real
        // gap - GetBoardFolder's mapping is Manufacturer/Hardware/Board -> folder, but nothing
        // WALKS the folder tree back the other way (Drafts/ carries no manifest of its own, by
        // design - one draft.json per board is the whole model, echoing WorklogManager's "one
        // file per record" convention). Matching against the known board list, one lookup per
        // entry, avoids inventing a second source of truth for "which boards exist" that Drafts/
        // would have to stay in sync with.
        //
        // A board whose draft folder exists but holds only an empty draft.json (BoardDraft.IsEmpty)
        // is excluded - see DraftDataStore.Load's own doc: a draft that was started (or is left
        // over from a discarded edit) but carries no actual rows is not "a board with local
        // changes" from the user's point of view.
        // ###########################################################################################
        // ###########################################################################################
        // *** THE "IS IT EMPTY" TEST IS GONE, AND THAT IS A REAL BEHAVIOUR CHANGE (Phase 6,
        // 2026-09-23). ***
        //
        // A BoardDraft could exist on disk with every section empty - started and then emptied -
        // and this used to exclude it, because "no local changes" is what the user would have said
        // about it.
        //
        // That test cannot survive the new model. A draft workbook is a full copy of the published
        // board, so it is never empty; asking whether it DIFFERS would mean parsing two workbooks
        // per board just to decide whether to list a row. And the answer would be wrong for the
        // case that matters most: a contributor who has seeded a draft and not yet edited it still
        // HAS a draft, and hiding it would leave them no way to discard it from inside the app.
        //
        // So a draft is listed because it EXISTS - one small marker file read per board. The
        // row's own summary says whether anything in it has actually changed.
        // ###########################################################################################
        public static List<HardwareBoardEntry> EnumerateDraftedBoards(IEnumerable<HardwareBoardEntry> hardwareBoards)
        {
            var result = new List<HardwareBoardEntry>();

            foreach (var entry in hardwareBoards)
            {
                if (DraftBoardSource.HasDraft(_draftsRoot, entry.ExcelDataFile))
                {
                    result.Add(entry);
                }
            }

            // ###########################################################################################
            // *** ONE ROW PER DRAFT FOLDER (code review, 2026-09-27). *** The marker is found by
            // FOLDER, so every listed entry in a draft's folder "has" it - including one under
            // another workbook name that cannot read it: a new board listed in BETA as
            // "... v2.0.0.xlsx" beside its own draft entry ("....xlsx"). That one showed a second row
            // with nothing changed. Where one entry of a folder reads the draft's workbook, the
            // entries that cannot are left out; where none can (a marker whose workbook has gone),
            // all stay, so the draft can still be discarded.
            // ###########################################################################################
            return result
                .GroupBy(entry => BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile), StringComparer.OrdinalIgnoreCase)
                .SelectMany(folder =>
                {
                    List<HardwareBoardEntry> reading = folder
                        .Where(entry => File.Exists(DraftFolderLayout.GetWorkbookPath(_draftsRoot, entry.ExcelDataFile)))
                        .ToList();

                    return reading.Count > 0 ? reading : folder.ToList();
                })
                .OrderBy(entry => result.IndexOf(entry))
                .ToList();
        }

        // ###########################################################################################
        // Every board that exists ONLY as a local draft - one created through "Add a new board"
        // (NewContributeStrategy.md Phase 2, session 2c, task 9), which the main Excel workbook
        // knows nothing about. DataManager merges these into HardwareBoards so a brand-new board
        // appears in the hardware/board drop-downs exactly like a synced one.
        //
        // This is the one thing in the app that walks Drafts/ BACKWARDS - every other path maps a
        // known ExcelDataFile to its folder via GetBoardFolder. It has to: a new board's identity
        // exists nowhere else yet, so there is nothing to look it up by. The walk is what makes a
        // separate registry file unnecessary (see NewBoardRegistration's own header for that
        // decision) - what is on disk IS the list, so a discarded folder cannot leave a stale entry
        // behind.
        //
        // Bounded to exactly three levels (Manufacturer/Hardware/Board), matching the layout
        // GetBoardFolder creates and that "Data/" itself uses, so this is a handful of directory
        // enumerations over at most a few dozen boards rather than an unbounded recursive walk.
        //
        // Fails soft on any IO error, returning whatever it found: the same rule Load() follows,
        // and for the same reason - a drafts folder that cannot be read must never stop the app
        // from starting.
        // ###########################################################################################
        public static List<HardwareBoardEntry> EnumerateDraftOnlyBoards()
        {
            var result = new List<HardwareBoardEntry>();

            if (string.IsNullOrWhiteSpace(_draftsRoot) || !Directory.Exists(_draftsRoot))
            {
                return result;
            }

            try
            {
                foreach (string manufacturerFolder in Directory.EnumerateDirectories(_draftsRoot))
                {
                    foreach (string hardwareFolder in Directory.EnumerateDirectories(manufacturerFolder))
                    {
                        foreach (string boardFolder in Directory.EnumerateDirectories(hardwareFolder))
                        {
                            // Read from the MARKER since Phase 6 - the one small file that makes a
                            // folder a draft, and where a new board's registration now lives.
                            DraftMarker? marker = DraftMarkerStore.Load(
                                Path.Combine(boardFolder, DraftFolderLayout.DraftMarkerFileName));

                            NewBoardRegistration? registration = marker?.NewBoard;

                            if (registration == null || string.IsNullOrWhiteSpace(registration.ExcelDataFile))
                            {
                                // Either an ordinary draft over an already-known board (the common
                                // case - it needs no registration, DataManager already lists it), or
                                // a folder that is not a draft at all.
                                continue;
                            }

                            result.Add(new HardwareBoardEntry
                            {
                                HardwareName = registration.HardwareName,
                                BoardName = registration.BoardName,
                                ExcelDataFile = registration.ExcelDataFile,
                                HardwareNotes = registration.HardwareNotes,
                                IsDraftOnly = true,
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to enumerate draft-only boards under [{_draftsRoot}] - [{ex.Message}]");
            }

            return result;
        }

        // ###########################################################################################
        // Makes a draft of every board folder put into Drafts/ by hand (owner request, 2026-09-27) -
        // see DraftFolderImport for what counts and which kind of draft each becomes.
        //
        // Called from DataManager.LoadMainExcel, the one moment the known boards are all in hand
        // and before anything is listed, so an imported folder shows on the Drafts tab and in the
        // drop-downs on the same launch. A folder dropped in while the application runs is picked
        // up at the next start.
        //
        // Every folder it looked at is logged: one it could NOT import is otherwise exactly the
        // silent "my board is not on the Drafts tab" this exists to end.
        // ###########################################################################################
        public static void ImportHandPlacedFolders(IEnumerable<KnownDraftBoard> knownBoards)
        {
            if (string.IsNullOrWhiteSpace(_draftsRoot))
            {
                return;
            }

            // ###########################################################################################
            // *** NEVER LETS AN EXCEPTION OUT (code review, 2026-09-27). *** DataManager calls this
            // while loading the main workbook, before HardwareBoards is set, inside that load's one
            // try: anything escaping here emptied the whole board list for an optional import. Each
            // folder already contains its own failures (DraftFolderImport.ImportOne); this contains
            // the walk around them.
            // ###########################################################################################
            IReadOnlyList<DraftFolderImportOutcome> outcomes;

            try
            {
                outcomes = DraftFolderImport.ImportUnmarkedFolders(_draftsRoot, knownBoards, DateTimeOffset.UtcNow);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Looking for board folders put into the drafts folder by hand failed - [{ex.Message}]");
                return;
            }

            foreach (DraftFolderImportOutcome outcome in outcomes)
            {
                if (!outcome.Imported)
                {
                    Logger.Warning($"Folder [{outcome.BoardFolder}] in the drafts folder was not taken in as a draft - {outcome.Reason}");
                    continue;
                }

                string kind = outcome.Kind == DraftFolderImportKind.NewBoard
                    ? "a new board"
                    : "a draft of the published board";

                string renamed = outcome.RenamedFrom.Length > 0
                    ? $", its workbook renamed from [{outcome.RenamedFrom}]"
                    : string.Empty;

                Logger.Info($"Took in folder [{outcome.BoardFolder}] as {kind} [{outcome.ExcelDataFile}]{renamed}");
            }
        }

        // ###########################################################################################
        // Permanently discards a board's entire draft - "per-board discard" (session 2b, task 7).
        // Deletes the WHOLE board folder under Drafts/, not just draft.json, so any blob a future
        // authoring session (2c) stores alongside it (a new schematic image, say) is discarded with
        // it too - one folder is the whole draft, the same model WorklogManager.DeleteWorkbook uses
        // for "one folder is the whole workbook".
        //
        // Silently does nothing for a board with no draft folder at all, so a double-discard (or a
        // discard racing a save) is never an error.
        //
        // *** ONE IMPLEMENTATION: DraftWorkbookStore.Discard. *** This used to run its own bare
        // recursive Directory.Delete, a second copy of that method without its IOException handling.
        // On Windows a draft workbook open in Excel is locked, so the delete removed everything up
        // to the locked file and then THREW - out of a fire-and-forget command, unobserved, leaving
        // a half-deleted draft and a Drafts tab that never refreshed. The store's version catches
        // that and says so.
        //
        // Returns whether the folder is GONE: true when it was deleted or never existed, false when
        // the delete failed part-way. A caller must tell the contributor on false - some of the
        // draft may still be on disk, and the usual cause (the workbook open in Excel) is one they
        // can fix and retry.
        // ###########################################################################################
        public static bool DiscardDraft(string excelDataFile)
        {
            string folder = GetBoardFolder(excelDataFile);
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                return true;
            }

            return DraftWorkbookStore.Discard(_draftsRoot, excelDataFile);
        }
    }
}
