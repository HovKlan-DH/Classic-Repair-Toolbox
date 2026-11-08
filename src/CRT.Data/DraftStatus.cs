using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // EVERYTHING A SURFACE NEEDS TO KNOW ABOUT ONE DRAFT, without loading a BoardDraft
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** THIS IS WHAT THE DRAFTS TAB READS NOW. *** It used to take a BoardDraft and count its
    // delta rows for the "N rows changed" chip, read NewBoard off it for the registration, and
    // read BaseRevision for the drift check. With the draft stored as a board workbook there is no
    // delta list to count - so the count is DERIVED, by comparing the draft workbook against the
    // published one (BoardDataDiffer), and the other two come off the marker.
    //
    // *** THE COUNT COSTS A BOARD PARSE, so it is opt-in. *** Resolve() answers the cheap
    // questions from the marker alone, which is all the tab needs to LIST a draft; the comparison
    // runs only when a caller asks for ChangeCount. Listing twenty drafted boards must not mean
    // parsing forty workbooks.
    // ###########################################################################################
    public sealed class DraftStatus
    {
        public string BoardKey { get; init; } = string.Empty;

        // The published RevisionDate this draft was taken from. Empty for a draft-only board,
        // which has no published counterpart to be based on.
        public string BaseRevision { get; init; } = string.Empty;

        // The published revision the contributor said "I have looked at this" to, for a draft with
        // no base (DraftMarker.DriftSeenRevision). Quiets the drift warning; BaseRevision stays true.
        public string DriftSeenRevision { get; init; } = string.Empty;

        // The published content it was said to (DraftMarker.DriftSeenContentHash).
        public string DriftSeenContentHash { get; init; } = string.Empty;

        // What the published files held when the draft was copied or last brought up to date
        // (DraftMarker.BaseContentHash). Empty for a draft from before it was kept.
        public string BaseContentHash { get; init; } = string.Empty;

        // Whether the draft keeps a base, so it can be brought up to date (DraftRebaser).
        public bool HasBase { get; init; }

        // Set only on a board that exists purely as a draft - "Add a new board".
        public NewBoardRegistration? NewBoard { get; init; }

        public bool IsNewBoard => this.NewBoard != null;

        public string WorkbookPath { get; init; } = string.Empty;

        // What the draft is compared against. Empty for a new board from Resolve; the LISTED
        // board's workbook from ResolveForBoard, which retirement uses.
        public string PublishedWorkbookPath { get; init; } = string.Empty;

        // When the draft was created, off its marker - null for a marker that does not say (one
        // written before markers carried it). The Drafts row's badge describes only submissions
        // sent from this draft, not from an earlier one (SubmissionReceiptPresenter.LatestForBoard).
        public DateTimeOffset? CreatedUtc { get; init; }

        public static DateTimeOffset? ParseCreated(string? createdUtc) =>
            DateTimeOffset.TryParse(
                createdUtc,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind,
                out DateTimeOffset parsed)
                ? parsed
                : null;
    }

    // ###########################################################################################
    // Resolves that status, and the change count when it is actually wanted.
    // ###########################################################################################
    public static partial class DraftStatusReader
    {
        // ###########################################################################################
        // The cheap answer: is there a draft, and what does its marker say?
        //
        // Null when there is none. Reads ONE small JSON file and touches no workbook, so a surface
        // listing every drafted board pays almost nothing.
        // ###########################################################################################
        //
        // *** THE KEY IS THE BOARD'S WORKBOOK PATH ("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
        // NOT A BOARD ID. *** Handed "Commodore/C64/250407" it strips the last segment as a file
        // name and looks in "Drafts/Commodore/C64" - finding nothing. A submission receipt names its
        // board by id, so go through ResolveForBoard for one.
        //
        // listedAsPublished: HardwareBoardEntry.IsPublished when the caller knows it - it decides what
        // the draft's changes are counted against (DraftBoardSource.ComparisonBaselineOf).
        public static DraftStatus? Resolve(string dataRoot, string draftsRoot, string excelDataFile, bool? listedAsPublished = null)
        {
            DraftMarker? marker = DraftMarkerStore.Load(
                DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile));

            if (marker is null)
            {
                return null;
            }

            return new DraftStatus
            {
                BoardKey = excelDataFile,
                BaseRevision = marker.BaseRevision,
                DriftSeenRevision = marker.DriftSeenRevision,
                DriftSeenContentHash = marker.DriftSeenContentHash,
                BaseContentHash = marker.BaseContentHash,
                HasBase = !marker.IsNewBoard && DraftBaseStore.Exists(draftsRoot, excelDataFile),
                NewBoard = marker.NewBoard,
                CreatedUtc = DraftStatus.ParseCreated(marker.CreatedUtc),
                WorkbookPath = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile),

                // Empty for a new board, so every one of its rows counts - see
                // DraftBoardSource.ComparisonBaselineOf for the legacy copy that made this matter.
                PublishedWorkbookPath = DraftBoardSource.ComparisonBaselineOf(dataRoot, excelDataFile, marker, listedAsPublished),
            };
        }

        // ###########################################################################################
        // The draft a SUBMISSION RECEIPT is about. A receipt names its board by id
        // ("Commodore/C64/250407"), built from the board's workbook path by
        // BoardDescriptorRules.BoardIdFromExcelDataFile when it was submitted; this finds that
        // workbook again among the boards the app knows, the same way round, and resolves it.
        //
        // It used to be Resolve called with the id itself, which looks one folder too high and finds
        // nothing - so a draft whose work had been published was never retired (owner report,
        // 2026-09-25: published to BETA, and the draft stayed in the list).
        //
        // Null when no known board has that id - the draft is then left alone, which is the safe way.
        //
        // *** THE DRAFT AND THE PUBLISHED BOARD CAN HAVE DIFFERENT WORKBOOK NAMES (2026-09-25). ***
        // A NEW board's draft is keyed by the name it was created with ("Data HW Board.xlsx"),
        // and its publish writes the tree's generation ("Data HW Board v2.0.0.xlsx"). So the draft
        // is found by its FOLDER (the marker) and keeps its own key (the marker's BoardKey), and the
        // published side is the board the app lists. For an existing board the two are the same.
        // Where several listed boards carry the id, one whose published copy is on disk wins.
        // ###########################################################################################
        public static DraftStatus? ResolveForBoard(
            string dataRoot,
            string draftsRoot,
            string boardId,
            IEnumerable<string>? excelDataFiles)
        {
            if (string.IsNullOrWhiteSpace(boardId))
            {
                return null;
            }

            string id = boardId.Trim();

            List<string> matching = (excelDataFiles ?? [])
                .Where(file => string.Equals(
                    BoardDescriptorRules.BoardIdFromExcelDataFile(file),
                    id,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();

            string? published = matching.FirstOrDefault(file => File.Exists(DraftBoardSource.PublishedPathOf(dataRoot, file)))
                ?? matching.FirstOrDefault();

            if (published is null)
            {
                return null;
            }

            DraftMarker? marker = DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(draftsRoot, published));

            if (marker is null)
            {
                return null;
            }

            // The draft's own key, when the marker names this same board; the listed board's
            // otherwise (a marker written before the key was recorded).
            string draftKey = !string.IsNullOrWhiteSpace(marker.BoardKey) &&
                string.Equals(BoardDescriptorRules.BoardIdFromExcelDataFile(marker.BoardKey), id, StringComparison.OrdinalIgnoreCase)
                    ? marker.BoardKey.Trim()
                    : published;

            return new DraftStatus
            {
                BoardKey = draftKey,
                BaseRevision = marker.BaseRevision,
                NewBoard = marker.NewBoard,
                CreatedUtc = DraftStatus.ParseCreated(marker.CreatedUtc),
                WorkbookPath = DraftFolderLayout.GetWorkbookPath(draftsRoot, draftKey),
                PublishedWorkbookPath = DraftBoardSource.PublishedPathOf(dataRoot, published),
            };
        }

        // ###########################################################################################
        // How many rows differ between the draft and the published board - the "N rows changed"
        // number.
        //
        // *** THIS PARSES TWO WORKBOOKS, so call it for the board being SHOWN, not for every row
        // in a list. *** Uncached on both sides deliberately: the whole point of the new layout is
        // that the contributor may have edited the workbook in Excel, and a cached count would go
        // on reporting the state from before they did.
        //
        // A draft-only board compares against nothing, so every row in it counts as an addition -
        // which is the literal truth, and is why the caller says that case differently ("New
        // board") rather than as "N rows changed".
        // ###########################################################################################
        public static int CountChanges(DraftStatus? status)
        {
            if (status is null || string.IsNullOrWhiteSpace(status.WorkbookPath))
            {
                return 0;
            }

            // Read through BoardReadsByStamp (code review, 2026-10-13): fresh whenever a workbook or
            // its sidecar moved - an Excel save included - and ONE parse shared with the problem
            // count and the fingerprint, which each parsed the same changed draft again.
            BoardData? draft = BoardReadsByStamp.Read(status.WorkbookPath);
            if (draft is null)
            {
                return 0;
            }

            BoardData? published = File.Exists(status.PublishedWorkbookPath)
                ? BoardReadsByStamp.Read(status.PublishedWorkbookPath)
                : null;

            return BoardDataDiffer.CountChanges(published, draft);
        }

        // ###########################################################################################
        // The same count, remembered per draft until one of the four files it is derived from
        // changes - the two workbooks and their two sidecars (the sidecar carries the highlights,
        // which the count includes).
        //
        // *** WHY THE DRAFTS TAB USES THIS, NOT CountChanges. *** RefreshDrafts runs after every
        // component save, label-editor save, table save, calibration save and board reload, and
        // it counted EVERY drafted board each time - two full workbook parses per board,
        // synchronously on the UI thread. Three drafted boards of a few thousand rows made every
        // save freeze the window for seconds, re-counting boards nobody had touched.
        //
        // *** STILL FRESH AFTER AN EXCEL EDIT, which is why CountChanges was uncached. *** The
        // cache key is each file's length and last-write time, so a workbook saved in Excel -
        // with the application open or closed - is a new stamp and is counted again. Only a
        // board whose four files are all exactly as they were is answered from memory, and for
        // that board the answer cannot have changed.
        // ###########################################################################################
        public static int CountChangesCached(DraftStatus? status)
        {
            if (status is null || string.IsNullOrWhiteSpace(status.WorkbookPath))
            {
                return 0;
            }

            string stamp = DraftStatusReader.StampOf(status);

            if (DraftStatusReader.CountCache.TryGetValue(status.WorkbookPath, out CachedCount cached)
                && string.Equals(cached.Stamp, stamp, StringComparison.Ordinal))
            {
                return cached.Count;
            }

            int count = DraftStatusReader.CountChanges(status);

            DraftStatusReader.CountCache[status.WorkbookPath] = new CachedCount(stamp, count);

            return count;
        }

        private readonly record struct CachedCount(string Stamp, int Count);

        // ###########################################################################################
        // *** THE DRAFT'S ERRORS AND WARNINGS, FOR ITS ROW ON THE DRAFTS TAB (owner report,
        // 2026-10-02: "It must check for errors when creating the draft, and if the board changes
        // "offline", outside of app"). *** The checks were only in the table, so a draft with
        // errors looked fine until its table was opened. This is the table's own count -
        // BoardDataChecks in its Everything scope, over the draft workbook and its sidecar, its files
        // looked for where a submit looks (DiskFileLookup) - so the row and the table agree.
        //
        // Remembered per draft like CountChangesCached, keyed on the workbook's and the sidecar's
        // length and last-write time: an edit in Excel is counted afresh the next time the tab
        // refreshes, an untouched draft costs nothing.
        //
        // *** A CITED FILE THAT WAS NOT THERE IS LOOKED FOR AGAIN (code review, 2026-10-04). *** The
        // stamp cannot see a picture arriving on its own - the background image sync finishing, or a
        // file dropped into the draft folder by hand - so a "file missing" error stayed on the row
        // for the session while the table and Submit no longer found it. Every path the check did
        // not find as cited is remembered with the count and asked about again on a cache hit:
        // usually none, so it costs nothing. A file that VANISHES is still only seen when the
        // workbook changes or the table opens - re-asking every found file (thousands on a large
        // board) on every refresh is the cost the cache exists to avoid.
        // ###########################################################################################
        public static BoardProblemCounts CountProblemsCached(DraftStatus? status, string dataRoot, string draftBoardFolder) =>
            DraftStatusReader.ProblemsCached(status, dataRoot, draftBoardFolder).Counts;

        // ###########################################################################################
        // The same read, with the problems themselves and the board they were found on - what says which
        // entry of the Contribute tab each is about (owner request, 2026-11-04). The very same object
        // while nothing moved, so a caller can tell that nothing did.
        // ###########################################################################################
        public static BoardProblemsRead ProblemsCached(DraftStatus? status, string dataRoot, string draftBoardFolder)
        {
            if (status is null || string.IsNullOrWhiteSpace(status.WorkbookPath))
            {
                return BoardProblemsRead.None;
            }

            return DraftStatusReader.ProblemsCachedOf(status.WorkbookPath, dataRoot, draftBoardFolder, boardIdentity: null);
        }

        // ###########################################################################################
        // *** A PUBLISHED BOARD'S OWN ERRORS AND WARNINGS, WHERE NO DRAFT HAS ANY TO SAY (owner decision, 2026-11-07: the
        // launch check of every board "needs to run for everyone and its findings should show just like any other error or
        // warning in the "Contribute" tab"). *** The very checks a draft gets (BoardCheckScope.Everything), on the
        // published board file as downloaded, its files looked for as the launch check looked for them - the downloaded data
        // by the board's identity - and kept the same way, until the file changes. Worked out for the board SHOWN, never every
        // board at launch: reading each one took the launch check some ten seconds.
        // ###########################################################################################
        public static BoardProblemsRead PublishedProblemsCached(string publishedBoardFile, string dataRoot, string draftBoardFolder, string boardIdentity)
        {
            if (string.IsNullOrWhiteSpace(publishedBoardFile) || !File.Exists(publishedBoardFile))
            {
                return BoardProblemsRead.None;
            }

            // One published board's read kept at a time - each holds its whole board, and a draft's own stays as it was.
            lock (DraftStatusReader.PublishedProblemsLock)
            {
                if (DraftStatusReader.LastPublishedProblemsFile is { } earlier &&
                    !string.Equals(earlier, publishedBoardFile, StringComparison.OrdinalIgnoreCase))
                {
                    DraftStatusReader.ProblemCache.TryRemove(earlier, out _);
                }

                DraftStatusReader.LastPublishedProblemsFile = publishedBoardFile;
            }

            return DraftStatusReader.ProblemsCachedOf(publishedBoardFile, dataRoot, draftBoardFolder, boardIdentity, published: true);
        }

        private static readonly object PublishedProblemsLock = new();

        private static string? LastPublishedProblemsFile;

        // Run as a published board's read is about to be kept - for a test to hold it there while another board is asked.
        internal static Action<string>? BeforePublishedReadIsKeptForTests { get; set; }

        private static BoardProblemsRead ProblemsCachedOf(string boardFile, string dataRoot, string draftBoardFolder, string? boardIdentity, bool published = false)
        {
            string stamp = string.Join(
                "|",
                DraftStatusReader.FileStamp(boardFile),
                DraftStatusReader.FileStamp(BoardComponentHighlightStorage.GetJsonPath(boardFile)),
                dataRoot,
                draftBoardFolder,
                boardIdentity ?? string.Empty);

            if (DraftStatusReader.ProblemCache.TryGetValue(boardFile, out CachedProblems cached)
                && string.Equals(cached.Stamp, stamp, StringComparison.Ordinal)
                && DraftStatusReader.StillNotFound(cached.NotFound, dataRoot, draftBoardFolder, boardIdentity))
            {
                return cached.Read;
            }

            BoardData? board = BoardReadsByStamp.Read(boardFile);
            var lookup = new RememberingLookup(new DiskFileLookup(dataRoot, draftBoardFolder, boardIdentity));

            IReadOnlyList<BoardDataProblem> problems = board is null
                ? []
                : BoardDataChecks.Check(BoardCheckRows.From(board), lookup, BoardCheckScope.Everything);

            var read = new BoardProblemsRead(BoardProblemCounts.Of(problems), problems, board);
            var kept = new CachedProblems(stamp, read, lookup.NotFound);

            if (!published)
            {
                DraftStatusReader.ProblemCache[boardFile] = kept;
                return read;
            }

            DraftStatusReader.BeforePublishedReadIsKeptForTests?.Invoke(boardFile);

            // ###########################################################################################
            // *** ONLY THE LATEST PUBLISHED BOARD'S READ IS KEPT - decided under the very lock that lets the earlier one go
            // (code review, 2026-11-08). *** A check still running as another board's began was stored after its eviction,
            // and stayed for the run - one more whole board for every such race. Its read is handed back all the same.
            // ###########################################################################################
            lock (DraftStatusReader.PublishedProblemsLock)
            {
                if (string.Equals(DraftStatusReader.LastPublishedProblemsFile, boardFile, StringComparison.OrdinalIgnoreCase))
                {
                    DraftStatusReader.ProblemCache[boardFile] = kept;
                }
            }

            return read;
        }

        // ###########################################################################################
        // *** A DRAFT THAT IS GONE IS LET GO (code review, 2026-11-04), as BoardReadsByStamp lets go. *** What
        // is kept per draft now holds its board, not a count, and a discarded, retired or renamed draft's
        // stayed for the rest of the run: tens of megabytes for a large board. Drops the entries whose file no
        // longer exists, one existence check each. Once per refresh of the drafts' list, NOT per draft's read
        // (code review, 2026-11-04: 20 drafts made that 400 checks on the UI thread, at every save).
        // ###########################################################################################
        public static void ForgetProblemsOfVanishedDrafts()
        {
            foreach (string path in DraftStatusReader.ProblemCache.Keys)
            {
                if (!File.Exists(path))
                {
                    DraftStatusReader.ProblemCache.TryRemove(path, out _);
                }
            }
        }

        // Whether what the checks found is kept for `workbookPath` - for tests.
        internal static bool HoldsProblemsForTests(string workbookPath) => DraftStatusReader.ProblemCache.ContainsKey(workbookPath);

        // The problems themselves, by the very checks and lookup the row counts them with - what "Board
        // status" says on each panel (owner request, 2026-11-01). On the draft as held in memory.
        public static IReadOnlyList<BoardDataProblem> ProblemsOf(BoardData draft, string dataRoot, string draftBoardFolder) =>
            BoardDataChecks.Check(BoardCheckRows.From(draft), new DiskFileLookup(dataRoot, draftBoardFolder), BoardCheckScope.Everything);

        // Whether every path remembered as not found (or found under other capitals) still answers
        // exactly so - the cached count is good only while it does.
        private static bool StillNotFound(
            IReadOnlyDictionary<string, BoardFileLookupResult> notFound,
            string dataRoot,
            string draftBoardFolder,
            string? boardIdentity)
        {
            if (notFound.Count == 0)
            {
                return true;
            }

            var lookup = new DiskFileLookup(dataRoot, draftBoardFolder, boardIdentity);

            return notFound.All(item => lookup.Check(item.Key) == item.Value);
        }

        private readonly record struct CachedProblems(
            string Stamp,
            BoardProblemsRead Read,
            IReadOnlyDictionary<string, BoardFileLookupResult> NotFound);

        // The checks' file lookup, remembering every answer that was not a plain "found".
        private sealed class RememberingLookup(IBoardFileLookup inner) : IBoardFileLookup
        {
            private readonly Dictionary<string, BoardFileLookupResult> thisNotFound = new(StringComparer.Ordinal);

            public IReadOnlyDictionary<string, BoardFileLookupResult> NotFound => this.thisNotFound;

            public string Where => inner.Where;

            public BoardFileLookupResult Check(string path)
            {
                BoardFileLookupResult result = inner.Check(path);

                if (result.State != BoardFileState.Found)
                {
                    this.thisNotFound[path] = result;
                }

                return result;
            }
        }

        private static readonly ConcurrentDictionary<string, CachedProblems> ProblemCache =
            new(StringComparer.OrdinalIgnoreCase);

        // Keyed by the draft workbook's path; case-insensitive to match how BoardDataReader keys
        // the same files. Thread-safe because RefreshDrafts is not the only possible caller.
        private static readonly ConcurrentDictionary<string, CachedCount> CountCache =
            new(StringComparer.OrdinalIgnoreCase);

        // Every input the count is derived from, as length + last-write time. A missing file
        // stamps as "-" so a published copy arriving (or vanishing) is a change too.
        private static string StampOf(DraftStatus status)
        {
            return string.Join(
                "|",
                DraftStatusReader.FileStamp(status.WorkbookPath),
                DraftStatusReader.FileStamp(BoardComponentHighlightStorage.GetJsonPath(status.WorkbookPath)),
                DraftStatusReader.FileStamp(status.PublishedWorkbookPath),
                DraftStatusReader.FileStamp(BoardComponentHighlightStorage.GetJsonPath(status.PublishedWorkbookPath)));
        }

        // The board reader's stamp, so every cache over a board's files agrees on "unchanged".
        private static string FileStamp(string? path) => BoardDataReader.FileStamp(path);

        // ###########################################################################################
        // The full change list, for the "what changed" window.
        //
        // Same cost and the same freshness rule as CountChanges; kept separate so a caller that
        // only wants the number does not build display labels for every row.
        // ###########################################################################################
        public static IReadOnlyList<BoardRowChange> DescribeChanges(DraftStatus? status)
        {
            if (status is null || string.IsNullOrWhiteSpace(status.WorkbookPath))
            {
                return [];
            }

            BoardData? draft = BoardDataReader.ReadWorkbookUncached(status.WorkbookPath);
            if (draft is null)
            {
                return [];
            }

            BoardData? published = File.Exists(status.PublishedWorkbookPath)
                ? BoardDataReader.ReadWorkbookUncached(status.PublishedWorkbookPath)
                : null;

            return BoardDataDiffer.Compare(published, draft);
        }
    }
}
