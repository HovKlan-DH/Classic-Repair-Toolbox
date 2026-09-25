using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // EVERYTHING A SURFACE NEEDS TO KNOW ABOUT ONE DRAFT, without loading a BoardDraft
    // (NewContributeStrategy.md Phase 6 - maintainer request, 2026-09-23).
    //
    // *** THIS IS WHAT THE DRAFTS TAB READS NOW. *** It used to take a BoardDraft and count its
    // delta rows for the "N rows changed" chip, read NewSystem off it for the registration, and
    // read BaseRevision for the drift check. With the draft stored as a board workbook there is no
    // delta list to count - so the count is DERIVED, by comparing the draft workbook against the
    // published one (BoardDataDiffer), and the other two come off the marker.
    //
    // *** THE COUNT COSTS A BOARD PARSE, so it is opt-in. *** Resolve() answers the cheap
    // questions from the marker alone, which is all the tab needs to LIST a draft; the comparison
    // runs only when a caller asks for ChangeCount. Listing twenty drafted systems must not mean
    // parsing forty workbooks.
    // ###########################################################################################
    public sealed class DraftStatus
    {
        public string SystemKey { get; init; } = string.Empty;

        // The published RevisionDate this draft was taken from. Empty for a draft-only system,
        // which has no published counterpart to be based on.
        public string BaseRevision { get; init; } = string.Empty;

        // Set only on a system that exists purely as a draft - "Add a new system".
        public NewSystemRegistration? NewSystem { get; init; }

        public bool IsNewSystem => this.NewSystem != null;

        public string WorkbookPath { get; init; } = string.Empty;

        public string PublishedWorkbookPath { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // Resolves that status, and the change count when it is actually wanted.
    // ###########################################################################################
    public static class DraftStatusReader
    {
        // ###########################################################################################
        // The cheap answer: is there a draft, and what does its marker say?
        //
        // Null when there is none. Reads ONE small JSON file and touches no workbook, so a surface
        // listing every drafted system pays almost nothing.
        // ###########################################################################################
        public static DraftStatus? Resolve(string dataRoot, string draftsRoot, string excelDataFile)
        {
            DraftMarker? marker = DraftMarkerStore.Load(
                DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile));

            if (marker is null)
            {
                return null;
            }

            return new DraftStatus
            {
                SystemKey = excelDataFile,
                BaseRevision = marker.BaseRevision,
                NewSystem = marker.NewSystem,
                WorkbookPath = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile),
                PublishedWorkbookPath = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile),
            };
        }

        // ###########################################################################################
        // How many rows differ between the draft and the published board - the "N rows changed"
        // number.
        //
        // *** THIS PARSES TWO WORKBOOKS, so call it for the system being SHOWN, not for every row
        // in a list. *** Uncached on both sides deliberately: the whole point of the new layout is
        // that the contributor may have edited the workbook in Excel, and a cached count would go
        // on reporting the state from before they did.
        //
        // A draft-only system compares against nothing, so every row in it counts as an addition -
        // which is the literal truth, and is why the caller words that case differently ("New
        // system, N rows so far") rather than as "N rows changed".
        // ###########################################################################################
        public static int CountChanges(DraftStatus? status)
        {
            if (status is null || string.IsNullOrWhiteSpace(status.WorkbookPath))
            {
                return 0;
            }

            BoardData? draft = BoardDataReader.ReadWorkbookUncached(status.WorkbookPath);
            if (draft is null)
            {
                return 0;
            }

            BoardData? published = File.Exists(status.PublishedWorkbookPath)
                ? BoardDataReader.ReadWorkbookUncached(status.PublishedWorkbookPath)
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
        // it counted EVERY drafted system each time - two full workbook parses per system,
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

        private static string FileStamp(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return "-";
            }

            try
            {
                var info = new FileInfo(path);

                return info.Exists
                    ? $"{info.Length}:{info.LastWriteTimeUtc.Ticks}"
                    : "-";
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
            {
                // Unreadable metadata is never a cache hit: a unique stamp forces a recount.
                return Guid.NewGuid().ToString("N");
            }
        }

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
