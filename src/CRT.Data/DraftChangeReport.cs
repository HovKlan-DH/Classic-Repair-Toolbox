using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE "WHAT CHANGED" WINDOW SHOWS for one drafted board
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23; the base, 2026-10-13).
    //
    // *** A DRAFT WITH A BASE (code review, 2026-10-13). *** A draft seeded since CRT keeps a copy of
    // what it was copied from (DraftBase) can say three things apart:
    //
    //   - Rows: the contributor's OWN changes - the base against the draft;
    //   - UpdateRows: what the published update changed since - the base against the published
    //     board as it is now;
    //   - Conflicts: where both changed the same row - what bringing the draft up to date would
    //     report (DraftRebase, run here without writing anything).
    //
    // and the window can offer to bring the draft up to date (CanBringUpToDate).
    //
    // *** A DRAFT WITHOUT ONE (seeded before 2026-10-13, or brought in from a folder) can only be
    // compared with the published board as it is NOW, so Rows mixes the contributor's changes with
    // the update's, reversed. That is said in the window rather than hidden: sending such a draft
    // would show the maintainer the update's changes as the contributor's.
    // ###########################################################################################
    public sealed class DraftChangeReport
    {
        // Whether the published board has moved since this draft was taken (DraftDriftCheck).
        public DraftDriftState State { get; init; }

        // The published revision this draft was seeded from (or last brought up to date to).
        public string BaseRevision { get; init; } = string.Empty;

        // The published revision as it reads NOW.
        public string OfficialRevision { get; init; } = string.Empty;

        // The contributor's changes: against the base when ComparedWithBase, else every row that
        // differs between the draft and the published board as it is now.
        public IReadOnlyList<BoardRowChange> Rows { get; init; } = [];

        // True when Rows were measured against the base, so they are the contributor's alone.
        public bool ComparedWithBase { get; init; }

        // What the published update changed since the draft was copied. Empty without a base.
        public IReadOnlyList<BoardRowChange> UpdateRows { get; init; } = [];

        // Where the update and the contributor changed the same row. Empty without a base.
        public IReadOnlyList<DraftRebaseConflict> Conflicts { get; init; } = [];

        // Whether "Bring my draft up to date" can run: a base, and a published board that moved.
        public bool CanBringUpToDate { get; init; }

        // ###########################################################################################
        // The report for one draft, read from disk: its workbook, the published board and the base.
        // Reads two workbooks, so callers run it under a wait (BusyOverlay.RunLocalAsync).
        // ###########################################################################################
        public static DraftChangeReport For(DraftStatus status, string draftsRoot, string publishedWorkbookPath)
        {
            System.ArgumentNullException.ThrowIfNull(status);

            DraftDriftState state = DraftDriftCheck.Resolve(status, publishedWorkbookPath);

            BoardData? draft = string.IsNullOrWhiteSpace(status.WorkbookPath)
                ? null
                : BoardDataReader.ReadWorkbookUncached(status.WorkbookPath);

            BoardData? published = File.Exists(publishedWorkbookPath)
                ? BoardDataReader.ReadWorkbookUncached(publishedWorkbookPath)
                : null;

            DraftBase? based = status.HasBase ? DraftBaseStore.Load(draftsRoot, status.BoardKey) : null;

            if (draft is null)
            {
                return new DraftChangeReport
                {
                    State = state,
                    BaseRevision = status.BaseRevision,
                    OfficialRevision = published?.RevisionDate ?? string.Empty,
                };
            }

            if (based is null)
            {
                return new DraftChangeReport
                {
                    State = state,
                    BaseRevision = status.BaseRevision,
                    OfficialRevision = published?.RevisionDate ?? string.Empty,
                    Rows = BoardDataDiffer.Compare(published, draft),
                };
            }

            bool moved = state is DraftDriftState.OfficialIsNewer or DraftDriftState.Changed;
            bool canMerge = moved && published is not null;

            IReadOnlyList<DraftRebaseConflict> conflicts = canMerge
                ? DraftRebase.Merge(
                    based.Board,
                    draft,
                    published!,
                    based.Calibrations,
                    DraftBoardSource.CollectCalibrations(status.WorkbookPath, draft),
                    DraftBoardSource.CollectCalibrations(publishedWorkbookPath, published)).Conflicts
                : [];

            return new DraftChangeReport
            {
                State = state,
                BaseRevision = status.BaseRevision,
                OfficialRevision = published?.RevisionDate ?? string.Empty,
                Rows = BoardDataDiffer.Compare(based.Board, draft),
                ComparedWithBase = true,
                UpdateRows = canMerge ? BoardDataDiffer.Compare(based.Board, published) : [],
                Conflicts = conflicts,
                CanBringUpToDate = canMerge,
            };
        }
    }
}
