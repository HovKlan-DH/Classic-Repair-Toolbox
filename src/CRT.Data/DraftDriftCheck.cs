using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HAS THE PUBLISHED BOARD MOVED SINCE THIS DRAFT WAS COPIED? (code review, 2026-10-13) - the one
    // rule the Drafts tab's row, the main window's notice and the drift window all ask.
    //
    // A draft that keeps the content hash of what it was copied from is compared by CONTENT: a
    // revision is only a date, and two publishes on one day read as no change by date. When the
    // content moved, the revisions still say whether the published one is newer.
    //
    // A draft from before content hashes were kept is compared by revision, as it always was - with
    // the revision the contributor said "I have looked at this" to standing in for its base, so the
    // warning stays quiet until the published board moves again (DraftMarker.DriftSeenRevision).
    // A draft compared by content is quiet the same way while the published files still hold what
    // "I have looked at this" was said to (DraftMarker.DriftSeenContentHash, code review 2026-10-13).
    // ###########################################################################################
    public static class DraftDriftCheck
    {
        public static DraftDriftState Resolve(DraftStatus? status, string publishedWorkbookPath)
        {
            if (status is null || status.IsNewBoard || string.IsNullOrWhiteSpace(publishedWorkbookPath))
                return DraftDriftState.Unknown;

            string officialRevision =
                BoardDataReader.TryGetCachedRevisionDate(publishedWorkbookPath) ??
                BoardDataReader.ReadRevisionDateOnly(publishedWorkbookPath);

            return DraftDriftCheck.Resolve(status, officialRevision, DraftBaseStore.PublishedContentHash(publishedWorkbookPath));
        }

        // The rule itself, given the published revision and content hash - pure, for the tests.
        public static DraftDriftState Resolve(DraftStatus status, string officialRevision, string officialContentHash)
        {
            ArgumentNullException.ThrowIfNull(status);

            if (status.IsNewBoard)
                return DraftDriftState.Unknown;

            if (status.BaseContentHash.Length > 0 && !string.IsNullOrEmpty(officialContentHash))
            {
                if (string.Equals(status.BaseContentHash, officialContentHash, StringComparison.Ordinal))
                    return DraftDriftState.InSync;

                // "I have looked at this", said to exactly this published content.
                if (string.Equals(status.DriftSeenContentHash, officialContentHash, StringComparison.Ordinal))
                    return DraftDriftState.InSync;

                DraftDriftState byRevision = DraftRevisionComparer.Compare(status.BaseRevision, officialRevision);

                // The same date with other content is a second publish that day.
                return byRevision is DraftDriftState.InSync or DraftDriftState.Unknown ? DraftDriftState.Changed : byRevision;
            }

            string seen = status.DriftSeenRevision.Length > 0 ? status.DriftSeenRevision : status.BaseRevision;

            return DraftRevisionComparer.Compare(seen, officialRevision);
        }
    }
}
