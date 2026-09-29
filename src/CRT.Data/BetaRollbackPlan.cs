using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ROLLING BETA BACK TO WHAT PRODUCTION HOLDS (owner decision, 2026-09-27).
    //
    // The maintainer asked for a way to "push back to queue" from the production window: a board
    // already merged into BETA turns out to need work, and the contributor has to be told. This is
    // the half that decides WHAT would be written; the server's BetaRollbackWriter performs it.
    //
    // *** WHY THIS IS POSSIBLE AT ALL - the two facts it rests on. ***
    //
    //   1. A MERGED SUBMISSION KEEPS ITS BYTES. SubmissionCollectionStates.Live includes Merged, so
    //      DeleteRetiredPayloadsAsync never touches it: the rows and blobs of everything in BETA are
    //      still on the server. (This was believed otherwise when the question was first asked, and
    //      the owner corrected it - "everything stays shadowed ... data is still there".)
    //   2. PRODUCTION HOLDS A COMPLETE BOARD, not a partial set. ProductionPromotionPlan copies the
    //      whole system folder plus the shared files it cites, so the last promoted state is a
    //      board that loads on its own.
    //
    // *** PER SYSTEM, NEVER PER SUBMISSION - the one limit that survives, and it is structural. ***
    // ProductionPromotionPlan's header puts it the other way round: "two submissions merged into one
    // board cannot be promoted separately, because the board's workbook already holds both". So a
    // rollback reverts ALL of them, and the plan NAMES them (owner decision: all back to `pending`).
    //
    // *** A SYSTEM NEVER PROMOTED IS A DIFFERENT OPERATION. *** Production has nothing of its own
    // folder to restore FROM, so its folder leaves the BETA tree entirely (owner decision). Kind says
    // which of the two a plan is, so a caller cannot perform one believing it is the other.
    //
    // *** "NEVER PROMOTED" IS THE RECORD'S WORD, NOT AN EMPTY FOLDER'S (code review, 2026-09-27). ***
    // An empty production listing also comes from a folder that is missing, renamed by hand, on a
    // mount that is down, or refused by SubmissionPathRules. Read as "never promoted", that wiped a
    // board that IS in production out of BETA, took its row out of BETA's lists and told the
    // maintainer "Nothing of this system is in production". So the caller says whether the system
    // was ever promoted (promotedBefore), and a promoted system with nothing readable in production
    // is ProductionUnreadable: a plan that does nothing and that the server refuses.
    //
    // *** PATHS ARE COMPARED ORDINALLY (code review, 2026-09-27). *** The server's trees are on a
    // case-sensitive filesystem and SubmissionPathRules resolves ordinally on purpose, so BETA's
    // "Images/Sheet.png" and production's "Images/sheet.png" are TWO files there. Pairing them
    // case-insensitively left BETA holding both after a rollback - the submission's own variant
    // neither restored over nor removed, and still served. A rollback makes the system's folder
    // IDENTICAL to production's, so a BETA-only spelling goes like any other BETA-only file.
    //
    // *** SHARED FILES - only the ones THESE submissions changed (code review, 2026-09-27). *** The
    // system's own folder is replaced wholesale, but a shared file reaches every board citing it, so
    // it is touched only when the rollback can prove the change is its own: BETA still holds exactly
    // the bytes a returning submission carried at that path. Then it is put back to production's
    // bytes. One production never had STAYS (owner decision, 2026-09-27 - AutomaticRemovalScope):
    // nothing outside the system's own folder is removed automatically, so it becomes an unused
    // file for Admin > Unused files. A shared file BETA holds DIFFERENT bytes for was written again
    // by somebody later; it is theirs and is left alone.
    //
    // Leaving the submission's shared change in BETA was the first version's behaviour, and it was a
    // leak: the next promotion of ANY board citing the file carried those un-reviewed bytes to
    // production while the submission that introduced them sat in the queue.
    //
    // Pure, and in CRT.Data because the Maintainer tab shows the same records before the
    // button is pressed that the server acts on - the SubmittedFileFact rule.
    // ###########################################################################################
    public static class BetaRollbackPlan
    {
        // ###########################################################################################
        // Builds the plan.
        //
        // betaOwnFiles       - every file under the system's folder in BETA, data-root-relative.
        // productionOwnFiles - the same under production. EMPTY means never promoted, which is what
        //                      makes the system's own folder a removal rather than a restore.
        // sameBytes          - true when both trees hold identical bytes at that path. The caller
        //                      hashes; this class reads no disk.
        // sharedFiles        - the shared files the returning submissions carried, with what the
        //                      caller found in each tree (see BetaRollbackSharedFile).
        // ###########################################################################################
        public static BetaRollbackPlanResult Build(
            IEnumerable<string> betaOwnFiles,
            IEnumerable<string> productionOwnFiles,
            Func<string, bool> sameBytes,
            IReadOnlyList<CarriedSubmission>? returning,
            IReadOnlyList<BetaRollbackSharedFile>? sharedFiles = null,
            bool promotedBefore = false)
        {
            ArgumentNullException.ThrowIfNull(betaOwnFiles);
            ArgumentNullException.ThrowIfNull(productionOwnFiles);
            ArgumentNullException.ThrowIfNull(sameBytes);

            List<string> beta = BetaRollbackPlan.Clean(betaOwnFiles);
            List<string> production = BetaRollbackPlan.Clean(productionOwnFiles);

            IReadOnlyList<CarriedSubmission> queued = returning ?? [];

            List<string> sharedRestored = BetaRollbackPlan.SharedRestored(sharedFiles);

            // ---- Promoted, but production cannot be read: nothing may happen -----------------
            if (production.Count == 0 && promotedBefore)
            {
                return new BetaRollbackPlanResult(
                    BetaRollbackKind.ProductionUnreadable,
                    Restored: [],
                    Removed: [],
                    Returning: queued,
                    SharedRestored: []);
            }

            // ---- Never promoted: the system's folder leaves BETA ------------------------------
            //
            // Not "restore nothing", which would silently leave the bad board exactly as it is.
            if (production.Count == 0)
            {
                return new BetaRollbackPlanResult(
                    BetaRollbackKind.RemoveFromBeta,
                    Restored: [],
                    Removed: beta.OrderBy(path => path, StringComparer.Ordinal).ToList(),
                    Returning: queued,
                    SharedRestored: sharedRestored);
            }

            // ---- Restore: production's bytes win, wholesale ------------------------------------
            //
            // A file production HAS is restored unless BETA already holds those exact bytes; a file
            // production does NOT have is removed, because it only exists in BETA through the work
            // being rolled back. Both halves are needed: restoring alone would leave every file a
            // rolled-back submission ADDED sitting in the tree, cited by nothing and shipped to
            // every BETA user.
            var productionPaths = new HashSet<string>(production, StringComparer.Ordinal);

            List<string> restored = production
                .Where(path => !sameBytes(path))
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();

            List<string> removed = beta
                .Where(path => !productionPaths.Contains(path))
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();

            return new BetaRollbackPlanResult(
                BetaRollbackKind.RestoreFromProduction,
                restored,
                removed,
                queued,
                sharedRestored);
        }

        // ###########################################################################################
        // Which shared files go back. See the header: only a file BETA still holds EXACTLY as a
        // returning submission wrote it is this rollback's to touch, and only production's own
        // version can replace it - one production never had stays.
        // ###########################################################################################
        private static List<string> SharedRestored(IReadOnlyList<BetaRollbackSharedFile>? sharedFiles)
        {
            var restored = new List<string>();

            foreach (BetaRollbackSharedFile file in (sharedFiles ?? [])
                .Where(file => !string.IsNullOrWhiteSpace(file.Path))
                .DistinctBy(file => file.Path.Trim(), StringComparer.Ordinal))
            {
                // Written again by somebody since - not this rollback's change to undo.
                if (!file.BetaHoldsTheSubmittedBytes)
                    continue;

                string path = file.Path.Trim();

                // A shared file the submission ADDED stays, unused (owner decision, 2026-09-27):
                // nothing outside the system's own folder is removed automatically - see
                // AutomaticRemovalScope. Admin > Unused files clears it on purpose.
                if (file.InProduction && !file.SameAsProduction)
                    restored.Add(path);

                // In production with the same bytes: another board's promotion already carried it
                // out, so it is public and reviewed - nothing to undo in BETA.
            }

            return restored.OrderBy(path => path, StringComparer.Ordinal).ToList();
        }

        // Ordinal: two spellings are two files on the server - see the header.
        private static List<string> Clean(IEnumerable<string> paths) =>
            paths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path.Trim())
                .Distinct(StringComparer.Ordinal)
                .ToList();
    }

    // ###########################################################################################
    // One shared file a returning submission carried, and what the caller found for it.
    //
    // BetaHoldsTheSubmittedBytes - BETA's bytes at the path are exactly the hash the submission
    //                              carried, i.e. nobody has written the file since.
    // InProduction               - production has a file at the path.
    // SameAsProduction           - and its bytes equal BETA's.
    // ###########################################################################################
    public sealed record BetaRollbackSharedFile(
        string Path,
        bool BetaHoldsTheSubmittedBytes,
        bool InProduction,
        bool SameAsProduction);

    // Which operation a rollback plan describes for the system's own folder. A caller must not
    // perform one believing it is the other: one puts a board back, the other takes it away.
    // ProductionUnreadable is neither - a promoted system whose production folder cannot be read,
    // which the server refuses rather than performs (see the class header).
    public enum BetaRollbackKind
    {
        RestoreFromProduction,
        RemoveFromBeta,
        ProductionUnreadable
    }

    // ###########################################################################################
    // What a rollback would do.
    //
    //   Restored                - the system's own files taken from production, overwriting BETA;
    //   Removed                 - the system's own files that leave BETA;
    //   Returning               - the submissions that go back to the queue - named, because a
    //                             rollback reverts the whole board and every one of them loses its
    //                             place in BETA;
    //   SharedRestored          - shared files this rollback's submissions changed, put back to
    //                             production's bytes (every board citing them sees it). No shared
    //                             file is ever REMOVED by a rollback - see the class header.
    // ###########################################################################################
    public sealed record BetaRollbackPlanResult(
        BetaRollbackKind Kind,
        IReadOnlyList<string> Restored,
        IReadOnlyList<string> Removed,
        IReadOnlyList<CarriedSubmission> Returning,
        IReadOnlyList<string> SharedRestored)
    {
        public bool TouchesSharedFiles => this.SharedRestored.Count > 0;

        public bool ChangesNothing =>
            this.Restored.Count == 0 &&
            this.Removed.Count == 0 &&
            !this.TouchesSharedFiles;

        // Every path whose bytes come from production - the system's own and the shared ones.
        public IEnumerable<string> AllRestored => this.Restored.Concat(this.SharedRestored);
    }
}
