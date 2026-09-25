using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // RUNS DRAFT RETIREMENT SAFELY: the slow check off the UI thread, the delete on it, and a
    // re-check in between (code review, 2026-09-25).
    //
    // WHETHER a draft may go is DraftRetirement's decision (in CRT.Data, pure, unit tested). This
    // class is only the sequencing around it, pulled out of Main so the sequencing is testable
    // too - Main is the one caller, and every step it needs is handed in as a delegate.
    //
    // *** WHAT WAS WRONG WITH DOING IT INLINE IN Main. ***
    //
    //   - The check parses two workbooks and reads every file of each retirable draft. It ran
    //     inside Dispatcher.UIThread.Post, so the window froze while it did.
    //   - Nothing re-checked the draft between deciding and deleting. With the check now off the
    //     UI thread there IS a gap, and a save landing in it would be deleted unseen - so the
    //     folder's stamp from before the check must still hold at the moment of deletion.
    //   - A draft whose table was open with unsaved edits was deleted under the table. Those edits
    //     are in memory, not on disk, so the folder being identical to the published board proves
    //     nothing about them.
    //   - The board cache was never cleared, so a board on screen kept reading as drafted from a
    //     folder that no longer existed.
    // ###########################################################################################
    public static class PublishedDraftRetirer
    {
        // ###########################################################################################
        // Finds the drafts whose work is now published, on a pool thread. The receipts are a
        // snapshot taken by the caller, so nothing here reads shared UI state.
        // ###########################################################################################
        public static Task<IReadOnlyList<RetirableDraft>> FindAsync(
            IReadOnlyList<SubmissionReceipt> receipts,
            Func<string, DraftStatus?> resolveStatus)
        {
            ArgumentNullException.ThrowIfNull(receipts);
            ArgumentNullException.ThrowIfNull(resolveStatus);

            return Task.Run(() => DraftRetirement.FindRetirableDrafts(receipts, resolveStatus));
        }

        // ###########################################################################################
        // Deletes each candidate that is still safe to delete, on the caller's (UI) thread.
        //
        //   isInUse      - true when something holds unsaved edits for the system (the Drafts
        //                  table); such a draft is kept, since those edits exist nowhere on disk.
        //   discard      - deletes the folder; false when it could not finish.
        //   afterDiscard - runs for every system a delete was ATTEMPTED on, finished or not, since
        //                  a partial delete still leaves the board cache describing files that
        //                  are gone. The caller clears its caches here.
        // ###########################################################################################
        public static DraftRetirementOutcome Retire(
            IEnumerable<RetirableDraft> candidates,
            Func<string, bool> isInUse,
            Func<string, bool> discard,
            Action<string> afterDiscard)
        {
            ArgumentNullException.ThrowIfNull(isInUse);
            ArgumentNullException.ThrowIfNull(discard);
            ArgumentNullException.ThrowIfNull(afterDiscard);

            var retired = new List<string>();
            var failed = new List<string>();

            foreach (RetirableDraft draft in candidates ?? [])
            {
                if (isInUse(draft.SystemId))
                {
                    Logger.Info($"Draft for [{draft.SystemId}] is published but has unsaved edits open - kept.");
                    continue;
                }

                if (!DraftRetirement.IsUnchangedSince(draft))
                {
                    Logger.Info($"Draft for [{draft.SystemId}] changed while it was being checked - kept.");
                    continue;
                }

                bool removed = discard(draft.SystemId);

                afterDiscard(draft.SystemId);

                if (removed)
                {
                    // Logged at Info rather than silently: a folder disappearing on its own is the
                    // kind of thing somebody will eventually want to find an explanation for.
                    Logger.Info(
                        $"Draft for [{draft.SystemId}] removed automatically - its changes are now in the " +
                        "published data.");

                    retired.Add(draft.SystemId);
                }
                else
                {
                    Logger.Warning($"Draft for [{draft.SystemId}] is published but could not be removed.");
                    failed.Add(draft.SystemId);
                }
            }

            return new DraftRetirementOutcome(retired, failed);
        }
    }

    // ###########################################################################################
    // What Retire did. Touched is every system whose folder a delete was attempted on - the ones
    // whose board, if on screen, has to be reloaded.
    // ###########################################################################################
    public sealed record DraftRetirementOutcome(IReadOnlyList<string> Retired, IReadOnlyList<string> Failed)
    {
        public IReadOnlyList<string> Touched => [.. this.Retired, .. this.Failed];
    }
}
