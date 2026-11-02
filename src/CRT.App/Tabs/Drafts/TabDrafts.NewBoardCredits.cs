using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A NEW BOARD'S CREDITS START WITH THEIR ROWS (owner request, 2026-11-02; the rule is CRT.Data's
    // NewBoardCredits): one per schematic image, "Board data", and "Oscilloscope baseline" for PAL and
    // NTSC, for the contributor to fill in or delete.
    //
    // Put in where the contributor goes to them: as "Board: Credits" opens, as the table opens and before
    // Submit - so the rows are there however the board got its images (the cards, a drop, the table), and
    // an image added since gets its own. ONE step of the draft when anything changes, so "Undo" takes it
    // back; what was offered stays in the draft's marker, so a row deleted or undone does not come back.
    // Only for a board with no published counterpart: the board's entry and the marker both say so - the
    // marker alone still says "new board" once its board has been published here and before its draft is
    // retired (code review, 2026-11-02).
    //
    // *** "BOARD STATUS" ONLY SAYS HOW MANY WOULD COME, NEVER PUTS THEM IN (code review, 2026-11-02). ***
    // Worked out again after every change, a read that wrote to the draft was behind two bugs: an Undo
    // there read the board again, which offered the very step back with the Redo gone, so nothing before
    // it could be undone; and its marker write ran into "Board: Name" moving the draft's folder.
    //
    // *** A STEP THAT TOOK A ROW AWAY, UNDONE, PUTS BACK WHAT WAS OFFERED (code review, 2026-11-02). ***
    // The rows are the step's and what was offered the marker's, so after an image's rename and the step
    // that followed it, undoing that step alone brought back the old name's row - no longer "offered", so
    // left for good - with the new name's row "offered", so never added, and Submit stopped on a credit for
    // an image that no longer had that name. So such a step says what was offered before and after it
    // (CreditsFollowedBoard), and an undo or redo of it writes that back: the next offer follows again. A
    // step that only added rows has none of that - undone, they stay gone.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        private sealed record CreditsFollowedBoard(IReadOnlyList<string> OfferedBefore, IReadOnlyList<string> OfferedAfter);

        internal void OfferNewBoardCredits(HardwareBoardEntry entry, DraftSession? held = null)
        {
            if (!entry.IsDraftOnly ||
                (held ?? DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) ?? DraftManager.Sessions.Get(DraftManager.DraftsRoot, entry.ExcelDataFile)) is not { } session)
            {
                return;
            }

            string path = DraftFolderLayout.GetMarkerPath(session.DraftsRoot, session.ExcelDataFile);

            if (DraftMarkerStore.Load(path) is not { IsNewBoard: true } marker)
            {
                return;
            }

            // What an editor holds first, so what is offered is said of the draft as it then is: an image
            // only the table has so far wants its row too.
            session.CollectHeldChanges(DraftStepWording.CreditsOfNewBoard);

            IReadOnlyList<string> offered = marker.CreditsOffered ?? [];
            NewBoardCreditOffer offer = NewBoardCredits.Offer(session.Content.Board, offered);

            // Nothing to do is no step, and nothing written.
            if (!ReferenceEquals(offer.Board, session.Content.Board))
            {
                if (offer.RowsWent)
                {
                    session.Changed -= this.OnCreditsStepMoved;
                    session.Changed += this.OnCreditsStepMoved;
                }

                session.Apply(
                    DraftStepWording.CreditsOfNewBoard,
                    content => NewBoardCredits.Offer(content.Board, offered).Board is BoardData board && !ReferenceEquals(board, content.Board)
                        ? content with { Board = board }
                        : null,
                    offer.RowsWent ? new CreditsFollowedBoard(offered, offer.Offered) : null);

                this.ShowCreditsOffered(entry);
            }

            if (!offer.Offered.SequenceEqual(offered))
            {
                TabDrafts.NoteCreditsOffered(path, offer.Offered);
            }
        }

        // The rows in, on the board shown: the left column counts them, and "Board status" says them when it
        // is what shows - whichever way they came, with nothing open beside it before Submit.
        private void ShowCreditsOffered(HardwareBoardEntry entry)
        {
            if (this.thisCurrentEntry is not { } shown || !TabDrafts.SameBoard(shown, entry))
            {
                return;
            }

            this.ShowBoardParts();

            if (this.StartPanel.IsVisible && this.NothingElseOpenOnTheRight)
            {
                this.ShowReadiness();
            }
        }

        // How many rows an offer would add now, with the draft left as it is - what "Board status" says.
        // Off the UI thread: the marker is read for it.
        private static int NewBoardCreditsToOffer(HardwareBoardEntry entry, DraftSession session) =>
            entry.IsDraftOnly &&
            DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(session.DraftsRoot, session.ExcelDataFile)) is { IsNewBoard: true } marker
                ? NewBoardCredits.RowsToAdd(session.Content.Board, marker.CreditsOffered)
                : 0;

        // `board` as an offer would leave it, with the draft left as it is - the same board when there is
        // nothing to offer. What another draft's errors are counted on before it is sent
        // (TabDrafts.SubmitOthers.cs). Off the UI thread: the marker is read for it.
        private static BoardData WithNewBoardCreditsOffered(HardwareBoardEntry entry, BoardData board) =>
            entry.IsDraftOnly &&
            DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, entry.ExcelDataFile)) is { IsNewBoard: true } marker
                ? NewBoardCredits.Offer(board, marker.CreditsOffered).Board
                : board;

        private void OnCreditsStepMoved(object? sender, DraftSessionChange change)
        {
            if (sender is DraftSession session && change.Origin is CreditsFollowedBoard step &&
                change.Kind is DraftSessionChangeKind.Undo or DraftSessionChangeKind.Redo)
            {
                TabDrafts.NoteCreditsOffered(
                    DraftFolderLayout.GetMarkerPath(session.DraftsRoot, session.ExcelDataFile),
                    change.Kind == DraftSessionChangeKind.Undo ? step.OfferedBefore : step.OfferedAfter);
            }
        }

        private static void NoteCreditsOffered(string path, IReadOnlyList<string> offered)
        {
            // Not kept, a row deleted may be offered once more - no reason to stop a Submit over.
            try
            {
                if (DraftMarkerStore.Load(path) is { } marker)
                {
                    DraftMarkerStore.Save(path, marker with { CreditsOffered = offered });
                }
            }
            catch (Exception ex) when (ex is System.IO.IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not note the credits offered in [{path}] - [{ex.Message}]");
            }
        }
    }
}
