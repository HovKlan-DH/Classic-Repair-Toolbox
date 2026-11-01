using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "BOARD: NAME" (owner request, 2026-11-01: "When I add a new board, then I get all the input fields
    // and I can add the new board - but how to edit those details again afterwards"): an entry of the left
    // column, a page on the right - ONE thing there at a time, as a component or "KiCad: Files" is.
    //
    //   - A NEW BOARD'S NAME AND NOTES CHANGE HERE, until it is sent for review. A change of name moves
    //     its draft (DraftBoardRename): what the draft holds saved first, and it let go of once moved - its
    //     session knows the file by the old identity, so its Undo steps go. Main then carries over what this
    //     machine keeps under the old name and shows the board under the new one (Main.AfterBoardRenamed).
    //     Notes alone move nothing.
    //   - A PUBLISHED BOARD'S NAME IS ONLY SHOWN: it is how every CRT and the server find the board.
    //     Showing it makes no draft.
    //   - SENT FOR REVIEW, it keeps its name while a maintainer has it: the server knows it by the
    //     name, and a submission under a new one would be a second board beside the first. Only once
    //     that is decided against can it change again (SubmissionReceiptPresenter.IsDecidedAgainst) -
    //     accepted or published is final too, and then the board is a published one.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // The page is what is open on the right - set by picking it, let go of as anything else opens
        // (ShowNavSelection).
        private bool thisBoardNamePicked;

        internal bool IsBoardNameOpen => this.thisBoardNamePicked;

        internal BoardNamePanel BoardNamePanelForTests => this.BoardNamePanel;

        // Stands in for Main's part of a save - a test has no Main.
        internal Action<HardwareBoardEntry, NewBoardRegistration>? BoardRenamedOverrideForTests { get; set; }

        private void OnBoardNameClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(this.OpenBoardNameAsync, "Board name");

        // ###########################################################################################
        // The page, in place of what is open on the right - each closed as its own close does, a picture's
        // name it cannot take keeping the Pictures part open.
        // ###########################################################################################
        internal async Task OpenBoardNameAsync()
        {
            if (this.thisCurrentEntry is not { } entry)
            {
                return;
            }

            if (this.IsComponentOpen)
            {
                if (this.ComponentPane.KeepsRefusedPictureName())
                {
                    return;
                }

                this.CloseComponent();
            }

            await this.CloseKiCadAsync();

            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            this.thisBoardNamePicked = true;

            this.ShowBoardNamePage(entry);
            this.ShowWorkspace();
        }

        private void ShowBoardNamePage(HardwareBoardEntry entry)
        {
            NewBoardRegistration? registration = entry.IsDraftOnly
                ? DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, entry.ExcelDataFile))?.NewBoard
                : null;

            this.BoardNamePanel.SaveRequested = wanted => this.SaveBoardNameAsync(entry, wanted);
            this.BoardNamePanel.Show(
                entry.ExcelDataFile,
                registration?.HardwareName ?? entry.HardwareName,
                registration?.BoardName ?? entry.BoardName,
                registration?.HardwareNotes ?? string.Empty,
                this.HardwareBoardsOverrideForTests ?? DataManager.HardwareBoards,
                this.WhyNoNewName(entry, registration),
                isNewBoard: registration is not null);
        }

        // Why this board's name cannot be changed here - null when it can.
        private string? WhyNoNewName(HardwareBoardEntry entry, NewBoardRegistration? registration)
        {
            if (registration is null)
            {
                return BoardNameWording.Published;
            }

            string boardId = BoardDescriptorRules.BuildBoardId(
                NewBoardIdentity.ExtractManufacturer(entry.ExcelDataFile),
                NewBoardIdentity.ExtractHardware(entry.ExcelDataFile),
                NewBoardIdentity.ExtractBoard(entry.ExcelDataFile));

            IEnumerable<SubmissionReceipt> receipts = this.ReceiptsOverrideForTests ?? SubmissionReceiptStore.All;

            List<SubmissionReceipt> sent = receipts
                .Where(receipt =>
                    (string.Equals(receipt.BoardId?.Trim(), boardId, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(receipt.BoardId?.Trim(), entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase)) &&
                    receipt.NotFoundUtc is null)
                .ToList();

            // Published, on a machine that has not synced it yet: it reads as a new board here, and under
            // a new name its next submission would be a second board beside the published one.
            if (sent.Any(receipt => SubmissionReceiptPresenter.IsInPublishedData(receipt.LastKnownState)))
            {
                return BoardNameWording.Published;
            }

            return sent.Any(receipt => !SubmissionReceiptPresenter.IsDecidedAgainst(receipt.LastKnownState))
                ? BoardNameWording.SentForReview
                : null;
        }

        // ###########################################################################################
        // What the page's button does: what the draft holds saved, its folder moved under the BusyOverlay,
        // it let go of, then Main's part. What stands in the way is said on the page - null when done.
        // ###########################################################################################
        private async Task<string?> SaveBoardNameAsync(HardwareBoardEntry entry, NewBoardRegistration wanted)
        {
            // A picture tool's Done writes into the draft as it holds it, and a board read again would
            // throw its work away - as for "Mark all automatically", it waits for the tool.
            if (this.PictureToolOnOverrideForTests?.Invoke() ?? this.thisMainWindow?.IsPictureToolOn == true)
            {
                return PictureToolWording.PictureToolOn;
            }

            if (!await this.SettleTableEditsAsync(entry))
            {
                return BoardNameWording.NotSaved;
            }

            string draftsRoot = DraftManager.DraftsRoot;
            string identity = entry.ExcelDataFile;

            // Asked again now: a submission may have gone meanwhile.
            NewBoardRegistration? registration = DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(draftsRoot, identity))?.NewBoard;

            if (this.WhyNoNewName(entry, registration) is { } whyNot)
            {
                return whyNot;
            }

            DraftRenameResult result = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.RenamingBoard,
                () => Task.Run(() => DraftBoardRename.Run(draftsRoot, identity, wanted)));

            if (!result.Renamed)
            {
                return result.Reason;
            }

            bool moved = !string.Equals(identity, wanted.ExcelDataFile, StringComparison.Ordinal);

            // ###########################################################################################
            // The draft held in memory is let go of only now, and only when its file has changed under it
            // (code review, 2026-11-01): closed before the rename was tried, a rename refused with
            // "Nothing was changed" and a save of notes alone lost its Undo steps too.
            // ###########################################################################################
            if (moved ||
                !string.Equals(registration?.HardwareName, wanted.HardwareName, StringComparison.Ordinal) ||
                !string.Equals(registration?.BoardName, wanted.BoardName, StringComparison.Ordinal))
            {
                DraftManager.Sessions.Close(draftsRoot, identity);
            }

            // Moved, the board shown is another board now: Board status follows, as for any other - and
            // until Main has it on screen nothing here may act on the folder that has gone.
            this.thisBoardNamePicked = !moved;
            this.ForgetKiCadCount();

            if (moved)
            {
                this.ShowBoard(null, null, this.thisCurrentRegion, null);
            }

            if (this.BoardRenamedOverrideForTests is { } renamed)
            {
                renamed(entry, wanted);
            }
            else
            {
                this.thisMainWindow?.AfterBoardRenamed(entry, wanted);
            }

            if (!moved)
            {
                this.ShowBoardNamePage(entry);
                this.BoardNamePanel.ShowSaved(BoardNameWording.Saved);
            }

            return null;
        }
    }
}
