using Avalonia.Input;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // FILES DROPPED ON THE CONTRIBUTE TAB (CRT 4.0 step 3, slice 3e - the plan: "Pictures: today's
    // Schematic images content, moved inline ... KiCad: drop the project folder and get a summary of
    // what was found, what was ignored, and what is left to do").
    //
    // Pictures of the board, or its KiCad project folder, dragged onto the tab anywhere are added to
    // the board shown: what was dropped decides which (BoardDrop). Pictures go into the Pictures part,
    // opened for them, as its "Add pictures" adds them (ComponentPane.AddPicturesAsync); a KiCad
    // project opens "KiCad data" with it being imported, as that window's own drop target imports it
    // (BoardFilesWindow.ImportOnceOpen) - its summary and all. The board's draft is made as for every
    // other way in (OnCurrentDraftAsync). Nothing usable dropped is said on the tab's problem line.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Stands in for adding what was dropped - the pictures, or the KiCad project's window.
        internal Func<HardwareBoardEntry, BoardDropKind, IReadOnlyList<string>, Task>? ManageDroppedFilesOverrideForTests { get; set; }

        // Stands in for reading what was dropped - a test holds it to change the board meanwhile.
        internal Func<IReadOnlyList<string>, Task<BoardDrop>>? ReadDropOverrideForTests { get; set; }

        private void WireDrop()
        {
            DragDrop.SetAllowDrop(this, true);
            this.AddHandler(DragDrop.DragOverEvent, this.OnDragOverTab);
            this.AddHandler(DragDrop.DropEvent, this.OnDropOnTab);
        }

        // Only a copy is offered, and only of files, onto a board - the window's own drop targets' rule -
        // and never onto "CRT has to be updated", which lies inside the tab (code review, 2026-10-28).
        private void OnDragOverTab(object? sender, DragEventArgs e)
        {
            e.DragEffects = this.thisCurrentEntry is not null && !this.IsUpdateRequiredShown && e.DataTransfer.Contains(DataFormat.File)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void OnDropOnTab(object? sender, DragEventArgs e)
        {
            e.Handled = true;
            List<string> paths = StorageItemPaths.LocalPaths(e.DataTransfer.TryGetFiles());
            this.RunGuarded(() => this.DropOnBoardAsync(paths), "Dropped files");
        }

        // ###########################################################################################
        // What was dropped, added to the board shown: the pictures, or the KiCad project - read off the
        // UI thread (a folder can be large) under the BusyOverlay, then its window opened with them.
        //
        // *** NOTHING UNDER "CRT HAS TO BE UPDATED", AND ONLY ON THE BOARD IT WAS DROPPED ON (code
        // review, 2026-10-28). *** The cover lies inside the tab, so a drop on it reached the tab and
        // opened a window over it; and a board shown while the drop was read - the read awaited, the
        // board taken after - took the files dropped on the one before.
        // ###########################################################################################
        internal async Task DropOnBoardAsync(IReadOnlyList<string> paths)
        {
            if (this.thisCurrentEntry is not { } droppedOn || paths.Count == 0 || this.IsUpdateRequiredShown)
            {
                return;
            }

            BoardDrop drop = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.ReadingDroppedFiles,
                () => this.ReadDropOverrideForTests?.Invoke(paths) ?? Task.Run(() => BoardDrop.OnDisk(paths)));

            if (this.thisCurrentEntry is not { } shown || !TabDrafts.SameBoard(shown, droppedOn) || this.IsUpdateRequiredShown)
            {
                return;
            }

            if (drop.Kind == BoardDropKind.Nothing)
            {
                this.ShowDraftProblem(drop.Problem);
                return;
            }

            this.ShowDraftProblem(null);

            // ###########################################################################################
            // What was dropped with it and left out, said on the tab's line (BoardDrop.Note) - ONCE
            // THE DRAFT IS THERE (code review, 2026-10-28): making a board's first draft clears that
            // line (Main.EnsureDraftAsync), and the note said before it went unseen on most boards.
            // ###########################################################################################
            void SayWhatWasLeftOut() => this.ShowDraftProblem(drop.Note.Length > 0 ? drop.Note : null);

            if (this.ManageDroppedFilesOverrideForTests is { } manage)
            {
                await this.OnCurrentDraftAsync(entry =>
                {
                    SayWhatWasLeftOut();
                    return manage(entry, drop.Kind, drop.Paths);
                });

                return;
            }

            if (drop.Kind == BoardDropKind.KiCad)
            {
                await this.OnCurrentDraftAsync(entry =>
                {
                    SayWhatWasLeftOut();
                    return this.ManageFilesAsync(entry, drop.Paths);
                });

                return;
            }

            await this.AddDroppedPicturesAsync(drop.Paths, SayWhatWasLeftOut);
        }

        // ###########################################################################################
        // Pictures dropped: the Pictures part opened on the board shown - its draft made first when it
        // has none - and the pictures added there, as "Add pictures" adds them (step 3 - the
        // "Schematic images" window went). `sayWhatWasLeftOut` once the part is open on its draft.
        // ###########################################################################################
        private async Task AddDroppedPicturesAsync(IReadOnlyList<string> pictures, Action sayWhatWasLeftOut)
        {
            await this.OpenBoardPartAsync(PaneBoardPart.Pictures);

            if (this.IsComponentOpen && this.ComponentPane.BoardPart == PaneBoardPart.Pictures)
            {
                sayWhatWasLeftOut();
                await this.ComponentPane.AddPicturesAsync(pictures);
            }
        }
    }
}
