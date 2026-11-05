using Avalonia.Input;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // FILES DROPPED ON "BOARD: SCHEMATIC IMAGES" (CRT 4.0 step 3, slice 3e - the plan: "Pictures: today's
    // Schematic images content, moved inline").
    //
    // *** ONLY THAT PAGE TAKES IMAGES (owner decision, 2026-11-03: "it should not allow to drag any
    // schematics into" the Overview page - "it is not the logical place to do this"). *** Until then
    // anything dropped anywhere on the tab was added, and opened what it belonged to. Images now go in
    // where they are seen, as "Add schematic image(s)" adds them (ComponentPane.AddPicturesAsync); a KiCad
    // project goes on "KiCad: Files", whose panel has its own drop target - dropped here it is told so.
    // Every other page of the tab refuses the drag, so the pointer says it before anything is let go.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Stands in for adding what was dropped.
        internal Func<HardwareBoardEntry, IReadOnlyList<string>, Task>? AddDroppedPicturesOverrideForTests { get; set; }

        // Stands in for reading what was dropped - a test holds it to change the board meanwhile.
        internal Func<IReadOnlyList<string>, Task<BoardDrop>>? ReadDropOverrideForTests { get; set; }

        private void WireDrop()
        {
            DragDrop.SetAllowDrop(this, true);
            this.AddHandler(DragDrop.DragOverEvent, this.OnDragOverTab);
            this.AddHandler(DragDrop.DropEvent, this.OnDropOnTab);
        }

        // The page that takes images is open on a board - and never under "CRT has to be updated", which
        // lies inside the tab (code review, 2026-10-28).
        private bool TakesDroppedPictures =>
            this.thisCurrentEntry is not null &&
            !this.IsUpdateRequiredShown &&
            this.IsComponentOpen &&
            this.ComponentPane.BoardPart == PaneBoardPart.Pictures;

        // Only a copy is offered, and only of files, onto that page.
        private void OnDragOverTab(object? sender, DragEventArgs e)
        {
            e.DragEffects = this.TakesDroppedPictures && e.DataTransfer.Contains(DataFormat.File)
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
        // What was dropped, added to the board shown - read off the UI thread (a folder can be large)
        // under the BusyOverlay, and only by the board it was dropped on: a board shown while the drop
        // was read - the read awaited, the board taken after - took the files dropped on the one before
        // (code review, 2026-10-28). The page has to be open still: nothing goes in out of sight.
        // ###########################################################################################
        internal async Task DropOnBoardAsync(IReadOnlyList<string> paths)
        {
            if (!this.TakesDroppedPictures || this.thisCurrentEntry is not { } droppedOn || paths.Count == 0)
            {
                return;
            }

            BoardDrop drop = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.ReadingDroppedFiles,
                () => this.ReadDropOverrideForTests?.Invoke(paths) ?? Task.Run(() => BoardDrop.OnDisk(paths)));

            if (!this.TakesDroppedPictures || this.thisCurrentEntry is not { } shown || !TabDrafts.SameBoard(shown, droppedOn))
            {
                return;
            }

            if (drop.Kind == BoardDropKind.Nothing)
            {
                this.ShowDraftProblem(drop.Problem);
                return;
            }

            // What was dropped with them and left out, said on the tab's line.
            this.ShowDraftProblem(drop.Note.Length > 0 ? drop.Note : null);

            if (this.AddDroppedPicturesOverrideForTests is { } add)
            {
                await add(shown, drop.Paths);
                return;
            }

            await this.ComponentPane.AddPicturesAsync(drop.Paths);
        }
    }
}
