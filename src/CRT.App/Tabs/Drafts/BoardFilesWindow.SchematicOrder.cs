using Avalonia.Input;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // DRAGGING THE SCHEMATIC IMAGES INTO A NEW ORDER (owner request, 2026-09-27) - "the same
    // move functionality which is already present in the worklog modal", with its placeholder.
    //
    // Press anywhere on a row but its Remove button, move a few pixels, and the row becomes a
    // dashed slot that travels with the pointer; the other rows stand in the order a drop will
    // produce. Releasing saves that order into the draft's "Board schematics" sheet - the BOARD's
    // own order, published with it as the default every user of this data sees the views in.
    //
    // *** IT NEVER TOUCHES A USER'S OWN ORDER (owner decision, 2026-09-27). *** Dragging
    // thumbnails on the Schematics tab saves a personal order per board (UserSettings
    // .SchematicsOrderByBoard), which overrides the board's order for that user alone. The two are
    // different things - the data's default for everyone, and one person's preference - so this
    // window leaves the personal one exactly as it is.
    //
    // *** THE POINTER HANDLING IS Controls/ListRowDrag *** - frozen slots, the re-entrancy guard,
    // pointer capture and edge auto-scroll, which is what stops the rows "rapidly switching
    // position and not settling" (the race the owner warned about). It was written here first and
    // lifted out when the Maintainer tab needed the same drag for placing a new board in
    // the drop-down lists, so both run the one copy. This file keeps only what is this window's:
    // saving the order a drop produced.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        private ListRowDrag<BoardSchematicRow>? thisSchematicRowDrag;

        // The window hosts the drag: the pointer spends part of it outside the list (over the drop
        // box above, or past the window's edge to scroll), and the window sees all of that.
        private void WireSchematicRowDrag() =>
            this.thisSchematicRowDrag = new ListRowDrag<BoardSchematicRow>(
                this,
                this.SchematicsItemsControl,
                this.ContentScrollViewer,
                this.Schematics,
                (row, from, to) => _ = this.SaveSchematicOrderAsync(row, to));

        // The row's handle (every part of the row but its Remove button, which handles its own press).
        private void OnSchematicRowPointerPressed(object? sender, PointerPressedEventArgs e) =>
            this.thisSchematicRowDrag?.Press(sender, e);

        // Ends any drag - the rows are about to be replaced.
        private void ResetSchematicRowDrag() => this.thisSchematicRowDrag?.Reset();

        // Lets a headless test run one auto-scroll tick - DispatcherTimer does not tick there.
        internal void AutoScrollSchematicDragForTests() => this.thisSchematicRowDrag?.AutoScrollForTests();

        // Lets a headless test take the capture away mid-drag the way the system does.
        internal void LoseSchematicDragCaptureForTests() => this.thisSchematicRowDrag?.LoseCaptureForTests();

        // ###########################################################################################
        // Saves the order ON SCREEN into the draft, by name (SchematicOrder). One of the window's
        // draft changes: queued, off the UI thread, under the "please wait"
        // (BoardFilesWindow.DraftChanges.cs) - the order is read when it runs, after any rename
        // queued before it has given its row the new name.
        //
        // The rows are NOT reloaded after a good save: they are already in the order just written,
        // and a reload would decode every preview again for a board of a couple of dozen large
        // scans. They ARE reloaded when the save failed, or when the draft on disk turned out to
        // hold different schematics than the list (changed in Excel meanwhile) - either way the
        // list must show what the draft really holds.
        // ###########################################################################################
        private Task SaveSchematicOrderAsync(BoardSchematicRow movedRow, int newIndex) =>
            this.ChangeDraftAsync(async () =>
            {
                // Each name with its image file, so two rows sharing a name keep the places they
                // were dragged to (SchematicOrder).
                List<(string SchematicName, string? ImageFile)> order = this.Schematics
                    .Select(row => (row.SchematicName, (string?)row.ImageFile))
                    .ToList();

                string draftsRoot = DraftManager.DraftsRoot;
                string excelDataFile = this.thisExcelDataFile;
                bool listMatchedTheDraft = false;

                bool saved = await this.WriteDraftAsync(CrtWaitWording.SavingSchematicOrder, () => DraftWorkbookStore.Edit(
                    draftsRoot,
                    excelDataFile,
                    board =>
                    {
                        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(board.Schematics, order);

                        listMatchedTheDraft = ordered.Count == order.Count && ordered
                            .Select((entry, i) =>
                                SchematicRename.SameName(entry.SchematicName, order[i].SchematicName) &&
                                SchematicRows.SameImageFile(entry.SchematicImageFile, order[i].ImageFile))
                            .All(same => same);

                        return BoardFilesWindow.WithSchematicList(board, ordered);
                    }));

                if (!saved)
                {
                    this.ShowStatus(
                        "The new order could not be saved. If this board's Excel file is open in Excel, close it there and try again.",
                        isError: true);

                    await this.ReloadSchematicsUnderOverlayAsync();
                    return;
                }

                this.ChangedTheDraft = true;

                if (!listMatchedTheDraft)
                {
                    this.ShowStatus(
                        "The order was saved, but the draft had changed outside this window, so the list has been read again. Check the order.",
                        isError: true);

                    await this.ReloadSchematicsUnderOverlayAsync();
                    return;
                }

                this.ShowStatus($"Moved [{movedRow.SchematicName}] to position {newIndex + 1} of {this.Schematics.Count}.");
            });
    }
}
