using Avalonia.Input;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

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
    // *** THE POINTER HANDLING IS CRT.UI's ListRowDrag *** - frozen slots, the re-entrancy guard,
    // pointer capture and edge auto-scroll, which is what stops the rows "rapidly switching
    // position and not settling" (the race the owner warned about). It was written here first and
    // lifted out when the Maintainer tab needed the same drag for placing a new system in
    // the drop-down lists, so both run the one copy. This file keeps only what is this window's:
    // saving the order a drop produced.
    // ###########################################################################################
    public partial class SystemFilesWindow
    {
        private ListRowDrag<SystemSchematicRow>? thisSchematicRowDrag;

        // The window hosts the drag: the pointer spends part of it outside the list (over the drop
        // box above, or past the window's edge to scroll), and the window sees all of that.
        private void WireSchematicRowDrag() =>
            this.thisSchematicRowDrag = new ListRowDrag<SystemSchematicRow>(
                this,
                this.SchematicsItemsControl,
                this.ContentScrollViewer,
                this.Schematics,
                (row, _, to) => this.SaveSchematicOrder(row, to));

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
        // Saves the order ON SCREEN into the draft, by name (SchematicOrder).
        //
        // The rows are NOT reloaded after a good save: they are already in the order just written,
        // and a reload would decode every preview again for a board of a couple of dozen large
        // scans. They ARE reloaded when the save failed, or when the draft on disk turned out to
        // hold different schematics than the list (changed in Excel meanwhile) - either way the
        // list must show what the draft really holds.
        // ###########################################################################################
        private void SaveSchematicOrder(SystemSchematicRow movedRow, int newIndex)
        {
            List<string> names = this.Schematics.Select(row => row.SchematicName).ToList();
            bool listMatchedTheDraft = false;

            bool saved = DraftWorkbookStore.Edit(
                DraftManager.DraftsRoot,
                this.thisExcelDataFile,
                board =>
                {
                    List<BoardSchematicEntry> ordered = SchematicOrder.Apply(board.Schematics, names);

                    listMatchedTheDraft = ordered
                        .Select(entry => entry.SchematicName?.Trim() ?? string.Empty)
                        .SequenceEqual(names.Select(name => name.Trim()), StringComparer.OrdinalIgnoreCase);

                    return SystemFilesWindow.WithSchematicList(board, ordered);
                });

            if (!saved)
            {
                this.ReloadSchematicsFromDraft();
                this.ShowStatus(
                    "The new order could not be saved. If this board's Excel file is open in Excel, close it there and try again.",
                    isError: true);
                return;
            }

            if (!listMatchedTheDraft)
            {
                this.ReloadSchematicsFromDraft();
                this.ShowStatus(
                    "The order was saved, but the draft had changed outside this window, so the list has been read again. Check the order.",
                    isError: true);
                return;
            }

            this.ShowStatus($"Moved [{movedRow.SchematicName}] to position {newIndex + 1} of {this.Schematics.Count}.");
        }
    }
}
