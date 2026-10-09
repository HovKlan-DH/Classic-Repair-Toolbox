using System;
using Handlers.Online;

namespace CRT
{
    // ###########################################################################################
    // "CRT HAS TO BE UPDATED" OVER THIS TAB (owner request, 2026-10-09) - when the server turns this
    // version of CRT away from submissions. Main decides when (Main.UpdateRequired.cs, from
    // AppUpdateRequirement) and what the button does; the overlay (UpdateRequiredOverlay) covers the
    // whole tab and cannot be closed.
    //
    // Nothing on disk is touched: the drafts stay as they are, an open table keeps its unsaved
    // edits, and quitting CRT - or installing the update from the overlay (code review, 2026-10-09;
    // Main.ConfirmLeavingTablesAsync) - still asks about them. Its Save works with the tab covered,
    // since it needs no server.
    //
    // *** NOTHING NEW IS SENT TO A COVERED TAB (code review, 2026-10-09). *** The Contribute tab's
    // "Edit board as draft" and "Add a new board" exist to work in this tab, so they make nothing
    // while it is covered and say why (Main.EditBoardAsDraft.cs, Main.NewBoard.cs). Another
    // editor's "Save to draft" still saves - its work is done already - but one held back by this
    // table's unsaved edits is told what settles them now that this tab cannot (SavingBlockedFor).
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Covers the tab, or - covered already - says `view` instead.
        internal void ShowUpdateRequired(AppUpdateRequiredView view) => this.UpdateRequired.Show(view);

        // ###########################################################################################
        // Whether another editor's save into the draft of `excelDataFile` must wait, and which notice
        // says so - null when it may go ahead. The table open on that draft with unsaved edits holds
        // it back (HasUnsavedTableEditsFor); covered, the notice says how they get settled instead.
        // ###########################################################################################
        internal UnsavedTableEditsPrompt? SavingBlockedFor(string? excelDataFile)
        {
            if (!this.HasUnsavedTableEditsFor(excelDataFile))
                return null;

            return this.IsUpdateRequiredShown
                ? UnsavedTableEditsPrompt.SavingElsewhereUpdateRequired
                : UnsavedTableEditsPrompt.SavingElsewhere;
        }

        internal bool IsUpdateRequiredShown => this.UpdateRequired.IsShown;

        // What the overlay says, or null while the tab is not covered.
        internal AppUpdateRequiredView? UpdateRequiredView => this.UpdateRequired.View;

        // The overlay's button was pressed.
        internal event EventHandler? UpdateRequiredActionClicked
        {
            add => this.UpdateRequired.ActionClicked += value;
            remove => this.UpdateRequired.ActionClicked -= value;
        }
    }
}
