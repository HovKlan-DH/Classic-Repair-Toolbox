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
    // Nothing on disk is touched: the drafts stay as they are, and an open table goes on saving
    // itself under the cover, since saving needs no server; a change it could not save is still
    // asked about when quitting CRT or installing the update from the overlay (code review,
    // 2026-10-09; Main.ConfirmLeavingTablesAsync).
    //
    // *** NOTHING NEW IS MADE UNDER A COVERED TAB (code review, 2026-10-09). *** A board's first
    // draft and "Add a new board" exist to work in this tab, so they make nothing while it is
    // covered and say why (Main.FirstDraft.cs, Main.NewBoard.cs). Since CRT 4.0 step 2 the cover
    // lies over the whole Contribute tab, its component list included. Another editor's "Save to
    // draft" still saves - its work is done already - onto the draft held in memory, the table's
    // changes included (CRT 4.0 step 1).
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Covers the tab, or - covered already - says `view` instead.
        internal void ShowUpdateRequired(AppUpdateRequiredView view) => this.UpdateRequired.Show(view);

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
