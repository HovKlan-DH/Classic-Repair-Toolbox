using System;
using Handlers.Online;

namespace CRT
{
    // ###########################################################################################
    // "CRT HAS TO BE UPDATED" OVER THIS TAB (owner request, 2026-10-09) - when the server turns this
    // version of CRT away from the review, administration and account routes, which is everything
    // this tab does. Main decides when (Main.UpdateRequired.cs, from AppUpdateRequirement) and what
    // the button does; the overlay (UpdateRequiredOverlay) covers the whole tab and cannot be closed.
    //
    // *** THE MINUTE CHECKS STOP, AND STAY STOPPED. *** Every one would be answered the same way, and
    // each costs the server a request it counts (API usage). The remembered sign-in is NOT forgotten:
    // a 426 is not a 401, and after the update the session works again as it was.
    //
    // *** A TABLE'S UNSAVED CHANGES ARE ASKED ABOUT WITHOUT A SAVE (code review, 2026-10-09). ***
    // Quitting or installing the update asks about both tables as ever, but a Save would go to the
    // server, which turns this CRT away - and its refusal was written under the cover, the quit
    // silently cancelled. So the question is Discard or Cancel (LeavingUpdateRequired), saying why.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        // Covers the tab, or - covered already - says `view` instead.
        internal void ShowUpdateRequired(AppUpdateRequiredView view)
        {
            this.StopQueueChecks();
            this.UpdateRequired.Show(view);
        }

        // ###########################################################################################
        // Whether the tab is covered - THE one record of it (code review, 2026-10-10: a flag beside
        // the overlay said the same and could drift from it). StartQueueChecks, every minute check
        // and both leaving prompts read it, so the checks stay stopped however the overlay came.
        // ###########################################################################################
        internal bool IsUpdateRequiredShown => this.UpdateRequired.IsShown;

        // What the overlay says, or null while the tab is not covered.
        internal AppUpdateRequiredView? UpdateRequiredView => this.UpdateRequired.View;

        // Whether the minute checks are running - for tests.
        internal bool QueueChecksRunningForTests => this.thisQueueTimer?.IsEnabled == true;

        // The overlay's button was pressed.
        internal event EventHandler? UpdateRequiredActionClicked
        {
            add => this.UpdateRequired.ActionClicked += value;
            remove => this.UpdateRequired.ActionClicked -= value;
        }
    }
}
