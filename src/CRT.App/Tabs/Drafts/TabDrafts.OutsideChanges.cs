using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // THE LIST FOLLOWS A DRAFT CHANGED OUTSIDE CRT (owner report, 2026-10-02: "It must check for
    // errors when creating the draft, and if the board changes "offline", outside of app").
    //
    // A draft's row says how many rows it changes and - since the same day - how many errors and
    // warnings it has. Both were worked out only when something IN CRT refreshed the list (a save,
    // a board load), so a draft edited in Excel kept showing its old numbers. Now the list is read
    // again whenever the tab is shown and whenever CRT's window comes back to the front while it
    // is - which is exactly the moment after an edit in Excel.
    //
    // Cheap: every row's numbers are remembered against the length and write time of the files
    // they come from (DraftStatusReader.CountChangesCached / CountProblemsCached), so a draft
    // nobody touched costs a few file-information reads, and only a changed one is read again.
    //
    // The OPEN TABLE reads its draft's file again at the same moments, when it changed and nothing
    // of the table's waits to be saved (ReadDraftFileAgain, CRT 4.0 step 1), and that refreshes the
    // list through ReloadedFromOutside. What it cannot see is a FILE the draft cites arriving on its
    // own - a picture dropped into the draft folder - so it is asked to run its checks again too
    // (RecheckFiles; code review, 2026-10-04): its "file missing" marks then agree with the draft's row.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // CRT's window while this tab is shown - held so its Activated handler is taken off the
        // same window it was put on. A tab not on screen is detached (a TabControl does that on
        // every switch), which clears it.
        private Window? thisActivationWindow;

        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);

            this.StopFollowingActivation();

            this.thisActivationWindow = TopLevel.GetTopLevel(this) as Window;

            if (this.thisActivationWindow is not null)
            {
                this.thisActivationWindow.Activated += this.OnWindowActivated;
            }

            // Shown: read again. Posted - the tab is still being laid out.
            Dispatcher.UIThread.Post(this.RefreshForOutsideChanges, DispatcherPriority.Background);
        }

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            this.StopFollowingActivation();

            // Leaving the tab: the board on screen catches up with what the table and the component
            // pane saved meanwhile - once the pane's last change is saved too (its Saved catches up,
            // the tab no longer being on screen).
            if (this.IsComponentOpen && this.ComponentPane.HasHeldOrUnsavedChanges())
            {
                this.RunGuarded(this.ComponentPane.SaveNowAsync, "Saving the component pane on leaving the tab");
            }
            else
            {
                this.CatchUpBoardOnScreen();
            }

            base.OnDetachedFromVisualTree(e);
        }

        private void StopFollowingActivation()
        {
            if (this.thisActivationWindow is not null)
            {
                this.thisActivationWindow.Activated -= this.OnWindowActivated;
            }

            this.thisActivationWindow = null;
        }

        private void OnWindowActivated(object? sender, EventArgs e) => this.RefreshForOutsideChanges();

        // The list, and an open table's checks - only while the tab is still shown when the posted
        // call runs.
        private void RefreshForOutsideChanges()
        {
            if (this.thisActivationWindow is not null)
            {
                this.RefreshDrafts();

                if (this.IsTableOpen)
                {
                    this.TableEditor.ReadDraftFileAgain();
                    this.TableEditor.RecheckFiles();
                }

                // The component pane's draft, written outside meanwhile - read again (an undoable
                // step), and the pane follows it.
                if (this.IsComponentOpen)
                {
                    this.ComponentPane.Session?.ReadFileIfChanged(DraftSaveWording.ReadAgainStep);
                }
            }
        }
    }
}
