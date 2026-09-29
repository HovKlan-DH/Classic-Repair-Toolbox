using System;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE QUEUE KEEPS ITSELF UP TO DATE - there is no "Refresh" button (owner request, 2026-09-26:
    // "I would like to make this one obsolete and remove it ... ideally it just updates if there are
    // new stuff, but of course it should not conflict if the maintainer is in the middle of
    // something ... I want as little UI/text as possible").
    //
    // WHEN is QueueRefreshRules': every minute while this tab is on screen with CRT's window in
    // front, and on coming back to either. WHAT a background check may change is the point of this file:
    //
    //   - THE LIST, freely: new submissions appear, decided ones go, badges and waits update. The
    //     selection stays on the same submission (ApplyQueue keeps it by id).
    //   - NEVER THE OPEN SUBMISSION'S TABLE. A check does not reload it, so an edit in progress - or
    //     just the place the maintainer is reading - is not disturbed.
    //   - The open submission's detail (who must approve, the decision bar) is read again only when
    //     its queue row CHANGED - another approver acted, say. That touches no table.
    //   - DECIDED ELSEWHERE: when the open submission has left the queue, it stays on screen with
    //     its decision buttons off and one line saying so, until the maintainer picks another. An
    //     explicit refresh (after a decision here) closes it, as before.
    //   - SIGNED OUT with unsaved table changes: the sign-in screen would throw them away, so the
    //     queue says so and waits; without unsaved changes it goes to the sign-in screen as always.
    //
    // THE OTHER SCREENS' LISTS COME WITH IT (2026-09-27): the BETA list and the systems are read on
    // the same check, whichever screen is shown, so the buttons' badges stay right - each keeping
    // its own "never disturb what is open" rule (TabMaintainer.Beta.cs, .Systems.cs).
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private DispatcherTimer? thisQueueTimer;
        private DateTimeOffset? thisQueueAskedUtc;
        private bool thisBackgroundRefreshRunning;

        // CRT's window while this tab is attached to it - held so its Activated handler is taken
        // off the same window it was put on.
        private Window? thisQueueCheckWindow;

        // ###########################################################################################
        // *** "IN FRONT" MEANS THIS TAB ON SCREEN AND CRT'S WINDOW ACTIVE (2026-09-29). *** As its
        // own application this checked while its window was active. As a tab, the window being
        // active is not enough: CRT is in front far more often than the Maintainer tab is on screen,
        // and each check extends the session - so a maintainer working on a schematic would be kept
        // signed in, and the server asked every minute, for a queue nobody is looking at. A tab
        // not on screen is DETACHED (a TabControl does that on every switch), which clears
        // thisQueueCheckWindow - so that is the test.
        // ###########################################################################################
        private bool IsOnScreenAndInFront => this.thisQueueCheckWindow is { IsActive: true };

        private void StartQueueChecks()
        {
            if (this.thisQueueTimer is null)
            {
                this.thisQueueTimer = new DispatcherTimer { Interval = QueueRefreshRules.PollInterval };
                this.thisQueueTimer.Tick += async (_, _) =>
                {
                    if (this.IsOnScreenAndInFront)
                        await this.RefreshQueueInBackgroundAsync();
                };
            }

            this.thisQueueTimer.Start();
        }

        private void StopQueueChecks() => this.thisQueueTimer?.Stop();

        // ###########################################################################################
        // Coming back to the tab, or to CRT's window while the tab is shown, checks at once - the
        // part the window's Activated handler played - unless it asked within ActivationGap. Both
        // are guarded by the timer running: signed out, the timer is stopped and nothing is asked.
        // Called from TabMaintainer.Session.cs's attach and detach.
        // ###########################################################################################
        private void AttachQueueChecks()
        {
            this.DetachQueueChecks();

            this.thisQueueCheckWindow = TopLevel.GetTopLevel(this) as Window;

            if (this.thisQueueCheckWindow is not null)
                this.thisQueueCheckWindow.Activated += this.OnQueueCheckWindowActivated;

            Dispatcher.UIThread.Post(async () => await this.CheckQueueIfDueAsync());
        }

        private void DetachQueueChecks()
        {
            if (this.thisQueueCheckWindow is not null)
                this.thisQueueCheckWindow.Activated -= this.OnQueueCheckWindowActivated;

            this.thisQueueCheckWindow = null;
        }

        private async void OnQueueCheckWindowActivated(object? sender, EventArgs e) => await this.CheckQueueIfDueAsync();

        private async System.Threading.Tasks.Task CheckQueueIfDueAsync()
        {
            if (this.thisQueueTimer is { IsEnabled: true } &&
                QueueRefreshRules.IsDue(this.thisQueueAskedUtc, DateTimeOffset.UtcNow, QueueRefreshRules.ActivationGap))
            {
                await this.RefreshQueueInBackgroundAsync();
            }
        }

        // One background check at a time; an explicit refresh never waits for one.
        private async System.Threading.Tasks.Task RefreshQueueInBackgroundAsync()
        {
            if (this.thisBackgroundRefreshRunning)
                return;

            this.thisBackgroundRefreshRunning = true;

            try
            {
                await this.RefreshQueueAsync(background: true);
                await this.RefreshOtherListsAsync(minuteCheck: true);
            }
            finally
            {
                this.thisBackgroundRefreshRunning = false;
            }
        }

        // The open submission is no longer in the queue - decided by someone else. It stays on
        // screen, and cannot be decided again.
        private void ShowDecidedElsewhere()
        {
            this.thisDecisionsClosed = true;
            this.thisApproveAllowed = false;
            this.SetDecisionButtonsEnabled(false);
            this.ShowDecisionMessage(ReviewDecisionWording.DecidedElsewhere, isError: true);
        }

        // Whether leaving for the sign-in screen now would throw away unsaved table changes.
        private bool WouldLoseTableChanges => this.IsTableOpen && this.TableEditor.HasUnsavedChanges;
    }
}
