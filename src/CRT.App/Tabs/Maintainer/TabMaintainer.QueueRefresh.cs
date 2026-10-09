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
    // THE OTHER SCREENS' LISTS COME WITH IT (2026-09-27): the BETA list and the boards are read on
    // the same check, whichever screen is shown, so the buttons' badges stay right - each keeping
    // its own "never disturb what is open" rule (TabMaintainer.Beta.cs, .Boards.cs).
    //
    // OFF SCREEN, THE BADGE'S TWO LISTS ONLY (owner request, 2026-09-30): the tab's badge in CRT's
    // row of tabs counts the queue and the BETA list, so while it can be seen those two are read
    // every minute with the tab on another tab or CRT behind another window - the same background
    // reads, the same rules. QueueRefreshRules.MinuteCheck decides which.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private DispatcherTimer? thisQueueTimer;
        private DateTimeOffset? thisQueueAskedUtc;
        private bool thisBackgroundRefreshRunning;

        // Whether the check running now is the badge's two lists only, and whether a full check was
        // asked for meanwhile - see RefreshQueueInBackgroundAsync.
        private bool thisBackgroundBadgesOnly;
        private bool thisEverythingPending;

        // When EVERYTHING was last read - not just the badge's two lists. Coming back to the tab
        // checks against this, so the first look after a launch spent off screen reads the Boards
        // overview and the drop-down listing, however recently the badge was brought up to date.
        private DateTimeOffset? thisEverythingAskedUtc;

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

        // This tab on screen, whether or not CRT's window is in front - see above.
        private bool IsOnScreen => this.thisQueueCheckWindow is not null;

        private void StartQueueChecks()
        {
            if (this.thisQueueTimer is null)
            {
                this.thisQueueTimer = new DispatcherTimer { Interval = QueueRefreshRules.PollInterval };
                this.thisQueueTimer.Tick += async (_, _) =>
                {
                    switch (QueueRefreshRules.MinuteCheck(this.IsOnScreenAndInFront, this.thisTabBadgeCanBeSeen?.Invoke() == true))
                    {
                        case QueueCheck.Everything:
                            await this.RefreshQueueInBackgroundAsync();
                            break;

                        case QueueCheck.BadgesOnly:
                            await this.RefreshBadgesInBackgroundAsync();

                            // The entry the tab will open on, read again when it changed
                            // (TabMaintainer.Prefetch.cs).
                            await this.PrefetchEntryToOpenAsync();
                            break;
                    }
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

            // Shown: the screen it opens on is decided anew - Boards when nothing waits
            // (TabMaintainer.OpenOnEntry.cs).
            if (this.thisQueueCheckWindow is not null)
                this.BeginOpening();

            if (this.thisQueueCheckWindow is not null)
                this.thisQueueCheckWindow.Activated += this.OnQueueCheckWindowActivated;

            // Coming back to the tab also opens the screen's entry when nothing is open - still on
            // screen by then, or it is not this tab's moment (TabMaintainer.OpenOnEntry.cs) - and
            // reads a Files view marked stale while the tab was away (TabMaintainer.Files.cs).
            //
            // *** A QUEUE ALREADY READ OPENS ITS ENTRY FIRST, THE CHECK AFTER (owner request,
            // 2026-10-02). *** The check reads the Boards overview too, which the server walks
            // both data trees for, and the entry waited for all of it. With the queue the badge
            // read - and the entry itself read ahead (TabMaintainer.Prefetch.cs) - it opens at
            // once; the check then runs under the rules every check keeps for what is open.
            Dispatcher.UIThread.Post(async () =>
            {
                if (this.IsOnScreen && this.thisQueueKnown)
                    this.SelectOnEntry();

                await this.CheckQueueIfDueAsync();

                if (this.IsOnScreen)
                {
                    this.SelectOnEntry();
                    await this.ShowFilesViewIfStaleAsync();
                }
            });
        }

        private void DetachQueueChecks()
        {
            if (this.thisQueueCheckWindow is not null)
                this.thisQueueCheckWindow.Activated -= this.OnQueueCheckWindowActivated;

            this.thisQueueCheckWindow = null;

            // Off screen, nothing is opened - an opening left undecided stays so.
            this.EndOpening();
        }

        private async void OnQueueCheckWindowActivated(object? sender, EventArgs e) => await this.CheckQueueIfDueAsync();

        private async System.Threading.Tasks.Task CheckQueueIfDueAsync()
        {
            if (this.thisQueueTimer is { IsEnabled: true } &&
                QueueRefreshRules.IsDue(this.thisEverythingAskedUtc, DateTimeOffset.UtcNow, QueueRefreshRules.ActivationGap))
            {
                await this.RefreshQueueInBackgroundAsync();
            }
        }

        // ###########################################################################################
        // One background check at a time; an explicit refresh never waits for one.
        //
        // *** A FULL CHECK THAT ARRIVES DURING A BADGES-ONLY ONE IS DEFERRED, NOT DROPPED (code
        // review, 2026-10-01). *** The two share the one guard. A maintainer who switched to the tab
        // while the off-screen badge check was mid-flight asked for everything (the Boards
        // overview, the drop-down listing) - and the guard returned at once with nothing read and
        // nothing remembered, so the screen they had just opened rendered from the last full check
        // until the next minute tick. It is now remembered, and runs the moment the badge check ends.
        // A full check blocked by ANOTHER full check is not remembered: that one reads the same.
        // ###########################################################################################
        private async System.Threading.Tasks.Task RefreshQueueInBackgroundAsync()
        {
            if (this.thisBackgroundRefreshRunning)
            {
                if (this.thisBackgroundBadgesOnly)
                    this.thisEverythingPending = true;

                return;
            }

            this.thisBackgroundRefreshRunning = true;
            this.thisEverythingAskedUtc = DateTimeOffset.UtcNow;

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

        // The badge's two lists, and nothing else - off screen, and at launch
        // (TabMaintainer.Session.cs). Shares the one-at-a-time guard with the full check.
        private async System.Threading.Tasks.Task RefreshBadgesInBackgroundAsync()
        {
            if (this.thisBackgroundRefreshRunning)
                return;

            this.thisBackgroundRefreshRunning = true;
            this.thisBackgroundBadgesOnly = true;

            try
            {
                await this.RefreshQueueAsync(background: true);
                await this.RefreshBetaAsync(background: true);
            }
            finally
            {
                this.thisBackgroundBadgesOnly = false;
                this.thisBackgroundRefreshRunning = false;
            }

            // The full check asked for meanwhile, now that the guard is free.
            //
            // *** AND THE ENTRY THE SCREEN OPENS ON (code review, 2026-10-01). *** The full check is
            // deferred because the tab was just shown - and AttachQueueChecks then ran SelectOnEntry
            // against a queue not read yet, so the first look after launch opened on nothing. The
            // lists are here now; still on screen, it is chosen now.
            if (this.thisEverythingPending)
            {
                this.thisEverythingPending = false;
                await this.RefreshQueueInBackgroundAsync();

                if (this.IsOnScreen)
                    this.SelectOnEntry();
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
