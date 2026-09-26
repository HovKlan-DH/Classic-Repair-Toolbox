using System;
using Avalonia.Threading;
using CRT.Maintainer.Handlers;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // THE QUEUE KEEPS ITSELF UP TO DATE - there is no "Refresh" button (owner request, 2026-09-26:
    // "I would like to make this one obsolete and remove it ... ideally it just updates if there are
    // new stuff, but of course it should not conflict if the maintainer is in the middle of
    // something ... I want as little UI/text as possible").
    //
    // WHEN is QueueRefreshRules': every minute while the window is in front, and on coming back to
    // it. WHAT a background check may change is the point of this file:
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
    // ###########################################################################################
    public partial class MaintainerMain
    {
        private DispatcherTimer? thisQueueTimer;
        private DateTimeOffset? thisQueueAskedUtc;
        private bool thisBackgroundRefreshRunning;

        private void StartQueueChecks()
        {
            if (this.thisQueueTimer is null)
            {
                this.thisQueueTimer = new DispatcherTimer { Interval = QueueRefreshRules.PollInterval };
                this.thisQueueTimer.Tick += async (_, _) =>
                {
                    if (this.IsActive)
                        await this.RefreshQueueInBackgroundAsync();
                };

                this.Activated += async (_, _) =>
                {
                    if (this.thisQueueTimer.IsEnabled &&
                        QueueRefreshRules.IsDue(this.thisQueueAskedUtc, DateTimeOffset.UtcNow, QueueRefreshRules.ActivationGap))
                    {
                        await this.RefreshQueueInBackgroundAsync();
                    }
                };
            }

            this.thisQueueTimer.Start();
        }

        private void StopQueueChecks() => this.thisQueueTimer?.Stop();

        // One background check at a time; an explicit refresh never waits for one.
        private async System.Threading.Tasks.Task RefreshQueueInBackgroundAsync()
        {
            if (this.thisBackgroundRefreshRunning)
                return;

            this.thisBackgroundRefreshRunning = true;

            try
            {
                await this.RefreshQueueAsync(background: true);
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
            this.thisApproveAllowed = false;
            this.SetDecisionButtonsEnabled(false);
            this.ShowDecisionMessage(ReviewDecisionWording.DecidedElsewhere, isError: true);
        }

        // Whether leaving for the sign-in screen now would throw away unsaved table changes.
        private bool WouldLoseTableChanges => this.IsTableOpen && this.TableEditor.HasUnsavedChanges;
    }
}
