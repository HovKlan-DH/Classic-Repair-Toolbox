using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // A SUBMISSION'S DECISIONS - Approve, Reject and Request changes. This part owns:
    //
    //   - sending one under the window's wait, and what is said of it (DecideAsync);
    //   - the decision bar: its buttons, which of them may be on, the approval and removal lines beside
    //     them (ShowDecisionPanel, ShowRemovalNote, BlockApproval);
    //   - each submission's own comment (SwapDecisionComment).
    //
    // Which submission they act on is thisSelectedId, and what they act on its detail (ShowSubmission,
    // TabMaintainer.axaml.cs). Moved out of that file as it passed the 1,500 lines a part may hold
    // (code review, 2026-11-02).
    // ###########################################################################################
    public partial class TabMaintainer
    {
        // True once the open submission was decided by someone else (ShowDecidedElsewhere).
        private bool thisDecisionsClosed;

        // True while a decision is on its way to the server.
        private bool thisDecisionInFlight;

        // How many decisions were sent - an answer read before the latest is older than it
        // (TabMaintainer.Prefetch.cs).
        private int thisDecisionsSent;

        // Turns Approve off for a reason the server's own answer does not cover (ApprovalGate). The
        // reason is the first line above the view switch; the button carries no tooltip (see
        // ShowDecisionPanel for why). The server refuses regardless.
        private void BlockApproval()
        {
            this.thisApproveAllowed = false;

            if (this.FindControl<Button>("ApproveButton") is Button approve)
                approve.IsEnabled = false;
        }

        // -----------------------------------------------------------------------------------
        // The three decisions (task 5)
        // -----------------------------------------------------------------------------------

        private async void OnApproveClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.DecideAsync(ReviewDecisionKind.Approve), this.ShowDecisionFailure);

        private async void OnRejectClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.DecideAsync(ReviewDecisionKind.Reject), this.ShowDecisionFailure);

        private async void OnRequestChangesClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.DecideAsync(ReviewDecisionKind.RequestChanges), this.ShowDecisionFailure);

        // A decision that failed in CRT itself: beside the decision, or in the queue's line when the
        // submission has left it (ShowDecisionOutcome).
        private void ShowDecisionFailure(string message) => this.ShowDecisionOutcome(message, isError: true);

        // ###########################################################################################
        // Sends one decision and reports what happened.
        //
        // *** THE BUTTONS ARE DISABLED FOR THE DURATION, and on APPROVE that is not cosmetic. ***
        // Approving publishes irreversibly; a second click while the first is in flight would be a
        // second publish, and the server's own interlock would refuse it as a conflict - which
        // reads to the maintainer as an error on something that actually worked.
        //
        // *** A COMMENT IS REQUIRED LOCALLY FOR TWO OF THE THREE, checked here as well as on the
        // server. *** Not because the client is trusted - the server refuses regardless - but
        // because a round trip to be told "write more" is a worse experience than being told
        // before anything is sent. The RULE is the server's; this only avoids the trip.
        // ###########################################################################################
        private async Task DecideAsync(ReviewDecisionKind kind)
        {
            // A cell still in its editor is a change too - the grid does not commit when a button takes the click.
            this.TableEditor.CommitCellEdit();

            // A decision is about what the server holds; unsaved table changes would not be in it.
            if (this.IsTableOpen && this.TableEditor.HasUnsavedChanges)
            {
                this.ShowDecisionMessage(ReviewTableWording.SaveTableBeforeDeciding, isError: true);
                return;
            }

            if (this.thisClient is null || this.thisSession is null || this.thisSelectedId is null)
                return;

            // A decision is on what was shown of THIS submission (code review, 2026-11-02): its removal
            // list, its version and its approval state all come from its detail - see ShowSubmission.
            if (this.thisShownDetail is not { } shown || shown.Submission.Id != this.thisSelectedId)
                return;

            string comment = this.FindControl<TextBox>("DecisionCommentTextBox")?.Text?.Trim() ?? string.Empty;

            // Optional for Approve - sent when written, and quoted in the contributor's mail (code
            // review, 2026-10-13: it used to be dropped). The other two are the contributor's only
            // feedback, so they need one.
            if (kind != ReviewDecisionKind.Approve &&
                !ReviewDecisionWording.IsUsableComment(comment, out string problem))
            {
                this.ShowDecisionMessage(problem, isError: true);
                return;
            }

            this.thisDecisionInFlight = true;
            this.thisDecisionsSent++;
            this.SetDecisionButtonsEnabled(false);
            this.ShowDecisionMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            long id = this.thisSelectedId.Value;
            IReadOnlyList<string> removals = this.thisShownRemovals;
            string? stateBefore = shown.Submission.State;

            // The version on screen (2026-10-13): the server refuses a decision on a submission
            // another maintainer has changed since - and the queue then reads it again.
            int? expectedVersion = shown.Submission.AmendmentVersion;

            string waiting = ReviewDecisionWording.Waiting(kind, shown.Approval, shown.Submission.BoardId);

            // ###########################################################################################
            // *** THE WHOLE WINDOW WAITS, not a line under the table (owner report, 2026-09-28). ***
            // A small green "Working..." was all an approval showed while it published, and the
            // screen read as hung. The window is HELD for the decision AND the lists read after it,
            // each step under its own two-minute limit, and let go however it ends.
            // ###########################################################################################
            try
            {
                await BusyOverlay.HoldAsync(this, waiting, async () =>
                {
                    ReviewApiResult<ReviewDecisionResult> result = await ServerWait.CallAsync(this, waiting, token => kind switch
                    {
                        ReviewDecisionKind.Approve => client.ApproveAsync(session, id, removals, token, expectedVersion, comment),
                        ReviewDecisionKind.Reject => client.RejectAsync(session, id, comment, token, expectedVersion),
                        _ => client.RequestChangesAsync(session, id, comment, token, expectedVersion)
                    });

                    // ###########################################################################################
                    // *** NO ANSWER IN TWO MINUTES IS NOT "IT FAILED" (2026-09-28). *** An approval goes
                    // on publishing after the window stops waiting, so the submission is read again
                    // and the maintainer is told what it now is - never invited to press Approve a
                    // second time over a publish that worked.
                    // ###########################################################################################
                    if (result.Failure == ReviewApiFailure.TimedOut)
                    {
                        ReviewApiResult<ReviewSubmissionDetail> now = await ServerWait.CallAsync(
                            this, WaitWording.Checking, token => client.GetSubmissionAsync(session, id, token));

                        string afterwards = MaintainerWaitWording.DecisionAfterTimeout(
                            kind, stateBefore, now.IsOk ? now.Value!.Submission.State : null);

                        // Decided - by this or by somebody else meanwhile: the comment has been
                        // used. Nothing moved: it stays, for sending again.
                        if (now.IsOk && !string.Equals(now.Value!.Submission.State, stateBefore, StringComparison.Ordinal))
                            this.ForgetDecisionComment(id);

                        await this.ReadListsAfterDecisionAsync();
                        this.ShowDecisionOutcome(afterwards, isError: !now.IsOk);
                        return;
                    }

                    if (!result.IsOk)
                    {
                        this.ShowDecisionMessage(result.Message, isError: true);

                        // ###########################################################################################
                        // A CONFLICT means somebody else decided it first, so the queue on screen is
                        // already stale. Refreshing is the only useful next step and doing it for them
                        // beats telling them to.
                        //
                        // *** WHERE THE REASON IS RE-SHOWN DEPENDS ON WHAT SURVIVED THE REFRESH (code
                        // review, 2026-09-26). *** The submission has usually left the queue, and then
                        // the refresh selects nothing and HIDES the decision panel - which is where
                        // DecisionMessageText lives, so writing the reason back into it left the
                        // maintainer looking at "Select a submission" with no explanation at all. It
                        // goes to the queue's own message instead, which stays on screen.
                        // ###########################################################################################
                        if (result.Failure == ReviewApiFailure.Conflict)
                        {
                            await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, () => this.RefreshQueueAsync());
                            this.ShowDecisionOutcome(result.Message, isError: true);
                        }

                        return;
                    }

                    this.ShowDecisionMessage(
                        ReviewDecisionWording.Describe(kind, result.Value!),
                        isError: false);

                    // Sent, so the box is empty again - before the lists are read, which move the
                    // selection off the decided submission and would otherwise keep the comment for it.
                    this.ForgetDecisionComment(id);

                    // The submission has left the queue, so the list on screen is now wrong - and an
                    // approval has put a board into BETA, which the BETA button counts.
                    await this.ReadListsAfterDecisionAsync();
                });
            }
            finally
            {
                this.thisDecisionInFlight = false;
                this.SetDecisionButtonsEnabled(true);

                // The lists were read while the decision was in flight, when the gate is left alone.
                this.ReapplyApprovalGate();
            }
        }

        // The queue and the other lists, read again after a decision - under the overlay the
        // decision is already holding.
        private Task ReadListsAfterDecisionAsync() =>
            ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, async () =>
            {
                await this.RefreshQueueAsync();
                await this.RefreshOtherListsAsync();
            });

        // A decision's outcome where it will still be seen: beside the decision when the submission
        // is still open, in the queue's own line when it has left the queue (see the Conflict note).
        private void ShowDecisionOutcome(string message, bool isError)
        {
            if (this.thisSelectedId is null)
                this.ShowQueueMessage(message, isError);
            else
                this.ShowDecisionMessage(message, isError);
        }

        private void SetDecisionButtonsEnabled(bool enabled)
        {
            foreach (string name in new[] { "ApproveButton", "RejectButton", "RequestChangesButton" })
            {
                var button = this.FindControl<Button>(name);

                // Approve comes back on only if this account may still approve - see
                // thisApproveAllowed.
                if (button is not null)
                    button.IsEnabled = enabled && (name != "ApproveButton" || this.thisApproveAllowed);
            }
        }

        // ###########################################################################################
        // *** THE DECISION COMMENT BELONGS TO ONE SUBMISSION (owner report, 2026-10-02: "I can see my
        // last rejection comment in the textarea field. This field should be blanked when
        // submitted/rejected"). *** The box was never emptied, so the comment written for one
        // submission sat ready to be sent - with one click - to the next contributor. Now a
        // decision that went through empties it, and moving to another submission puts away what
        // was typed for this one and shows that one's own (nothing, unless something was typed for
        // it earlier), so a half-written comment is still there on coming back. For as long as
        // CRT runs.
        // ###########################################################################################
        private readonly Dictionary<long, string> thisDecisionCommentBySubmission = [];

        private TextBox? DecisionCommentBox => this.FindControl<TextBox>("DecisionCommentTextBox");

        // Called as the panel moves from thisSelectedId to `next` (null for none).
        private void SwapDecisionComment(long? next)
        {
            if (this.DecisionCommentBox is not TextBox box)
                return;

            if (this.thisSelectedId is long current)
            {
                if (string.IsNullOrWhiteSpace(box.Text))
                    this.thisDecisionCommentBySubmission.Remove(current);
                else
                    this.thisDecisionCommentBySubmission[current] = box.Text;
            }

            box.Text = next is long id && this.thisDecisionCommentBySubmission.TryGetValue(id, out string? kept)
                ? kept
                : string.Empty;
        }

        // A decision on `submissionId` went through: its comment has been used.
        private void ForgetDecisionComment(long submissionId)
        {
            this.thisDecisionCommentBySubmission.Remove(submissionId);

            if (this.thisSelectedId == submissionId && this.DecisionCommentBox is TextBox box)
                box.Text = string.Empty;
        }

        private void ShowDecisionMessage(string? message, bool isError)
        {
            WindowMessage.ShowOutcome(this.FindControl<TextBlock>("DecisionMessageText"), message, isError);
        }

        // ###########################################################################################
        // Shows or hides the decision bar, and says whether this account may publish.
        //
        // *** THE APPROVE BUTTON IS DISABLED RATHER THAN HIDDEN for an account that may not
        // publish this submission. *** Since Phase 6 that should not happen - the queue is already
        // filtered to what the account may decide - but an older server, or a pool changed while
        // the queue was open, can still answer no. A missing button reads as a broken screen; a
        // disabled one beside a sentence saying why does not. The server enforces it regardless -
        // this is explanation, never enforcement.
        // ###########################################################################################
        private void ShowDecisionPanel(bool visible, bool canPublish, ApprovalStatus? approval)
        {
            var panel = this.FindControl<StackPanel>("DecisionPanel");

            if (panel is not null)
                panel.IsVisible = visible;

            // The Approve button says what pressing it does - publish, or record one of the two
            // approvals a shared-file change needs - and is off when this account's part is done.
            // All from the server's ApprovalStatus; the server enforces it regardless.
            this.thisApproveAllowed = canPublish && ApprovalWording.CanApprove(approval);

            var approve = this.FindControl<Button>("ApproveButton");

            // ###########################################################################################
            // *** NO TOOLTIP ON APPROVE - IT ATE THE CLICK (owner report, 2026-09-30: "Sometimes I do
            // feel that I need to click the "Approve and publish to BETA" multiple times for it to
            // react ... I once needed to click it 3 times"). *** Avalonia puts a tooltip 20 px below
            // the pointer; at the bottom of the window there is no room, so it flips ABOVE - and the
            // 20 px offset is kept, which moves it back down OVER the pointer. The next click then
            // lands on the tooltip (its own window in CRT), not on the button; the tooltip closes,
            // and after a moment's hover opens again. What it said is in the status line instead,
            // and the button's own text says what it does. TabMaintainerDecisionClickTests
            // reproduces it. Every other tooltip in CRT is safe from it since the same day: each
            // opens at its control's edge, never under the pointer (ToolTipPlacement).
            // ###########################################################################################
            if (approve is not null)
            {
                approve.IsEnabled = this.thisApproveAllowed;
                approve.Content = ApprovalWording.ApproveButton(approval, "BETA");
            }

            var status = this.FindControl<TextBlock>("ApprovalStatusText");

            if (status is not null)
            {
                // With the other boards a replaced shared file reaches, once the file tree has named
                // them (2026-10-13; TabMaintainer.Files.cs).
                string? line = ApprovalWording.WithAlsoUsedBy(ApprovalWording.StatusLine(approval, canPublish), this.thisAlsoUsedByLine);
                status.Text = line ?? string.Empty;
                status.IsVisible = visible && line is not null;
            }
        }

        // ###########################################################################################
        // The files approving the submission on screen would remove, as the server listed them -
        // said beside the button and sent back with the approval, so what goes is what was shown.
        // `kept`: the names the submission's board lists - see FileRemovalWording.Unused.
        // ###########################################################################################
        private void ShowRemovalNote(FileRemovalPreview? removals, IEnumerable<string>? kept = null)
        {
            this.thisShownRemovals = removals?.Files ?? [];

            if (this.FindControl<TextBlock>("RemovalStatusText") is TextBlock note)
            {
                string? text = FileRemovalWording.ApproveNote(removals, kept);
                note.Text = text ?? string.Empty;
                note.IsVisible = text is not null;
            }
        }

        private IReadOnlyList<string> thisShownRemovals = [];

        // Whether the Approve button may be on for the submission on screen - re-applied after a
        // decision re-enables the buttons, so it never comes back on for an account whose part of
        // a two-person approval is already done.
        private bool thisApproveAllowed;
    }
}
