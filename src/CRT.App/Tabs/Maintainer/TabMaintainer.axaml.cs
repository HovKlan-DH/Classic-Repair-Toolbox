using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // The Maintainer tab: sign in, then four screens chosen by the tab strip across the top
    // (owner request, 2026-09-27; a strip since 2026-10-01) - Boards, Contributor Submissions,
    // Beta > Prod and Account (Admin until 2026-10-04) - each a list on the left and
    // what is chosen in it on the right. Review is the queue and the selected submission's TABLE
    // (NewContributeStrategy.md Phase 5, tasks 2 and 3); the other three replaced the
    // "Production", "Maintainers" and "Unused files" windows.
    //
    // *** THE TABLE IS THE SUBMISSION VIEW (owner request, 2026-09-26). *** A change summary -
    // section lines, field diffs, the file list, pictures side by side, moved highlights drawn on
    // their schematic - filled this panel until then, with the table a button away. The project
    // owner found it "confusing to look at" and asked for it to go. What the table cannot show
    // (highlights, calibration points, the automatic checks' warnings) stays as a few lines above
    // it - ReviewNotInTable.
    //
    // *** A TAB SINCE 2026-09-29, NOT A WINDOW. *** This was CRT Maintainer's main window, a
    // separate application. It is now a tab in CRT's window, shown when "Enable Maintainer tab" is
    // ticked in the Configuration tab (Main.Maintainer.cs, which also collapses the sidebar and
    // the worklog bar while it is selected). What the window did and a tab cannot, and where each
    // went:
    //   - the remembered sign-in, restored in OnOpened -> at CRT's LAUNCH, quietly, for the tab's
    //     badge (2026-09-30), or else the first time the tab is shown (TabMaintainer.Session.cs).
    //     ReviewSessionStore.Initialise runs at CRT's start-up.
    //   - the minute queue check "while the window is in front" -> everything while this tab is
    //     on screen AND CRT's window is in front; only the badge's two lists (the queue and the
    //     BETA list) while the tab's badge can be seen otherwise (TabMaintainer.QueueRefresh.cs).
    //   - asking about unsaved table changes on Closing -> Main.OnWindowClosing asks, through
    //     HasUnsavedTableEdits / ConfirmLeavingTableAsync (TabMaintainer.Table.cs).
    //   - its own "please wait" overlay -> CRT's window's one (Main.axaml), found up the tree.
    //   - its window placement -> gone; CRT's window remembers its own. The table's
    //     filter (its colour-key pills; "Show changes only" until 2026-10-02) is
    //     UserSettings.MaintainerTableFilter (TabMaintainer.Table.cs).
    //
    // FILE MAP - this file (sign in, the queue, the submission shown), TabMaintainer.Decisions.cs
    // (Approve, Reject and Request changes, and the decision bar), TabMaintainer.Session.cs (the
    // remembered sign-in, restored at launch for the badge or when the tab is first shown, and
    // sharing it with the rest of CRT, and sliding its expiry as the server does), TabMaintainer.QueueItems.cs (the
    // queue list grouped by board, and every list on the tab changed in place), TabMaintainer.QueueRefresh.cs (the queue checking itself - no
    // Refresh button - and the other lists with it, or only the badge's off screen),
    // TabMaintainer.OpenOnEntry.cs (the entry "Contributor Submissions" and "Beta > Prod" open
    // on), TabMaintainer.Prefetch.cs (that entry read ahead while the tab is away, so opening it
    // needs no wait), TabMaintainer.Table.cs (the table, and asking
    // before unsaved changes in it are left), TabMaintainer.Modes.cs (which screen is shown,
    // the buttons' badges, handing the panels the session), TabMaintainer.Beta.cs (the BETA list),
    // TabMaintainer.Boards.cs (the Boards list), TabMaintainer.Account.cs (the Account screen:
    // "My account" and "Server version" for every maintainer and the administrator's padlocked
    // entries, the "Signed in as" line, and the remembered sign-in's name and address read again at
    // launch), TabMaintainer.Invitation.cs ("I have an invitation" on the sign-in screen),
    // TabMaintainer.SubmissionViews.cs (a submission's three views - Board data, Files,
    // Contributor - and which is shown), TabMaintainer.Files.cs (the Files view: the submission's
    // file tree, and the count on its button), TabMaintainer.Contributor.cs (the Contributor
    // view: who sent it and everything they sent before) and TabMaintainer.UpdateRequired.cs
    // ("CRT has to be updated" over the whole tab, when the server turns this version of CRT away).
    // The right-hand panels of the other three screens are controls of their own: BetaView,
    // BoardDetailView (a board's five views - Board data, Files, Contributor, Maintainer, Statistics -
    // with its own file map, and BoardPlacementView), and under Account MyAccountView (the name,
    // address, password and Sign out) and the administrator's MaintainerPoolView, BoardOrderView,
    // UnusedFilesView, RebuildManifestsView, BoardDeletionView (with DeleteBoardWindow, its
    // confirmation), ApiUsageView, OscilloscopeModelsView and DataResetView. A
    // board's table can hold a change not sent, so HasUnsavedTableEdits
    // and sign-out ask about it too. Every async event handler of the tab and of those panels runs
    // its work through HandlerGuard.cs, so an exception is logged and said, never the end of CRT.
    //
    // *** THE LOGIC IS IN Handlers/, NOT HERE. *** How a queue row reads is ReviewQueueDisplay's;
    // what the table cannot show is ReviewNotInTable's;
    // what the server answered is ReviewApiParser's. All are unit tested. This file resolves
    // controls, calls one of them, and shows the result - which is why the screens that decide
    // whether a change gets looked at have real coverage rather than being verified by eye.
    //
    // *** WHAT CHANGED IS NOT SHOWN IN THE QUEUE, AND CANNOT BE. *** Working it out needs the
    // published board AND the submitted manifest; the queue endpoint returns neither, because
    // doing so would mean loading two full BoardData per queued row on the server for a list the
    // maintainer scrolls past. So the queue shows what the CONTRIBUTOR said, and the rest arrives
    // when a submission is opened. See ReviewQueueDisplay's header.
    // ###########################################################################################
    public partial class TabMaintainer : UserControl
    {
        private readonly List<ReviewQueueRow> thisQueue = [];

        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // Which submission the decision buttons act on. Held rather than re-read off the list at
        // press time, so a decision cannot be sent for a row the selection moved to while a
        // request was in flight.
        private long? thisSelectedId;

        // True while code - not the maintainer - is moving the queue's selection (keeping it across a
        // refresh, or putting it back on the table's submission), so the handler does not treat it
        // as a new choice.
        private bool thisSuppressQueueSelection;

        // True while the maintainer is being asked about the table's unsaved changes after choosing
        // another submission - a refresh then leaves the list's selection alone (ApplyQueue).
        private bool thisLeavingTable;

        public TabMaintainer()
        {
            this.InitializeComponent();
            this.WireTable();
            this.WireAccountScreen();
            this.ShowNotRememberedNote(ReviewSessionStore.IsSupported);
            this.ShowTestServerNote(AppConfig.UsesTestServer);
        }

        // ###########################################################################################
        // The queue and the other lists, read under the "please wait" overlay (2026-09-28) - after
        // signing in, when the tab is first shown on a remembered session. The queue's own minute check
        // reads them WITHOUT it: nobody pressed anything, so nobody is waiting.
        // ###########################################################################################
        private async Task ReadListsAsync()
        {
            this.thisEverythingAskedUtc = DateTimeOffset.UtcNow;

            // Signing in, or the tab first shown on a remembered session, opens it: on Boards when
            // nothing waits (TabMaintainer.OpenOnEntry.cs).
            this.BeginOpening();

            bool answered = await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, async () =>
            {
                await this.RefreshQueueAsync();
                await this.RefreshOtherListsAsync();
            });

            if (!answered)
                this.ShowQueueMessage(WaitWording.NoAnswer, isError: true);

            // Signed in, or the tab first shown on a remembered session: Boards when nothing waits,
            // else the queue on the submission looked at last, or its first (TabMaintainer
            // .OpenOnEntry.cs).
            this.SelectOnEntry();
        }

        // -----------------------------------------------------------------------------------
        // Sign in
        // -----------------------------------------------------------------------------------

        // Every async handler runs through HandlerGuard: what throws is logged and said, never the
        // end of CRT (code review, 2026-10-13).
        private async void OnSignInClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.SignInAsync, this.ShowHandlerFailure);

        // Enter in the password box signs in. Unlike CRT's own destructive dialogs - where Enter
        // deliberately CANCELS - this is a safe, repeatable action, and making someone reach for
        // the mouse after typing a password is the kind of friction that has no upside.
        private async void OnPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await HandlerGuard.RunAsync(this.SignInAsync, this.ShowHandlerFailure);
            }
        }

        private async Task SignInAsync()
        {
            var email = this.FindControl<TextBox>("EmailTextBox");
            var password = this.FindControl<TextBox>("PasswordTextBox");
            var button = this.FindControl<Button>("SignInButton");

            if (email is null || password is null || button is null)
                return;

            // *** NOT ASKED FOR. *** The Maintainer tab talks to one service and always will, so the
            // address is a constant rather than a field somebody retypes on every launch - and
            // gets wrong with a trailing slash or a missing scheme.
            string address = ReviewApiRoutes.DefaultBaseAddress;

            // Disabled for the duration, so a second click cannot start a second sign-in while
            // the first is in flight - which would race two sessions and leave whichever finished
            // last in place, not necessarily the one the maintainer waited for.
            button.IsEnabled = false;
            this.ShowSignInMessage(null);

            try
            {
                this.thisClient?.Dispose();

                var client = new ReviewApiClient(address);
                this.UseClient(client);
                string enteredEmail = email.Text?.Trim() ?? string.Empty;
                string enteredPassword = password.Text ?? string.Empty;

                // A timeout is an ordinary failure here: signing in changes nothing worth checking.
                ReviewApiResult<ReviewSession> result = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.SigningIn,
                    token => client.LoginAsync(enteredEmail, enteredPassword, token));

                if (!result.IsOk)
                {
                    this.ShowSignInMessage(result.Message);
                    return;
                }

                this.UseSession(result.Value);

                // Fresh from the server: nothing for the launch's account read to bring up to date
                // (TabMaintainer.Account.cs).
                this.thisAccountRead = true;

                // *** REMEMBERED HERE, ONCE, ON A SUCCESSFUL SIGN-IN. *** Storing it anywhere else
                // would mean storing a token that has not been proved to work. The store itself
                // declines on any machine where it cannot be encrypted, so this call is safe to
                // make unconditionally - see ReviewSessionStore.Remember.
                ReviewSessionStore.Remember(this.thisSession!);

                // The password is cleared the moment it is no longer needed. It is not a secret
                // this window has any further use for, and a populated password box left on a
                // signed-in machine is a credential sitting in plain sight.
                password.Text = string.Empty;

                this.ShowQueuePanel();
                await this.ReadListsAsync();
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        // ###########################################################################################
        // The sign-in panel's one message line.
        //
        // isError defaults TRUE because almost every use is a failure - but the password-reset
        // answer is good news, and colouring that red would read as a failure to somebody who has
        // just asked for help.
        // ###########################################################################################
        private void ShowSignInMessage(string? message, bool isError = true)
        {
            var text = this.FindControl<TextBlock>("SignInMessageText");

            if (text is null)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = isError ? Brushes.IndianRed : Brushes.SeaGreen;
        }

        // ###########################################################################################
        // "I forgot my password".
        //
        // *** THE ANSWER IS THE SAME WHETHER OR NOT THE ADDRESS EXISTS. *** The server always
        // answers 202 for exactly that reason, and this must not try to be more helpful:
        // distinguishing the two would turn the sign-in screen into a tool for discovering which
        // addresses hold maintainer accounts - the accounts that can publish to every user.
        //
        // The reset itself happens through the emailed link, not in this app. Building a
        // reset-token screen here would be a second place that flow lives, for an action taken
        // once in a career.
        // ###########################################################################################
        private async void OnForgotPasswordClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.ForgotPasswordAsync, this.ShowHandlerFailure);

        private async Task ForgotPasswordAsync()
        {
            var email = this.FindControl<TextBox>("EmailTextBox");
            var button = this.FindControl<Button>("ForgotPasswordButton");

            if (email is null || button is null)
                return;

            string address = email.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(address))
            {
                // The one thing worth checking locally - the server cannot email nothing, and
                // saying so costs no round trip.
                this.ShowSignInMessage("Enter your email address first, then ask for a reset.");
                return;
            }

            button.IsEnabled = false;

            try
            {
                // A client of its own: the shared one is created by signing in, and this runs
                // before anybody has.
                using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);

                ReviewApiResult<string> result = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.AskingForResetCode,
                    token => client.ForgotPasswordAsync(address, token));

                // A code may still be mailed after the window stopped waiting - so the place to
                // paste it opens then too.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowSignInMessage(MaintainerWaitWording.ResetCodeNoAnswer);
                    this.ShowResetPanel();
                    return;
                }

                this.ShowSignInMessage(
                    result.IsOk ? result.Value : result.Message,
                    isError: !result.IsOk);

                // *** THE PANEL OPENS ON SUCCESS, WHICH IS THE ANSWER FOR ANY ADDRESS. *** The
                // server deliberately answers 202 whether or not the address is known, so this
                // reveals somewhere to paste a code that may never arrive. That is correct: a
                // panel that appeared only for registered addresses would turn this screen into a
                // way of discovering which addresses hold maintainer accounts - the exact
                // enumeration oracle the neutral 202 exists to prevent.
                if (result.IsOk)
                    this.ShowResetPanel();
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        // ###########################################################################################
        // Reveals the "paste your code" panel and puts the cursor in it.
        //
        // Idempotent: asking for a second code while the panel is already open must not clear a
        // half-typed password, so nothing here resets the fields.
        // ###########################################################################################
        private void ShowResetPanel()
        {
            var panel = this.FindControl<StackPanel>("ResetPanel");

            if (panel is null)
                return;

            panel.IsVisible = true;

            this.FindControl<TextBox>("ResetCodeTextBox")?.Focus();
        }

        // Enter in the password box submits, matching the sign-in box above it. Safe and
        // repeatable, so there is no reason to make somebody reach for the mouse.
        private async void OnNewPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await HandlerGuard.RunAsync(this.ResetPasswordAsync, this.ShowHandlerFailure);
            }
        }

        private async void OnSetPasswordClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.ResetPasswordAsync, this.ShowHandlerFailure);

        // ###########################################################################################
        // Completes the reset: the code from the mail, plus the chosen password.
        //
        // *** ON SUCCESS THE PANEL CLOSES AND THE PASSWORD BOX IS FILLED IN, NOT THE QUEUE
        // OPENED. *** Resetting a password is not signing in - the server issues no session here,
        // and pretending otherwise would mean holding the password to log in with it behind the
        // person's back. Filling the sign-in box means one click from "password set" to working,
        // without inventing a session that was never granted.
        // ###########################################################################################
        private async Task ResetPasswordAsync()
        {
            var code = this.FindControl<TextBox>("ResetCodeTextBox");
            var password = this.FindControl<TextBox>("NewPasswordTextBox");
            var button = this.FindControl<Button>("SetPasswordButton");

            if (code is null || password is null || button is null)
                return;

            string enteredCode = code.Text?.Trim() ?? string.Empty;
            string enteredPassword = password.Text ?? string.Empty;

            // Both checked locally, because the round trip can say nothing useful about an empty
            // box that this cannot say immediately.
            if (string.IsNullOrWhiteSpace(enteredCode))
            {
                this.ShowSignInMessage("Paste the code from the email first.");
                return;
            }

            if (string.IsNullOrEmpty(enteredPassword))
            {
                this.ShowSignInMessage("Type the password you want to use.");
                return;
            }

            button.IsEnabled = false;

            try
            {
                // Its own client, like the forgot-password request: the shared one is created by
                // signing in, and this runs before anybody has.
                using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);

                ReviewApiResult<string> result = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.SettingPassword,
                    token => client.ResetPasswordAsync(enteredCode, enteredPassword, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowSignInMessage(MaintainerWaitWording.PasswordNoAnswer);
                    return;
                }

                this.ShowSignInMessage(
                    result.IsOk ? result.Value : result.Message,
                    isError: !result.IsOk);

                if (!result.IsOk)
                    return;

                // The code is spent and the fields have served their purpose. Clearing them stops
                // a second press resubmitting a code the server has already consumed, which would
                // answer "that link has already been used" and read as a failure.
                code.Text = string.Empty;
                password.Text = string.Empty;

                var panel = this.FindControl<StackPanel>("ResetPanel");

                if (panel is not null)
                    panel.IsVisible = false;

                var signInPassword = this.FindControl<TextBox>("PasswordTextBox");

                if (signInPassword is not null)
                {
                    signInPassword.Text = enteredPassword;
                    signInPassword.Focus();
                }
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private void ShowQueuePanel()
        {
            var signIn = this.FindControl<StackPanel>("SignInPanel");
            var queue = this.FindControl<Grid>("QueuePanel");

            if (signIn is null || queue is null)
                return;

            signIn.IsVisible = false;
            queue.IsVisible = true;

            // A remembered sign-in can be restored while the invitation is on screen; the next
            // sign-out must find the sign-in form, not that (TabMaintainer.Invitation.cs).
            this.CloseInvitationView();

            this.ShowSignedInAs();

            this.InitialiseScreens();
            this.ApplyModeVisibility();
            this.StartQueueChecks();
        }

        // -----------------------------------------------------------------------------------
        // The queue
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // Asks for the queue and shows it. `background` is the queue's own check
        // (TabMaintainer.QueueRefresh.cs): it updates the list and leaves the open submission alone.
        // Without it - after a decision here, a changed pool - the open submission is read again,
        // table included unless it holds unsaved changes.
        // ###########################################################################################
        private async Task RefreshQueueAsync(bool background = false)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            this.thisQueueAskedUtc = DateTimeOffset.UtcNow;

            // *** AN EXPIRED SESSION IS CAUGHT BEFORE THE REQUEST, not after the 401. *** The
            // server would refuse it anyway, but "your session has expired" is a better answer
            // than "the server answered 401", and it is the same answer either way - so say the
            // true one.
            if (!this.thisSession.IsUsableAt(DateTimeOffset.UtcNow))
            {
                // ###########################################################################################
                // *** A DEAD SESSION UNDER UNSAVED EDITS STOPS THE CHECKS AND SAYS WHAT WILL HAPPEN
                // TO THEM (code review, 2026-09-26). *** The panel is not switched while a table
                // has unsaved changes a check nobody asked for would throw away - but the timer
                // used to keep firing, repainting the same message every minute with no route that
                // kept the edits: saving needs the dead session too. So the checks stop, and the
                // message says the edits cannot be saved and what signing in again would cost.
                // Signing in again is the maintainer's own decision, taken from the Sign out button.
                // ###########################################################################################
                if (background && this.WouldLoseAnyTableChanges)
                {
                    this.KeepTablesAfterSessionEnded(expired: true);
                    return;
                }

                this.ShowQueueMessage("Your session has expired. Sign in again.", isError: true);
                this.ShowSignInPanel();

                return;
            }

            ReviewQueueRow? before = this.thisSelectedId is long selected
                ? this.thisQueue.FirstOrDefault(row => row.Id == selected)
                : null;

            ReviewSession asked = this.thisSession;

            ReviewApiResult<ReviewQueueResponse> result =
                await this.thisClient.GetQueueAsync(asked);

            // ###########################################################################################
            // *** AN ANSWER FOR A SESSION NO LONGER IN USE IS DROPPED (code review, 2026-11-02). *** The
            // minute check's answer landing after the sign-in screen refilled the hidden lists, set "known"
            // and the administrator's flag, and redrew the tab's badge - which then stayed, the timer being
            // stopped - and the next account started with the old queue and the administrator's entries.
            // ###########################################################################################
            if (!this.IsSessionInUse(asked))
                return;

            if (!result.IsOk)
            {
                // ###########################################################################################
                // *** A 401 UNDER UNSAVED EDITS IS THE EXPIRED SESSION'S CASE (code review, 2026-10-13). ***
                // The server's "no" - signed out elsewhere, revoked, the account locked - kept the
                // panel for the edits but left the checks running, so the same "Sign in to continue."
                // came back every minute with no way out said. It now stops them and says what the
                // local expiry above says.
                // ###########################################################################################
                if (result.Failure == ReviewApiFailure.NotSignedIn && background && this.WouldLoseAnyTableChanges)
                {
                    this.KeepTablesAfterSessionEnded(expired: false);
                    return;
                }

                // A 403 is the server's answer, not "could not ask" - see EmptyListsNoLongerAllowed.
                if (result.Failure == ReviewApiFailure.NotPermitted && !this.WouldLoseAnyTableChanges)
                    this.EmptyListsNoLongerAllowed();

                // *** THE LIST IS LEFT ALONE ON A FAILURE. *** Clearing it would turn "I could not
                // ask" into "nothing is waiting", which is the one wrong answer a review queue
                // must never give - a backlog would go unnoticed for as long as nobody checked.
                this.ShowQueueMessage(result.Message, isError: true);

                if (result.Failure == ReviewApiFailure.NotSignedIn)
                    this.ShowSignInPanel();

                return;
            }

            if (this.ApplyQueueResponse(result.Value!, background) is not ReviewQueueRow kept)
                return;

            if (background)
            {
                // Its detail only, and only when its row changed - never its table. Or when no detail of
                // it is held (code review, 2026-11-03): a read that failed left no decision bar, and with
                // the row as it was nothing read it again - one submission in the queue has no other to
                // click and back, so it could not be decided until CRT restarted.
                if (!Equals(before, kept) || this.thisShownDetail?.Submission.Id != kept.Id)
                    await this.LoadSubmissionAsync(kept);

                return;
            }

            // Still selected: read again, since whatever prompted the refresh (a changed pool, a
            // refused decision, another maintainer's change) may have changed what it says. Its
            // table too - unless it holds unsaved changes, which a reload would throw away.
            await Task.WhenAll(
                this.LoadSubmissionAsync(kept),
                this.WouldLoseTableChanges ? Task.CompletedTask : this.LoadTableAsync(message: null));
        }

        // ###########################################################################################
        // Puts the server's queue answer on screen, and returns the submission still selected.
        //
        // *** THE ADMINISTRATOR'S ENTRIES ON THE ACCOUNT SCREEN ARE OFFERED ONLY WHEN THE SERVER SAYS
        // THIS ACCOUNT IS AN ADMINISTRATOR *** (TabMaintainer.Account.cs). The server refuses
        // everyone else regardless; an entry that could only ever be refused would be clutter.
        // ###########################################################################################
        internal ReviewQueueRow? ApplyQueueResponse(ReviewQueueResponse response, bool background = false)
        {
            ArgumentNullException.ThrowIfNull(response);

            this.thisQueue.Clear();
            this.thisQueue.AddRange(response.Submissions);

            this.SetAdministrator(response.IsAdministrator);
            this.thisQueueKnown = true;

            // The administrator's oscilloscope models waiting count as work too (code review, 2026-10-29).
            this.thisOscilloscopeModelsWaiting = response.IsAdministrator ? Math.Max(0, response.OscilloscopeModelsWaiting) : 0;

            ReviewQueueRow? kept = this.ApplyQueue(background);

            // The Review button counts the boards waiting for this account - the list just built.
            this.UpdateModeBadges();

            return kept;
        }

        // ###########################################################################################
        // Drops back to the sign-in screen and FORGETS the remembered session.
        //
        // *** THIS IS THE ONE FUNNEL FOR EVERY "NO LONGER SIGNED IN" PATH, and that is why the
        // forget lives here rather than at each call site. *** It is reached from an expired
        // session, from a 401 (which is what a revoked session, a signed-out session or a LOCKED
        // ACCOUNT looks like from here), and from signing out. All of them mean the stored token
        // is worthless, and a call site that forgot to clear it would leave the app retrying a
        // dead credential on every launch - the exact "it keeps asking me to sign in and does not
        // work" complaint that is hardest to diagnose.
        //
        // The project owner named locking and invalidation specifically. Neither is something this
        // app can detect on its own; both arrive as a 401 from the server, which lands here.
        // ###########################################################################################
        private void ShowSignInPanel()
        {
            var signIn = this.FindControl<StackPanel>("SignInPanel");
            var queue = this.FindControl<Grid>("QueuePanel");

            // Cleared BEFORE the early return: a null control is a markup problem, and it must not
            // leave a rejected token sitting on disk to be retried next launch.
            this.UseSession(null);
            this.thisQueueKnown = false;
            this.thisOscilloscopeModelsWaiting = 0;
            ReviewSessionStore.Forget();

            // The session is over, so an open table could not be saved any more.
            this.CloseTableNow();
            this.StopQueueChecks();

            // ###########################################################################################
            // *** THE PREVIOUS SESSION'S QUEUE AND SELECTION GO WITH IT (code review, 2026-09-26). ***
            // Left standing, the next sign-in - possibly as a DIFFERENT account on the same machine -
            // found the old `thisSelectedId` in the fresh queue and silently re-selected it, then
            // re-applied the old detail's badges, which were judged for the old account. Nothing
            // opened its table, because only a real selection change does that, so the panel showed
            // a decision bar for a submission the screen said was not selected.
            // ###########################################################################################
            this.thisQueue.Clear();
            this.thisSelectedId = null;
            this.thisShownDetail = null;

            // And every comment typed for it - the next account must not find one ready to send.
            this.thisDecisionCommentBySubmission.Clear();

            if (this.DecisionCommentBox is TextBox comment)
                comment.Text = string.Empty;

            // What the badges count, too - a 401 reaches here without ApplyQueue, and the tab's own
            // badge in CRT's row of tabs would otherwise go on counting the last account's queue.
            this.thisQueueEntries.Clear();

            // The other screens' lists and panels too, and the next sign-in starts on Review.
            this.ResetScreens();

            if (signIn is null || queue is null)
                return;

            queue.IsVisible = false;
            signIn.IsVisible = true;
        }

        // ###########################################################################################
        // Signing out deliberately - "Sign out" under "My account" (MyAccountView; a button under
        // the lists until 2026-10-04).
        //
        // Revokes server-side FIRST, while the token is still in hand, then clears locally. The
        // order matters: clearing first would leave nothing to revoke with, and the session would
        // stay live for the rest of its sliding lifetime.
        //
        // The server call cannot fail in a way worth reporting (see ReviewApiClient.LogoutAsync),
        // so there is no error path here - signing out of your own machine must always work, and
        // it is most wanted precisely when the server cannot be reached.
        // ###########################################################################################
        internal async Task SignOutAsync()
        {
            // Unsaved table changes are asked about first - after signing out nothing can save them.
            // Both tables: the submission's, and a board's on the Boards screen (2026-10-03).
            if (!await this.CloseTableAsync() || !await this.BoardDetail.MayLeaveTableAsync())
                return;

            // Under the overlay, and signed out locally whatever the answer - even none at all.
            if (this.thisClient is { } client && this.thisSession is { } session)
                await ServerWait.RunAsync(this, MaintainerWaitWording.SigningOut, () => client.LogoutAsync(session));

            this.thisQueue.Clear();
            this.ApplyQueue();
            this.ShowSignInPanel();
            this.ShowSignInMessage("You are signed out.", isError: false);
        }

        // ###########################################################################################
        // Puts the queue on screen, KEEPING the selected submission selected when it is still in it
        // (2026-09-26). Replacing the list used to drop the selection and blank the panel on every
        // refresh - which, with the table open in that panel, would have thrown away the table and
        // any unsaved change in it. Returns the submission still selected, or null when it left the
        // queue (decided, here or by somebody else) - the panel, and a table on it, then close.
        // ###########################################################################################
        private ReviewQueueRow? ApplyQueue(bool background = false)
        {
            var list = this.FindControl<ListBox>("QueueList");

            if (list is null)
                return null;

            DateTimeOffset now = DateTimeOffset.UtcNow;

            long? keep = this.thisSelectedId;
            ReviewQueueRow? kept = keep is null ? null : this.thisQueue.FirstOrDefault(row => row.Id == keep.Value);

            this.thisSuppressQueueSelection = true;

            try
            {
                // Grouped by board, with a heading per board, changed in place - see
                // TabMaintainer.QueueItems.cs.
                this.ShowQueueItems(list, this.thisQueue, now);

                // ###########################################################################################
                // *** NOT WHILE THE "LEAVE THE TABLE?" PROMPT IS UP (code review, 2026-10-13). *** The
                // maintainer has clicked another submission and is being asked about the table's
                // changes; putting the selection back on the table's submission here split the list
                // from the panel - after Discard the panel opened the clicked one while the list
                // showed the old. The list is left as the maintainer clicked it, and the prompt's
                // answer decides (OnQueueSelectionChanged).
                // ###########################################################################################
                if (!this.thisLeavingTable)
                    this.SelectQueueRow(kept?.Id);
            }
            finally
            {
                this.thisSuppressQueueSelection = false;
            }

            // An empty queue is a perfectly good answer and says so plainly - distinct from the
            // error wording used above, which the maintainer must be able to tell apart.
            this.ShowQueueMessage(
                this.thisQueue.Count == 0 ? "Nothing waiting for review." : null,
                isError: false);

            if (kept is null)
            {
                // Decided by someone else while open, found by the queue's own check: it stays on
                // screen, undecidable - see TabMaintainer.QueueRefresh.cs.
                if (background && this.thisSelectedId is not null)
                {
                    this.ShowDecidedElsewhere();
                    return null;
                }

                this.ShowSubmission(null);
                return null;
            }

            // The rebuilt entry says what the opened submission itself said.
            if (this.thisShownDetail is { } detail && detail.Submission.Id == kept.Id)
            {
                this.UpdateQueueEntry(
                    kept,
                    detail.Changes?.IsNewBoard ?? kept.IsNewBoard,
                    ReviewQueueDisplay.AwaitsYou(detail.CanPublish, detail.Approval));
            }

            return kept;
        }

        private void ShowQueueMessage(string? message, bool isError)
        {
            var text = this.FindControl<TextBlock>("QueueMessageText");

            if (text is null)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = isError ? Brushes.IndianRed : Brushes.Gray;
        }

        private async void OnQueueSelectionChanged(object? sender, SelectionChangedEventArgs e) =>
            await HandlerGuard.RunAsync(this.QueueSelectionChangedAsync, this.ShowHandlerFailure);

        private async Task QueueSelectionChangedAsync()
        {
            if (this.thisSuppressQueueSelection)
                return;

            // By ITEM: a board's heading sits before its submissions, so a position in the list is
            // not a position in the queue.
            ReviewQueueRow? chosen = this.SelectedQueueRow;

            // *** LEAVING A TABLE WITH UNSAVED CHANGES ASKS FIRST. *** It sits in this panel, so
            // choosing another submission replaces it. Cancel puts the selection back on it.
            if (this.IsTableOpen && chosen?.Id != this.thisTableRow!.Id)
            {
                bool mayLeave;

                this.thisLeavingTable = true;

                try
                {
                    mayLeave = await this.CloseTableAsync();
                }
                finally
                {
                    this.thisLeavingTable = false;
                }

                if (!mayLeave)
                {
                    this.ReselectTableRow();
                    return;
                }

                // ###########################################################################################
                // *** THE SUBMISSION SELECTED NOW, NOT THE ONE CLICKED (code review, 2026-10-13). *** A
                // minute check can run while the prompt waits for an answer, and can take the clicked
                // submission out of the queue - the list then selects nothing, and so does the panel.
                // ###########################################################################################
                chosen = this.SelectedQueueRow;
            }

            if (chosen is null)
            {
                this.ShowSubmission(null);
                return;
            }

            ReviewQueueRow row = chosen;

            // What the screen opens on next time, this run or the next (TabMaintainer.OpenOnEntry.cs).
            this.RememberSubmission(row.Id);

            // The header is shown IMMEDIATELY from what the queue already knows, and the rest fills
            // in when it arrives. Waiting for the requests before drawing anything would leave the
            // panel blank on every click over a slow link, which reads as the app having lost the
            // selection.
            this.ShowSubmission(row);

            // Read ahead while the tab was not on screen (TabMaintainer.Prefetch.cs): shown at once,
            // with no wait.
            if (await this.TryOpenPrefetchedAsync(row))
                return;

            // The submission (who must approve, what approving removes, what the table cannot show)
            // and its table, side by side - under the overlay (2026-09-28), so a slow link reads as
            // a wait rather than as a panel that never fills in.
            bool answered = await ServerWait.RunAsync(
                this,
                MaintainerWaitWording.OpeningSubmission(row.BoardId),
                () => Task.WhenAll(this.LoadSubmissionAsync(row), this.OpenTableAsync(row)));

            if (!answered && this.SelectedQueueRow?.Id == row.Id)
                this.ShowNotInTable([new ReviewNoteLine(WaitWording.NoAnswer, ReviewNoteKind.Error)]);
        }

        // ###########################################################################################
        // Fetches one submission and shows what the server says about it.
        //
        // *** THE ANSWER IS DISCARDED IF THE SELECTION MOVED WHILE IT WAS IN FLIGHT. *** A maintainer
        // arrowing down the queue starts a request per row, and they can finish out of order.
        // Without this check a slow earlier response would overwrite a faster later one, and the
        // panel would describe a different submission from the one highlighted - which is exactly
        // the state in which somebody approves the wrong thing.
        // ###########################################################################################
        private async Task LoadSubmissionAsync(ReviewQueueRow row)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiResult<ReviewSubmissionDetail> result =
                await this.thisClient.GetSubmissionAsync(this.thisSession, row.Id);

            if (this.SelectedQueueRow?.Id != row.Id)
                return;

            if (!result.IsOk)
            {
                this.ShowContributorHistory(null);
                this.ShowNotInTable([new ReviewNoteLine(result.Message, ReviewNoteKind.Error)]);
                return;
            }

            this.ShowDetail(result.Value!);
        }

        // -----------------------------------------------------------------------------------
        // One submission
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // Makes the selected submission the one the panel is about - its details are in its queue
        // row. Null empties the panel.
        // ###########################################################################################
        private void ShowSubmission(ReviewQueueRow? row)
        {
            if (this.FindControl<TextBlock>("NoSubmissionText") is TextBlock none)
                none.IsVisible = row is null;

            // Another submission - or none - opens on Board data, with nothing of the last one's
            // files or contributor in the other two views (TabMaintainer.SubmissionViews.cs), and
            // with its own decision comment rather than the last one's.
            if (row is null || row.Id != this.thisSelectedId)
            {
                this.ResetSubmissionViews();
                this.SwapDecisionComment(row?.Id);

                // ###########################################################################################
                // *** THE DECISION BAR GOES WITH THE DETAIL IT WAS SHOWN FROM (code review, 2026-11-02). ***
                // It stayed live while the next submission's detail was on its way - and for good when
                // that read failed - so Approve sent the approval for B with A's removal list and no
                // expected version, past the version check and ApprovalGate's reason. A publish cannot be
                // undone. ShowDetail brings it back, read ahead (TryOpenPrefetchedAsync) or not.
                // ###########################################################################################
                this.ShowDecisionPanel(visible: false, canPublish: false, approval: null);
                this.ShowRemovalNote(null);
            }

            this.ShowNotInTable([]);
            this.ShowBeforeApproving([]);

            if (row is null)
            {
                this.thisSelectedId = null;
                this.thisShownDetail = null;

                // Nothing selected: the submission a table was open on has left the queue, or the
                // queue is empty. Its changes could not be saved now, so it closes.
                this.CloseTableNow();
                return;
            }

            this.thisSelectedId = row.Id;

            // ###########################################################################################
            // *** THE PREVIOUS SUBMISSION'S DETAIL IS DROPPED HERE, NOT WHEN THE NEW ONE ARRIVES
            // (code review, 2026-09-26). *** The table's file card reads `thisShownDetail` for the
            // submitted files' hashes, and the table usually opens before the detail answer comes
            // back - so between the two, hovering a file cell resolved it against the PREVIOUS
            // submission's files and showed its bytes labelled as this one's, then cached them.
            // Two submissions of the same board sit next to each other in the queue, which is
            // exactly when that is easiest to do and hardest to notice.
            //
            // The source treats "no detail yet" as "no submitted file at this path" and falls back
            // to the published side, which is the safe reading while the answer is in flight.
            // ###########################################################################################
            if (this.thisShownDetail is { } shown && shown.Submission.Id != row.Id)
                this.thisShownDetail = null;
        }

        // ###########################################################################################
        // Shows what the server says about the submission: the badges, who must approve, what
        // approving removes, and what the table cannot show.
        // ###########################################################################################
        internal void ShowDetail(ReviewSubmissionDetail detail)
        {
            ArgumentNullException.ThrowIfNull(detail);

            // A detail read again may carry other files (a save, another maintainer's change): the
            // Files view reads its tree anew the next time it is shown. The same detail shown again
            // (ReapplyApprovalGate) changes nothing there.
            if (!ReferenceEquals(this.thisShownDetail, detail))
                this.FilesViewIsStale();

            // The table's file preview reads the submitted files' hashes from it.
            this.thisShownDetail = detail;

            // The table can be changed only by an account that may publish it (TabMaintainer.Table.cs).
            this.ApplyTableReadOnly();

            // A detail shown is a submission that can be decided - see ReapplyApprovalGate.
            this.thisDecisionsClosed = false;

            // *** canPublish COMES FROM THE SERVER, never from anything this app worked out. ***
            // It is the same answer the server will enforce when the button is pressed, so the
            // screen cannot promise something the API then refuses.
            this.ShowDecisionPanel(visible: true, canPublish: detail.CanPublish, approval: detail.Approval);
            this.SetDecisionButtonsEnabled(true);
            this.ShowDecisionMessage(null, isError: false);
            this.ShowRemovalNote(detail.Removals, detail.SubmittedFiles.Select(file => file.Path));

            // The list says what the submission itself says - see TabMaintainer.QueueItems.cs.
            this.UpdateQueueEntry(
                detail.Submission,
                detail.Changes?.IsNewBoard ?? detail.Submission.IsNewBoard,
                ReviewQueueDisplay.AwaitsYou(detail.CanPublish, detail.Approval));

            if (this.FindControl<TextBlock>("AmendmentText") is TextBlock amended)
            {
                string? line = ReviewTableWording.AmendedLine(detail.Amendment);
                amended.Text = line ?? string.Empty;
                amended.IsVisible = line is not null;
            }

            // The other two views' facts come with the detail: how many files the submission changes
            // (the Files button's count - the one place that still says a file was replaced under
            // its own name before anything is opened) and the contributor's record.
            this.ShowFilesCount(SubmissionViews.ChangingFiles(detail.SubmittedFiles, detail.Removals));
            this.ShowContributorHistory(detail.Contributor);

            // *** WHAT THE TABLE CANNOT SHOW IS NOT OPTIONAL. *** Highlights and calibration points
            // publish with the approval, and the automatic checks' warnings are what the machine
            // already worked out - see ReviewNotInTable.
            this.ShowNotInTable(ReviewNotInTable.Lines(detail.Changes, detail.Findings));

            // ###########################################################################################
            // *** WHAT TO SETTLE BEFORE APPROVING, ABOVE ALL THREE VIEWS (code review, 2026-10-01). ***
            // These two were the first lines above the TABLE, inside the Board data view - so with
            // Files or Contributor on screen, Approve was greyed out with no reason anywhere. They
            // sit above the view switch now (BeforeApprovingPanel), still before anything else.
            // ###########################################################################################
            List<ReviewNoteLine> beforeApproving = [];

            // Approve is OFF, with the reason, for a new board with no place in the drop-down lists
            // yet, or a board with an earlier submission still in BETA (2026-09-27) - said before
            // the maintainer reads the submission, not after pressing Approve.
            string? blocked = ApprovalGate.Blocked(detail.Submission.BoardId, this.thisListing, this.thisBeta);
            this.thisAppliedGate = blocked;

            if (blocked is not null)
            {
                beforeApproving.Add(new ReviewNoteLine(blocked, ReviewNoteKind.Warning));
                this.BlockApproval();
            }

            // *** THE CONTRIBUTOR DISCARDED THE DRAFT (owner request, 2026-09-28). *** They may have
            // changed their mind, and the maintainer should ask before approving. See
            // DraftDiscardWording.
            if (detail.Submission.DraftDiscardedUtc is DateTimeOffset discarded)
            {
                beforeApproving.Add(new ReviewNoteLine(
                    DraftDiscardWording.SubmissionWarning(discarded, detail.Submission.ContactEmail),
                    ReviewNoteKind.Warning));
            }

            // *** MADE FROM OLDER DATA THAN BETA HAS NOW (code review, 2026-10-13). *** The draft is
            // a whole copy of the board, so what was published since its base shows - and would
            // be published - as removed. Said, not blocked: the maintainer compares and decides.
            if (detail.Submission.BasedOnOlderRevision == true)
                beforeApproving.Add(new ReviewNoteLine(ApprovalWording.MadeFromOlderData, ReviewNoteKind.Warning));

            // *** A NEWER ONE FROM THE SAME CONTRIBUTOR WAITS BESIDE IT (owner decision, 2026-10-13). ***
            // An amended or half-approved submission is not replaced by the newer one, so the two
            // wait side by side. Said, not blocked: this one may be the right one to publish.
            if (detail.Submission.NewerSubmissionId is not null || detail.Submission.NewerSubmissionUtc is not null)
            {
                beforeApproving.Add(new ReviewNoteLine(
                    ApprovalWording.NewerFromSameContributor(detail.Submission.NewerSubmissionId, detail.Submission.NewerSubmissionUtc),
                    ReviewNoteKind.Warning));
            }

            this.ShowBeforeApproving(beforeApproving);
            this.OfferPlacement(detail.Submission.BoardId);

            // A shared-file replacement's other boards, on the approval status line (2026-10-13).
            this.ReadFilesForApproval(detail);
        }

        // ###########################################################################################
        // *** THE GATE FOLLOWS THE LISTS IT READS (code review, 2026-09-27). *** ApprovalGate reads
        // the drop-down listing and the "Beta > Prod" list, which are replaced on the minute check
        // and after every decision - but it was applied only when a detail was SHOWN, and a detail
        // is read again only when its queue row changed. So Approve stayed off, with a stale reason,
        // after another maintainer promoted the board; and at sign-in a detail that arrived before
        // the lists was never gated at all. Called whenever either list is replaced; the detail is
        // shown again only when the gate's answer changed. Never while a decision is being sent, or
        // for a submission decided elsewhere, whose decisions stay off whatever the lists say.
        // ###########################################################################################
        private void ReapplyApprovalGate()
        {
            if (this.thisShownDetail is not { } detail ||
                this.thisSelectedId != detail.Submission.Id ||
                this.thisDecisionsClosed ||
                this.thisDecisionInFlight)
            {
                return;
            }

            string? blocked = ApprovalGate.Blocked(detail.Submission.BoardId, this.thisListing, this.thisBeta);

            if (!string.Equals(blocked, this.thisAppliedGate, StringComparison.Ordinal))
                this.ShowDetail(detail);
        }

        // The gate's reason as last applied to the open detail - null for none.
        private string? thisAppliedGate;

        // The short lines above the table - nothing at all when there is nothing to say.
        private void ShowNotInTable(IReadOnlyList<ReviewNoteLine> lines) => this.ShowNoteLines("NotInTablePanel", lines);

        // ###########################################################################################
        // What to settle before approving, above the three views, as THE AMBER NOTICE PANEL (owner
        // request, 2026-10-09: "this should look the same as the previous highlighted panel ... The
        // UI should have a uniform and consistent look") - App.axaml's Border.Notice, the Boards
        // screen's ReadOnlyNotice's look: one NoticeText line per reason, the panel gone with none.
        // Not bold since 2026-10-30 (owner request, of the Contribute tab's notice: "please do not bold
        // the text") - the style's weight; a person named in a line keeps the bold of its own run.
        // ###########################################################################################
        private void ShowBeforeApproving(IReadOnlyList<ReviewNoteLine> lines)
        {
            if (this.FindControl<StackPanel>("BeforeApprovingPanel") is not StackPanel panel)
                return;

            panel.Children.Clear();

            foreach (ReviewNoteLine line in lines)
            {
                var block = new TextBlock();
                block.Classes.Add("NoticeText");

                TabMaintainer.ShowLine(block, line);
                panel.Children.Add(block);
            }

            if (this.FindControl<Border>("BeforeApprovingNotice") is Border notice)
                notice.IsVisible = lines.Count > 0;
        }

        // One panel of note lines, coloured by kind - hidden when there is nothing to say.
        private void ShowNoteLines(string panelName, IReadOnlyList<ReviewNoteLine> lines)
        {
            if (this.FindControl<StackPanel>(panelName) is not StackPanel panel)
                return;

            panel.Children.Clear();

            foreach (ReviewNoteLine line in lines)
            {
                var block = new TextBlock
                {
                    FontSize = 12,
                    TextWrapping = TextWrapping.Wrap,
                    // A change under its section's heading ("Component highlights have [2] changes:").
                    Margin = new Avalonia.Thickness(16 * line.Indent, 0, 0, 0),
                    FontWeight = line.Kind == ReviewNoteKind.Change ? FontWeight.Normal : FontWeight.SemiBold,
                    Foreground = line.Kind switch
                    {
                        ReviewNoteKind.Error => Brushes.IndianRed,
                        ReviewNoteKind.Warning => Brushes.DarkOrange,
                        _ => Brushes.SteelBlue
                    }
                };

                TabMaintainer.ShowLine(block, line);
                panel.Children.Add(block);
            }

            panel.IsVisible = lines.Count > 0;
        }

        // ###########################################################################################
        // Puts one line into a TextBlock. A count's number is bold ("[2]"), so a line with one is
        // built of runs - its Text is then null (a TextBlock cannot mix weights within one Text).
        // Whatever the block showed before is replaced.
        // ###########################################################################################
        private static void ShowLine(TextBlock block, ReviewNoteLine line) =>
            TabMaintainer.ShowRuns(block, line.Runs.Select(run => (run.Text, run.IsCount)).ToList());

        // A status line with its to-do pieces in bold - see StatusPart (BoardsDisplay.cs).
        internal static void ShowParts(TextBlock block, IReadOnlyList<StatusPart> parts) =>
            TabMaintainer.ShowRuns(block, parts.Select(part => (part.Text, part.IsToDo)).ToList());

        // A line of counts, each number in bold as the review notes show theirs ("[12]").
        internal static void ShowCounts(TextBlock block, IReadOnlyList<ReviewNoteRun> runs) =>
            TabMaintainer.ShowRuns(block, runs.Select(run => (run.Text, run.IsCount)).ToList());

        // Text in pieces, the marked ones bold - Handlers/Theme/BoldRuns, which Main's notice cards
        // write through too. With a bold piece Text is NULL - read one through TextOf.
        private static void ShowRuns(TextBlock block, IReadOnlyList<(string Text, bool IsBold)> runs) =>
            Handlers.Theming.BoldRuns.Show(block, runs);

        // What a block says, whether it holds plain Text or bold pieces in Inlines - for tests.
        internal static string TextOf(TextBlock block) => Handlers.Theming.BoldRuns.TextOf(block);
    }
}
