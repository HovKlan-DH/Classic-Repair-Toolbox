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
    // The Maintainer tab: sign in, then four screens chosen by the buttons at the top left
    // (owner request, 2026-09-27) - Review, BETA, Systems and Admin - each a list on the left and
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
    //   - the remembered sign-in, restored in OnOpened -> the FIRST time the tab is shown
    //     (TabMaintainer.Session.cs). ReviewSessionStore.Initialise runs at CRT's start-up.
    //   - the minute queue check "while the window is in front" -> while this tab is on screen
    //     AND CRT's window is in front (TabMaintainer.QueueRefresh.cs).
    //   - asking about unsaved table changes on Closing -> Main.OnWindowClosing asks, through
    //     HasUnsavedTableEdits / ConfirmLeavingTableAsync (TabMaintainer.Table.cs).
    //   - its own "please wait" overlay -> CRT's window's one (Main.axaml), found up the tree.
    //   - its window placement -> gone; CRT's window remembers its own. "Show changes only" is
    //     UserSettings.MaintainerShowChangesOnly (TabMaintainer.Table.cs).
    //
    // FILE MAP - this file (sign in, the queue, the decisions), TabMaintainer.Session.cs (the
    // remembered sign-in, restored when the tab is first shown), TabMaintainer.QueueItems.cs (the
    // queue list grouped by board), TabMaintainer.QueueRefresh.cs (the queue checking itself - no
    // Refresh button - and the other lists with it), TabMaintainer.Table.cs (the table, and asking
    // before unsaved changes in it are left), TabMaintainer.Modes.cs (which screen is shown,
    // the buttons' badges, handing the panels the session), TabMaintainer.Beta.cs (the BETA list),
    // TabMaintainer.Systems.cs (the Systems list), TabMaintainer.Admin.cs (the Admin list),
    // TabMaintainer.Invitation.cs ("I have an invitation" on the sign-in screen) and
    // TabMaintainer.Files.cs (a submission's file tree, in its own window).
    // The right-hand panels of the last three screens are controls of their own: BetaView,
    // SystemView (with its SystemView.Maintainers part) and UnusedFilesView.
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

        public TabMaintainer()
        {
            this.InitializeComponent();
            this.WireTable();
        }

        // ###########################################################################################
        // The queue and the other lists, read under the "please wait" overlay (2026-09-28) - after
        // signing in, when the tab is first shown on a remembered session. The queue's own minute check
        // reads them WITHOUT it: nobody pressed anything, so nobody is waiting.
        // ###########################################################################################
        private async Task ReadListsAsync()
        {
            bool answered = await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, async () =>
            {
                await this.RefreshQueueAsync();
                await this.RefreshOtherListsAsync();
            });

            if (!answered)
                this.ShowQueueMessage(WaitWording.NoAnswer, isError: true);
        }

        // -----------------------------------------------------------------------------------
        // Sign in
        // -----------------------------------------------------------------------------------

        private async void OnSignInClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.SignInAsync();

        // Enter in the password box signs in. Unlike CRT's own destructive dialogs - where Enter
        // deliberately CANCELS - this is a safe, repeatable action, and making someone reach for
        // the mouse after typing a password is the kind of friction that has no upside.
        private async void OnPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await this.SignInAsync();
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
                this.thisClient = new ReviewApiClient(address);

                ReviewApiClient client = this.thisClient;
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

                this.thisSession = result.Value;

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
        private async void OnForgotPasswordClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
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
                await this.ResetPasswordAsync();
            }
        }

        private async void OnSetPasswordClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.ResetPasswordAsync();

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
            var signedInAs = this.FindControl<TextBlock>("SignedInAsText");

            if (signIn is null || queue is null)
                return;

            signIn.IsVisible = false;
            queue.IsVisible = true;

            if (signedInAs is not null && this.thisSession is not null)
            {
                // Named, because somebody with both a maintainer and an administrator account
                // needs to know which one they are acting as before they publish anything.
                signedInAs.Text = $"{this.thisSession.DisplayName} ({this.thisSession.Email})";
            }

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
                if (background && this.WouldLoseTableChanges)
                {
                    this.StopQueueChecks();

                    this.ShowQueueMessage(
                        "Your session has expired, so the changes in the table below can no longer be saved. " +
                        "Copy anything you need, then sign out and in again to carry on.",
                        isError: true);

                    return;
                }

                this.ShowQueueMessage("Your session has expired. Sign in again.", isError: true);
                this.ShowSignInPanel();

                return;
            }

            ReviewQueueRow? before = this.thisSelectedId is long selected
                ? this.thisQueue.FirstOrDefault(row => row.Id == selected)
                : null;

            ReviewApiResult<ReviewQueueResponse> result =
                await this.thisClient.GetQueueAsync(this.thisSession);

            if (!result.IsOk)
            {
                // *** THE LIST IS LEFT ALONE ON A FAILURE. *** Clearing it would turn "I could not
                // ask" into "nothing is waiting", which is the one wrong answer a review queue
                // must never give - a backlog would go unnoticed for as long as nobody checked.
                this.ShowQueueMessage(result.Message, isError: true);

                if (result.Failure == ReviewApiFailure.NotSignedIn && !(background && this.WouldLoseTableChanges))
                    this.ShowSignInPanel();

                return;
            }

            if (this.ApplyQueueResponse(result.Value!, background) is not ReviewQueueRow kept)
                return;

            if (background)
            {
                // Its detail only, and only when its row changed - never its table.
                if (!Equals(before, kept))
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
        // *** THE ADMIN SCREEN IS OFFERED ONLY WHEN THE SERVER SAYS THIS ACCOUNT IS AN
        // ADMINISTRATOR *** (TabMaintainer.Modes.cs). The server refuses everyone else regardless; a
        // button that could only ever be refused would be clutter.
        // ###########################################################################################
        internal ReviewQueueRow? ApplyQueueResponse(ReviewQueueResponse response, bool background = false)
        {
            ArgumentNullException.ThrowIfNull(response);

            this.thisQueue.Clear();
            this.thisQueue.AddRange(response.Submissions);

            this.SetAdministrator(response.IsAdministrator);

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
            this.thisSession = null;
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

            // The other screens' lists and panels too, and the next sign-in starts on Review.
            this.ResetScreens();

            if (signIn is null || queue is null)
                return;

            queue.IsVisible = false;
            signIn.IsVisible = true;
        }

        // ###########################################################################################
        // Signing out deliberately.
        //
        // Revokes server-side FIRST, while the token is still in hand, then clears locally. The
        // order matters: clearing first would leave nothing to revoke with, and the session would
        // stay live for the rest of its sliding lifetime.
        //
        // The server call cannot fail in a way worth reporting (see ReviewApiClient.LogoutAsync),
        // so there is no error path here - signing out of your own machine must always work, and
        // it is most wanted precisely when the server cannot be reached.
        // ###########################################################################################
        private async void OnSignOutClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            // Unsaved table changes are asked about first - after signing out nothing can save them.
            if (!await this.CloseTableAsync())
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
                // Grouped by board, with a heading per board - see TabMaintainer.QueueItems.cs.
                list.ItemsSource = this.BuildQueueItems(this.thisQueue, now);
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
                    detail.Changes?.IsNewSystem ?? kept.IsNewSystem,
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

        private async void OnQueueSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisSuppressQueueSelection)
                return;

            // By ITEM: a board's heading sits before its submissions, so a position in the list is
            // not a position in the queue.
            ReviewQueueRow? chosen = this.SelectedQueueRow;

            // *** LEAVING A TABLE WITH UNSAVED CHANGES ASKS FIRST. *** It sits in this panel, so
            // choosing another submission replaces it. Cancel puts the selection back on it.
            if (this.IsTableOpen && chosen?.Id != this.thisTableRow!.Id && !await this.CloseTableAsync())
            {
                this.ReselectTableRow();
                return;
            }

            if (chosen is null)
            {
                this.ShowSubmission(null);
                return;
            }

            ReviewQueueRow row = chosen;

            // The header is shown IMMEDIATELY from what the queue already knows, and the rest fills
            // in when it arrives. Waiting for the requests before drawing anything would leave the
            // panel blank on every click over a slow link, which reads as the app having lost the
            // selection.
            this.ShowSubmission(row);

            // The submission (who must approve, what approving removes, what the table cannot show)
            // and its table, side by side - under the overlay (2026-09-28), so a slow link reads as
            // a wait rather than as a panel that never fills in.
            bool answered = await ServerWait.RunAsync(
                this,
                MaintainerWaitWording.OpeningSubmission(row.SystemId),
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
                this.ShowContributor(null);
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

            this.ShowContributor(null);
            this.ShowNotInTable([]);

            if (row is null)
            {
                this.thisSelectedId = null;
                this.thisShownDetail = null;
                this.ShowDecisionPanel(visible: false, canPublish: false, approval: null);
                this.ShowRemovalNote(null);

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

            // The table's file preview reads the submitted files' hashes from it.
            this.thisShownDetail = detail;

            // A detail shown is a submission that can be decided - see ReapplyApprovalGate.
            this.thisDecisionsClosed = false;

            // *** canPublish COMES FROM THE SERVER, never from anything this app worked out. ***
            // It is the same answer the server will enforce when the button is pressed, so the
            // screen cannot promise something the API then refuses.
            this.ShowDecisionPanel(visible: true, canPublish: detail.CanPublish, approval: detail.Approval);
            this.SetDecisionButtonsEnabled(true);
            this.ShowDecisionMessage(null, isError: false);
            this.ShowRemovalNote(detail.Removals);

            // The list says what the submission itself says - see TabMaintainer.QueueItems.cs.
            this.UpdateQueueEntry(
                detail.Submission,
                detail.Changes?.IsNewSystem ?? detail.Submission.IsNewSystem,
                ReviewQueueDisplay.AwaitsYou(detail.CanPublish, detail.Approval));

            if (this.FindControl<TextBlock>("AmendmentText") is TextBlock amended)
            {
                string? line = ReviewTableWording.AmendedLine(detail.Amendment);
                amended.Text = line ?? string.Empty;
                amended.IsVisible = line is not null;
            }

            // *** WHAT THE TABLE CANNOT SHOW IS NOT OPTIONAL. *** Highlights and calibration points
            // publish with the approval, and the automatic checks' warnings are what the machine
            // already worked out - see ReviewNotInTable.
            this.ShowContributor(detail.Contributor);
            IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(detail.Changes, detail.Findings, detail.SubmittedFiles);

            // ###########################################################################################
            // *** THE CONTRIBUTOR DISCARDED THEIR OWN DRAFT (owner request, 2026-09-28). *** Said
            // above the table, before anything else is read - they may have changed their mind, and
            // the maintainer should ask them before approving. See DraftDiscardWording.
            // ###########################################################################################
            if (detail.Submission.DraftDiscardedUtc is DateTimeOffset discarded)
            {
                lines =
                [
                    new ReviewNoteLine(
                        DraftDiscardWording.SubmissionWarning(discarded, detail.Submission.ContactEmail),
                        ReviewNoteKind.Warning),
                    .. lines
                ];
            }

            // Approve is OFF, with the reason above the table, for a new system with no place in the
            // drop-down lists yet, or a system with an earlier submission still in BETA (2026-09-27)
            // - said before the maintainer reads the whole table, not after pressing Approve.
            string? blocked = ApprovalGate.Blocked(detail.Submission.SystemId, this.thisListing, this.thisBeta);
            this.thisAppliedGate = blocked;

            if (blocked is not null)
            {
                lines = [new ReviewNoteLine(blocked, ReviewNoteKind.Warning), .. lines];
                this.BlockApproval(blocked);
            }

            this.ShowNotInTable(lines);
        }

        // ###########################################################################################
        // *** THE GATE FOLLOWS THE LISTS IT READS (code review, 2026-09-27). *** ApprovalGate reads
        // the drop-down listing and the "Beta > Prod" list, which are replaced on the minute check
        // and after every decision - but it was applied only when a detail was SHOWN, and a detail
        // is read again only when its queue row changed. So Approve stayed off, with a stale reason,
        // after another maintainer promoted the system; and at sign-in a detail that arrived before
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

            string? blocked = ApprovalGate.Blocked(detail.Submission.SystemId, this.thisListing, this.thisBeta);

            if (!string.Equals(blocked, this.thisAppliedGate, StringComparison.Ordinal))
                this.ShowDetail(detail);
        }

        // The gate's reason as last applied to the open detail - null for none.
        private string? thisAppliedGate;

        // True once the open submission was decided by someone else (ShowDecidedElsewhere).
        private bool thisDecisionsClosed;

        // True while a decision is on its way to the server.
        private bool thisDecisionInFlight;

        // Turns Approve off for a reason the server's own answer does not cover (ApprovalGate), and
        // says why on the button itself too. The server refuses regardless.
        private void BlockApproval(string reason)
        {
            this.thisApproveAllowed = false;

            if (this.FindControl<Button>("ApproveButton") is Button approve)
            {
                approve.IsEnabled = false;
                ToolTip.SetTip(approve, reason);
            }
        }

        // The short lines above the table - nothing at all when there is nothing to say.
        private void ShowNotInTable(IReadOnlyList<ReviewNoteLine> lines)
        {
            if (this.FindControl<StackPanel>("NotInTablePanel") is not StackPanel panel)
                return;

            panel.Children.Clear();

            foreach (ReviewNoteLine line in lines)
            {
                var block = new TextBlock
                {
                    FontSize = 12,
                    TextWrapping = TextWrapping.Wrap,
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

        // Who sent the submission and how their other submissions went - above the lines the
        // table cannot show. Null hides it (nothing selected, or an older server).
        private void ShowContributor(ReviewContributorFacts? facts)
        {
            if (this.FindControl<TextBlock>("ContributorText") is not TextBlock block)
                return;

            ReviewNoteLine? line = ReviewContributorLine.For(facts);

            if (line is not null)
                TabMaintainer.ShowLine(block, line);

            block.IsVisible = line is not null;
        }

        // ###########################################################################################
        // Puts one line into a TextBlock. A count's number is bold ("[2]"), so a line with one is
        // built of runs - its Text is then null (a TextBlock cannot mix weights within one Text).
        // Whatever the block showed before is replaced.
        // ###########################################################################################
        private static void ShowLine(TextBlock block, ReviewNoteLine line) =>
            TabMaintainer.ShowRuns(block, line.Runs.Select(run => (run.Text, run.IsCount)).ToList());

        // A status line with its to-do pieces in bold - see StatusPart (SystemsDisplay.cs).
        internal static void ShowParts(TextBlock block, IReadOnlyList<StatusPart> parts) =>
            TabMaintainer.ShowRuns(block, parts.Select(part => (part.Text, part.IsToDo)).ToList());

        // A line of counts, each number in bold as the review notes show theirs ("[12]").
        internal static void ShowCounts(TextBlock block, IReadOnlyList<ReviewNoteRun> runs) =>
            TabMaintainer.ShowRuns(block, runs.Select(run => (run.Text, run.IsCount)).ToList());

        // ###########################################################################################
        // Text in pieces, the marked ones bold. With nothing marked it is plain Text, which lays out
        // cheaper. With a bold piece the words are in Inlines and Text is NULL - read one through
        // TextOf, never through Text.
        // ###########################################################################################
        private static void ShowRuns(TextBlock block, IReadOnlyList<(string Text, bool IsBold)> runs)
        {
            block.Inlines?.Clear();
            block.Text = null;

            if (!runs.Any(run => run.IsBold))
            {
                block.Text = string.Concat(runs.Select(run => run.Text));
                return;
            }

            block.Inlines ??= new Avalonia.Controls.Documents.InlineCollection();

            foreach ((string text, bool isBold) in runs)
            {
                block.Inlines.Add(new Avalonia.Controls.Documents.Run(text)
                {
                    FontWeight = isBold ? FontWeight.Bold : block.FontWeight
                });
            }
        }

        // What a block says, whether it holds plain Text or bold pieces in Inlines - for tests.
        internal static string TextOf(TextBlock block) =>
            block.Text ?? string.Concat(block.Inlines?.OfType<Avalonia.Controls.Documents.Run>().Select(run => run.Text) ?? []);

        // -----------------------------------------------------------------------------------
        // The three decisions (task 5)
        // -----------------------------------------------------------------------------------

        private async void OnApproveClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.DecideAsync(ReviewDecisionKind.Approve);

        private async void OnRejectClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.DecideAsync(ReviewDecisionKind.Reject);

        private async void OnRequestChangesClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.DecideAsync(ReviewDecisionKind.RequestChanges);

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
            // A decision is about what the server holds; unsaved table changes would not be in it.
            if (this.IsTableOpen && this.TableEditor.HasUnsavedChanges)
            {
                this.ShowDecisionMessage(ReviewTableWording.SaveTableBeforeDeciding, isError: true);
                return;
            }

            if (this.thisClient is null || this.thisSession is null || this.thisSelectedId is null)
                return;

            string comment = this.FindControl<TextBox>("DecisionCommentTextBox")?.Text?.Trim() ?? string.Empty;

            // Approve carries no comment; the published tree and the summary already say what
            // happened. The other two are the contributor's only feedback.
            if (kind != ReviewDecisionKind.Approve &&
                !ReviewDecisionWording.IsUsableComment(comment, out string problem))
            {
                this.ShowDecisionMessage(problem, isError: true);
                return;
            }

            this.thisDecisionInFlight = true;
            this.SetDecisionButtonsEnabled(false);
            this.ShowDecisionMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            long id = this.thisSelectedId.Value;
            IReadOnlyList<string> removals = this.thisShownRemovals;
            string? stateBefore = this.thisShownDetail?.Submission.State;

            string waiting = ReviewDecisionWording.Waiting(
                kind, this.thisShownDetail?.Approval, this.thisShownDetail?.Submission.SystemId);

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
                        ReviewDecisionKind.Approve => client.ApproveAsync(session, id, removals, token),
                        ReviewDecisionKind.Reject => client.RejectAsync(session, id, comment, token),
                        _ => client.RequestChangesAsync(session, id, comment, token)
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

        private void ShowDecisionMessage(string? message, bool isError)
        {
            var text = this.FindControl<TextBlock>("DecisionMessageText");

            if (text is null)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = isError ? Brushes.IndianRed : Brushes.SeaGreen;
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

            if (approve is not null)
            {
                approve.IsEnabled = this.thisApproveAllowed;
                approve.Content = ApprovalWording.ApproveButton(approval, "BETA");

                ToolTip.SetTip(
                    approve,
                    canPublish
                        ? "Publishes this submission into the BETA data - once every approval it needs is given. Everyone gets it once it is published from Production. This cannot be undone."
                        : "This account is not a maintainer of this system. Ask the administrator.");
            }

            var status = this.FindControl<TextBlock>("ApprovalStatusText");

            if (status is not null)
            {
                string? line = ApprovalWording.StatusLine(approval);
                status.Text = line ?? string.Empty;
                status.IsVisible = visible && line is not null;
            }
        }

        // ###########################################################################################
        // The files approving the submission on screen would remove, as the server listed them -
        // said beside the button and sent back with the approval, so what goes is what was shown.
        // ###########################################################################################
        private void ShowRemovalNote(FileRemovalPreview? removals)
        {
            this.thisShownRemovals = removals?.Files ?? [];

            if (this.FindControl<TextBlock>("RemovalStatusText") is TextBlock note)
            {
                string? text = FileRemovalWording.ApproveNote(removals);
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
