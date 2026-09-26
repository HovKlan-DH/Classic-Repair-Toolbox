using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // The review window: sign in, then the queue on the left and the selected submission's TABLE
    // on the right (NewContributeStrategy.md Phase 5, tasks 2 and 3).
    //
    // *** THE TABLE IS THE SUBMISSION VIEW (owner request, 2026-09-26). *** A change summary -
    // section lines, field diffs, the file list, pictures side by side, moved highlights drawn on
    // their schematic - filled this panel until then, with the table a button away. The project
    // owner found it "confusing to look at" and asked for it to go. What the table cannot show
    // (highlights, calibration points, the automatic checks' warnings) stays as a few lines above
    // it - ReviewNotInTable.
    //
    // FILE MAP - this file (sign in, the queue, the decisions), MaintainerMain.QueueItems.cs (the
    // queue list grouped by board), MaintainerMain.QueueRefresh.cs (the queue checking itself - no
    // Refresh button), MaintainerMain.Table.cs (the table, and asking before unsaved changes in it
    // are left), and MaintainerMain.Settings.cs (the window's place and "Show changes only",
    // remembered between runs).
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
    public partial class MaintainerMain : Window
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

        public MaintainerMain()
        {
            this.InitializeComponent();
            this.WireTable();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // ###########################################################################################
        // RESTORES A REMEMBERED SESSION, so a signed-in maintainer never sees the password box.
        //
        // *** IN OnOpened RATHER THAN IN THE CONSTRUCTOR. *** Restoring means fetching the queue,
        // and doing that during construction would hold the window off screen behind a network
        // call - so a slow or unreachable server would look like an app that failed to start.
        // Here the window is already up and the fetch fills it in.
        //
        // *** THE STORED SESSION IS NOT TRUSTED, ONLY OFFERED. *** Recall refuses an expired one,
        // and RefreshQueueAsync then asks the SERVER, which is the only authority on whether the
        // session is still live: signed out, revoked, or the account locked all come back as a
        // 401 and land in ShowSignInPanel, which clears the file. So the worst case for a stale
        // token is one failed request and a normal sign-in screen.
        // ###########################################################################################
        protected override async void OnOpened(EventArgs e)
        {
            base.OnOpened(e);

            this.SettleRestoredPlacement();

            ReviewSessionStore.Initialise();

            ReviewSession? remembered = ReviewSessionStore.Recall(DateTimeOffset.UtcNow);

            if (remembered is null)
                return;

            this.thisSession = remembered;
            this.thisClient = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);

            this.ShowQueuePanel();
            await this.RefreshQueueAsync();
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

            // *** NOT ASKED FOR. *** The maintainer app talks to one service and always will, so the
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

                ReviewApiResult<ReviewSession> result = await this.thisClient.LoginAsync(
                    email.Text?.Trim() ?? string.Empty,
                    password.Text ?? string.Empty);

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
                await this.RefreshQueueAsync();
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

                ReviewApiResult<string> result = await client.ForgotPasswordAsync(address);

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

                ReviewApiResult<string> result =
                    await client.ResetPasswordAsync(enteredCode, enteredPassword);

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

            this.StartQueueChecks();
        }

        // -----------------------------------------------------------------------------------
        // The queue
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // Asks for the queue and shows it. `background` is the queue's own check
        // (MaintainerMain.QueueRefresh.cs): it updates the list and leaves the open submission alone.
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
        // *** THE ADMINISTRATOR'S SCREENS ARE OFFERED ONLY WHEN THE SERVER SAYS THIS ACCOUNT IS ONE ***
        // (Maintainers, Unused files). The server refuses everyone else regardless; a button that
        // could only ever be refused would be clutter.
        // ###########################################################################################
        internal ReviewQueueRow? ApplyQueueResponse(ReviewQueueResponse response, bool background = false)
        {
            ArgumentNullException.ThrowIfNull(response);

            this.thisQueue.Clear();
            this.thisQueue.AddRange(response.Submissions);

            if (this.FindControl<Button>("MaintainersButton") is Button maintainers)
                maintainers.IsVisible = response.IsAdministrator;

            if (this.FindControl<Button>("UnusedFilesButton") is Button unused)
                unused.IsVisible = response.IsAdministrator;

            return this.ApplyQueue(background);
        }

        // ###########################################################################################
        // Opens the "Production" window - BETA to production (2026-09-25). Modal, like Maintainers.
        // ###########################################################################################
        private async void OnProductionClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            var window = new ProductionWindow();
            window.Initialize(this.thisClient, this.thisSession);

            await window.ShowDialog(this);
        }

        // ###########################################################################################
        // Opens the administrator's "Unused files" window (2026-09-25). Modal, like Maintainers.
        // ###########################################################################################
        private async void OnUnusedFilesClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            var window = new UnusedFilesWindow();
            window.Initialize(this.thisClient, this.thisSession);

            await window.ShowDialog(this);
        }

        // ###########################################################################################
        // Opens the administrator's "Maintainers" window over this one. Modal, so the queue cannot be
        // acted on while pools are being changed under it; the queue is refreshed afterwards,
        // because assigning a maintainer to the administrator's own account changes nothing but
        // assigning somebody ELSE may have been prompted by a submission still on screen.
        // ###########################################################################################
        private async void OnMaintainersClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            var window = new MaintainersWindow();
            window.Initialize(this.thisClient, this.thisSession);

            await window.ShowDialog(this);
            await this.RefreshQueueAsync();
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

            if (this.thisClient is not null && this.thisSession is not null)
                await this.thisClient.LogoutAsync(this.thisSession);

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
                // Grouped by board, with a heading per board - see MaintainerMain.QueueItems.cs.
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
                // screen, undecidable - see MaintainerMain.QueueRefresh.cs.
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
            // and its table, side by side.
            await Task.WhenAll(this.LoadSubmissionAsync(row), this.OpenTableAsync(row));
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

            // *** canPublish COMES FROM THE SERVER, never from anything this app worked out. ***
            // It is the same answer the server will enforce when the button is pressed, so the
            // screen cannot promise something the API then refuses.
            this.ShowDecisionPanel(visible: true, canPublish: detail.CanPublish, approval: detail.Approval);
            this.SetDecisionButtonsEnabled(true);
            this.ShowDecisionMessage(null, isError: false);
            this.ShowRemovalNote(detail.Removals);

            // The list says what the submission itself says - see MaintainerMain.QueueItems.cs.
            this.UpdateQueueEntry(
                detail.Submission,
                detail.Changes?.IsNewSystem ?? detail.Submission.IsNewSystem,
                ReviewQueueDisplay.AwaitsYou(detail.CanPublish, detail.Approval));

            if (this.FindControl<TextBlock>("AmendmentText") is TextBlock amended)
            {
                string? line = Handlers.ReviewTableWording.AmendedLine(detail.Amendment);
                amended.Text = line ?? string.Empty;
                amended.IsVisible = line is not null;
            }

            // *** WHAT THE TABLE CANNOT SHOW IS NOT OPTIONAL. *** Highlights and calibration points
            // publish with the approval, and the automatic checks' warnings are what the machine
            // already worked out - see ReviewNotInTable.
            this.ShowContributor(detail.Contributor);
            this.ShowNotInTable(ReviewNotInTable.Lines(detail.Changes, detail.Findings, detail.SubmittedFiles));
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

                MaintainerMain.ShowLine(block, line);
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
                MaintainerMain.ShowLine(block, line);

            block.IsVisible = line is not null;
        }

        // ###########################################################################################
        // Puts one line into a TextBlock. A count's number is bold ("[2]"), so a line with one is
        // built of runs - its Text is then null (a TextBlock cannot mix weights within one Text).
        // Whatever the block showed before is replaced.
        // ###########################################################################################
        private static void ShowLine(TextBlock block, ReviewNoteLine line)
        {
            block.Inlines?.Clear();
            block.Text = null;

            if (!line.Runs.Any(run => run.IsCount))
            {
                block.Text = line.Text;
                return;
            }

            block.Inlines ??= new Avalonia.Controls.Documents.InlineCollection();

            foreach (ReviewNoteRun run in line.Runs)
            {
                block.Inlines.Add(new Avalonia.Controls.Documents.Run(run.Text)
                {
                    FontWeight = run.IsCount ? FontWeight.Bold : block.FontWeight
                });
            }
        }

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
                this.ShowDecisionMessage(Handlers.ReviewTableWording.SaveTableBeforeDeciding, isError: true);
                return;
            }

            if (this.thisClient is null || this.thisSession is null || this.thisSelectedId is null)
                return;

            string comment = this.FindControl<TextBox>("DecisionCommentTextBox")?.Text?.Trim() ?? string.Empty;

            // Approve carries no comment; the published tree and the summary already say what
            // happened. The other two are the contributor's only feedback.
            if (kind != ReviewDecisionKind.Approve &&
                !Handlers.ReviewDecisionWording.IsUsableComment(comment, out string problem))
            {
                this.ShowDecisionMessage(problem, isError: true);
                return;
            }

            this.SetDecisionButtonsEnabled(false);
            this.ShowDecisionMessage("Working...", isError: false);

            try
            {
                long id = this.thisSelectedId.Value;

                ReviewApiResult<ReviewDecisionResult> result = kind switch
                {
                    ReviewDecisionKind.Approve =>
                        await this.thisClient.ApproveAsync(this.thisSession, id, this.thisShownRemovals),
                    ReviewDecisionKind.Reject =>
                        await this.thisClient.RejectAsync(this.thisSession, id, comment),
                    _ =>
                        await this.thisClient.RequestChangesAsync(this.thisSession, id, comment)
                };

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
                        await this.RefreshQueueAsync();

                        if (this.thisSelectedId is null)
                            this.ShowQueueMessage(result.Message, isError: true);
                        else
                            this.ShowDecisionMessage(result.Message, isError: true);
                    }

                    return;
                }

                this.ShowDecisionMessage(
                    Handlers.ReviewDecisionWording.Describe(kind, result.Value!),
                    isError: false);

                // The submission has left the queue, so the list on screen is now wrong.
                await this.RefreshQueueAsync();
            }
            finally
            {
                this.SetDecisionButtonsEnabled(true);
            }
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
            this.thisApproveAllowed = canPublish && Handlers.ApprovalWording.CanApprove(approval);

            var approve = this.FindControl<Button>("ApproveButton");

            if (approve is not null)
            {
                approve.IsEnabled = this.thisApproveAllowed;
                approve.Content = Handlers.ApprovalWording.ApproveButton(approval, "BETA");

                ToolTip.SetTip(
                    approve,
                    canPublish
                        ? "Publishes this submission into the BETA data - once every approval it needs is given. Everyone gets it once it is published from Production. This cannot be undone."
                        : "This account is not a maintainer of this system. Ask the administrator.");
            }

            var status = this.FindControl<TextBlock>("ApprovalStatusText");

            if (status is not null)
            {
                string? line = Handlers.ApprovalWording.StatusLine(approval);
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
                string? text = Handlers.FileRemovalWording.ApproveNote(removals);
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
