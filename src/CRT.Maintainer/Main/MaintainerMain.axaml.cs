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
    // The review window: sign in, then the queue on the left and a submission's SUMMARY on the
    // right (NewContributeStrategy.md Phase 5, tasks 2 and 3).
    //
    // FILE MAP - one file for now. When it outgrows ~1,500 lines, split into MaintainerMain.<Area>.cs
    // partials and list them here, as TabSchematics and Main do in CRT.App.
    //
    // *** THE LOGIC IS IN Handlers/, NOT HERE. *** What the summary says is
    // ReviewSummaryPresenter's decision; how a queue row reads is ReviewQueueDisplay's; what the
    // server answered is ReviewApiParser's. All three are unit tested. This file resolves
    // controls, calls one of them, and shows the result - which is why the screens that decide
    // whether a change gets looked at have real coverage rather than being verified by eye.
    //
    // *** THE CHANGE SUMMARY IS NOT SHOWN IN THE QUEUE, AND CANNOT BE. *** Computing it needs the
    // published board AND the submitted manifest; the queue endpoint returns neither, because
    // doing so would mean loading two full BoardData per queued row on the server for a list the
    // maintainer scrolls past. So the queue shows what the CONTRIBUTOR said, and the summary
    // arrives when a submission is opened. See ReviewQueueDisplay's header.
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

        public MaintainerMain()
        {
            this.InitializeComponent();
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
                signedInAs.Text = $"Signed in as {this.thisSession.DisplayName} ({this.thisSession.Email})";
            }
        }

        // -----------------------------------------------------------------------------------
        // The queue
        // -----------------------------------------------------------------------------------

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.RefreshQueueAsync();

        private async Task RefreshQueueAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            // *** AN EXPIRED SESSION IS CAUGHT BEFORE THE REQUEST, not after the 401. *** The
            // server would refuse it anyway, but "your session has expired" is a better answer
            // than "the server answered 401", and it is the same answer either way - so say the
            // true one.
            if (!this.thisSession.IsUsableAt(DateTimeOffset.UtcNow))
            {
                this.ShowQueueMessage("Your session has expired. Sign in again.", isError: true);
                this.ShowSignInPanel();
                return;
            }

            ReviewApiResult<ReviewQueueResponse> result =
                await this.thisClient.GetQueueAsync(this.thisSession);

            if (!result.IsOk)
            {
                // *** THE LIST IS LEFT ALONE ON A FAILURE. *** Clearing it would turn "I could not
                // ask" into "nothing is waiting", which is the one wrong answer a review queue
                // must never give - a backlog would go unnoticed for as long as nobody checked.
                this.ShowQueueMessage(result.Message, isError: true);

                if (result.Failure == ReviewApiFailure.NotSignedIn)
                    this.ShowSignInPanel();

                return;
            }

            this.thisQueue.Clear();
            this.thisQueue.AddRange(result.Value!.Submissions);

            // The administrator's screen is offered only when the SERVER says this account is one.
            var maintainers = this.FindControl<Button>("MaintainersButton");

            if (maintainers is not null)
                maintainers.IsVisible = result.Value.IsAdministrator;

            if (this.FindControl<Button>("UnusedFilesButton") is Button unused)
                unused.IsVisible = result.Value.IsAdministrator;

            this.ApplyQueue();
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
            if (this.thisClient is not null && this.thisSession is not null)
                await this.thisClient.LogoutAsync(this.thisSession);

            this.thisQueue.Clear();
            this.ApplyQueue();
            this.ShowSignInPanel();
            this.ShowSignInMessage("You are signed out.", isError: false);
        }

        private void ApplyQueue()
        {
            var list = this.FindControl<ListBox>("QueueList");

            if (list is null)
                return;

            DateTimeOffset now = DateTimeOffset.UtcNow;

            list.ItemsSource = this.thisQueue
                .Select(row => $"{ReviewQueueDisplay.Title(row)}\n{ReviewQueueDisplay.Subtitle(row, now)}")
                .ToList();

            // An empty queue is a perfectly good answer and says so plainly - distinct from the
            // error wording used above, which the maintainer must be able to tell apart.
            this.ShowQueueMessage(
                this.thisQueue.Count == 0 ? "Nothing waiting for review." : null,
                isError: false);

            if (this.thisQueue.Count == 0)
                this.ShowSubmission(null);
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
            var list = this.FindControl<ListBox>("QueueList");

            if (list is null)
                return;

            int index = list.SelectedIndex;

            if (index < 0 || index >= this.thisQueue.Count)
            {
                this.ShowSubmission(null);
                return;
            }

            ReviewQueueRow row = this.thisQueue[index];

            // The row is shown IMMEDIATELY from what the queue already knows, and the summary
            // fills in when it arrives. Waiting for the request before drawing anything would
            // leave the panel blank on every click over a slow link, which reads as the app
            // having lost the selection.
            this.ShowSubmission(row);

            await this.LoadSubmissionAsync(row);
        }

        // ###########################################################################################
        // Fetches one submission and draws its change summary.
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

            var list = this.FindControl<ListBox>("QueueList");

            if (list is null)
                return;

            int index = list.SelectedIndex;

            if (index < 0 || index >= this.thisQueue.Count || this.thisQueue[index].Id != row.Id)
                return;

            if (!result.IsOk)
            {
                this.ShowHeadline(result.Message);
                return;
            }

            this.ShowDetail(result.Value!);
        }

        // -----------------------------------------------------------------------------------
        // One submission
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // Shows the selected submission.
        //
        // THE CHANGE SUMMARY IS NOT FILLED IN YET - that needs GET /api/review/submissions/{id}
        // plus the published board to compare against, which is the next piece of task 4. The
        // headline says so ROUNDLY rather than sitting blank, because a blank line where a summary
        // belongs reads as "no changes", which would be a lie about a submission nobody has
        // examined.
        // ###########################################################################################
        private void ShowSubmission(ReviewQueueRow? row)
        {
            var title = this.FindControl<TextBlock>("SubmissionTitleText");
            var subtitle = this.FindControl<TextBlock>("SubmissionSubtitleText");
            var headline = this.FindControl<TextBlock>("SummaryLineText");
            var sections = this.FindControl<StackPanel>("SummarySectionsPanel");

            if (title is null || subtitle is null || headline is null || sections is null)
                return;

            sections.Children.Clear();

            if (row is null)
            {
                title.Text = "Select a submission";
                subtitle.Text = string.Empty;
                headline.Text = string.Empty;

                this.thisSelectedId = null;
                this.ShowDecisionPanel(visible: false, canPublish: false, approval: null);
                this.ShowRemovalNote(null);
                return;
            }

            this.thisSelectedId = row.Id;

            title.Text = ReviewQueueDisplay.Title(row);
            subtitle.Text = ReviewQueueDisplay.Subtitle(row, DateTimeOffset.UtcNow);

            // Says it is LOADING rather than sitting blank. A blank line where a summary belongs
            // reads as "no changes" about something nobody has examined yet.
            headline.Text = "Loading changes...";
        }

        // ###########################################################################################
        // Draws a change summary. Not reachable from the UI yet - see ShowSubmission - but kept
        // and exercised by tests, because it is the rendering half of task 3 and the piece the
        // next step plugs into rather than replaces.
        // ###########################################################################################
        internal void ShowDetail(ReviewSubmissionDetail detail)
        {
            ArgumentNullException.ThrowIfNull(detail);

            var sections = this.FindControl<StackPanel>("SummarySectionsPanel");

            if (sections is null)
                return;

            sections.Children.Clear();

            // *** canPublish COMES FROM THE SERVER, never from anything this app worked out. ***
            // It is the same answer the server will enforce when the button is pressed, so the
            // screen cannot promise something the API then refuses.
            this.ShowDecisionPanel(visible: true, canPublish: detail.CanPublish, approval: detail.Approval);
            this.ShowDecisionMessage(null, isError: false);
            this.ShowRemovalNote(detail.Removals);

            if (this.FindControl<TextBlock>("AmendmentText") is TextBlock amended)
            {
                string? line = Handlers.ReviewTableWording.AmendedLine(detail.Amendment);
                amended.Text = line ?? string.Empty;
                amended.IsVisible = line is not null;
            }

            if (detail.Changes is null)
            {
                // The server could not build a summary - an unloadable payload. Saying so beats a
                // blank panel, and the findings below are where the reason will be.
                this.ShowHeadline("The changes in this submission could not be compared.");
            }
            else
            {
                this.ShowHeadline(ReviewSummaryPresenter.BuildHeadline(detail.Changes));

                foreach (ReviewSummaryLine line in ReviewSummaryPresenter.BuildLines(detail.Changes))
                {
                    sections.Children.Add(MaintainerMain.BuildSectionLine(line));

                    // *** THE FIELD-LEVEL DIFF (task 4). *** "U8 changed" is not something a
                    // maintainer can act on; "Part-number: 906114 -> 251715-01" is the whole
                    // decision. Without it they would have to open the board in CRT and hunt for
                    // what moved.
                    foreach (Control field in MaintainerMain.BuildFieldLines(detail.Changes, line.Section))
                    {
                        sections.Children.Add(field);
                    }
                }
            }

            // *** FINDINGS COME AFTER THE CHANGES BUT ARE NOT OPTIONAL. *** They are why automated
            // validation flagged this submission, and a maintainer who never scrolls to them is
            // deciding without the one thing the machine already worked out.
            foreach (ReviewFindingView finding in detail.Findings)
            {
                sections.Children.Add(MaintainerMain.BuildFindingLine(finding));
            }

            // The pictures go LAST because they are the tallest thing on the panel; a maintainer
            // scrolling past a screen of images to reach a one-line finding would miss it. They
            // arrive asynchronously - see LoadImageAsync.
            // EVERY file that would change on the server, written out, BEFORE any picture - the
            // pictures cover images only, and a file that cannot be drawn must still be seen.
            // (security review, 2026-09-25)
            this.ShowFileChanges(detail);
            this.ShowScopeSettingChanges(detail);
            this.ShowMovedHighlights(detail);
            this.ShowImagePairs(detail);
        }

        // ###########################################################################################
        // The complete written list of files this submission changes on the server.
        //
        // *** NOT ONLY IMAGES. *** The picture panel below draws what it can decode; a PDF, a text
        // file or anything else used to appear nowhere at all, so it was approved unseen. What each
        // line says, and which warnings it carries, is ReviewFileComparison's decision - tested
        // there. This only lays the lines out.
        // ###########################################################################################
        private void ShowFileChanges(ReviewSubmissionDetail detail)
        {
            var sections = this.FindControl<StackPanel>("SummarySectionsPanel");

            if (sections is null)
                return;

            IReadOnlyList<ReviewFileLine> lines = ReviewFileComparison.Plan(detail.SubmittedFiles, detail.PublishedFiles, detail.Removals);

            if (lines.Count == 0)
                return;

            sections.Children.Add(new TextBlock
            {
                Text = lines.Count == 1 ? "1 file changes on the server" : $"{lines.Count} files change on the server",
                FontWeight = FontWeight.SemiBold,
                Margin = new Avalonia.Thickness(0, 16, 0, 0)
            });

            // What publishing removes - or that it removes nothing, or why it cannot tell. From the
            // server's own list, the one the approval sends back. (2026-09-25)
            string? removals = Handlers.FileRemovalWording.Headline(detail.Removals, "the BETA data");

            if (removals is not null)
            {
                bool any = detail.Removals!.Files.Count > 0;

                var headline = new TextBlock
                {
                    Text = removals,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Avalonia.Thickness(12, 4, 0, 0),
                    FontWeight = any ? FontWeight.SemiBold : FontWeight.Normal
                };

                if (any)
                    headline.Foreground = Brushes.IndianRed;

                sections.Children.Add(headline);
            }

            foreach (ReviewFileLine line in lines)
            {
                sections.Children.Add(new TextBlock
                {
                    Text = ReviewFileComparison.Describe(line),
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Avalonia.Thickness(12, 4, 0, 0),
                    FontWeight = line.Change == ReviewFileChange.Removed ? FontWeight.SemiBold : FontWeight.Normal,
                    Foreground = line.Change switch
                    {
                        ReviewFileChange.Removed => Brushes.IndianRed,
                        ReviewFileChange.NoLongerUsed => Brushes.Gray,
                        ReviewFileChange.Added => Brushes.SeaGreen,
                        _ => Brushes.SteelBlue
                    }
                });

                foreach (string warning in ReviewFileComparison.Warnings(line))
                {
                    sections.Children.Add(new TextBlock
                    {
                        Text = warning,
                        TextWrapping = TextWrapping.Wrap,
                        Margin = new Avalonia.Thickness(28, 0, 0, 0),
                        Foreground = Brushes.DarkOrange,
                        FontWeight = FontWeight.SemiBold
                    });
                }
            }
        }

        // ###########################################################################################
        // Task 4's SCOPE BASELINES - specifically, that the settings under them moved.
        //
        // *** THE PICTURE DOES NOT SHOW THIS, WHICH IS THE ENTIRE REASON IT IS HERE. *** A baseline
        // is stored as an image, so the side-by-side comparison below already shows both waveforms.
        // What it cannot show is that one was captured at 2 V/div and the other at 5 - the same
        // signal at a different scale, which reads as a change in the circuit.
        //
        // *** LISTED BY COMPONENT RATHER THAN PINNED TO ITS IMAGE PANEL, deliberately. *** A
        // component image's row key is BoardLabel|Region|Pin|Name and carries no file name, so
        // matching a row to the picture it produced is not something this can do RELIABLY - and a
        // warning attached to the wrong trace is worse than one listed separately. It sits
        // immediately above the pictures instead, where it is read before them.
        // ###########################################################################################
        private void ShowScopeSettingChanges(ReviewSubmissionDetail detail)
        {
            var sections = this.FindControl<StackPanel>("SummarySectionsPanel");

            if (sections is null || detail.Changes is null)
                return;

            ReviewSectionView? images = detail.Changes.Sections
                .FirstOrDefault(section => section.Section == ReviewScopeBaseline.SectionName);

            if (images is null)
                return;

            var changed = new List<(string Key, ReviewScopeSettingsChange Change)>();

            foreach (string key in images.Changed)
            {
                if (ReviewScopeBaseline.TryReadChange(images, key, out ReviewScopeSettingsChange change))
                    changed.Add((key, change));
            }

            if (changed.Count == 0)
                return;

            sections.Children.Add(new TextBlock
            {
                Text = changed.Count == 1
                    ? "1 scope baseline was recaptured at different settings"
                    : $"{changed.Count} scope baselines were recaptured at different settings",
                FontWeight = FontWeight.SemiBold,
                Margin = new Avalonia.Thickness(0, 16, 0, 0),

                // Amber rather than red: this is not necessarily wrong, it is something the
                // maintainer has to take into account when looking at the traces below.
                Foreground = Brushes.DarkOrange
            });

            foreach ((string key, ReviewScopeSettingsChange change) in changed)
            {
                var row = new StackPanel
                {
                    Spacing = 2,
                    Margin = new Avalonia.Thickness(24, 4, 0, 0)
                };

                row.Children.Add(new TextBlock
                {
                    // The raw key joins its parts with U+241F, which renders as a box or as
                    // nothing - so it is spelled out readably.
                    Text = ReviewScopeBaseline.DescribeRowKey(key),
                    FontWeight = FontWeight.SemiBold,
                    TextWrapping = TextWrapping.Wrap
                });

                row.Children.Add(new TextBlock
                {
                    Text = ReviewScopeBaseline.Describe(change),
                    Opacity = 0.8,
                    TextWrapping = TextWrapping.Wrap
                });

                sections.Children.Add(row);
            }
        }

        // ###########################################################################################
        // Task 4's MOVED HIGHLIGHT, drawn on the schematic before and after.
        //
        // *** THIS IS THE CHANGE A ROW DIFF CANNOT ANSWER. *** "X: 100 -> 400" is the same
        // information and tells nobody whether the new position is right - which is the only
        // question a maintainer actually has. Seeing the old and new rectangles on the board itself
        // answers it in a glance.
        //
        // A highlight's key is SchematicName|BoardLabel, so X/Y/Width/Height are compared FIELDS -
        // a move arrives as an ordinary changed row with a field diff, and the work here is
        // finding the right schematic image and putting the rectangles back on it.
        // ###########################################################################################
        private void ShowMovedHighlights(ReviewSubmissionDetail detail)
        {
            var sections = this.FindControl<StackPanel>("SummarySectionsPanel");

            if (sections is null || detail.Changes is null)
                return;

            ReviewSectionView? highlights = detail.Changes.Sections
                .FirstOrDefault(section => section.Section == ReviewHighlightGeometry.SectionName);

            if (highlights is null)
                return;

            var moves = new List<(string Key, ReviewHighlightMove Move)>();

            foreach (string key in highlights.Changed)
            {
                if (ReviewHighlightGeometry.TryReadMove(highlights, key, out ReviewHighlightMove move))
                    moves.Add((key, move));
            }

            if (moves.Count == 0)
                return;

            sections.Children.Add(new TextBlock
            {
                Text = moves.Count == 1 ? "1 highlight moved" : $"{moves.Count} highlights moved",
                FontWeight = FontWeight.SemiBold,
                Margin = new Avalonia.Thickness(0, 16, 0, 0)
            });

            foreach ((string key, ReviewHighlightMove move) in moves)
            {
                sections.Children.Add(this.BuildMovedHighlightPanel(detail, key, move));
            }
        }

        // ###########################################################################################
        // One moved highlight: which component on which schematic, and the board with both
        // rectangles drawn on it.
        //
        // *** BOTH RECTANGLES GO ON ONE COPY OF THE BOARD, not two side by side. *** A move is
        // small relative to a board scan, and two images a few hundred pixels apart require the
        // maintainer to hold one in their head while looking at the other. Overlaying them makes the
        // distance itself the thing on screen.
        // ###########################################################################################
        private Control BuildMovedHighlightPanel(
            ReviewSubmissionDetail detail,
            string rowKey,
            ReviewHighlightMove move)
        {
            var panel = new StackPanel
            {
                Spacing = 4,
                Margin = new Avalonia.Thickness(0, 12, 0, 0)
            };

            bool named = ReviewHighlightGeometry.TryReadKeyParts(
                rowKey, out string schematicName, out string boardLabel);

            panel.Children.Add(new TextBlock
            {
                Text = named ? $"{boardLabel} on {schematicName}" : rowKey,
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                Foreground = Brushes.SteelBlue
            });

            // The numbers are shown as well as the picture. The drawing answers "is the new
            // position right"; the numbers are what a maintainer needs if they go and look at the
            // board file by hand afterwards.
            panel.Children.Add(new TextBlock
            {
                Text = MaintainerMain.DescribeMove(move),
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            });

            if (!named || !detail.SchematicImages.TryGetValue(schematicName, out string? imageFile))
            {
                // The schematic's picture is not known - a submission can move a highlight on a
                // schematic whose own row it did not touch and which is not published either.
                // Said plainly, because a silently absent picture reads as the app failing.
                panel.Children.Add(new TextBlock
                {
                    Text = "The schematic image for this board is not available to draw on.",
                    Opacity = 0.6,
                    FontStyle = FontStyle.Italic,
                    TextWrapping = TextWrapping.Wrap
                });

                return panel;
            }

            var canvas = new ReviewHighlightCanvas { Move = move, MaxHeight = 320 };

            var status = new TextBlock
            {
                Text = "Loading...",
                Opacity = 0.6,
                TextWrapping = TextWrapping.Wrap
            };

            panel.Children.Add(canvas);
            panel.Children.Add(status);

            // Deliberately not awaited - the panel is being built and the picture fills itself in.
            _ = this.LoadHighlightBoardAsync(detail, imageFile, canvas, status);

            return panel;
        }

        // ###########################################################################################
        // Fetches the schematic a highlight moved on.
        //
        // *** THE PUBLISHED IMAGE IS TRIED FIRST, THEN THE SUBMITTED ONE. *** The "before"
        // rectangle only means anything against the board as it is TODAY. But a submission can add
        // a schematic and place highlights on it in one go, and there is no published copy of that
        // board at all - so the submitted image is the fallback rather than the first choice.
        // ###########################################################################################
        private async Task LoadHighlightBoardAsync(
            ReviewSubmissionDetail detail,
            string imageFile,
            ReviewHighlightCanvas canvas,
            TextBlock status)
        {
            if (this.thisClient is null || this.thisSession is null)
            {
                status.Text = "Not signed in.";
                return;
            }

            long submissionId = detail.Submission.Id;

            ReviewApiResult<byte[]> result =
                await this.thisClient.GetPublishedAssetAsync(this.thisSession, submissionId, imageFile);

            if (!result.IsOk && result.Failure == ReviewApiFailure.NotFound)
            {
                // Not published - a schematic this submission is adding. Fall back to the
                // submitted copy, which is the only board this highlight has ever sat on.
                string? hash = detail.Assets.Files
                    .FirstOrDefault(file => string.Equals(file.Path, imageFile, StringComparison.Ordinal))
                    ?.Sha256;

                if (!string.IsNullOrEmpty(hash))
                {
                    result = await this.thisClient.GetSubmittedAssetAsync(
                        this.thisSession, submissionId, hash);
                }
            }

            if (!result.IsOk)
            {
                status.Text = result.Failure == ReviewApiFailure.NotFound
                    ? "The schematic image could not be found."
                    : result.Message;

                return;
            }

            try
            {
                using var stream = new System.IO.MemoryStream(result.Value!);

                canvas.Board = new Avalonia.Media.Imaging.Bitmap(stream);
                status.Text = string.Empty;
                status.IsVisible = false;
            }
            catch (Exception exception)
            {
                // Caught broadly for the same reason LoadImageAsync does - a malformed upload
                // must be reported in the panel rather than ending the review session.
                status.Text = $"The schematic could not be shown ({exception.GetType().Name}).";
            }
        }

        // ###########################################################################################
        // The move in words, beside the drawing.
        //
        // Only the parts that actually changed are named. A field diff omits what did not move, so
        // listing all four would print "Y: -> " for a purely horizontal slide.
        // ###########################################################################################
        private static string DescribeMove(ReviewHighlightMove move)
        {
            var parts = new List<string>();

            void Add(string name, string before, string after)
            {
                if (before.Length > 0 || after.Length > 0)
                    parts.Add($"{name} {before} -> {after}");
            }

            Add("X", move.BeforeX, move.AfterX);
            Add("Y", move.BeforeY, move.AfterY);
            Add("Width", move.BeforeWidth, move.AfterWidth);
            Add("Height", move.BeforeHeight, move.AfterHeight);

            return string.Join("   ", parts);
        }

        // ###########################################################################################
        // Task 4's IMAGES SIDE BY SIDE.
        //
        // *** THE PLAN IS DRAWN FIRST AND THE PICTURES FILL IN. *** Each pair is one or two HTTP
        // fetches of a full-resolution board scan, so waiting for them all before drawing anything
        // would leave the panel empty for seconds on exactly the submissions worth looking at. The
        // captions are the part that says what happened, and they need no bytes at all.
        //
        // *** A PAIR WITH NO PUBLISHED COUNTERPART IS CAPTIONED, NOT LEFT BLANK. *** An empty
        // panel beside a full one reads as an image that failed to load.
        // ###########################################################################################
        private void ShowImagePairs(ReviewSubmissionDetail detail)
        {
            var sections = this.FindControl<StackPanel>("SummarySectionsPanel");

            if (sections is null)
                return;

            // The published side's paths come from the SERVER, which holds the published board.
            // They are deliberately not derived from `detail.Changes` - a summary's row keys are
            // natural keys, not file paths, and pairing against those matches nothing while
            // drawing perfectly. See ReviewApiParser.ParsePublishedFiles.
            // The hashes are what let identical files drop out. Without them every file the
            // submission carries is drawn as replaced, which for a rows-only change is the entire
            // board (reported as "1178 images to compare" for a one-line description edit).
            IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
                detail.Assets,
                detail.PublishedFiles,
                detail.PublishedHashes);

            if (pairs.Count == 0)
                return;

            sections.Children.Add(new TextBlock
            {
                Text = pairs.Count == 1 ? "1 image to compare" : $"{pairs.Count} images to compare",
                FontWeight = FontWeight.SemiBold,
                Margin = new Avalonia.Thickness(0, 16, 0, 0)
            });

            foreach (ReviewImagePair pair in pairs)
            {
                sections.Children.Add(this.BuildImagePairPanel(detail.Submission.Id, pair));
            }
        }

        // ###########################################################################################
        // One comparison: the caption, then the two pictures beside each other.
        //
        // The fetch is started and NOT awaited - this is called while building a panel, and each
        // picture fills itself in when it arrives. A failure replaces that side's picture with the
        // reason rather than leaving a hole.
        // ###########################################################################################
        private Control BuildImagePairPanel(long submissionId, ReviewImagePair pair)
        {
            var panel = new StackPanel
            {
                Spacing = 4,
                Margin = new Avalonia.Thickness(0, 12, 0, 0)
            };

            panel.Children.Add(new TextBlock
            {
                Text = ReviewSummaryPresenter.DescribeImagePair(pair),
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                Foreground = pair.Change switch
                {
                    ReviewImageChange.Removed => Brushes.IndianRed,
                    ReviewImageChange.Added => Brushes.SeaGreen,
                    _ => Brushes.SteelBlue
                }
            });

            var side = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };

            side.Children.Add(this.BuildImageSide(submissionId, pair, isBeforeSide: true));
            side.Children.Add(this.BuildImageSide(submissionId, pair, isBeforeSide: false));

            panel.Children.Add(side);

            return panel;
        }

        // ###########################################################################################
        // One side of a comparison - "Before" (published) or "After" (submitted).
        //
        // *** THE SIDES ARE ALWAYS LABELLED, EVEN WHEN BOTH CARRY A PICTURE. *** Two board scans
        // side by side are near-identical by definition; without a label a maintainer has no way to
        // tell which one is the proposal, and half of them would guess the wrong way round.
        // ###########################################################################################
        private Control BuildImageSide(long submissionId, ReviewImagePair pair, bool isBeforeSide)
        {
            var column = new StackPanel { Spacing = 4, Width = 320 };

            column.Children.Add(new TextBlock
            {
                Text = isBeforeSide ? "Before (published)" : "After (submitted)",
                Opacity = 0.7
            });

            bool hasPicture = isBeforeSide ? pair.HasBefore : pair.HasAfter;

            if (!hasPicture)
            {
                column.Children.Add(new TextBlock
                {
                    Text = ReviewSummaryPresenter.DescribeMissingSide(pair.Change, isBeforeSide),
                    Opacity = 0.6,
                    FontStyle = FontStyle.Italic,
                    TextWrapping = TextWrapping.Wrap
                });

                return column;
            }

            var image = new Image
            {
                // Uniform so a board scan is never distorted - a stretched schematic would make a
                // highlight appear to have moved when it has not.
                Stretch = Stretch.Uniform,
                MaxHeight = 240
            };

            var status = new TextBlock
            {
                Text = "Loading...",
                Opacity = 0.6,
                TextWrapping = TextWrapping.Wrap
            };

            column.Children.Add(image);
            column.Children.Add(status);

            // Deliberately not awaited: the panel is being built, and each picture fills itself in.
            _ = this.LoadImageAsync(submissionId, pair, isBeforeSide, image, status);

            return column;
        }

        // ###########################################################################################
        // Fetches one side's bytes and decodes them.
        //
        // *** A DECODE FAILURE MUST NOT CRASH THE WINDOW. *** These bytes are contributor-supplied
        // and may be a truncated upload or something that is not an image at all despite its name.
        // Avalonia's Bitmap constructor throws on both, and this runs on the UI thread from a
        // panel build - an unhandled exception here takes the maintainer app down mid-review.
        // ###########################################################################################
        private async Task LoadImageAsync(
            long submissionId,
            ReviewImagePair pair,
            bool isBeforeSide,
            Image target,
            TextBlock status)
        {
            if (this.thisClient is null || this.thisSession is null)
            {
                status.Text = "Not signed in.";
                return;
            }

            ReviewApiResult<byte[]> result = isBeforeSide
                ? await this.thisClient.GetPublishedAssetAsync(this.thisSession, submissionId, pair.Path)
                : await this.thisClient.GetSubmittedAssetAsync(this.thisSession, submissionId, pair.SubmittedHash);

            if (!result.IsOk)
            {
                // A 404 on the published side is ORDINARY - the comparison expected a file that is
                // not there - and is worded as a fact rather than as an error.
                status.Text = result.Failure == ReviewApiFailure.NotFound
                    ? "No published file at this path."
                    : result.Message;

                return;
            }

            try
            {
                using var stream = new System.IO.MemoryStream(result.Value!);

                target.Source = new Avalonia.Media.Imaging.Bitmap(stream);
                status.Text = string.Empty;
                status.IsVisible = false;
            }
            catch (Exception exception)
            {
                // Caught broadly on purpose: Avalonia's decoder reports a malformed image through
                // more than one exception type depending on the platform's imaging backend, and
                // which ones is not part of its contract. What matters is that a bad upload is
                // reported in the panel rather than ending the session.
                status.Text = $"This file could not be shown as an image ({exception.GetType().Name}).";
            }
        }

        // -----------------------------------------------------------------------------------
        // The three decisions (task 5)
        // -----------------------------------------------------------------------------------

        private async void OnApproveClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.DecideAsync(ReviewDecisionKind.Approve);

        // ###########################################################################################
        // "View in table format" (2026-09-25): the submission's rows in the shared table editor.
        // Modal, so the decision buttons cannot act on a submission while it is being changed; when
        // a change was saved the submission is loaded again - its summary, the files it removes and
        // who must still approve all follow the changed content.
        // ###########################################################################################
        private async void OnViewTableClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (this.thisClient is null || this.thisSession is null || this.thisSelectedId is null)
                return;

            ReviewQueueRow? row = this.thisQueue.FirstOrDefault(candidate => candidate.Id == this.thisSelectedId.Value);

            if (row is null)
                return;

            var window = new ReviewTableWindow();
            window.Initialize(this.thisClient, this.thisSession, row);

            await window.ShowDialog(this);

            if (window.WasAmended)
            {
                await this.RefreshQueueAsync();
                await this.LoadSubmissionAsync(row);
            }
        }

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

                    // A CONFLICT means somebody else decided it first, so the queue on screen is
                    // already stale. Refreshing is the only useful next step and doing it for them
                    // beats telling them to.
                    if (result.Failure == ReviewApiFailure.Conflict)
                        await this.RefreshQueueAsync();

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

        private void ShowHeadline(string text)
        {
            var headline = this.FindControl<TextBlock>("SummaryLineText");

            if (headline is not null)
                headline.Text = text;
        }

        // ###########################################################################################
        // The per-row field diffs for one section, indented under its line.
        //
        // Rows are drawn in the order the section reports them, which is already sorted, so two
        // maintainers looking at the same submission see the same list in the same order.
        // ###########################################################################################
        private static IEnumerable<Control> BuildFieldLines(ReviewChangeSummaryView summary, string sectionName)
        {
            ReviewSectionView? section = summary.Sections
                .FirstOrDefault(candidate => candidate.Section == sectionName);

            if (section is null)
                yield break;

            foreach (string key in section.Changed)
            {
                if (!section.FieldChanges.TryGetValue(key, out IReadOnlyList<ReviewFieldChangeView>? fields))
                    continue;

                foreach (ReviewFieldChangeView field in fields)
                {
                    var row = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Spacing = 8,

                        // Indented so the detail reads as belonging to the section above it rather
                        // than as another section.
                        Margin = new Avalonia.Thickness(24, 0, 0, 0)
                    };

                    row.Children.Add(new TextBlock
                    {
                        // Spelled out readably: a natural key joins its parts with U+241F, which
                        // renders as a box or as nothing at all. This has been drawing raw keys
                        // since the field diff landed.
                        Text = ReviewScopeBaseline.DescribeRowKey(key),
                        Width = 176,
                        Opacity = 0.7,
                        TextWrapping = TextWrapping.Wrap
                    });

                    row.Children.Add(new TextBlock
                    {
                        Text = ReviewSummaryPresenter.DescribeFieldChange(field),
                        TextWrapping = TextWrapping.Wrap
                    });

                    yield return row;
                }
            }
        }

        private static Control BuildFindingLine(ReviewFindingView finding)
        {
            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 8,
                Margin = new Avalonia.Thickness(0, 4, 0, 0)
            };

            row.Children.Add(new TextBlock
            {
                Text = finding.IsError ? "Error" : "Warning",
                FontWeight = FontWeight.SemiBold,
                Width = 200,
                Foreground = finding.IsError ? Brushes.IndianRed : Brushes.DarkOrange
            });

            row.Children.Add(new TextBlock
            {
                // The message is written for the contributor and reads for a maintainer too; the
                // subject is appended only when it adds something the message does not already say.
                Text = string.IsNullOrWhiteSpace(finding.Subject)
                    ? finding.Message
                    : $"{finding.Message} [{finding.Subject}]",
                TextWrapping = TextWrapping.Wrap
            });

            return row;
        }

        private static Control BuildSectionLine(ReviewSummaryLine line)
        {
            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 8
            };

            row.Children.Add(new TextBlock
            {
                Text = line.Section,
                FontWeight = FontWeight.SemiBold,
                Width = 200,
                TextWrapping = TextWrapping.Wrap
            });

            foreach (ReviewSummaryPart part in line.Parts)
            {
                row.Children.Add(new TextBlock
                {
                    Text = part.Describe(),
                    Foreground = MaintainerMain.BrushFor(part.Kind)
                });
            }

            return row;
        }

        // ###########################################################################################
        // The colour per change kind.
        //
        // Hardcoded FOR NOW, deliberately: CRT resolves its colours through a two-step
        // Application.Current + ThemeVariant lookup against keys in App.axaml, and this app has no
        // palette yet. When it gets one, these move there - do not grow more hardcoded colours in
        // the meantime.
        //
        // Removals are red because they are the least recoverable change a submission can make.
        // ###########################################################################################
        private static IBrush BrushFor(ReviewChangeKind kind) => kind switch
        {
            ReviewChangeKind.Removed => Brushes.IndianRed,
            ReviewChangeKind.Added => Brushes.SeaGreen,
            ReviewChangeKind.Renamed => Brushes.DarkOrange,
            _ => Brushes.SteelBlue
        };
    }
}
