using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Online;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CRT
{
    // ###########################################################################################
    // THE "MY SUBMISSIONS" VIEW (NewContributeStrategy.md Phase 4, task 6) - what this computer
    // has sent, where each one stands, and anything the maintainer said.
    //
    // *** THE LIST IS LOCAL, AND THAT IS NOT A SHORTCUT. *** Contributing needs no account, so the
    // server cannot answer "what did I send" - it has no idea who is asking. What proves ownership
    // is the capability token returned once when a submission is created, which CRT keeps as a
    // receipt (see SubmissionReceipt / SubmissionReceiptStore). This window reads its own receipts
    // and asks the server about each one by id and token.
    //
    // The consequence is stated to the user rather than hidden: a reinstall or a second computer
    // starts empty. The email the contributor gets is the channel that does not depend on this
    // file, and the header says so.
    //
    // NOTHING HERE DECIDES ANYTHING. The state vocabulary is SubmissionReceiptPresenter's (pure,
    // unit tested), the HTTP is SubmissionClient's, the storage is SubmissionReceiptStore's. This
    // file builds rows and marshals back to the UI thread.
    // ###########################################################################################
    public partial class MySubmissionsWindow : Window
    {
        private CancellationTokenSource? thisRefresh;

        // Lets a headless test drive the refresh without a server. null (the default) is the
        // shipped path: a real SubmissionClient against the configured base URL.
        internal Func<long, string, CancellationToken, Task<SubmissionStatus?>>? StatusLookupForTests { get; set; }

        public ObservableCollection<SubmissionListItem> Submissions { get; } = new();

        public MySubmissionsWindow()
        {
            this.InitializeComponent();

            this.SubmissionsItemsControl.ItemsSource = this.Submissions;

            // Tunnel, not bubbling - a focused Button consumes Enter itself. Same trap the other
            // dialogs in this folder document.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            this.Closing += (_, _) => this.CancelRefresh();
        }

        // ###########################################################################################
        // Builds the list from the receipts on disk. Shows immediately from the cached state, so
        // the window is useful with no connection at all; the server is asked separately.
        // ###########################################################################################
        internal void Initialize()
        {
            this.RebuildRows();
        }

        // ###########################################################################################
        // "Mark as read" on one row.
        //
        // *** THE COMMENT STAYS ON SCREEN. *** Only its "New" flag and outline go. Marking feedback
        // as read must never hide what was said - the contributor is very likely about to act on
        // it, and having it vanish the moment they acknowledge it would be the worst possible
        // moment to lose it.
        //
        // Rebuilds the whole list rather than mutating the row, because these rows are immutable
        // and carry no change notification; that is the same reason a refresh rebuilds too.
        // ###########################################################################################
        private void MarkRead(SubmissionListItem row)
        {
            if (row is null)
            {
                return;
            }

            SubmissionReceiptStore.AcknowledgeComment(row.SubmissionId);

            this.RebuildRows();
        }

        private void RebuildRows()
        {
            this.Submissions.Clear();

            foreach (SubmissionReceipt receipt in SubmissionReceiptStore.All)
                this.Submissions.Add(new SubmissionListItem(receipt, this.ForgetAsync, this.MarkRead));

            bool hasAny = this.Submissions.Count > 0;

            this.ListScrollViewer.IsVisible = hasAny;
            this.EmptyStateText.IsVisible = !hasAny;
            this.RefreshButton.IsEnabled = hasAny;
        }

        // ###########################################################################################
        // Asks the server about every submission that could still move.
        //
        // A DECIDED submission is skipped: its state cannot change again, so re-asking is a request
        // that can only return what is already on screen. That is what keeps this cheap for someone
        // with a long history.
        //
        // One row failing does not fail the refresh - GetStatusAsync answers null for anything it
        // cannot reach, and that row simply keeps its last known state.
        // ###########################################################################################
        private async void OnRefreshClick(object? sender, RoutedEventArgs e)
        {
            await this.RefreshAsync();
        }

        // Drives the real refresh from a headless test. The click handler is `async void` (it must
        // be, as an event handler), which a test cannot await - so the work lives here and the
        // handler is one line. This is the shipped path, not a parallel implementation.
        internal void RefreshForTests()
        {
            this.RefreshAsync().GetAwaiter().GetResult();
        }

        private async Task RefreshAsync()
        {
            this.CancelRefresh();
            this.thisRefresh = new CancellationTokenSource();

            CancellationToken token = this.thisRefresh.Token;

            this.RefreshButton.IsEnabled = false;
            this.ShowStatus(string.Empty);

            try
            {
                // Under this window's "please wait" overlay (2026-09-28). Each contribution asked about
                // starts the two minutes again, so a long list is never cut off while it is moving.
                WaitResult<(int Checked, int Updated, int Unreachable, int NotFound)> waited = await BusyOverlay.RunAsync(
                    this,
                    CrtWaitWording.CheckingSubmissions,
                    context => this.CheckReceiptsAsync(context, token));

                // Whatever was checked before the limit is kept - the rows show it.
                this.RebuildRows();

                if (waited.IsTimedOut)
                {
                    this.ShowStatus(CrtWaitWording.SubmissionsNoAnswer);
                    return;
                }

                (int checkedCount, int updated, int unreachable, int notFound) = waited.Value;
                this.ShowRefreshOutcome(checkedCount, updated, unreachable, notFound);
            }
            catch (ClientOutdatedException ex)
            {
                // The server answered "update CRT" (code review, 2026-10-04): its words, not "the
                // server could not be reached" - which would send the contributor looking for a
                // network problem that is not there. Whatever was checked before it is kept.
                this.RebuildRows();
                this.ShowStatus(SubmissionStatusRefresh.DescribeOutdated(ex.Message));
            }
            catch (OperationCanceledException)
            {
                // The window is closing, or a newer refresh replaced this one. Nothing to say.
            }
            finally
            {
                this.RefreshButton.IsEnabled = this.Submissions.Count > 0;
            }
        }

        // The asking itself - one request per contribution still open. `closing` is the window's own
        // token (closed, or replaced by a newer refresh); the overlay's is its two-minute limit.
        private async Task<(int Checked, int Updated, int Unreachable, int NotFound)> CheckReceiptsAsync(WaitContext context, CancellationToken closing)
        {
            using CancellationTokenSource both = CancellationTokenSource.CreateLinkedTokenSource(closing, context.Token);
            CancellationToken token = both.Token;

            var client = new SubmissionClient();

            // *** THE LAUNCH CHECK'S "which rows are worth asking about" RULE, WITHOUT ITS TIME
            // LIMIT *** (SubmissionStatusRefresh.RefreshAsync). This loop is kept rather than
            // delegated because this screen reports "checked N, updated M, could not reach K, L no
            // longer on the server", and the shared method deliberately returns only how many
            // CHANGED - a count that is right for "should anything be redrawn" and wrong for a
            // status line somebody is reading.
            //
            // The launch check stops asking about a submission in BETA after
            // SubmissionReceiptPresenter.MergedRecheckWindow, so it costs nothing for ever
            // (code review, 2026-09-25). This button does not: the contributor pressed it to
            // ask, and one request per such row, once, is what they asked for.
            List<SubmissionReceipt> toCheck = SubmissionReceiptStore.All
                .Where(receipt => SubmissionReceiptPresenter.IsStillOpen(receipt.LastKnownState))
                .ToList();

            int updated = 0;
            int unreachable = 0;
            int notFound = 0;

            foreach (SubmissionReceipt receipt in toCheck)
            {
                token.ThrowIfCancellationRequested();

                // Both a sentence and a sign of life: each one starts the two minutes again.
                context.Report(CrtWaitWording.CheckingSubmission(updated + unreachable + notFound + 1, toCheck.Count));

                SubmissionStatus? status;

                try
                {
                    status = this.StatusLookupForTests is not null
                        ? await this.StatusLookupForTests(receipt.SubmissionId, receipt.UploadToken, token)
                        : await client.GetStatusAsync(receipt.SubmissionId, receipt.UploadToken, token);
                }
                catch (SubmissionNotFoundException)
                {
                    // The server ANSWERED: it does not know it any more (deleted with its board, or
                    // by a reset of the contribution data). The row says "No longer on the server",
                    // and the minute check stops asking every minute. Counted apart from the rows
                    // that could not be reached (code review, 2026-10-04): counted with them, the
                    // status line said the server could not be reached while it was answering.
                    SubmissionReceiptStore.NoteNotFound(receipt.SubmissionId, DateTimeOffset.UtcNow);
                    notFound++;
                    continue;
                }

                if (status is null)
                {
                    unreachable++;
                    continue;
                }

                SubmissionReceiptStore.UpdateState(
                    receipt.SubmissionId,
                    status.State,
                    status.MaintainerComment,
                    DateTimeOffset.UtcNow,
                    status.DecidedUtc,
                    status.AmendedByMaintainer);

                updated++;
            }

            return (toCheck.Count, updated, unreachable, notFound);
        }

        // ###########################################################################################
        // Says what the refresh actually achieved, in those terms.
        //
        // "Nothing to check" is its own message rather than silence: a user who presses a button
        // and sees nothing happen assumes it is broken, when in fact every submission is already
        // decided and there is nothing left to ask about.
        //
        // A contribution the server no longer knows WAS answered for, so it counts as checked and
        // is named apart - never as "could not be reached" (code review, 2026-10-04).
        // ###########################################################################################
        private void ShowRefreshOutcome(int checkedCount, int updated, int unreachable, int notFound)
        {
            if (checkedCount == 0)
            {
                this.ShowStatus("Everything here has already been decided - there is nothing left to check.");
                return;
            }

            int answered = updated + notFound;

            string gone = notFound switch
            {
                0 => string.Empty,
                1 => " 1 is no longer on the server.",
                _ => $" {notFound} are no longer on the server."
            };

            if (unreachable == 0)
            {
                this.ShowStatus((answered == 1 ? "Checked 1 contribution." : $"Checked {answered} contributions.") + gone);
                return;
            }

            if (answered == 0)
            {
                this.ShowStatus(
                    "The server could not be reached. The states below are the last ones known, " +
                    "and nothing has been lost.");

                return;
            }

            this.ShowStatus(
                $"Checked {answered}, but {unreachable} could not be reached. Those rows show their last known state." + gone);
        }

        private void ShowStatus(string message)
        {
            this.StatusText.Text = message;
            this.StatusText.IsVisible = true;
        }

        // ###########################################################################################
        // Removes one receipt from this computer, after saying plainly what that does and does not
        // do - "Remove" beside a queued contribution reads as withdrawing it, and it is not that.
        // ###########################################################################################
        private async Task ForgetAsync(SubmissionListItem item)
        {
            var confirm = new ForgetSubmissionWindow();
            confirm.Initialize(item.BoardDisplayName);

            bool? confirmed = await confirm.ShowDialog<bool?>(this);
            if (confirmed != true)
                return;

            SubmissionReceiptStore.Forget(item.SubmissionId);

            this.RebuildRows();
            this.ShowStatus("Removed from this list. The contribution itself is unaffected.");
        }

        private void CancelRefresh()
        {
            try
            {
                this.thisRefresh?.Cancel();
                this.thisRefresh?.Dispose();
            }
            catch (ObjectDisposedException)
            {
            }
            finally
            {
                this.thisRefresh = null;
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }
    }

    // ###########################################################################################
    // One row: a receipt rendered in the words the contributor reads.
    //
    // Every displayed string comes from SubmissionReceiptPresenter rather than being formatted
    // here, so the vocabulary is pinned by unit tests instead of living in a data template.
    // ###########################################################################################
    public sealed class SubmissionListItem
    {
        public long SubmissionId { get; }
        public string BoardDisplayName { get; }
        public string Summary { get; }
        public string StateText { get; }
        public string SentText { get; }
        public string CheckedText { get; }
        public string MaintainerComment { get; }
        public IReadOnlyList<string> Findings { get; }
        public ICommand ForgetCommand { get; }
        public ICommand MarkReadCommand { get; }

        // "Replied 22 September 2026" - when the maintainer decided, empty until one has. Distinct
        // from CheckedText, which is this computer's own bookkeeping and says nothing about the
        // submission.
        public string DecidedText { get; }

        // "A maintainer changed some of the details ..." - empty when nobody did (2026-09-25).
        public string AmendedText { get; }

        public bool HasAmendedText => !string.IsNullOrWhiteSpace(this.AmendedText);

        // How to try it, while it is in the BETA source (2026-10-03) - empty otherwise.
        public string BetaTryText { get; }

        public bool HasBetaTryText => !string.IsNullOrWhiteSpace(this.BetaTryText);

        public bool HasSummary => !string.IsNullOrWhiteSpace(this.Summary);
        public bool HasCheckedText => !string.IsNullOrWhiteSpace(this.CheckedText);
        public bool HasDecidedText => !string.IsNullOrWhiteSpace(this.DecidedText);
        public bool HasMaintainerComment => !string.IsNullOrWhiteSpace(this.MaintainerComment);
        public bool HasFindings => this.Findings.Count > 0;

        // ###########################################################################################
        // Whether this row's comment is one the contributor has not acknowledged.
        //
        // Read off the receipt through the shared presenter rather than recomputed here, so the
        // row, the badge on the Drafts tab and the tests all answer this question the same way -
        // a row that disagreed with the badge would be worse than no badge at all.
        // ###########################################################################################
        public bool HasUnreadComment { get; }

        // ###########################################################################################
        // Is there anything on this row the contributor has not seen - a comment, a decision, or
        // both? (owner report, 2026-09-23.)
        //
        // *** DISTINCT FROM HasUnreadComment, and the difference is a row like a silent publish. ***
        // That one has a new OUTCOME and nothing written, so it carries no comment panel at all -
        // and the "Mark as read" button lives inside that panel. Without this the badge counted the
        // row (correctly, it is news) while the window offered no way whatever to dismiss it.
        //
        // Drives a dismiss button OUTSIDE the comment panel, so every row the badge counts can be
        // cleared by hand. Opening the window still clears nothing: see MarkRead.
        // ###########################################################################################
        public bool HasUnreadNews { get; }

        // A row whose only news is the decision. The comment panel already carries its own "New"
        // badge and button, so showing a second one there would say the same thing twice.
        public bool HasUnreadDecisionOnly => this.HasUnreadNews && !this.HasUnreadComment;

        // ###########################################################################################
        // THE STATUS COLOUR, resolved once from the shared classification.
        //
        // *** THE COLOUR AND THE WORDS COME OFF THE SAME SWITCH *** (SubmissionReceiptPresenter),
        // so a card can never be painted green while its own text reads "Changes requested".
        //
        // These are the app's existing accent colours rather than new ones: IndianRed is what the
        // unread badge and every destructive action already use, and SeaGreen is the Workbooks
        // tab's "Open/good" green.
        // ###########################################################################################
        public IBrush StatusAccentBrush { get; }

        // The state line itself, in the same colour as the edge, so the two read as one statement
        // rather than as a stripe next to some text.
        public IBrush StatusTextBrush => this.StatusAccentBrush;

        // ###########################################################################################
        // A faint fill so each card reads as a panel against the window (owner request).
        //
        // *** RESOLVED FROM THE THEME, NOT HARDCODED GREY. *** A literal #F5F5F5 would be invisible
        // in the light theme's own white and would glare in the dark one - this app defines both,
        // and Table_AlternateRow is the key that already means "a surface one step off the
        // background" in each.
        // ###########################################################################################
        public IBrush CardBackgroundBrush { get; }

        // The unread row is outlined in the badge's own IndianRed so it can be picked out of a
        // list of several; a read one falls back to the ordinary form border. Exposed as brushes
        // rather than as a style trigger because this row is built in code, like every other
        // theme-aware value in this codebase.
        public IBrush CommentBorderBrush => this.HasUnreadComment
            ? this.StatusAccentBrush
            : SubmissionListItem.Theme("Form_Border", Brushes.Gray);

        // ###########################################################################################
        // One theme lookup, so every brush on this row resolves the same way.
        //
        // *** IT MUST BE THE TWO-STEP Application.Current + ActualThemeVariant FORM. *** These rows
        // are built in code rather than by a DataTemplate precisely because a template binding
        // cannot express that, which is the same reason Main builds the worklog bar's pill in code
        // (see CLAUDE.md). The fallback covers a test that runs without an Application.
        // ###########################################################################################
        private static IBrush Theme(string key, IBrush fallback) =>
            Application.Current?.FindResource(Application.Current.ActualThemeVariant, key) as IBrush
                ?? fallback;

        // ###########################################################################################
        // The accent for each outcome kind.
        //
        // Reuses the app's EXISTING status colours rather than introducing a private palette:
        // Text_Fail_Fg is IndianRed, the same red the unread badge and every destructive action
        // already use, and Text_Success_Fg is the green the rest of the app reads as "good".
        //
        // *** NeedsAction IS ORANGE, NOT RED (owner request, 2026-09-23). *** It used to share
        // Text_Fail_Fg with Bad, on the reasoning that both mean "this did not go through" and that
        // the words beside them ("Changes requested" vs "Not accepted") carry the difference. That
        // was wrong in practice and was reported: on a list of four cards the two states were
        // indistinguishable at a glance, and the one that mattered - the one the contributor can
        // still rescue - looked exactly like the one that is dead.
        //
        // The colour now carries the distinction the words alone did not: a WARNING rather than a
        // failure. Four colours across four buckets, which is the legend being one-to-one rather
        // than something extra to learn.
        // ###########################################################################################
        // internal: the Drafts tab's submission badge colours itself with this same mapping, so a
        // state cannot be green in one place and orange in the other.
        internal static IBrush AccentFor(SubmissionOutcomeKind kind) => kind switch
        {
            SubmissionOutcomeKind.Good => SubmissionListItem.Theme("Text_Success_Fg", Brushes.SeaGreen),
            SubmissionOutcomeKind.Bad => SubmissionListItem.Theme("Text_Fail_Fg", Brushes.IndianRed),
            SubmissionOutcomeKind.NeedsAction => SubmissionListItem.Theme("Text_Warning_Fg", Brushes.DarkOrange),
            _ => SubmissionListItem.Theme("Text_Waiting_Fg", Brushes.Goldenrod)
        };

        public Thickness CommentBorderThickness => new(this.HasUnreadComment ? 2 : 1);

        public SubmissionListItem(
            SubmissionReceipt receipt,
            Func<SubmissionListItem, Task> forget,
            Action<SubmissionListItem>? markRead = null)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            this.SubmissionId = receipt.SubmissionId;

            // The board's own identity is an ExcelDataFile key ("Commodore/C64/250407/Data...xlsx").
            // Shown as the folder segments only - a file name nobody typed is noise on a row whose
            // job is to say which board this was.
            this.BoardDisplayName = SubmissionListItem.DescribeBoard(receipt.BoardId);

            this.Summary = receipt.Summary ?? string.Empty;
            this.StateText = SubmissionReceiptPresenter.DescribeReceiptState(receipt);
            this.SentText = SubmissionReceiptPresenter.DescribeSent(receipt.SentUtc);

            // Through the presenter like every other date on this row, so all three share one
            // format. This one used to be built here with its own format string, which is how it
            // ended up being the odd one out.
            this.CheckedText = SubmissionReceiptPresenter.DescribeLastChecked(receipt.LastCheckedUtc);

            // Through the shared presenter, like every other displayed string on this row, so the
            // wording is pinned by unit tests rather than living in a data template.
            this.DecidedText = SubmissionReceiptPresenter.DescribeDecided(receipt.DecidedUtc);
            this.AmendedText = SubmissionReceiptPresenter.DescribeAmended(receipt.AmendedByMaintainer);
            this.BetaTryText = SubmissionReceiptPresenter.DescribeReceiptBetaTry(receipt);

            this.MaintainerComment = receipt.MaintainerComment ?? string.Empty;
            this.HasUnreadComment = SubmissionReceiptPresenter.HasUnreadComment(receipt);
            this.HasUnreadNews = SubmissionReceiptPresenter.HasUnreadNews(receipt);

            // The colour and the words both come off SubmissionReceiptPresenter, so they cannot
            // disagree about what this row's state means.
            this.StatusAccentBrush = SubmissionListItem.AccentFor(
                SubmissionReceiptPresenter.ClassifyReceipt(receipt));

            this.CardBackgroundBrush = SubmissionListItem.Theme("Card_Bg", Brushes.WhiteSmoke);

            // Findings are not stored in the receipt - they belong to the server's answer and can
            // change. An empty list keeps the row's shape honest until a refresh supplies them.
            this.Findings = [];

            this.ForgetCommand = new ActionCommand(
                () => Dispatcher.UIThread.InvokeAsync(() => forget(this)));

            // Null-safe so a row can still be constructed by a test that only cares about the
            // displayed text, the same way ForgetCommand's own callback is injected.
            this.MarkReadCommand = new ActionCommand(
                () => Dispatcher.UIThread.InvokeAsync(() => markRead?.Invoke(this)));
        }

        // ###########################################################################################
        // "Commodore/C128/310378 Open128" -> "Commodore C128 310378 Open128" - through the shared
        // presenter, which the Drafts tab's "switch back from BETA" notice names boards with too.
        // It used to drop the last segment as a file name, which on a real receipt was the BOARD.
        // ###########################################################################################
        private static string DescribeBoard(string? boardId) =>
            SubmissionReceiptPresenter.DescribeBoard(boardId);
    }
}
