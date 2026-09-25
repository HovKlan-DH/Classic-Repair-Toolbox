using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // The maintainer's table (owner request, 2026-09-25) - "View in table format" on a
    // submission.
    //
    // *** THE LOGIC IS ELSEWHERE, AS EVERYWHERE IN THIS APPLICATION. *** The table itself is
    // CRT.UI's BoardTableEditor in document mode, its rules are CRT.Data's (BoardTableDocument),
    // what a save sends is ReviewTableWording.RowsToSave, and what may change - and who may change
    // it, and when - is the server's (AmendSubmissionFlow). This window loads, sends, and asks
    // before unsaved changes are lost.
    //
    // *** A TABLE OPENED IS A VERSION SENT. *** The save carries the amendment version the table
    // was opened at, and the server refuses it when somebody else changed the submission since -
    // so one maintainer's change is never silently written over another's.
    // ###########################################################################################
    public partial class ReviewTableWindow : Window
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;
        private ReviewQueueRow? thisRow;

        // What the table was opened on - the version a save is checked against, and the rows the
        // edits are applied to.
        private ReviewTableData? thisTable;

        // Set once closing has been settled (saved, discarded, or nothing to lose), so the Closing
        // handler lets the window go the second time round.
        private bool thisClosingSettled;

        public ReviewTableWindow()
        {
            this.InitializeComponent();

            this.TableEditor.SaveRequested += async (_, _) => await this.SaveAsync();
            this.Closing += this.OnWindowClosing;
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        private BoardTableEditor TableEditor => this.FindControl<BoardTableEditor>("Editor")!;

        // Whether a change was saved - the submission view reloads when it was.
        public bool WasAmended { get; private set; }

        public void Initialize(ReviewApiClient client, ReviewSession session, ReviewQueueRow row)
        {
            ArgumentNullException.ThrowIfNull(client);
            ArgumentNullException.ThrowIfNull(session);
            ArgumentNullException.ThrowIfNull(row);

            this.thisClient = client;
            this.thisSession = session;
            this.thisRow = row;

            this.Title = ReviewTableWording.WindowTitle(row);

            if (this.FindControl<TextBlock>("HeadingText") is TextBlock heading)
                heading.Text = $"#{row.Id} - {row.SystemId}";
        }

        protected override async void OnOpened(EventArgs e)
        {
            base.OnOpened(e);
            await this.LoadAsync(message: null);
        }

        // ###########################################################################################
        // Fetches the submission's rows and the published board, and opens the table on them.
        // `message` is shown under the toolbar afterwards - "Saved - ..." after a save.
        // ###########################################################################################
        private async Task LoadAsync(string? message)
        {
            if (this.thisClient is null || this.thisSession is null || this.thisRow is null)
                return;

            ReviewApiResult<ReviewTableData> result = await this.thisClient.GetTableAsync(this.thisSession, this.thisRow.Id);

            if (!result.IsOk)
            {
                this.ShowLoadMessage(result.Message);
                return;
            }

            this.thisTable = result.Value!;
            this.ShowLoadMessage(null);

            BoardData? published = this.thisTable.Published is null ? null : SubmissionRowsBoard.ToBoard(this.thisTable.Published);
            BoardData submitted = SubmissionRowsBoard.ToBoard(this.thisTable.Submitted);

            this.TableEditor.Open(
                BoardTableDocument.Create(published, submitted),
                message ?? ReviewTableWording.OpenedMessage(this.thisTable));
        }

        // ###########################################################################################
        // "Save changes": the table's rows, sent as an amendment at the version it was opened at.
        // Returns whether it was saved, for the closing prompt's Save.
        // ###########################################################################################
        private async Task<bool> SaveAsync()
        {
            if (this.thisClient is null || this.thisSession is null || this.thisRow is null || this.thisTable is null)
                return false;

            BoardTableDocument? document = this.TableEditor.CommitAndGetDocument();

            if (document is null || !document.HasUnsavedChanges)
                return true;

            this.TableEditor.ShowMessage("Saving...");

            ReviewApiResult<ReviewAmendResult> result = await this.thisClient.AmendAsync(
                this.thisSession,
                this.thisRow.Id,
                this.thisTable.Version,
                ReviewTableWording.RowsToSave(document, this.thisTable.Submitted));

            if (!result.IsOk)
            {
                this.TableEditor.ShowMessage(ReviewTableWording.NotSaved(result.Message));
                return false;
            }

            this.WasAmended = true;

            // Opened again on what the server now holds, so what is on screen is what was stored.
            await this.LoadAsync(ReviewTableWording.Saved(result.Value!));

            return true;
        }

        private void OnCloseClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => this.Close();

        // ###########################################################################################
        // Unsaved changes are asked about before the window goes - Save, Discard or Cancel - with the
        // same prompt the Drafts tab uses, worded for a submission.
        // ###########################################################################################
        private async void OnWindowClosing(object? sender, WindowClosingEventArgs e)
        {
            if (this.thisClosingSettled || !this.TableEditor.HasUnsavedChanges)
                return;

            e.Cancel = true;

            var prompt = new UnsavedTableEditsWindow();
            prompt.Initialize(UnsavedTableEditsPrompt.LeavingSubmission);

            UnsavedTableEditsChoice? choice = await prompt.ShowDialog<UnsavedTableEditsChoice?>(this);

            bool mayClose = choice switch
            {
                UnsavedTableEditsChoice.Save => await this.SaveAsync(),
                UnsavedTableEditsChoice.Discard => true,
                _ => false
            };

            if (mayClose)
            {
                this.thisClosingSettled = true;
                this.Close();
            }
        }

        private void ShowLoadMessage(string? message)
        {
            if (this.FindControl<TextBlock>("LoadMessageText") is not TextBlock text)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = Brushes.IndianRed;
        }
    }
}
