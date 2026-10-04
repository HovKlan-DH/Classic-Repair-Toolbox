using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // A SUBMISSION'S FILES (owner request, 2026-09-28) - the Files view (2026-09-30), this part owns:
    //
    //   - reading the tree - the BETA data after approving the submission, against BETA now - the
    //     first time the view is shown for a submission, and showing it;
    //   - the count on its button (SubmissionViews.ChangingFiles), which the submission detail
    //     answers without reading the tree.
    //
    // *** A VIEW, NOT A WINDOW ANY MORE (owner request, 2026-09-30). *** "Files..." opened it in a
    // window of its own, in the corner of the panel, and it "can easily be missed, and then this
    // great possibility may not be used". It is one of the submission's three views now
    // (TabMaintainer.SubmissionViews.cs), and its button says how many files change before it is
    // opened - which is also what the lines above the table used to say about files, and no longer
    // do (ReviewNotInTable).
    //
    // *** READ WHEN FIRST SHOWN, NOT WITH THE SUBMISSION. *** The server walks the board's folder
    // and builds the publish plan for it (SubmissionFileTreeFlow), and a submission is opened on
    // every click in the queue. Once read, the tree stays until another submission is chosen or
    // this one's detail is read again (a save, another maintainer's change) - then it is read anew
    // the next time the view is shown, never under the maintainer's feet while it is on screen.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        // The submission the Files view holds a tree for - null for none read yet.
        private long? thisFilesTreeFor;

        private FileTreeView FilesViewTree => this.FindControl<FileTreeView>("SubmissionFileTree")!;

        // Reads the chosen submission's tree, unless the view already holds it.
        private async Task LoadFilesViewAsync()
        {
            if (this.thisClient is null || this.thisSession is null ||
                this.thisSelectedId is not long id || this.thisFilesTreeFor == id)
            {
                return;
            }

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            // *** THE SUBMISSION'S OWN RECORD, NOT THE QUEUE'S ROW (code review, 2026-10-01). *** A
            // submission another maintainer decided stays open with its decisions off, and its row
            // has left the queue list - SelectedQueueRow is then null and the overlay named no
            // system ("Working out 's files..."). The detail read for this very submission always
            // carries it; the row stays as the fallback for the moment before that arrives.
            ReviewQueueRow? row = this.thisShownDetail is { } shown && shown.Submission.Id == id
                ? shown.Submission
                : this.SelectedQueueRow;
            string systemId = row?.SystemId ?? string.Empty;

            ReviewApiResult<SubmissionFilesAnswer> answer = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingSubmissionFiles(systemId),
                token => client.GetSubmissionFilesAsync(session, id, token));

            // Another submission chosen while it was asked for: that one's view is not this answer's.
            if (this.thisSelectedId != id)
                return;

            FileTreeView tree = this.FilesViewTree;

            if (!answer.IsOk)
            {
                tree.Files = null;
                tree.Clear();
                tree.ShowMessage(answer.Message, isError: true);
                return;
            }

            SubmissionFilesAnswer files = answer.Value!;

            if (this.FindControl<TextBlock>("FilesViewHeadingText") is TextBlock heading)
            {
                heading.Text = row is null || string.IsNullOrWhiteSpace(row.Summary)
                    ? "The BETA data after approving this submission, compared with BETA now."
                    : $"The BETA data after approving \"{row.Summary.Trim()}\", compared with BETA now.";
            }

            tree.Files = new FileTreeFiles(this, client, session, id, files.BetaDataUrl, productionDataUrl: null);
            tree.Show(files.Files);

            this.thisFilesTreeFor = id;
        }

        // ###########################################################################################
        // The tree no longer matches the submission (its detail was read again): read anew.
        //
        // *** AND NOW, IF THE VIEW IS ON SCREEN (code review, 2026-10-01). *** It used to wait for
        // the next time the view was shown - but a maintainer who simply stays on the Files view
        // never shows it again, so the tree built BEFORE another maintainer's amendment stayed up
        // and they could approve on the strength of a list that no longer described the
        // submission. Reading it again is the same wait a click on "Files" costs, and it only
        // happens when the detail really changed (ShowDetail's ReferenceEquals guard). Posted, not
        // awaited: ShowDetail is synchronous and has not stored the new detail yet.
        //
        // *** ONLY WHILE THE TAB IS ON SCREEN (code review, 2026-10-01). *** The detail is also read
        // again by the badge's minute check while CRT shows another tab, and LoadFilesViewAsync
        // waits under MAIN's "please wait" - which then dimmed and blocked a schematic nobody had
        // asked anything of. Off screen the tree is only marked stale, and coming back to the tab
        // reads it (ShowFilesViewIfStale, from AttachQueueChecks).
        // ###########################################################################################
        private void FilesViewIsStale()
        {
            this.thisFilesTreeFor = null;

            if (this.IsOnScreen && this.thisSubmissionView == SubmissionView.Files)
                Dispatcher.UIThread.Post(async () => await this.LoadFilesViewAsync());
        }

        // Coming back to the tab on the Files view: a tree marked stale while it was off screen is
        // read now. Nothing when it is still current (LoadFilesViewAsync's own guard).
        private async Task ShowFilesViewIfStaleAsync()
        {
            if (this.IsOnScreen && this.thisSubmissionView == SubmissionView.Files)
                await this.LoadFilesViewAsync();
        }

        // Nothing of any submission's files - another one chosen, or signed out.
        private void ClearFilesView()
        {
            this.thisFilesTreeFor = null;

            FileTreeView tree = this.FilesViewTree;
            tree.Files = null;
            tree.Clear();

            if (this.FindControl<TextBlock>("FilesViewHeadingText") is TextBlock heading)
                heading.Text = string.Empty;

            this.ShowFilesCount(null);
        }

        // The count on the Files button - none before the detail is read, or with nothing changing.
        private void ShowFilesCount(int? changing)
        {
            string? badge = changing is int count ? SubmissionViews.FilesBadge(count) : null;

            if (this.FindControl<TextBlock>("FilesViewBadgeText") is TextBlock text)
                text.Text = badge ?? string.Empty;

            this.SetShown("FilesViewBadge", badge is not null);
        }

        // The count on the Files button as drawn, or null for none - for tests.
        internal string? FilesCountForTests =>
            this.FindControl<Border>("FilesViewBadge") is { IsVisible: true } badge && badge.Child is TextBlock text
                ? text.Text
                : null;
    }
}
