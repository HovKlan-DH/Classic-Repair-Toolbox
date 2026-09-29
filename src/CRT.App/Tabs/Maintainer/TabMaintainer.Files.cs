using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // A SYSTEM'S FILES (owner request, 2026-09-28) - this part owns:
    //
    //   - "Files..." on the Contributor Submissions screen: the BETA data after approving the
    //     submission, in FileTreeWindow (one at a time; another submission reuses it).
    //
    // An administrator's "Open BETA folder" on this computer went again the same day (owner
    // request, 2026-09-28: "with this new folder view you can scrap that") - pointing at a file
    // shows it, and double-clicking opens it, for every maintainer.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private FileTreeWindow? thisFilesWindow;

        private async void OnSubmissionFilesClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisClient is null || this.thisSession is null ||
                this.thisSelectedId is not long id || this.thisShownDetail is not ReviewSubmissionDetail detail)
            {
                return;
            }

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            string systemId = detail.Submission.SystemId;

            ReviewApiResult<SubmissionFilesAnswer> answer = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingSubmissionFiles(systemId),
                token => client.GetSubmissionFilesAsync(session, id, token));

            if (!answer.IsOk)
            {
                this.ShowTableLoadMessage(answer.Message);
                return;
            }

            SubmissionFilesAnswer files = answer.Value!;

            FileTreeWindow window = this.thisFilesWindow ??= this.NewFilesWindow();

            window.ShowFiles(
                files.SystemId,
                $"\"{detail.Submission.Summary}\"",
                files.Files,
                host => new FileTreeFiles(host, client, session, id, files.BetaDataUrl, productionDataUrl: null));

            // Owned by CRT's window, so it stays above it and closes with it (2026-09-29: the owner
            // was the Maintainer tab's own window until then). No window, nothing to show
            // it over - the tab is not on screen.
            if (window.IsVisible)
                window.Activate();
            else if (TopLevel.GetTopLevel(this) is Window owner)
                window.Show(owner);
        }

        private FileTreeWindow NewFilesWindow()
        {
            var window = new FileTreeWindow();
            window.Closed += (_, _) => this.thisFilesWindow = null;
            return window;
        }

        // Signed out: nothing of the account's stays on screen.
        private void CloseFilesWindow() => this.thisFilesWindow?.Close();
    }
}
