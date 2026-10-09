using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // A BOARD'S FILES (owner request, 2026-10-03: "Files (should not show changed files - just list
    // all files)") - every file the board uses as BETA holds it: its own folder, and the shared
    // files its board cites (the server's BoardFilesFlow). The same tree the two queues show, as a
    // LISTING (FileTreeView.ShowListing): no "Show only changed files", a count of files, and the
    // board's own folder open. Pointing at a file shows it and double-clicking opens it, from BETA's
    // public address, exactly as the queues' trees do (FileTreeFiles).
    //
    // Read when the view is first shown for a board (BoardDetailView.Sections.cs), and kept until another
    // board is chosen - or until BETA's board moves (2026-10-04, CatchUpWithBetaAsync).
    // ###########################################################################################
    public partial class BoardDetailView
    {
        // The board the Files view holds a list for - null for none read yet.
        private string? thisFilesFor;

        // The board as it was when the list was read, held against it as it is now
        // (QueueRefreshRules.BetaBoardChanged).
        private BoardOverviewEntry? thisFilesReadAt;

        private FileTreeView BoardFiles => this.FindControl<FileTreeView>("BoardFileTree")!;

        private bool HoldsFilesFor(string? boardId) =>
            this.thisFilesFor is not null && string.Equals(this.thisFilesFor, boardId, StringComparison.Ordinal);

        // Reads the board on screen's files and shows them. The caller holds the wait.
        private async Task LoadFilesAsync()
        {
            if (this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session ||
                this.ShownBoard is not { } board)
            {
                return;
            }

            ReviewApiResult<BoardFilesAnswer> answer = await client.GetBoardFilesAsync(session, board.BoardId);

            if (!string.Equals(this.ShownBoard?.BoardId, board.BoardId, StringComparison.Ordinal))
                return;

            if (!answer.IsOk)
            {
                this.ShowFiles(null, null, board.BoardId);

                // "Not in BETA" is an answer about the board, not a failure.
                this.BoardFiles.ShowMessage(answer.Message, isError: answer.Failure != ReviewApiFailure.NotFound);
                return;
            }

            this.ShowFiles(answer.Value!, client, board.BoardId);
            this.thisFilesFor = board.BoardId;
            this.thisFilesReadAt = board;
        }

        private void ShowFiles(BoardFilesAnswer? files, ReviewApiClient? client, string boardId)
        {
            FileTreeView tree = this.BoardFiles;

            tree.Files = files is null || client is null
                ? null
                : new FileTreeFiles(this, client, this.thisSession, submissionId: null, files.BetaDataUrl, productionDataUrl: null);

            tree.ShowListing(files?.Files, openFolder: boardId);

            if (this.FindControl<TextBlock>("FilesHeadingText") is TextBlock heading)
                heading.Text = files is null ? string.Empty : BoardSections.FilesHeading;
        }

        // Shows a list without a server - for tests. `readAt` is the board as it was when it was read.
        internal void ShowFilesForTests(BoardFilesAnswer files, BoardOverviewEntry? readAt = null)
        {
            this.ShowFiles(files, client: null, files.BoardId);

            if (readAt is not null)
            {
                this.thisFilesFor = files.BoardId;
                this.thisFilesReadAt = readAt;
            }
        }

        internal bool HoldsFilesForTests(string boardId) => this.HoldsFilesFor(boardId);

        internal FileTreeView FileTreeForTests => this.BoardFiles;

        // Nothing of any board's files - another one chosen, or signed out.
        private void ClearFiles()
        {
            this.thisFilesFor = null;
            this.thisFilesReadAt = null;

            FileTreeView tree = this.BoardFiles;
            tree.Files = null;
            tree.ShowListing(null, openFolder: null);

            if (this.FindControl<TextBlock>("FilesHeadingText") is TextBlock heading)
                heading.Text = string.Empty;
        }
    }
}
