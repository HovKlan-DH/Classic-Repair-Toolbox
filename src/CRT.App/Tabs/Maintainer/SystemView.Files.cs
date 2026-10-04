using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // A SYSTEM'S FILES (owner request, 2026-10-03: "Files (should not show changed files - just list
    // all files)") - every file the system uses as BETA holds it: its own folder, and the shared
    // files its board cites (the server's SystemFilesFlow). The same tree the two queues show, as a
    // LISTING (FileTreeView.ShowListing): no "Show only changed files", a count of files, and the
    // system's own folder open. Pointing at a file shows it and double-clicking opens it, from BETA's
    // public address, exactly as the queues' trees do (FileTreeFiles).
    //
    // Read when the view is first shown for a system (SystemView.Sections.cs), and kept until another
    // system is chosen - or until BETA's board moves (2026-10-04, CatchUpWithBetaAsync).
    // ###########################################################################################
    public partial class SystemView
    {
        // The system the Files view holds a list for - null for none read yet.
        private string? thisFilesFor;

        // The system as it was when the list was read, held against it as it is now
        // (QueueRefreshRules.BetaBoardChanged).
        private SystemOverviewEntry? thisFilesReadAt;

        private FileTreeView SystemFiles => this.FindControl<FileTreeView>("SystemFileTree")!;

        private bool HoldsFilesFor(string? systemId) =>
            this.thisFilesFor is not null && string.Equals(this.thisFilesFor, systemId, StringComparison.Ordinal);

        // Reads the system on screen's files and shows them. The caller holds the wait.
        private async Task LoadFilesAsync()
        {
            if (this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session ||
                this.ShownSystem is not { } system)
            {
                return;
            }

            ReviewApiResult<SystemFilesAnswer> answer = await client.GetSystemFilesAsync(session, system.SystemId);

            if (!string.Equals(this.ShownSystem?.SystemId, system.SystemId, StringComparison.Ordinal))
                return;

            if (!answer.IsOk)
            {
                this.ShowFiles(null, null, system.SystemId);

                // "Not in BETA" is an answer about the system, not a failure.
                this.SystemFiles.ShowMessage(answer.Message, isError: answer.Failure != ReviewApiFailure.NotFound);
                return;
            }

            this.ShowFiles(answer.Value!, client, system.SystemId);
            this.thisFilesFor = system.SystemId;
            this.thisFilesReadAt = system;
        }

        private void ShowFiles(SystemFilesAnswer? files, ReviewApiClient? client, string systemId)
        {
            FileTreeView tree = this.SystemFiles;

            tree.Files = files is null || client is null
                ? null
                : new FileTreeFiles(this, client, this.thisSession, submissionId: null, files.BetaDataUrl, productionDataUrl: null);

            tree.ShowListing(files?.Files, openFolder: systemId);

            if (this.FindControl<TextBlock>("FilesHeadingText") is TextBlock heading)
                heading.Text = files is null ? string.Empty : SystemSections.FilesHeading;
        }

        // Shows a list without a server - for tests. `readAt` is the system as it was when it was read.
        internal void ShowFilesForTests(SystemFilesAnswer files, SystemOverviewEntry? readAt = null)
        {
            this.ShowFiles(files, client: null, files.SystemId);

            if (readAt is not null)
            {
                this.thisFilesFor = files.SystemId;
                this.thisFilesReadAt = readAt;
            }
        }

        internal bool HoldsFilesForTests(string systemId) => this.HoldsFilesFor(systemId);

        internal FileTreeView FileTreeForTests => this.SystemFiles;

        // Nothing of any system's files - another one chosen, or signed out.
        private void ClearFiles()
        {
            this.thisFilesFor = null;
            this.thisFilesReadAt = null;

            FileTreeView tree = this.SystemFiles;
            tree.Files = null;
            tree.ShowListing(null, openFolder: null);

            if (this.FindControl<TextBlock>("FilesHeadingText") is TextBlock heading)
                heading.Text = string.Empty;
        }
    }
}
