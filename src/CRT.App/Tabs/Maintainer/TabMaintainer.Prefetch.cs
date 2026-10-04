using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE ENTRY "CONTRIBUTOR SUBMISSIONS" OPENS ON, READ BEFORE THE TAB IS SHOWN (owner request,
    // 2026-10-02: "When I launched CRT app, then it was idle for some minutes. Then I did go into the
    // "Maintainer" tab, as it did have a badge for that, and then I saw that "Wait" screen ... Will
    // it be possible to load the data for this first entry already at application launch, but as the
    // very last thing").
    //
    // WHEN: at launch, after the badge's two lists (TabMaintainer.Session.cs - the last thing CRT's
    // start-up does), and after every off-screen badge check (TabMaintainer.QueueRefresh.cs), which
    // reads again only when the entry to open, or its queue row, changed - SubmissionPrefetch has
    // the rule. Never while the tab is on screen: there the entry is simply opened.
    //
    // QUIETLY: no overlay, nothing on screen changes, and a failure keeps nothing - the open then
    // asks the server under the overlay exactly as before.
    //
    // OPENING IT: the detail and the table are shown at once, with no wait, and then read once more
    // in the background - the queue row says when a submission was decided or its approvals moved,
    // but not when another maintainer saved a change to its table. A table read at a newer version
    // replaces the one on screen, unless something has been typed into it.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private readonly SubmissionPrefetch thisPrefetch = new();
        private bool thisPrefetchRunning;

        // ###########################################################################################
        // Reads the entry "Contributor Submissions" would open on, when the tab is not on screen,
        // that screen is the one it will show, and nothing is open on it already.
        // ###########################################################################################
        private async Task PrefetchEntryToOpenAsync()
        {
            if (this.thisPrefetchRunning ||
                this.IsOnScreen ||
                this.ScreenOnNextOpening != MaintainerMode.Review ||
                this.thisSelectedId is not null ||
                this.IsTableOpen ||
                this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session)
            {
                return;
            }

            ReviewQueueRow? entry = this.EntryToOpenOnReview();

            if (entry is null || !this.thisPrefetch.NeedsReading(entry))
                return;

            this.thisPrefetchRunning = true;

            try
            {
                Task<ReviewApiResult<ReviewSubmissionDetail>> detail = client.GetSubmissionAsync(session, entry.Id);
                Task<ReviewApiResult<ReviewTableData>> table = client.GetTableAsync(session, entry.Id);

                await Task.WhenAll(detail, table);

                // Kept only as a complete pair, for the account that asked, and only while the moment
                // for it is still ahead - the tab shown meanwhile has opened the entry itself.
                if (detail.Result.IsOk && table.Result.IsOk &&
                    ReferenceEquals(this.thisSession, session) &&
                    !this.IsOnScreen &&
                    this.thisSelectedId is null)
                {
                    this.thisPrefetch.Keep(new PrefetchedSubmission(entry, detail.Result.Value!, table.Result.Value!));
                }
            }
            catch (Exception)
            {
                // Nothing kept. Reading ahead is a convenience: the open asks again, under the
                // overlay, and reports whatever goes wrong then.
            }
            finally
            {
                this.thisPrefetchRunning = false;
            }
        }

        // The submission the Review screen opens on now - MaintainerModes.EntryToOpen over the queue
        // as listed - or null for an empty queue.
        private ReviewQueueRow? EntryToOpenOnReview() =>
            this.FindControl<ListBox>("QueueList") is ListBox queue &&
            MaintainerModes.EntryToOpen(this.QueueIdsInListOrder(queue), this.thisRememberedSubmissionId) is long id
                ? this.thisQueue.FirstOrDefault(row => row.Id == id)
                : null;

        // ###########################################################################################
        // Opens `row` from what was read ahead, when that was read for exactly this queue row. True
        // when it did - the caller then has nothing to wait for.
        // ###########################################################################################
        private async Task<bool> TryOpenPrefetchedAsync(ReviewQueueRow row)
        {
            if (this.thisPrefetch.Take(row) is not PrefetchedSubmission ready ||
                this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session)
            {
                return false;
            }

            this.ShowDetail(ready.Detail);

            if (this.thisTableRow?.Id != row.Id)
            {
                this.BeginTable(row, client, session);
                this.ShowTable(ready.Table, message: null);
            }

            await this.ReadPrefetchedAgainAsync(row, ready.Table.Version, client, session);
            return true;
        }

        // ###########################################################################################
        // The one quiet re-read after opening from what was read ahead - see the header. An answer
        // for a submission no longer selected, or a failure, changes nothing: what is on screen was
        // a good answer a moment ago, and the minute check goes on reading the queue.
        // ###########################################################################################
        private async Task ReadPrefetchedAgainAsync(ReviewQueueRow row, int tableVersion, ReviewApiClient client, ReviewSession session)
        {
            try
            {
                Task<ReviewApiResult<ReviewSubmissionDetail>> detail = client.GetSubmissionAsync(session, row.Id);
                Task<ReviewApiResult<ReviewTableData>> table = client.GetTableAsync(session, row.Id);

                await Task.WhenAll(detail, table);

                if (detail.Result.IsOk && this.SelectedQueueRow?.Id == row.Id)
                    this.ShowDetail(detail.Result.Value!);

                // Only over the very table that was read ahead - not one a save has replaced since -
                // and never over something typed into it.
                if (table.Result.IsOk &&
                    table.Result.Value!.Version != tableVersion &&
                    this.thisTableRow?.Id == row.Id &&
                    this.thisTable?.Version == tableVersion &&
                    !this.TableEditor.HasUnsavedChanges)
                {
                    this.ShowTable(table.Result.Value!, message: null);
                }
            }
            catch (Exception)
            {
                // As above: nothing changes on screen.
            }
        }

        // The queue row of what is held, for tests.
        internal ReviewQueueRow? PrefetchedRowForTests => this.thisPrefetch.HeldRow;

        internal Task PrefetchEntryToOpenForTests() => this.PrefetchEntryToOpenAsync();
    }
}
