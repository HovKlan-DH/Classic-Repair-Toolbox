using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // THE SUBMISSION "CONTRIBUTOR SUBMISSIONS" WILL OPEN ON, READ BEFORE ANYBODY ASKS (owner request,
    // 2026-10-02: "Will it be possible to load the data for this first entry already at application
    // launch, but as the very last thing, after UI has launched and everything? Then there would be
    // no wait screen shown to the user").
    //
    // Opening a submission asks the server for two things - its detail and its table - under the
    // "please wait" overlay. The screen opens on a known entry (MaintainerModes.EntryToOpen), so
    // both can be asked for while the tab is not on screen: at launch after the badge's lists, and
    // again whenever the badge's minute check finds that entry changed. Opening it then needs no
    // request at all, and no overlay.
    //
    // *** HELD FOR EXACTLY ONE QUEUE ROW. *** The answers are kept with the queue row they were read
    // for, and used only while the queue still says EXACTLY that (a record: id, state, waiting-for,
    // a discarded draft, ...). Anything the queue reports changing makes them stale, and the next
    // check reads them again.
    //
    // *** USED ONCE. *** Taking them empties the holder, whichever submission was opened: from then
    // on that submission is open and reads itself like any other (and the tab re-reads it quietly
    // once, for what the queue row cannot show - an amendment by another maintainer).
    // ###########################################################################################
    public sealed record PrefetchedSubmission(ReviewQueueRow Row, ReviewSubmissionDetail Detail, ReviewTableData Table);

    public sealed class SubmissionPrefetch
    {
        private PrefetchedSubmission? thisHeld;

        // The queue row the held answers were read for - null when nothing is held.
        public ReviewQueueRow? HeldRow => this.thisHeld?.Row;

        // ###########################################################################################
        // Whether `entry` - the submission the screen would open on now - needs reading: there is
        // one, and what is held was not read for exactly that row.
        // ###########################################################################################
        public bool NeedsReading(ReviewQueueRow? entry) =>
            entry is not null && !Equals(this.thisHeld?.Row, entry);

        public void Keep(PrefetchedSubmission prefetched) => this.thisHeld = prefetched;

        // ###########################################################################################
        // The answers for `row`, when they were read for exactly that queue row - and the holder is
        // empty afterwards either way: a different submission being opened means the moment for
        // the held one has passed.
        // ###########################################################################################
        public PrefetchedSubmission? Take(ReviewQueueRow row)
        {
            PrefetchedSubmission? held = this.thisHeld;
            this.thisHeld = null;

            return held is not null && Equals(held.Row, row) ? held : null;
        }

        // Signed out, or another account: answers judged for the last one are worthless.
        public void Forget() => this.thisHeld = null;
    }
}
