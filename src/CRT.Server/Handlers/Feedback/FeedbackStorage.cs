namespace CRT.Server.Handlers.Feedback
{
    // ###########################################################################################
    // WHAT EVERY FEEDBACK WITH FILES SHARES (code review, 2026-10-04) - one per service:
    //
    //   - ONE UNPACK AT A TIME. Unpacking is synchronous file work - up to 2 GB of it - and each
    //     one held a request thread for its whole length, so a few large feedbacks at once tied up
    //     the threads of a small box. Feedback is rare (ten an hour per address); waiting a turn
    //     costs nobody anything. It also means two feedbacks never both pass the folder total's
    //     check and then both write - because what an unpack kept is added (Saved) BEFORE the gate
    //     lets the next one in, not after the mail (code review, 2026-10-04), and taken off again
    //     (Removed) when the mail fails and the files are deleted.
    //
    //   - THE FOLDER'S TOTAL, REMEMBERED. Each feedback used to stat every file of every saved
    //     feedback (hundreds of thousands near the 20 GiB cap) to total the folder. The total is now
    //     walked at most every RescanAfter, and what each feedback saves is added to it meanwhile.
    //     Walked again on that interval, because the project owner deletes old feedback through a
    //     network share, which this service never hears about - so room made there counts within
    //     a few minutes.
    //
    // The walk and the clock are handed in, so a test counts the walks and moves the time.
    // ###########################################################################################
    public sealed class FeedbackStorage
    {
        public static readonly TimeSpan RescanAfter = TimeSpan.FromMinutes(5);

        private readonly Func<long?> thisWalk;
        private readonly Func<DateTimeOffset> thisClock;
        private readonly object thisLock = new();

        private long? thisTotal;
        private DateTimeOffset thisWalkedAt;

        public FeedbackStorage(Func<long?> walk, Func<DateTimeOffset>? clock = null)
        {
            ArgumentNullException.ThrowIfNull(walk);

            this.thisWalk = walk;
            this.thisClock = clock ?? (() => DateTimeOffset.UtcNow);
        }

        public SemaphoreSlim UnpackGate { get; } = new(1, 1);

        // ###########################################################################################
        // What the saved feedback takes now: the remembered total while it is fresh, else a new walk.
        // Null when the folder cannot be read - an unknown total is room (FeedbackFlow.StoredBytesUnder),
        // and is not remembered, so the next feedback walks again.
        // ###########################################################################################
        public long? StoredBytes()
        {
            lock (this.thisLock)
            {
                DateTimeOffset now = this.thisClock();

                if (this.thisTotal is long remembered && now - this.thisWalkedAt < FeedbackStorage.RescanAfter)
                    return remembered;

                this.thisTotal = this.thisWalk();
                this.thisWalkedAt = now;

                return this.thisTotal;
            }
        }

        // A feedback's files were kept: they count until the next walk counts them itself.
        public void Saved(long bytes)
        {
            lock (this.thisLock)
            {
                if (this.thisTotal is long remembered)
                    this.thisTotal = remembered + Math.Max(0, bytes);
            }
        }

        // A feedback's files were deleted again (its mail failed): they stop counting at once.
        public void Removed(long bytes)
        {
            lock (this.thisLock)
            {
                if (this.thisTotal is long remembered)
                    this.thisTotal = Math.Max(0, remembered - Math.Max(0, bytes));
            }
        }
    }
}
