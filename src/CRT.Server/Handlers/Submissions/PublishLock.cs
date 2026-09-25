namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // ONE WRITE TO A PUBLISHED TREE AT A TIME (2026-09-25).
    //
    // A production promotion copies files OUT of BETA; a publish writes files INTO BETA. If the two
    // overlapped, the promotion could copy half a publish - the new images with the old workbook -
    // to every user. The promotion also re-checks the BETA content hash the reviewer looked at, and
    // that check means nothing unless no publish can land between it and the copy.
    //
    // A single lock for every write rather than one per system: publishes and promotions are rare,
    // seconds long, and done by a handful of people, so the cost of serialising them is nothing,
    // while a per-system lock would miss the shared files two boards cite.
    //
    // A reviewer's AMENDMENT takes it too (code review, 2026-09-25), though it writes no tree: it
    // rewrites what a publish reads, and landing mid-publish left the tree with the old rows while
    // the database and the contributor's mail said the amended ones were published.
    // ###########################################################################################
    public sealed class PublishLock
    {
        private readonly SemaphoreSlim thisGate = new(1, 1);

        public async Task<IDisposable> EnterAsync(CancellationToken cancellationToken = default)
        {
            await this.thisGate.WaitAsync(cancellationToken).ConfigureAwait(false);

            return new Releaser(this.thisGate);
        }

        private sealed class Releaser(SemaphoreSlim gate) : IDisposable
        {
            private SemaphoreSlim? thisGate = gate;

            public void Dispose() => Interlocked.Exchange(ref this.thisGate, null)?.Release();
        }
    }
}
