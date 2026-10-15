using System;
using System.Threading;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // A time limit on NO PROGRESS, not on the whole job (code review, 2026-10-13).
    //
    // A data file was downloaded under HttpClient.Timeout - a TOTAL budget of 30 seconds - so any
    // file bigger than about 20 MB never arrived on a slower line, however steadily it came in:
    // six shipped files are bigger than that, two KiCad boards at 44 MB. The manifest fetch met the
    // same trap before (AppConfig.ManifestTimeout). Now a download may take as long as it takes, and
    // gives up only when NOTHING has arrived for the limit: every chunk read calls Progress, which
    // starts the limit over.
    //
    // Token is cancelled when the limit runs out, or when the outer token is; Stalled says which.
    // The clock is a TimeProvider so the tests can move time on by hand.
    // ###########################################################################################
    internal sealed class NoProgressWatchdog : IDisposable
    {
        private readonly TimeSpan thisLimit;
        private readonly CancellationTokenSource thisStall;
        private readonly CancellationTokenSource thisLinked;

        internal NoProgressWatchdog(TimeSpan limit, CancellationToken outer = default, TimeProvider? time = null)
        {
            this.thisLimit = limit;
            this.thisStall = new CancellationTokenSource(limit, time ?? TimeProvider.System);
            this.thisLinked = CancellationTokenSource.CreateLinkedTokenSource(this.thisStall.Token, outer);
        }

        internal CancellationToken Token => this.thisLinked.Token;

        // True when the token was cancelled because nothing arrived for the limit.
        internal bool Stalled => this.thisStall.IsCancellationRequested;

        // Something arrived: the limit starts over. Does nothing once it has run out.
        internal void Progress()
        {
            if (this.thisStall.IsCancellationRequested)
            {
                return;
            }

            try
            {
                this.thisStall.CancelAfter(this.thisLimit);
            }
            catch (ObjectDisposedException)
            {
                // Disposed under a read still finishing - nothing left to watch.
            }
        }

        public void Dispose()
        {
            this.thisLinked.Dispose();
            this.thisStall.Dispose();
        }
    }
}
