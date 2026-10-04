using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // An upload that reports how far it has got - the Feedback tab's "Sending to server... 42%".
    //
    // *** IT STREAMS (2026-10-03). *** It used to copy the whole inner content into a MemoryStream
    // before sending a byte - harmless at a few MB, but feedback may now carry 250 MB packed
    // (owner request: "one could potentially zip the entire system"), and that was a third copy
    // in memory beside the zip's own two. Now the inner content writes straight through a counting
    // stream to the connection, and the percentage is reported as it moves.
    // Moved out of TabFeedback.axaml.cs for its test (ProgressableStreamContentTests).
    //
    // *** REPORTED EVERY ReportEveryBytes TOO, NOT ONLY PER WHOLE PERCENT (code review,
    // 2026-10-04). *** Each report is the sign of life that keeps the "please wait" overlay's
    // two-minute limit from firing (WaitLimit). One percent of a 250 MB zip is 2.5 MB, which a line
    // under about 170 kbit/s takes longer than two minutes to send - so an upload still moving was
    // cut off. The same percentage may therefore be reported more than once.
    // ###########################################################################################
    public sealed class ProgressableStreamContent : HttpContent
    {
        // A report at least this often, whatever the percentage does: a few seconds of a slow line.
        public const long ReportEveryBytes = 256 * 1024;

        private readonly HttpContent thisInner;
        private readonly Action<int> thisProgress;

        public ProgressableStreamContent(HttpContent innerContent, Action<int> progress)
        {
            ArgumentNullException.ThrowIfNull(innerContent);
            ArgumentNullException.ThrowIfNull(progress);

            this.thisInner = innerContent;
            this.thisProgress = progress;

            foreach (var header in innerContent.Headers)
            {
                this.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
        }

        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) =>
            this.SerializeToStreamAsync(stream, context, CancellationToken.None);

        protected override async Task SerializeToStreamAsync(Stream stream, TransportContext? context, CancellationToken cancellationToken)
        {
            // Not disposed: the connection's stream belongs to HttpClient.
            var counting = new ProgressCountingStream(stream, this.thisInner.Headers.ContentLength ?? -1, this.thisProgress);

            await this.thisInner.CopyToAsync(counting, context, cancellationToken);
        }

        protected override bool TryComputeLength(out long length)
        {
            length = this.thisInner.Headers.ContentLength ?? -1;
            return length != -1;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                this.thisInner.Dispose();
            }

            base.Dispose(disposing);
        }

        // ###########################################################################################
        // Writes through to the real stream, counting - write only, as an upload is.
        // ###########################################################################################
        private sealed class ProgressCountingStream(Stream inner, long total, Action<int> progress) : Stream
        {
            private long thisWritten;
            private long thisReportedAt;
            private int thisLastPercent = -1;

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() => inner.Flush();

            public override Task FlushAsync(CancellationToken cancellationToken) => inner.FlushAsync(cancellationToken);

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count)
            {
                inner.Write(buffer, offset, count);
                this.Counted(count);
            }

            public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
                this.WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();

            public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                await inner.WriteAsync(buffer, cancellationToken);
                this.Counted(buffer.Length);
            }

            private void Counted(int count)
            {
                this.thisWritten += count;

                if (total <= 0)
                    return;

                int percent = (int)Math.Min(100, this.thisWritten * 100 / total);

                if (percent != this.thisLastPercent ||
                    this.thisWritten - this.thisReportedAt >= ProgressableStreamContent.ReportEveryBytes)
                {
                    this.thisLastPercent = percent;
                    this.thisReportedAt = this.thisWritten;
                    progress(percent);
                }
            }
        }
    }
}
