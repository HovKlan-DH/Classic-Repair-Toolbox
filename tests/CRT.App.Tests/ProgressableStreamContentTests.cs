using System.IO;
using System.Net.Http;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The Feedback tab's upload progress (2026-10-03, when feedback grew to 250 MB packed): every byte
// goes through unchanged, and the percentage climbs to 100 - per whole percent, and at least every
// ReportEveryBytes on a large upload - streamed, not copied into memory first as it used to be.
// ###########################################################################################
public sealed class ProgressableStreamContentTests
{
    [Fact]
    public async Task Every_byte_goes_through_and_the_percentage_climbs_once_per_percent_to_100()
    {
        byte[] payload = new byte[1_000_000];
        new Random(7).NextBytes(payload);

        var reported = new List<int>();

        using var content = new ProgressableStreamContent(new StreamContent(new MemoryStream(payload), bufferSize: 4096), reported.Add);
        using var target = new MemoryStream();

        await content.CopyToAsync(target, TestContext.Current.CancellationToken);

        Assert.Equal(payload, target.ToArray());
        Assert.Equal(payload.Length, content.Headers.ContentLength);
        Assert.Equal(100, reported[^1]);
        Assert.Equal(reported.Distinct().Count(), reported.Count);
        Assert.Equal(reported.Order(), reported);
    }

    // ###########################################################################################
    // *** A REPORT AT LEAST EVERY ReportEveryBytes (code review, 2026-10-04). *** Each report keeps
    // the overlay's two-minute limit from firing. Once per whole percent was 2.5 MB apart for a
    // 250 MB zip - over two minutes on a slow line - so an upload still moving was cut off. A 100 MB
    // upload here: no two reports further apart than ReportEveryBytes and one write, the
    // percentage never going backwards and ending at 100.
    // ###########################################################################################
    [Fact]
    public async Task A_large_upload_is_reported_at_least_every_ReportEveryBytes_not_only_per_percent()
    {
        const long size = 100L * 1024 * 1024;
        const int chunk = 81920;

        var upload = new ZeroContent(size, chunk);
        var reportedAt = new List<long>();
        var reported = new List<int>();

        using var content = new ProgressableStreamContent(upload, percent =>
        {
            reported.Add(percent);
            reportedAt.Add(upload.Written);
        });

        await content.CopyToAsync(Stream.Null, TestContext.Current.CancellationToken);

        // A report comes with the write that crosses the mark, so at most one write late.
        Assert.True(reported.Count >= size / (ProgressableStreamContent.ReportEveryBytes + chunk), $"{reported.Count} reports.");
        Assert.Equal(reported.Order(), reported);
        Assert.Equal(100, reported[^1]);
        Assert.All(
            reportedAt.Zip(reportedAt.Skip(1), (before, after) => after - before),
            gap => Assert.True(gap <= ProgressableStreamContent.ReportEveryBytes + chunk, $"{gap} bytes between two reports."));
    }

    [Fact]
    public async Task With_no_known_length_nothing_is_reported_and_everything_still_goes_through()
    {
        var reported = new List<int>();

        using var content = new ProgressableStreamContent(new NoLengthContent([1, 2, 3]), reported.Add);
        using var target = new MemoryStream();

        await content.CopyToAsync(target, TestContext.Current.CancellationToken);

        Assert.Equal([1, 2, 3], target.ToArray());
        Assert.Empty(reported);
    }

    // `length` zero bytes in writes of `chunk`, never held in memory; Written is how many have been
    // handed to the stream, counted as each write begins.
    private sealed class ZeroContent(long length, int chunk) : HttpContent
    {
        public long Written { get; private set; }

        protected override async Task SerializeToStreamAsync(Stream stream, System.Net.TransportContext? context)
        {
            byte[] buffer = new byte[chunk];

            while (this.Written < length)
            {
                int count = (int)Math.Min(chunk, length - this.Written);
                this.Written += count;
                await stream.WriteAsync(buffer.AsMemory(0, count));
            }
        }

        protected override bool TryComputeLength(out long computed)
        {
            computed = length;
            return true;
        }
    }

    private sealed class NoLengthContent(byte[] bytes) : HttpContent
    {
        protected override Task SerializeToStreamAsync(Stream stream, System.Net.TransportContext? context) =>
            stream.WriteAsync(bytes, 0, bytes.Length);

        protected override bool TryComputeLength(out long length)
        {
            length = -1;
            return false;
        }
    }
}
