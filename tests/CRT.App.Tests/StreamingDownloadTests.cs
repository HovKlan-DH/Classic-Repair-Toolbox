using System.Security.Cryptography;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// StreamingDownload - a download's body copied to its file a chunk at a time and hashed as it goes
// (code review, 2026-10-13: it was read whole into memory under a total time budget). The network
// half (OnlineServices.DownloadFileAsync) needs a server and is not tested, per CLAUDE.md.
public sealed class StreamingDownloadTests
{
    [Fact]
    public async Task Everything_is_copied_and_the_hash_is_the_files_sha256()
    {
        byte[] content = new byte[StreamingDownload.ChunkSize * 3 + 123];
        new Random(7).NextBytes(content);

        using var watchdog = new NoProgressWatchdog(TimeSpan.FromMinutes(5));
        using var source = new MemoryStream(content);
        using var target = new MemoryStream();

        string hash = await StreamingDownload.CopyAndHashAsync(source, target, watchdog);

        Assert.Equal(content, target.ToArray());
        Assert.Equal(Convert.ToHexStringLower(SHA256.HashData(content)), hash);
    }

    [Fact]
    public async Task An_empty_body_hashes_as_empty()
    {
        using var watchdog = new NoProgressWatchdog(TimeSpan.FromMinutes(5));

        string hash = await StreamingDownload.CopyAndHashAsync(new MemoryStream(), new MemoryStream(), watchdog);

        Assert.Equal(Convert.ToHexStringLower(SHA256.HashData(Array.Empty<byte>())), hash);
    }

    // A body that stops delivering is given up when the watchdog runs out - with what had come in
    // still in the target, which the caller deletes (it is the .tmp file, never moved into place).
    [Fact]
    public async Task A_body_that_stops_arriving_is_given_up_as_a_stall()
    {
        var time = new ManualTimeProvider();
        using var watchdog = new NoProgressWatchdog(TimeSpan.FromSeconds(30), time: time);
        using var source = new StallingStream(firstChunk: [1, 2, 3]);
        using var target = new MemoryStream();

        Task<string> copy = StreamingDownload.CopyAndHashAsync(source, target, watchdog);
        await source.Stalled;

        time.Advance(TimeSpan.FromSeconds(31));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => copy);
        Assert.True(watchdog.Stalled);
        Assert.Equal([1, 2, 3], target.ToArray());
    }

    // Delivers one chunk, then waits for ever (until its token is cancelled).
    private sealed class StallingStream(byte[] firstChunk) : Stream
    {
        private readonly TaskCompletionSource thisStalled = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private bool thisDelivered;

        public Task Stalled => this.thisStalled.Task;

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (!this.thisDelivered)
            {
                this.thisDelivered = true;
                firstChunk.CopyTo(buffer);
                return firstChunk.Length;
            }

            this.thisStalled.TrySetResult();
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return 0;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
