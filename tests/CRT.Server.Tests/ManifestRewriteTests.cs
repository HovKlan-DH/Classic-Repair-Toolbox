using CRT.Server.Handlers.Submissions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ManifestRewrite (code review, 2026-10-13): a manifest rebuilt after a write that let
    // go of the publish lock - a BETA publish, a placement, a promotion - takes the lock again for
    // its scan, so it never describes a tree another publish is half way through writing. The
    // writer is a fake: what is under test is WHEN it may run, and that nothing escapes.
    // ###########################################################################################
    public sealed class ManifestRewriteTests
    {
        [Fact]
        public async Task A_rebuild_waits_for_a_publish_holding_the_lock()
        {
            var publishLock = new PublishLock();
            bool scanned = false;

            Task<int> rebuilding;

            using (await publishLock.EnterAsync())
            {
                rebuilding = ManifestRewrite.RewriteAsync(publishLock, "root", "url", "manifest", (_, _, _) =>
                {
                    scanned = true;
                    return 3;
                });

                await Task.Delay(100);

                Assert.False(scanned);
                Assert.False(rebuilding.IsCompleted);
            }

            Assert.Equal(3, await rebuilding);
            Assert.True(scanned);
        }

        // Its callers run after a write that has already happened: a failure is -1, never a throw.
        [Fact]
        public async Task A_rebuild_that_throws_answers_minus_one()
        {
            int written = await ManifestRewrite.RewriteAsync(
                new PublishLock(), "root", "url", "manifest", (_, _, _) => throw new IOException("The disk went away."));

            Assert.Equal(-1, written);
        }

        // The lock is let go afterwards - the next publish is not kept out for ever.
        [Fact]
        public async Task The_lock_is_free_again_after_a_rebuild()
        {
            var publishLock = new PublishLock();

            await ManifestRewrite.RewriteAsync(publishLock, "root", "url", "manifest", (_, _, _) => 1);

            using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(5));
            using IDisposable held = await publishLock.EnterAsync(timeout.Token);
        }
    }
}
