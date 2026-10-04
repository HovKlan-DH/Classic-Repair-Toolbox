using CRT.Server.Handlers.Usage;
using CRT.Server.Tests.Fakes;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ApiUsageFlusher.FlushAsync - one write of the API usage counted in memory. What matters:
    // a write that worked empties the counter, and one the database refused loses nothing - the
    // tallies wait in the counter for the next write - and never throws (an exception escaping a
    // background service stops the whole contribution service).
    // ###########################################################################################
    public sealed class ApiUsageFlusherTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

        [Fact]
        public async Task A_write_that_works_stores_every_tally_and_empties_the_counter()
        {
            var counter = new ApiUsageCounter();
            var store = new FakeApiUsageStore();

            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageFlusherTests.Now);
            counter.Count("POST", "/api/usage/check-in", "2.5.0", ApiUsageFlusherTests.Now);

            Assert.True(await ApiUsageFlusher.FlushAsync(counter, store, NullLogger.Instance));

            Assert.Equal(2, store.Added.Count);
            Assert.Empty(counter.TakeAll());
        }

        [Fact]
        public async Task A_write_the_database_refuses_keeps_every_tally_for_the_next_one()
        {
            var counter = new ApiUsageCounter();
            var store = new FakeApiUsageStore { Failure = new InvalidOperationException("The database is away.") };

            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageFlusherTests.Now);
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageFlusherTests.Now);

            Assert.False(await ApiUsageFlusher.FlushAsync(counter, store, NullLogger.Instance));
            Assert.Empty(store.Added);

            store.Failure = null;

            Assert.True(await ApiUsageFlusher.FlushAsync(counter, store, NullLogger.Instance));
            Assert.Equal(2, Assert.Single(store.Added).Calls);
        }

        // ###########################################################################################
        // *** A WRITE WAITS WHILE A RESET HOLDS THE WRITES (code review, 2026-10-04). *** A reset of
        // the contribution data holds them around its transaction; a write that took its tallies
        // meanwhile would insert calls from before the reset after it. Given up while it waits, it
        // writes nothing and the tallies stay counted.
        // ###########################################################################################
        [Fact]
        public async Task A_write_waits_while_the_writes_are_held_and_a_write_given_up_keeps_the_tallies()
        {
            var counter = new ApiUsageCounter();
            var store = new FakeApiUsageStore();

            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageFlusherTests.Now);

            IDisposable reset = await counter.HoldWritesAsync();

            using (var giveUp = new CancellationTokenSource(TimeSpan.FromMilliseconds(50)))
                Assert.False(await ApiUsageFlusher.FlushAsync(counter, store, NullLogger.Instance, giveUp.Token));

            Task<bool> waiting = ApiUsageFlusher.FlushAsync(counter, store, NullLogger.Instance);

            await Task.Delay(100);
            Assert.False(waiting.IsCompleted);
            Assert.Empty(store.Added);

            reset.Dispose();

            Assert.True(await waiting.WaitAsync(TimeSpan.FromSeconds(10)));
            Assert.Equal(1, Assert.Single(store.Added).Calls);
        }

        // Nothing counted, nothing written - not even an empty write.
        [Fact]
        public async Task With_nothing_counted_nothing_is_written()
        {
            var store = new FakeApiUsageStore { Failure = new InvalidOperationException("Must not be asked.") };

            Assert.True(await ApiUsageFlusher.FlushAsync(new ApiUsageCounter(), store, NullLogger.Instance));
        }
    }
}
