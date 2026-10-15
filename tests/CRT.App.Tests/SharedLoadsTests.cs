using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// SharedLoads - board workbook reads that overlap share one parse (code review, 2026-10-13: the
// catalogue, the launch validation and the board on screen asking for one workbook at once parsed
// it two or three times over).
public sealed class SharedLoadsTests
{
    [Fact]
    public async Task Requests_while_a_load_runs_share_it()
    {
        var loads = new SharedLoads<string>(StringComparer.OrdinalIgnoreCase);
        var gate = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        int started = 0;

        Task<string> Load()
        {
            Interlocked.Increment(ref started);
            return gate.Task;
        }

        Task<string> first = loads.RunAsync("board.xlsx", Load);
        Task<string> second = loads.RunAsync("BOARD.xlsx", Load);

        gate.SetResult("parsed");

        Assert.Equal("parsed", await first);
        Assert.Equal("parsed", await second);
        Assert.Equal(1, started);
    }

    // Once a load is done nothing is remembered - that is BoardDataReader's cache, which knows
    // when a file changed. The next request loads again.
    [Fact]
    public async Task A_request_after_the_load_finished_starts_a_new_one()
    {
        var loads = new SharedLoads<int>(StringComparer.Ordinal);
        int started = 0;

        Task<int> Load() => Task.FromResult(Interlocked.Increment(ref started));

        Assert.Equal(1, await loads.RunAsync("a", Load));
        await WaitUntilNothingRuns(loads);
        Assert.Equal(2, await loads.RunAsync("a", Load));
    }

    [Fact]
    public async Task Different_keys_load_separately()
    {
        var loads = new SharedLoads<string>(StringComparer.Ordinal);

        Task<string> a = loads.RunAsync("a", () => Task.FromResult("A"));
        Task<string> b = loads.RunAsync("b", () => Task.FromResult("B"));

        Assert.Equal("A", await a);
        Assert.Equal("B", await b);
    }

    // A failure reaches everybody who joined - and the next request tries again.
    [Fact]
    public async Task A_failed_load_fails_for_everybody_who_joined_and_is_not_kept()
    {
        var loads = new SharedLoads<string>(StringComparer.Ordinal);
        var gate = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);

        Task<string> first = loads.RunAsync("a", () => gate.Task);
        Task<string> second = loads.RunAsync("a", () => Task.FromResult("never"));

        gate.SetException(new IOException("locked"));

        await Assert.ThrowsAsync<IOException>(() => first);
        await Assert.ThrowsAsync<IOException>(() => second);

        await WaitUntilNothingRuns(loads);
        Assert.Equal("fine", await loads.RunAsync("a", () => Task.FromResult("fine")));
    }

    [Fact]
    public async Task A_load_that_throws_at_once_fails_its_task_rather_than_the_caller()
    {
        var loads = new SharedLoads<string>(StringComparer.Ordinal);

        Task<string> load = loads.RunAsync("a", () => throw new InvalidOperationException("bad"));

        await Assert.ThrowsAsync<InvalidOperationException>(() => load);
    }

    // The entry goes just after the result is set - wait for it rather than race it.
    private static async Task WaitUntilNothingRuns<T>(SharedLoads<T> loads)
    {
        for (int i = 0; i < 200 && loads.RunningCount > 0; i++)
        {
            await Task.Delay(5);
        }

        Assert.Equal(0, loads.RunningCount);
    }
}
