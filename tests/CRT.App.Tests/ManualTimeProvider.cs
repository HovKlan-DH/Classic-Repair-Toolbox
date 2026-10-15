namespace ClassicRepairToolbox.Tests;

// A clock the test moves on by hand (Advance), for code that takes a TimeProvider - a timer fires
// only when the test says time has passed, so nothing depends on how fast the machine is.
internal sealed class ManualTimeProvider : TimeProvider
{
    private readonly object thisSync = new();
    private readonly List<ManualTimer> thisTimers = new();
    private DateTimeOffset thisNow = new(2026, 10, 13, 9, 0, 0, TimeSpan.Zero);

    public override DateTimeOffset GetUtcNow()
    {
        lock (this.thisSync)
        {
            return this.thisNow;
        }
    }

    public override ITimer CreateTimer(TimerCallback callback, object? state, TimeSpan dueTime, TimeSpan period)
    {
        var timer = new ManualTimer(this, callback, state);
        timer.Change(dueTime, period);

        lock (this.thisSync)
        {
            this.thisTimers.Add(timer);
        }

        return timer;
    }

    // Moves time on and fires every timer that came due on the way.
    public void Advance(TimeSpan by)
    {
        List<ManualTimer> due;

        lock (this.thisSync)
        {
            this.thisNow += by;
            due = this.thisTimers.Where(timer => timer.DueAt is { } at && at <= this.thisNow).ToList();
        }

        foreach (ManualTimer timer in due)
        {
            timer.Fire();
        }
    }

    private sealed class ManualTimer(ManualTimeProvider owner, TimerCallback callback, object? state) : ITimer
    {
        public DateTimeOffset? DueAt { get; private set; }

        public bool Change(TimeSpan dueTime, TimeSpan period)
        {
            this.DueAt = dueTime == Timeout.InfiniteTimeSpan ? null : owner.GetUtcNow() + dueTime;
            return true;
        }

        public void Fire()
        {
            this.DueAt = null;
            callback(state);
        }

        public void Dispose() => this.DueAt = null;

        public ValueTask DisposeAsync()
        {
            this.Dispose();
            return ValueTask.CompletedTask;
        }
    }
}
