using CRT.Server.Handlers.Usage;

namespace CRT.Server.Tests.Fakes
{
    // An in-memory ICheckInStore, so the check-in flow tests with no database: every row it is
    // handed, in order - and, for the API usage screen, the launches it is told to answer with.
    public sealed class FakeCheckInStore : ICheckInStore
    {
        public List<CheckInRow> Rows { get; } = [];

        public List<CheckInLaunchRow> Launches { get; } = [];

        // When set, CountLaunchesAsync throws it - a table that cannot be read.
        public Exception? LaunchFailure { get; set; }

        public DateOnly? FirstDayAsked { get; private set; }

        public Task RecordAsync(CheckInRow row, CancellationToken cancellationToken = default)
        {
            this.Rows.Add(row);
            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<CheckInLaunchRow>> CountLaunchesAsync(DateOnly firstDay, CancellationToken cancellationToken = default)
        {
            this.FirstDayAsked = firstDay;

            if (this.LaunchFailure is not null)
                throw this.LaunchFailure;

            return Task.FromResult<IReadOnlyList<CheckInLaunchRow>>([.. this.Launches]);
        }
    }
}
