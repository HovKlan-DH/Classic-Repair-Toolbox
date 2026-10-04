using CRT.Server.Handlers.Usage;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // An in-memory IApiUsageStore: every tally added, in order, and the rows ReadSinceAsync answers
    // with. Failure, when set, is thrown by AddAsync with nothing kept - a database that is away.
    // ###########################################################################################
    public sealed class FakeApiUsageStore : IApiUsageStore
    {
        public List<ApiUsageTally> Added { get; } = [];

        public List<ApiUsageRow> Rows { get; } = [];

        public Exception? Failure { get; set; }

        public DateOnly? FirstDayAsked { get; private set; }

        public Task AddAsync(IReadOnlyList<ApiUsageTally> tallies, CancellationToken cancellationToken = default)
        {
            if (this.Failure is not null)
                throw this.Failure;

            this.Added.AddRange(tallies);
            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<ApiUsageRow>> ReadSinceAsync(DateOnly firstDay, CancellationToken cancellationToken = default)
        {
            this.FirstDayAsked = firstDay;
            return Task.FromResult<IReadOnlyList<ApiUsageRow>>([.. this.Rows]);
        }
    }
}
