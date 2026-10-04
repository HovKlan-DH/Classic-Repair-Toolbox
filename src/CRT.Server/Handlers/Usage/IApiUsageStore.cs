namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // Where the API usage counts are kept - crt_api_calls (migration 0018). MySqlApiUsageStore is
    // the real one, an untested I/O boundary like the other MySql* stores; the tests' fake keeps
    // rows in a list. Emptied by a reset of the contribution data (MySqlDataResetStore).
    // ###########################################################################################
    public interface IApiUsageStore
    {
        // Adds each tally to its day's row, making the row when it is the first of that day.
        Task AddAsync(IReadOnlyList<ApiUsageTally> tallies, CancellationToken cancellationToken = default);

        // Every route and version called on `firstDay` or later, the days added up.
        Task<IReadOnlyList<ApiUsageRow>> ReadSinceAsync(DateOnly firstDay, CancellationToken cancellationToken = default);
    }
}
