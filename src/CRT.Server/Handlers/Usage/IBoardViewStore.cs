using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // Where board views are kept - `crt_board_views` (migration 0013). See CRT.Data's
    // BoardViewContract for what a view is and why nothing in it identifies a user.
    //
    // The seam the flows are tested through: MySqlBoardViewStore is the real one, an untested I/O
    // boundary like the other MySql* stores; the tests' fake keeps rows in a list.
    // ###########################################################################################
    public interface IBoardViewStore
    {
        // ###########################################################################################
        // Stores one batch's views - ONCE. False when this batch id was stored before, and nothing
        // is written then: CRT sends a batch again when it never heard the answer, and the second
        // copy must not count its views twice.
        // ###########################################################################################
        Task<bool> RecordAsync(
            Guid batchId,
            IReadOnlyList<BoardViewRow> rows,
            DateTimeOffset receivedUtc,
            CancellationToken cancellationToken = default);

        // Views per system since `sinceUtc`, BETA-source views left out - the Systems list's count.
        Task<IReadOnlyDictionary<string, int>> CountSinceBySystemAsync(
            DateTimeOffset sinceUtc,
            CancellationToken cancellationToken = default);

        // One system's views since `sinceUtc`, counted per UTC day, source and country - what
        // BoardViewStatisticsRules turns into the Systems screen's numbers.
        Task<IReadOnlyList<BoardViewFact>> FactsForSystemAsync(
            string systemId,
            DateTimeOffset sinceUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Deletes every view of one system - a system the administrator deleted (owner decision,
        // 2026-10-03: its statistics go with it, so the Fun facts page stops counting a board that
        // no longer exists). The number of rows deleted.
        // ###########################################################################################
        Task<int> DeleteForSystemAsync(string systemId, CancellationToken cancellationToken = default);
    }

    // One row of crt_board_views, as the flow decided it (names from the published listing, country
    // looked up, text clipped). ViewedUtc is already BoardViewRules.StoredTime. FromLocalNetwork is
    // a view from the server's own network (ServerOptions.CountLocalNetworkBoardViews).
    public sealed record BoardViewRow(
        DateTimeOffset ViewedUtc,
        string SystemId,
        string HardwareName,
        string BoardName,
        string Version,
        string? OsHighlevel,
        string? OsVersion,
        string? Cpu,
        string? CountryCode,
        string? CountryName,
        bool FromBeta,
        bool FromLocalNetwork);

    // How many views one system had on one UTC day, from one source, in one country (null: unknown).
    public sealed record BoardViewFact(DateOnly Day, bool FromBeta, string? CountryCode, string? CountryName, int Views);
}
