namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // Where launch check-ins are kept - `crt_update`, the table the app-checkin PHP page wrote
    // until 2026-10-03. NOT one of this service's tables: no migration creates it; the project
    // owner moved it into crt_review by hand (2026-09-27) with crt_statistics and crt_versions, and
    // the Fun facts pages and a nightly statistics job read it. So a row must look exactly like
    // the PHP page's - see CheckInFlow.
    //
    // The seam CheckInFlow is tested through: MySqlCheckInStore is the real one, an untested I/O
    // boundary like the other MySql* stores; the tests' fake keeps rows in a list.
    // ###########################################################################################
    public interface ICheckInStore
    {
        Task RecordAsync(CheckInRow row, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Per version text, the installations (distinct addresses) and launches from the start of
        // `firstDay` (a UTC day) until now - for Account > "API usage" (2026-10-04), the same whole UTC
        // days the route calls beside them are counted over. Counted in the database; no address
        // leaves it.
        // ###########################################################################################
        Task<IReadOnlyList<CheckInLaunchRow>> CountLaunchesAsync(DateOnly firstDay, CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // One row of crt_update, as the flow decided it. Every text is a string, never null - the PHP
    // page always wrote one ("" for anything missing), so a column that is NOT NULL keeps working.
    // The time and versionMajor are not here: the store writes NOW() and 2, as the PHP page did.
    //
    //   IpAddress  - the sender's public address. Stored, unlike a board view's: every Fun facts
    //                chart counts DISTINCT ipaddr, and the Wiki's "Information collected" says so.
    //   Version    - the User-Agent, "CRT 2026.10.0".
    //   ApiJson    - what the country lookup answered, as JSON; "{}" when it found nothing.
    // ###########################################################################################
    public sealed record CheckInRow(
        string IpAddress,
        string Version,
        string OsHighlevel,
        string OsVersion,
        string Cpu,
        string CountryCode,
        string CountryName,
        string ApiJson);
}
