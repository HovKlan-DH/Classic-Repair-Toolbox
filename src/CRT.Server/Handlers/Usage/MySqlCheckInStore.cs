using CRT.Server.Configuration;
using MySqlConnector;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // ICheckInStore over MariaDB - crt_update. An I/O boundary, untested like the other MySql*
    // stores; every decision is made before it is called (CheckInFlow).
    //
    // *** THE INSERT CHECK-INS HAVE ALWAYS BEEN WRITTEN WITH, word for word. *** createDateTime is
    // the database's NOW() - the server's local time, as every row before this one - and
    // versionMajor is 2, which is what tells this CRT's check-ins from the old Commodore Repair
    // Toolbox's (1) in the same table.
    // ###########################################################################################
    public sealed class MySqlCheckInStore : ICheckInStore
    {
        private readonly string thisConnectionString;

        public MySqlCheckInStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public async Task RecordAsync(CheckInRow row, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(row);

            await using var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            await using MySqlCommand insert = connection.CreateCommand();

            insert.CommandText =
                "INSERT INTO crt_update SET createDateTime = NOW(), ipaddr = @ip, versionMajor = 2, version = @version, " +
                "osHighlevel = @os, osVersion = @osVersion, cpu = @cpu, countryCode = @code, countryName = @country, apiJson = @json;";

            insert.Parameters.AddWithValue("@ip", row.IpAddress);
            insert.Parameters.AddWithValue("@version", row.Version);
            insert.Parameters.AddWithValue("@os", row.OsHighlevel);
            insert.Parameters.AddWithValue("@osVersion", row.OsVersion);
            insert.Parameters.AddWithValue("@cpu", row.Cpu);
            insert.Parameters.AddWithValue("@code", row.CountryCode);
            insert.Parameters.AddWithValue("@country", row.CountryName);
            insert.Parameters.AddWithValue("@json", row.ApiJson);

            await insert.ExecuteNonQueryAsync(cancellationToken);
        }

        // ###########################################################################################
        // This CRT's launches only (versionMajor 2 - the old Commodore Repair Toolbox's are 1), from
        // midnight UTC of `firstDay` on. The addresses are counted here and never read out.
        //
        // *** THE SAME UTC DAYS AS THE ROUTE CALLS (code review, 2026-10-04). *** createDateTime is
        // written with NOW(), the DATABASE's local time, while crt_api_calls counts whole UTC days -
        // and this used to be a rolling "NOW() - INTERVAL n DAY", so with one day asked for, the
        // routes covered today since UTC midnight and the launches the last 24 hours in local time.
        // The UTC midnight is turned into the database's local time by its own current offset
        // (NOW() and UTC_TIMESTAMP() are the same instant within a statement), which keeps the
        // index on createDateTime usable; across a daylight saving change inside the window the
        // first day's edge is an hour off, which a count over days does not notice.
        // ###########################################################################################
        public async Task<IReadOnlyList<CheckInLaunchRow>> CountLaunchesAsync(DateOnly firstDay, CancellationToken cancellationToken = default)
        {
            await using var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT version, COUNT(DISTINCT ipaddr), COUNT(*)
                FROM crt_update
                WHERE versionMajor = 2
                  AND createDateTime >= TIMESTAMPADD(SECOND, TIMESTAMPDIFF(SECOND, UTC_TIMESTAMP(), NOW()), @firstUtc)
                GROUP BY version;
                """;
            command.Parameters.AddWithValue("@firstUtc", firstDay.ToDateTime(TimeOnly.MinValue));

            var rows = new List<CheckInLaunchRow>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                rows.Add(new CheckInLaunchRow(
                    reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                    Convert.ToInt32(reader.GetValue(1)),
                    Convert.ToInt32(reader.GetValue(2))));
            }

            return rows;
        }
    }
}
