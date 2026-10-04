using CRT.Server.Configuration;
using MySqlConnector;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // IApiUsageStore over MariaDB - crt_api_calls (migration 0018). An I/O boundary, untested like
    // the other MySql* stores; what is counted is decided before it is called (ApiUsageRules) and
    // what is shown after it answers (ApiUsageRules.Build).
    // ###########################################################################################
    public sealed class MySqlApiUsageStore : IApiUsageStore
    {
        private readonly string thisConnectionString;

        public MySqlApiUsageStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public async Task AddAsync(IReadOnlyList<ApiUsageTally> tallies, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(tallies);

            if (tallies.Count == 0)
                return;

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            foreach (ApiUsageTally tally in tallies)
            {
                await using MySqlCommand command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO crt_api_calls (callDate, method, route, version, calls, lastUtc)
                    VALUES (@day, @method, @route, @version, @calls, @last)
                    ON DUPLICATE KEY UPDATE
                        calls = calls + VALUES(calls),
                        lastUtc = GREATEST(lastUtc, VALUES(lastUtc));
                    """;

                command.Parameters.AddWithValue("@day", tally.Day.ToDateTime(TimeOnly.MinValue));
                command.Parameters.AddWithValue("@method", tally.Method);
                command.Parameters.AddWithValue("@route", tally.Route);
                command.Parameters.AddWithValue("@version", tally.Version);
                command.Parameters.AddWithValue("@calls", tally.Calls);
                command.Parameters.AddWithValue("@last", tally.LastUtc.UtcDateTime);

                await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
        }

        public async Task<IReadOnlyList<ApiUsageRow>> ReadSinceAsync(DateOnly firstDay, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT method, route, version, SUM(calls), MAX(lastUtc)
                FROM crt_api_calls
                WHERE callDate >= @first
                GROUP BY method, route, version;
                """;
            command.Parameters.AddWithValue("@first", firstDay.ToDateTime(TimeOnly.MinValue));

            var rows = new List<ApiUsageRow>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                rows.Add(new ApiUsageRow(
                    reader.GetString(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    Convert.ToInt64(reader.GetValue(3)),
                    new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(4), DateTimeKind.Utc))));
            }

            return rows;
        }

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            return connection;
        }
    }
}
