using System.Globalization;
using System.Text;
using CRT.Server.Configuration;
using MySqlConnector;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // IBoardViewStore over MariaDB - crt_board_views and crt_board_view_batches (migration 0013).
    // An I/O boundary, untested like the other MySql* stores; every decision is made before it is
    // called (BoardViewFlows) or after it answers (BoardViewStatisticsRules).
    // ###########################################################################################
    public sealed class MySqlBoardViewStore : IBoardViewStore
    {
        // How long a batch id is remembered. CRT stops sending a view after 30 days
        // (BoardViewRules.MaxAge), so a retry never arrives later than that; twice it is margin.
        private static readonly TimeSpan BatchMemory = TimeSpan.FromDays(60);

        private readonly string thisConnectionString;

        public MySqlBoardViewStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public async Task<bool> RecordAsync(
            Guid batchId,
            IReadOnlyList<BoardViewRow> rows,
            DateTimeOffset receivedUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(rows);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            // The batch first: a second copy of it inserts nothing here, and then nothing at all.
            await using (MySqlCommand claim = connection.CreateCommand())
            {
                claim.Transaction = transaction;
                claim.CommandText = "INSERT IGNORE INTO crt_board_view_batches (batchId, receivedUtc) VALUES (@id, @received);";
                claim.Parameters.AddWithValue("@id", batchId.ToByteArray());
                claim.Parameters.AddWithValue("@received", receivedUtc.UtcDateTime);

                if (await claim.ExecuteNonQueryAsync(cancellationToken) == 0)
                {
                    await transaction.RollbackAsync(cancellationToken);
                    return false;
                }
            }

            if (rows.Count > 0)
            {
                await using MySqlCommand insert = connection.CreateCommand();
                insert.Transaction = transaction;

                var sql = new StringBuilder(
                    "INSERT INTO crt_board_views (viewedUtc, systemId, hardwareName, boardName, version, " +
                    "osHighlevel, osVersion, cpu, countryCode, countryName, fromBeta, fromLocalNetwork) VALUES ");

                for (int index = 0; index < rows.Count; index++)
                {
                    BoardViewRow row = rows[index];
                    string n = index.ToString(CultureInfo.InvariantCulture);

                    if (index > 0)
                        sql.Append(", ");

                    sql.Append($"(@at{n}, @system{n}, @hardware{n}, @board{n}, @version{n}, @os{n}, @osVersion{n}, @cpu{n}, @code{n}, @country{n}, @beta{n}, @local{n})");

                    insert.Parameters.AddWithValue($"@at{n}", row.ViewedUtc.UtcDateTime);
                    insert.Parameters.AddWithValue($"@system{n}", row.SystemId);
                    insert.Parameters.AddWithValue($"@hardware{n}", row.HardwareName);
                    insert.Parameters.AddWithValue($"@board{n}", row.BoardName);
                    insert.Parameters.AddWithValue($"@version{n}", row.Version);
                    insert.Parameters.AddWithValue($"@os{n}", (object?)row.OsHighlevel ?? DBNull.Value);
                    insert.Parameters.AddWithValue($"@osVersion{n}", (object?)row.OsVersion ?? DBNull.Value);
                    insert.Parameters.AddWithValue($"@cpu{n}", (object?)row.Cpu ?? DBNull.Value);
                    insert.Parameters.AddWithValue($"@code{n}", (object?)row.CountryCode ?? DBNull.Value);
                    insert.Parameters.AddWithValue($"@country{n}", (object?)row.CountryName ?? DBNull.Value);
                    insert.Parameters.AddWithValue($"@beta{n}", row.FromBeta);
                    insert.Parameters.AddWithValue($"@local{n}", row.FromLocalNetwork);
                }

                insert.CommandText = sql.Append(';').ToString();
                await insert.ExecuteNonQueryAsync(cancellationToken);
            }

            // Forget batch ids nobody can send again - cheap, and keeps the table small for good.
            await using (MySqlCommand forget = connection.CreateCommand())
            {
                forget.Transaction = transaction;
                forget.CommandText = "DELETE FROM crt_board_view_batches WHERE receivedUtc < @cutoff;";
                forget.Parameters.AddWithValue("@cutoff", (receivedUtc - MySqlBoardViewStore.BatchMemory).UtcDateTime);
                await forget.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return true;
        }

        public async Task<IReadOnlyDictionary<string, int>> CountSinceBySystemAsync(
            DateTimeOffset sinceUtc,
            CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                "SELECT systemId, COUNT(*) FROM crt_board_views " +
                "WHERE viewedUtc >= @since AND fromBeta = 0 GROUP BY systemId;";
            command.Parameters.AddWithValue("@since", sinceUtc.UtcDateTime);

            var counts = new Dictionary<string, int>(StringComparer.Ordinal);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                counts[reader.GetString(0)] = (int)reader.GetInt64(1);

            return counts;
        }

        public async Task<IReadOnlyList<BoardViewFact>> FactsForSystemAsync(
            string systemId,
            DateTimeOffset sinceUtc,
            CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                "SELECT DATE(viewedUtc), fromBeta, countryCode, countryName, COUNT(*) FROM crt_board_views " +
                "WHERE systemId = @system AND viewedUtc >= @since " +
                "GROUP BY DATE(viewedUtc), fromBeta, countryCode, countryName;";
            command.Parameters.AddWithValue("@system", systemId);
            command.Parameters.AddWithValue("@since", sinceUtc.UtcDateTime);

            var facts = new List<BoardViewFact>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                facts.Add(new BoardViewFact(
                    DateOnly.FromDateTime(reader.GetDateTime(0)),
                    reader.GetBoolean(1),
                    reader.IsDBNull(2) ? null : reader.GetString(2),
                    reader.IsDBNull(3) ? null : reader.GetString(3),
                    (int)reader.GetInt64(4)));
            }

            return facts;
        }

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            return connection;
        }
    }
}
