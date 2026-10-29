using CRT.Server.Configuration;
using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // IOscilloscopeModelStore on MySQL (migration 0024's oscilloscope_submissions). An I/O boundary
    // like the other MySql* stores - its rules are OscilloscopeModelFlows', tested against a fake.
    // ###########################################################################################
    public sealed class MySqlOscilloscopeModelStore : IOscilloscopeModelStore
    {
        private const string Columns =
            "id, created_utc, vendor, series_or_model, values_json, contact_email, contact_name, account_id, " +
            "message, full_test_suite_passed, client_version, state";

        private readonly string thisConnectionString;

        public MySqlOscilloscopeModelStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public async Task<long> AddAsync(NewOscilloscopeModelSubmission submission, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            // The same contributor's earlier model of this name that is still pending is replaced - by
            // the account when signed in, else by the address typed AND the internet connection (an
            // address alone is no proof: code review, 2026-10-29). One in BETA never is: it would
            // leave the list while BETA keeps it. The contributor's rows are found here, their NAMES
            // compared by OscilloscopeModelRules.SameName - never the collation (see the interface).
            var sameName = new List<long>();

            await using (MySqlCommand select = connection.CreateCommand())
            {
                select.Transaction = transaction;
                select.CommandText =
                    "SELECT id, vendor, series_or_model FROM oscilloscope_submissions " +
                    "WHERE state = @pending AND (" + (submission.AccountId is null
                        ? "account_id IS NULL AND contact_email = @email AND sender_address = @sender"
                        : "account_id = @account") + ") FOR UPDATE;";
                select.Parameters.AddWithValue("@pending", OscilloscopeModelStates.Pending);
                select.Parameters.AddWithValue("@sender", (object?)submission.SenderAddress ?? DBNull.Value);
                select.Parameters.AddWithValue("@email", (object?)submission.ContactEmail ?? DBNull.Value);
                select.Parameters.AddWithValue("@account", (object?)submission.AccountId ?? DBNull.Value);

                await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

                while (await reader.ReadAsync(cancellationToken))
                {
                    if (OscilloscopeModelRules.SameName(reader.GetString(1), reader.GetString(2), submission.Vendor, submission.SeriesOrModel))
                        sameName.Add(Convert.ToInt64(reader.GetValue(0)));
                }
            }

            foreach (long replacedId in sameName)
                await MySqlOscilloscopeModelStore.SetStateAsync(connection, transaction, replacedId, OscilloscopeModelStates.Pending, OscilloscopeModelStates.Replaced, cancellationToken);

            long id;

            await using (MySqlCommand insert = connection.CreateCommand())
            {
                insert.Transaction = transaction;
                insert.CommandText =
                    "INSERT INTO oscilloscope_submissions " +
                    "(created_utc, vendor, series_or_model, values_json, contact_email, contact_name, account_id, message, " +
                    "full_test_suite_passed, client_version, sender_address, state) VALUES " +
                    "(@created, @vendor, @model, @values, @email, @name, @account, @message, @passed, @client, @sender, @pending); " +
                    "SELECT LAST_INSERT_ID();";
                insert.Parameters.AddWithValue("@created", submission.CreatedUtc.UtcDateTime);
                insert.Parameters.AddWithValue("@vendor", submission.Vendor);
                insert.Parameters.AddWithValue("@model", submission.SeriesOrModel);
                insert.Parameters.AddWithValue("@values", submission.ValuesJson);
                insert.Parameters.AddWithValue("@email", (object?)submission.ContactEmail ?? DBNull.Value);
                insert.Parameters.AddWithValue("@name", (object?)submission.ContactName ?? DBNull.Value);
                insert.Parameters.AddWithValue("@account", (object?)submission.AccountId ?? DBNull.Value);
                insert.Parameters.AddWithValue("@message", (object?)submission.Message ?? DBNull.Value);
                insert.Parameters.AddWithValue("@passed", submission.FullTestSuitePassed);
                insert.Parameters.AddWithValue("@client", (object?)submission.ClientVersion ?? DBNull.Value);
                insert.Parameters.AddWithValue("@sender", (object?)submission.SenderAddress ?? DBNull.Value);
                insert.Parameters.AddWithValue("@pending", OscilloscopeModelStates.Pending);
                id = Convert.ToInt64(await insert.ExecuteScalarAsync(cancellationToken));
            }

            await transaction.CommitAsync(cancellationToken);

            return id;
        }

        public async Task<int> CountFromAddressSinceAsync(string senderAddress, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = connection.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM oscilloscope_submissions WHERE sender_address = @sender AND created_utc > @since;";
            count.Parameters.AddWithValue("@sender", senderAddress);
            count.Parameters.AddWithValue("@since", sinceUtc.UtcDateTime);

            return Convert.ToInt32(await count.ExecuteScalarAsync(cancellationToken));
        }

        public async Task<IReadOnlyList<OscilloscopeModelSubmissionRecord>> ListPendingAsync(CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText =
                $"SELECT {MySqlOscilloscopeModelStore.Columns} FROM oscilloscope_submissions " +
                "WHERE state = @pending OR state = @inBeta ORDER BY created_utc, id;";
            select.Parameters.AddWithValue("@pending", OscilloscopeModelStates.Pending);
            select.Parameters.AddWithValue("@inBeta", OscilloscopeModelStates.InBeta);

            var records = new List<OscilloscopeModelSubmissionRecord>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                records.Add(MySqlOscilloscopeModelStore.Read(reader));

            return records;
        }

        public async Task<int> CountWaitingAsync(CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = connection.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM oscilloscope_submissions WHERE state = @pending OR state = @inBeta;";
            count.Parameters.AddWithValue("@pending", OscilloscopeModelStates.Pending);
            count.Parameters.AddWithValue("@inBeta", OscilloscopeModelStates.InBeta);

            return Convert.ToInt32(await count.ExecuteScalarAsync(cancellationToken));
        }

        public async Task<OscilloscopeModelSubmissionRecord?> FindAsync(long id, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText = $"SELECT {MySqlOscilloscopeModelStore.Columns} FROM oscilloscope_submissions WHERE id = @id;";
            select.Parameters.AddWithValue("@id", id);

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            return await reader.ReadAsync(cancellationToken) ? MySqlOscilloscopeModelStore.Read(reader) : null;
        }

        public async Task<bool> MarkInBetaAsync(long id, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand update = connection.CreateCommand();

            update.CommandText =
                "UPDATE oscilloscope_submissions SET state = @inBeta WHERE id = @id AND (state = @pending OR state = @replaced);";
            update.Parameters.AddWithValue("@inBeta", OscilloscopeModelStates.InBeta);
            update.Parameters.AddWithValue("@id", id);
            update.Parameters.AddWithValue("@pending", OscilloscopeModelStates.Pending);
            update.Parameters.AddWithValue("@replaced", OscilloscopeModelStates.Replaced);

            return await update.ExecuteNonQueryAsync(cancellationToken) == 1;
        }

        public async Task<IReadOnlyList<OscilloscopeModelSubmissionRecord>> ReplaceInBetaOfNameAsync(string vendor, string seriesOrModel, long exceptId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            var inBeta = new List<OscilloscopeModelSubmissionRecord>();

            await using (MySqlCommand select = connection.CreateCommand())
            {
                select.Transaction = transaction;
                select.CommandText = $"SELECT {MySqlOscilloscopeModelStore.Columns} FROM oscilloscope_submissions WHERE state = @inBeta AND id <> @id FOR UPDATE;";
                select.Parameters.AddWithValue("@inBeta", OscilloscopeModelStates.InBeta);
                select.Parameters.AddWithValue("@id", exceptId);

                await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

                while (await reader.ReadAsync(cancellationToken))
                {
                    OscilloscopeModelSubmissionRecord record = MySqlOscilloscopeModelStore.Read(reader);

                    if (OscilloscopeModelRules.SameName(record.Vendor, record.SeriesOrModel, vendor, seriesOrModel))
                        inBeta.Add(record);
                }
            }

            var replaced = new List<OscilloscopeModelSubmissionRecord>(inBeta.Count);

            foreach (OscilloscopeModelSubmissionRecord record in inBeta)
            {
                if (await MySqlOscilloscopeModelStore.SetStateAsync(connection, transaction, record.Id, OscilloscopeModelStates.InBeta, OscilloscopeModelStates.Replaced, cancellationToken))
                    replaced.Add(record);
            }

            await transaction.CommitAsync(cancellationToken);

            return replaced;
        }

        // One row from `from` to `to` - false when it was no longer `from`.
        private static async Task<bool> SetStateAsync(MySqlConnection connection, MySqlTransaction transaction, long id, string from, string to, CancellationToken cancellationToken)
        {
            await using MySqlCommand update = connection.CreateCommand();

            update.Transaction = transaction;
            update.CommandText = "UPDATE oscilloscope_submissions SET state = @to WHERE id = @id AND state = @from;";
            update.Parameters.AddWithValue("@to", to);
            update.Parameters.AddWithValue("@id", id);
            update.Parameters.AddWithValue("@from", from);

            return await update.ExecuteNonQueryAsync(cancellationToken) == 1;
        }

        public async Task<bool> DecideAsync(long id, string state, long decidedByAccountId, string? note, DateTimeOffset whenUtc, bool evenIfReplaced = false, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand update = connection.CreateCommand();

            update.CommandText =
                "UPDATE oscilloscope_submissions SET state = @state, decided_utc = @when, decided_by_account_id = @by, decision_note = @note " +
                "WHERE id = @id AND (state = @pending OR state = @inBeta" + (evenIfReplaced ? " OR state = @replaced" : string.Empty) + ");";
            update.Parameters.AddWithValue("@inBeta", OscilloscopeModelStates.InBeta);
            update.Parameters.AddWithValue("@replaced", OscilloscopeModelStates.Replaced);
            update.Parameters.AddWithValue("@state", state);
            update.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
            update.Parameters.AddWithValue("@by", decidedByAccountId);
            update.Parameters.AddWithValue("@note", (object?)note ?? DBNull.Value);
            update.Parameters.AddWithValue("@id", id);
            update.Parameters.AddWithValue("@pending", OscilloscopeModelStates.Pending);

            return await update.ExecuteNonQueryAsync(cancellationToken) == 1;
        }

        private static OscilloscopeModelSubmissionRecord Read(MySqlDataReader reader) => new(
            Convert.ToInt64(reader.GetValue(0)),
            new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(1), DateTimeKind.Utc)),
            reader.GetString(2),
            reader.GetString(3),
            reader.GetString(4),
            reader.IsDBNull(5) ? null : reader.GetString(5),
            reader.IsDBNull(6) ? null : reader.GetString(6),
            reader.IsDBNull(7) ? null : Convert.ToInt64(reader.GetValue(7)),
            reader.IsDBNull(8) ? null : reader.GetString(8),
            Convert.ToBoolean(reader.GetValue(9)),
            reader.IsDBNull(10) ? null : reader.GetString(10),
            reader.GetString(11));

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            return connection;
        }
    }
}
