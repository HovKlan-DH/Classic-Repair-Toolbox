using System.Globalization;
using System.Text.Json;
using CRT.Server.Configuration;
using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // MySqlSubmissionStore, part two: the `systems` rows - listing, a system's place in the drop-down
    // lists, its production state and a BETA rollback's bookkeeping. Split out of MySqlSubmissionStore.cs
    // when it passed the project's ~1,500 lines (code review, 2026-09-27); that file's header holds here.
    // ###########################################################################################
    public sealed partial class MySqlSubmissionStore
    {
        public async Task<IReadOnlyList<SystemRecord>> ListSystemsAsync(CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                $"SELECT {MySqlSubmissionStore.SystemColumns} FROM systems ORDER BY system_id;";

            var systems = new List<SystemRecord>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                systems.Add(MySqlSubmissionStore.ReadSystem(reader));

            return systems;
        }

        private const string SystemColumns =
            "system_id, manufacturer, hardware, board, current_revision, is_accepting, " +
            "content_hash, production_revision, production_content_hash, production_published_utc";

        private static SystemRecord ReadSystem(MySqlDataReader reader)
        {
            return new SystemRecord(
                reader.GetString(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetString(3),
                reader.IsDBNull(4) ? null : reader.GetString(4),
                Convert.ToInt32(reader.GetValue(5), CultureInfo.InvariantCulture) != 0,
                reader.IsDBNull(6) ? null : reader.GetString(6),
                reader.IsDBNull(7) ? null : reader.GetString(7),
                reader.IsDBNull(8) ? null : reader.GetString(8),
                MySqlSubmissionStore.ReadUtc(reader, 9));
        }

        public async Task<SystemRecord?> FindSystemAsync(string systemId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                $"SELECT {MySqlSubmissionStore.SystemColumns} FROM systems WHERE system_id = @systemId LIMIT 1;";
            command.Parameters.AddWithValue("@systemId", systemId);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            return await reader.ReadAsync(cancellationToken) ? MySqlSubmissionStore.ReadSystem(reader) : null;
        }

        public async Task<SystemPlacement?> GetPlacementAsync(string systemId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT listing_hardware_name, listing_board_name, listing_notes, listing_after
                  FROM systems
                 WHERE system_id = @systemId
                 LIMIT 1;
                """;
            command.Parameters.AddWithValue("@systemId", systemId);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            // A NULL hardware name is "nobody has placed it" - see 0011_system_listing.sql.
            if (!await reader.ReadAsync(cancellationToken) || reader.IsDBNull(0))
                return null;

            return new SystemPlacement(
                reader.GetString(0),
                reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                reader.IsDBNull(3) ? null : reader.GetString(3));
        }

        public async Task<bool> SetPlacementAsync(
            string systemId,
            SystemPlacement placement,
            long setByAccountId,
            DateTimeOffset setUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(placement);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                UPDATE systems
                   SET listing_hardware_name = @hardware,
                       listing_board_name = @board,
                       listing_notes = @notes,
                       listing_after = @after,
                       listing_set_by = @by,
                       listing_set_utc = @when
                 WHERE system_id = @systemId;
                """;

            command.Parameters.AddWithValue("@hardware", placement.HardwareName);
            command.Parameters.AddWithValue("@board", placement.BoardName);
            command.Parameters.AddWithValue("@notes", placement.Notes ?? string.Empty);
            command.Parameters.AddWithValue("@after", (object?)placement.AfterExcelDataFile ?? DBNull.Value);
            command.Parameters.AddWithValue("@by", setByAccountId);
            command.Parameters.AddWithValue("@when", setUtc.UtcDateTime);
            command.Parameters.AddWithValue("@systemId", systemId);

            // Rows MATCHED, not changed - the client's default (UseAffectedRows is off) - so saving the
            // same placement twice is still "there is such a row".
            return await command.ExecuteNonQueryAsync(cancellationToken) > 0;
        }

        public Task SetSystemInProductionAsync(
            string systemId,
            string? revision,
            string? contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE systems
                   SET production_revision = @revision,
                       production_content_hash = @hash,
                       production_published_utc = @when
                 WHERE system_id = @systemId;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@revision", (object?)revision ?? DBNull.Value);
                    command.Parameters.AddWithValue("@hash", (object?)contentHash ?? DBNull.Value);
                    command.Parameters.AddWithValue("@when", publishedUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@systemId", systemId);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // A BETA rollback's bookkeeping, all or nothing - see ISubmissionStore.RecordRollbackAsync.
        // ###########################################################################################
        public async Task RecordRollbackAsync(
            string systemId,
            IReadOnlyList<long> returningSubmissionIds,
            long decidedByAccountId,
            string comment,
            string? betaRevision,
            string? betaContentHash,
            DateTimeOffset decidedUtc,
            bool reject = false,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(returningSubmissionIds);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            MySqlCommand Command(string sql)
            {
                MySqlCommand command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = sql;
                return command;
            }

            foreach (long id in returningSubmissionIds)
            {
                // Only a row still MERGED moves: one decided again meanwhile is left as it is.
                await using (MySqlCommand row = Command(
                    """
                    UPDATE submissions
                       SET state = @state,
                           decided_utc = @when,
                           decided_by = @by,
                           decision_comment = @comment
                     WHERE id = @id AND state = @merged;
                    """))
                {
                    row.Parameters.AddWithValue("@state", reject ? SubmissionState.Rejected : SubmissionState.Pending);
                    row.Parameters.AddWithValue("@merged", SubmissionState.Merged);
                    row.Parameters.AddWithValue("@when", decidedUtc.UtcDateTime);
                    row.Parameters.AddWithValue("@by", decidedByAccountId);
                    row.Parameters.AddWithValue("@comment", comment);
                    row.Parameters.AddWithValue("@id", id);

                    // Nothing moved (decided again meanwhile): nothing to record either.
                    if (await row.ExecuteNonQueryAsync(cancellationToken) == 0)
                        continue;
                }

                // RETURNED to the queue, recorded (migration 0015) at the same instant as its
                // decided_utc - so it reads as "returned" until a later decision moves that on. Not
                // for a rejection: that is its own final state.
                if (!reject)
                {
                    await using MySqlCommand returned = Command(
                        """
                        INSERT INTO submission_beta_returns (submission_id, returned_utc)
                        VALUES (@id, @when)
                        ON DUPLICATE KEY UPDATE returned_utc = VALUES(returned_utc);
                        """);

                    returned.Parameters.AddWithValue("@id", id);
                    returned.Parameters.AddWithValue("@when", decidedUtc.UtcDateTime);
                    await returned.ExecuteNonQueryAsync(cancellationToken);
                }

                await using (MySqlCommand approvals = Command("DELETE FROM submission_approvals WHERE submission_id = @id;"))
                {
                    approvals.Parameters.AddWithValue("@id", id);
                    await approvals.ExecuteNonQueryAsync(cancellationToken);
                }
            }

            await using (MySqlCommand system = Command(
                """
                UPDATE systems
                   SET current_revision = @revision,
                       content_hash = @hash
                 WHERE system_id = @systemId;
                """))
            {
                system.Parameters.AddWithValue("@revision", (object?)betaRevision ?? DBNull.Value);
                system.Parameters.AddWithValue("@hash", (object?)betaContentHash ?? DBNull.Value);
                system.Parameters.AddWithValue("@systemId", systemId);
                await system.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
        }

        // -----------------------------------------------------------------------------------
        // Approvals (migration 0008).
        // -----------------------------------------------------------------------------------

        // -----------------------------------------------------------------------------------
        // A maintainer's amendment (migration 0009). See ISubmissionStore.AmendAsync.
        // -----------------------------------------------------------------------------------
    }
}
