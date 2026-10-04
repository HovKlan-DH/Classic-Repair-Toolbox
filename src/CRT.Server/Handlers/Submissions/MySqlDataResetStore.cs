using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using MySqlConnector;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // IDataResetStore over MariaDB - every table the contribution service fills about PEOPLE, emptied
    // in one transaction (owner request, 2026-10-04). An I/O boundary, untested like the other MySql*
    // stores; what goes and what stays is decided in DataResetFlow's header, and this is its SQL.
    //
    // *** DELETE, NEVER TRUNCATE. *** TRUNCATE starts the ids again at 1, and installed CRTs still
    // hold receipts naming submission ids: a new submission #12 must never be mistaken for the
    // deleted #12. (The status check is proved by the submission's own token, so a reused id would
    // answer "not found" anyway - but nothing else would have to be right for that to hold.)
    // TRUNCATE also refuses a table other tables point at, and is not transactional.
    //
    // THE ORDER follows the foreign keys, though ON DELETE CASCADE / SET NULL would cope with any:
    // submissions first (their files, payloads, findings, approvals, amendments, discards, BETA
    // returns and change summaries go with them - CASCADE), then what hangs off systems, then the
    // systems, then the accounts (their sessions and tokens go with them - CASCADE), then the
    // history and the usage tables. The administrators' accounts and sessions stay.
    //
    // NOT TOUCHED: crt_board_view_batches (the batch ids already stored - kept, so a CRT sending a
    // stored batch again after the reset is still not counted twice), crt_update and the other
    // check-in tables (no migration owns them, and they hold CRT 2.x's history), schema_migrations.
    // ###########################################################################################
    public sealed class MySqlDataResetStore : IDataResetStore
    {
        private readonly string thisConnectionString;

        public MySqlDataResetStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public async Task<DataResetCounts> CountAsync(CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);

            return await MySqlDataResetStore.CountAsync(connection, transaction: null, cancellationToken);
        }

        public async Task<DataResetResult?> ResetAsync(
            Func<DataResetCounts, bool> isAsShown,
            Func<DataResetCounts, AuditEntry> makeRecord,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(isAsShown);
            ArgumentNullException.ThrowIfNull(makeRecord);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);

            // REPEATABLE READ (MariaDB's default, named so nothing depends on the server's setting):
            // the locking reads below then take gap locks too, which is what keeps a NEW row out.
            await using MySqlTransaction transaction =
                await connection.BeginTransactionAsync(System.Data.IsolationLevel.RepeatableRead, cancellationToken);

            // ###########################################################################################
            // *** EVERY TABLE THE FINGERPRINT HOLDS IS LOCKED BEFORE IT IS COUNTED (code review,
            // 2026-10-04). *** A plain count reads a snapshot, and DELETE then removes whatever is
            // there by then - a submission or an account committed in between would be deleted
            // unseen. A locking read of every row holds off other writers to these tables until the
            // commit, so the counts below are exactly what the deletes remove. A writer that waits is
            // a contributor's create or an invitation accepted, for the length of the reset.
            // ###########################################################################################
            string[] locks =
            [
                "SELECT 1 FROM submissions FOR UPDATE;",
                "SELECT 1 FROM accounts WHERE is_administrator = 0 FOR UPDATE;",
                "SELECT 1 FROM maintainers FOR UPDATE;",
                "SELECT 1 FROM maintainer_invitations FOR UPDATE;",
                "SELECT 1 FROM production_approvals FOR UPDATE;",
                "SELECT 1 FROM systems FOR UPDATE;"
            ];

            foreach (string sql in locks)
            {
                await using MySqlCommand lockRows = connection.CreateCommand();
                lockRows.Transaction = transaction;
                lockRows.CommandText = sql;

                await lockRows.ExecuteNonQueryAsync(cancellationToken);
            }

            DataResetCounts before = await MySqlDataResetStore.CountAsync(connection, transaction, cancellationToken);

            if (!isAsShown(before))
            {
                await transaction.RollbackAsync(CancellationToken.None);
                return null;
            }

            AuditEntry record = makeRecord(before);

            var submissionIds = new List<long>();

            await using (MySqlCommand ids = connection.CreateCommand())
            {
                ids.Transaction = transaction;
                ids.CommandText = "SELECT id FROM submissions;";

                await using MySqlDataReader reader = await ids.ExecuteReaderAsync(cancellationToken);

                while (await reader.ReadAsync(cancellationToken))
                    submissionIds.Add(Convert.ToInt64(reader.GetValue(0)));
            }

            string[] deletes =
            [
                "DELETE FROM submissions;",
                "DELETE FROM production_approvals;",
                "DELETE FROM maintainer_invitations;",
                "DELETE FROM maintainers;",
                "DELETE FROM systems;",
                "DELETE FROM accounts WHERE is_administrator = 0;",
                "DELETE FROM audit;",
                "DELETE FROM crt_board_views;",
                "DELETE FROM crt_api_calls;"
            ];

            foreach (string sql in deletes)
            {
                await using MySqlCommand delete = connection.CreateCommand();
                delete.Transaction = transaction;
                delete.CommandText = sql;

                await delete.ExecuteNonQueryAsync(cancellationToken);
            }

            // The first row of the new history: who reset, and what went.
            await using (MySqlCommand audit = connection.CreateCommand())
            {
                audit.Transaction = transaction;
                audit.CommandText = """
                    INSERT INTO audit (actor_account_id, actor_label, action, subject, detail, at_utc)
                    VALUES (@actorId, @actorLabel, @action, @subject, @detail, @at);
                    """;
                audit.Parameters.AddWithValue("@actorId", (object?)record.ActorAccountId ?? DBNull.Value);
                audit.Parameters.AddWithValue("@actorLabel", record.ActorLabel);
                audit.Parameters.AddWithValue("@action", record.Action);
                audit.Parameters.AddWithValue("@subject", (object?)record.Subject ?? DBNull.Value);
                audit.Parameters.AddWithValue("@detail", (object?)record.Detail ?? DBNull.Value);
                audit.Parameters.AddWithValue("@at", record.AtUtc.UtcDateTime);

                await audit.ExecuteNonQueryAsync(cancellationToken);
            }

            // Not cancellable: a commit the server carried out but the client was told was
            // cancelled would leave the deleted submissions' partial uploads uncleared for good.
            await transaction.CommitAsync(CancellationToken.None);

            return new DataResetResult(before, submissionIds);
        }

        private static async Task<DataResetCounts> CountAsync(
            MySqlConnection connection,
            MySqlTransaction? transaction,
            CancellationToken cancellationToken)
        {
            await using MySqlCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = """
                SELECT
                    (SELECT COUNT(*) FROM submissions),
                    (SELECT COALESCE(MAX(id), 0) FROM submissions),
                    (SELECT COUNT(*) FROM accounts WHERE is_administrator = 0),
                    (SELECT COALESCE(MAX(id), 0) FROM accounts WHERE is_administrator = 0),
                    (SELECT COUNT(*) FROM accounts WHERE is_administrator <> 0),
                    (SELECT COUNT(*) FROM maintainers),
                    (SELECT COUNT(*) FROM maintainer_invitations),
                    (SELECT COUNT(*) FROM production_approvals),
                    (SELECT COUNT(*) FROM systems),
                    (SELECT COUNT(*) FROM audit),
                    (SELECT COUNT(*) FROM crt_board_views),
                    (SELECT COUNT(*) FROM crt_api_calls);
                """;

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            if (!await reader.ReadAsync(cancellationToken))
                return new DataResetCounts(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            int Count(int column) => Convert.ToInt32(reader.GetValue(column));
            long Id(int column) => Convert.ToInt64(reader.GetValue(column));

            return new DataResetCounts(
                Submissions: Count(0),
                LastSubmissionId: Id(1),
                Accounts: Count(2),
                LastAccountId: Id(3),
                Administrators: Count(4),
                Maintainers: Count(5),
                Invitations: Count(6),
                ProductionApprovals: Count(7),
                Systems: Count(8),
                HistoryEntries: Count(9),
                BoardViews: Count(10),
                ApiUsageRows: Count(11));
        }

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            return connection;
        }
    }
}
