using System.Globalization;
using System.Text.Json;
using CRT.Server.Configuration;
using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // MySqlSubmissionStore, part three: a maintainer's amendments and the approvals given for BETA and
    // for production. Split out of MySqlSubmissionStore.cs when it passed the project's ~1,500 lines
    // (code review, 2026-09-27); that file's header holds here.
    // ###########################################################################################
    public sealed partial class MySqlSubmissionStore
    {
        public async Task<AmendStoreResult> AmendAsync(
            long submissionId,
            int expectedVersion,
            SubmissionManifest amended,
            bool touchesSharedFiles,
            long? accountId,
            string accountLabel,
            DateTimeOffset amendedUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(amended);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            MySqlCommand Command(string sql)
            {
                MySqlCommand command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = sql;
                command.Parameters.AddWithValue("@id", submissionId);
                return command;
            }

            // ---- Still amendable, and still at the version the maintainer opened? ------------------
            //
            // Both read with LOCKING reads, after which nothing another transaction commits can
            // change the answer before this one does: the submission row first, so a second
            // amendment - or a decision - waits here until this one is done, then reads what it
            // left. Leaving without committing rolls the transaction back, so a refusal changes
            // nothing. (code review, 2026-09-25)
            string? state;

            await using (MySqlCommand read = Command("SELECT state FROM submissions WHERE id = @id FOR UPDATE;"))
            {
                state = await read.ExecuteScalarAsync(cancellationToken) as string;
            }

            if (!SubmissionState.CanBeAmended(state))
                return new AmendStoreResult(AmendStoreOutcome.NotAmendable, 0);

            int latestVersion;

            await using (MySqlCommand read = Command(
                "SELECT version FROM submission_amendments WHERE submission_id = @id ORDER BY version DESC LIMIT 1 FOR UPDATE;"))
            {
                object? latest = await read.ExecuteScalarAsync(cancellationToken);
                latestVersion = latest is null or DBNull ? 0 : Convert.ToInt32(latest, CultureInfo.InvariantCulture);
            }

            if (latestVersion != expectedVersion)
                return new AmendStoreResult(AmendStoreOutcome.VersionChanged, latestVersion);

            // What is replaced, read under a lock so two amendments cannot interleave.
            string previousRows;

            await using (MySqlCommand read = Command("SELECT rows_json FROM submission_payloads WHERE submission_id = @id FOR UPDATE;"))
            {
                previousRows = await read.ExecuteScalarAsync(cancellationToken) as string ?? "{}";
            }

            var previousFiles = new List<SubmissionFile>();

            await using (MySqlCommand read = Command("SELECT path, sha256, size_bytes FROM submission_files WHERE submission_id = @id ORDER BY path;"))
            await using (MySqlDataReader reader = await read.ExecuteReaderAsync(cancellationToken))
            {
                while (await reader.ReadAsync(cancellationToken))
                {
                    previousFiles.Add(new SubmissionFile
                    {
                        Path = reader.GetString(0),
                        Sha256 = reader.GetString(1),
                        SizeBytes = Convert.ToInt64(reader.GetValue(2), CultureInfo.InvariantCulture)
                    });
                }
            }

            int version = latestVersion + 1;

            await using (MySqlCommand keep = Command(
                """
                INSERT INTO submission_amendments
                    (submission_id, version, previous_rows_json, previous_files_json, account_id, account_label, amended_utc)
                VALUES (@id, @version, @rows, @files, @accountId, @label, @when);
                """))
            {
                keep.Parameters.AddWithValue("@version", version);
                keep.Parameters.AddWithValue("@rows", previousRows);
                keep.Parameters.AddWithValue("@files", JsonSerializer.Serialize(previousFiles, MySqlSubmissionStore.JsonOptions));
                keep.Parameters.AddWithValue("@accountId", (object?)accountId ?? DBNull.Value);
                keep.Parameters.AddWithValue("@label", MySqlSubmissionStore.Cut(accountLabel, 255));
                keep.Parameters.AddWithValue("@when", amendedUtc.UtcDateTime);
                await keep.ExecuteNonQueryAsync(cancellationToken);
            }

            await using (MySqlCommand rows = Command("UPDATE submission_payloads SET rows_json = @rows WHERE submission_id = @id;"))
            {
                rows.Parameters.AddWithValue("@rows", JsonSerializer.Serialize(amended.Rows, MySqlSubmissionStore.JsonOptions));
                await rows.ExecuteNonQueryAsync(cancellationToken);
            }

            await using (MySqlCommand clear = Command("DELETE FROM submission_files WHERE submission_id = @id;"))
            {
                await clear.ExecuteNonQueryAsync(cancellationToken);
            }

            // Every file is already held - the caller imported any new one - so each is uploaded.
            foreach (SubmissionFile file in amended.Files)
            {
                await using MySqlCommand insert = Command(
                    """
                    INSERT INTO submission_files (submission_id, path, sha256, size_bytes, is_uploaded)
                    VALUES (@id, @path, @sha, @size, 1);
                    """);
                insert.Parameters.AddWithValue("@path", file.Path);
                insert.Parameters.AddWithValue("@sha", file.Sha256);
                insert.Parameters.AddWithValue("@size", file.SizeBytes);
                await insert.ExecuteNonQueryAsync(cancellationToken);
            }

            await using (MySqlCommand approvals = Command("DELETE FROM submission_approvals WHERE submission_id = @id;"))
            {
                await approvals.ExecuteNonQueryAsync(cancellationToken);
            }

            await using (MySqlCommand row = Command(
                """
                UPDATE submissions
                   SET touches_shared_files = @touches,
                       state = CASE WHEN state = @approved THEN @pending ELSE state END
                 WHERE id = @id;
                """))
            {
                row.Parameters.AddWithValue("@touches", touchesSharedFiles);
                row.Parameters.AddWithValue("@approved", SubmissionState.Approved);
                row.Parameters.AddWithValue("@pending", SubmissionState.Pending);
                await row.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);

            return new AmendStoreResult(AmendStoreOutcome.Amended, version);
        }

        public async Task<SubmissionAmendment?> GetLatestAmendmentAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT version, account_label, amended_utc
                FROM submission_amendments
                WHERE submission_id = @id
                ORDER BY version DESC
                LIMIT 1;
                """;
            command.Parameters.AddWithValue("@id", submissionId);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            if (!await reader.ReadAsync(cancellationToken))
                return null;

            return new SubmissionAmendment(
                reader.GetInt32(0),
                reader.GetString(1),
                MySqlSubmissionStore.ReadUtc(reader, 2)!.Value);
        }

        public Task SetTouchesSharedFilesAsync(long submissionId, bool touchesSharedFiles, CancellationToken cancellationToken = default) =>
            this.ExecuteAsync(
                "UPDATE submissions SET touches_shared_files = @touches WHERE id = @id;",
                command =>
                {
                    command.Parameters.AddWithValue("@id", submissionId);
                    command.Parameters.AddWithValue("@touches", touchesSharedFiles);
                },
                cancellationToken);

        public async Task<IReadOnlyList<GivenApproval>> GetApprovalsAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                "SELECT role, account_label, approved_utc, account_id FROM submission_approvals WHERE submission_id = @id ORDER BY approved_utc;";
            command.Parameters.AddWithValue("@id", submissionId);

            return await MySqlSubmissionStore.ReadApprovalsAsync(command, cancellationToken);
        }

        public Task AddApprovalAsync(
            long submissionId,
            ApproverRole role,
            long accountId,
            string accountLabel,
            DateTimeOffset approvedUtc,
            CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                INSERT IGNORE INTO submission_approvals (submission_id, role, account_id, account_label, approved_utc)
                VALUES (@id, @role, @accountId, @label, @when);
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@id", submissionId);
                    command.Parameters.AddWithValue("@role", MySqlSubmissionStore.RoleText(role));
                    command.Parameters.AddWithValue("@accountId", accountId);
                    command.Parameters.AddWithValue("@label", MySqlSubmissionStore.Cut(accountLabel, 255));
                    command.Parameters.AddWithValue("@when", approvedUtc.UtcDateTime);
                },
                cancellationToken);
        }

        public async Task<IReadOnlyList<GivenApproval>> GetProductionApprovalsAsync(
            string boardId,
            string betaContentHash,
            CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT role, account_label, approved_utc, account_id FROM production_approvals
                WHERE board_id = @boardId AND beta_content_hash = @hash ORDER BY approved_utc;
                """;
            command.Parameters.AddWithValue("@boardId", boardId);
            command.Parameters.AddWithValue("@hash", betaContentHash);

            return await MySqlSubmissionStore.ReadApprovalsAsync(command, cancellationToken);
        }

        // See ISubmissionStore.GetProductionApprovalsForAsync - one query for every board on the list.
        public async Task<IReadOnlyDictionary<string, IReadOnlyList<GivenApproval>>> GetProductionApprovalsForAsync(
            IReadOnlyCollection<(string BoardId, string BetaContentHash)> states,
            CancellationToken cancellationToken = default)
        {
            var found = new Dictionary<string, List<GivenApproval>>(StringComparer.Ordinal);

            if (states is null || states.Count == 0)
                return new Dictionary<string, IReadOnlyList<GivenApproval>>(StringComparer.Ordinal);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            // One pair of parameters per board - never the values written into the text.
            List<string> pairs = [];
            int index = 0;

            foreach ((string boardId, string hash) in states.Distinct())
            {
                string board = FormattableString.Invariant($"@s{index}");
                string beta = FormattableString.Invariant($"@h{index++}");
                pairs.Add($"({board}, {beta})");
                command.Parameters.AddWithValue(board, boardId);
                command.Parameters.AddWithValue(beta, hash);
            }

            command.CommandText =
                "SELECT board_id, role, account_label, approved_utc, account_id FROM production_approvals " +
                $"WHERE (board_id, beta_content_hash) IN ({string.Join(", ", pairs)}) ORDER BY approved_utc;";

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                ApproverRole? role = reader.GetString(1) switch
                {
                    "administrator" => ApproverRole.Administrator,
                    "maintainer" => ApproverRole.Maintainer,
                    _ => null
                };

                if (role is null)
                    continue;

                long? accountId = reader.IsDBNull(4) ? null : reader.GetInt64(4);
                string boardId = reader.GetString(0);

                if (!found.TryGetValue(boardId, out List<GivenApproval>? list))
                    found[boardId] = list = [];

                list.Add(new GivenApproval(role.Value, reader.GetString(2), MySqlSubmissionStore.ReadUtc(reader, 3)!.Value, accountId));
            }

            return found.ToDictionary(pair => pair.Key, pair => (IReadOnlyList<GivenApproval>)pair.Value, StringComparer.Ordinal);
        }

        public Task AddProductionApprovalAsync(
            string boardId,
            string betaContentHash,
            ApproverRole role,
            long accountId,
            string accountLabel,
            DateTimeOffset approvedUtc,
            CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                INSERT IGNORE INTO production_approvals
                    (board_id, beta_content_hash, role, account_id, account_label, approved_utc)
                VALUES (@boardId, @hash, @role, @accountId, @label, @when);
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@boardId", boardId);
                    command.Parameters.AddWithValue("@hash", betaContentHash);
                    command.Parameters.AddWithValue("@role", MySqlSubmissionStore.RoleText(role));
                    command.Parameters.AddWithValue("@accountId", accountId);
                    command.Parameters.AddWithValue("@label", MySqlSubmissionStore.Cut(accountLabel, 255));
                    command.Parameters.AddWithValue("@when", approvedUtc.UtcDateTime);
                },
                cancellationToken);
        }

        // The text the CHECK constraint allows. A role this build does not know reads back as
        // nothing rather than as the wrong role.
        private static string RoleText(ApproverRole role) =>
            role == ApproverRole.Administrator ? "administrator" : "maintainer";

        private static string Cut(string? value, int length)
        {
            string text = value ?? string.Empty;
            return text.Length <= length ? text : text[..length];
        }

        private static async Task<IReadOnlyList<GivenApproval>> ReadApprovalsAsync(MySqlCommand command, CancellationToken cancellationToken)
        {
            var approvals = new List<GivenApproval>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                ApproverRole? role = reader.GetString(0) switch
                {
                    "administrator" => ApproverRole.Administrator,
                    "maintainer" => ApproverRole.Maintainer,
                    _ => null
                };

                // account_id is NULL once the account is deleted (ON DELETE SET NULL).
                long? accountId = reader.IsDBNull(3) ? null : reader.GetInt64(3);

                if (role is not null)
                    approvals.Add(new GivenApproval(role.Value, reader.GetString(1), MySqlSubmissionStore.ReadUtc(reader, 2)!.Value, accountId));
            }

            return approvals;
        }
    }
}
