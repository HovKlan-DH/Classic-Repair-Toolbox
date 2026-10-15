using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // MySqlSubmissionStore, part four (code review, 2026-10-13): the state and decision writes that
    // happen only from the state the caller saw, a maintainer's own queue filtered in the query, and
    // the files a submission left exactly as published (migration 0021). Split out so
    // MySqlSubmissionStore.cs stays under the project's ~1,500 lines; that file's header holds here.
    // ###########################################################################################
    public sealed partial class MySqlSubmissionStore
    {
        // See ISubmissionStore.TrySetStateAsync: the WHERE carries the state the caller saw, and the
        // affected-rows count says whether this call moved it.
        public async Task<bool> TrySetStateAsync(long submissionId, string fromState, string state, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            bool isTerminal = state is not SubmissionState.Uploading;

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = isTerminal
                ? "UPDATE submissions SET state = @state, decided_utc = @when WHERE id = @id AND state = @from;"
                : "UPDATE submissions SET state = @state WHERE id = @id AND state = @from;";

            command.Parameters.AddWithValue("@state", state);
            command.Parameters.AddWithValue("@id", submissionId);
            command.Parameters.AddWithValue("@from", fromState);

            if (isTerminal)
                command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);

            return await command.ExecuteNonQueryAsync(cancellationToken) > 0;
        }

        // See ISubmissionStore.TrySetDecisionAsync - SetDecisionAsync's columns, with the states the
        // row may still be in as parameters of the WHERE (never written into the text).
        public async Task<bool> TrySetDecisionAsync(
            long submissionId,
            IReadOnlyCollection<string> fromStates,
            string state,
            long decidedByAccountId,
            string? comment,
            DateTimeOffset decidedUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(fromStates);

            if (fromStates.Count == 0)
                return false;

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            List<string> names = [];
            int index = 0;

            foreach (string from in fromStates.Distinct(StringComparer.Ordinal))
            {
                string name = FormattableString.Invariant($"@from{index++}");
                names.Add(name);
                command.Parameters.AddWithValue(name, from);
            }

            command.CommandText = $"""
                UPDATE submissions
                   SET state = @state,
                       decided_utc = @when,
                       decided_by = @by,
                       decision_comment = @comment
                 WHERE id = @id AND state IN ({string.Join(", ", names)});
                """;

            command.Parameters.AddWithValue("@state", state);
            command.Parameters.AddWithValue("@when", decidedUtc.UtcDateTime);
            command.Parameters.AddWithValue("@by", decidedByAccountId);
            command.Parameters.AddWithValue("@comment", string.IsNullOrWhiteSpace(comment) ? DBNull.Value : comment);
            command.Parameters.AddWithValue("@id", submissionId);

            return await command.ExecuteNonQueryAsync(cancellationToken) > 0;
        }

        public async Task<IReadOnlyDictionary<string, string>> GetLeftAsPublishedAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = "SELECT path, sha256 FROM submission_left_as_published WHERE submission_id = @id;";
            command.Parameters.AddWithValue("@id", submissionId);

            var files = new Dictionary<string, string>(StringComparer.Ordinal);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                files[reader.GetString(0)] = reader.GetString(1);

            return files;
        }

        // The queue of `boardIds` only - see ISubmissionStore.GetQueueForBoardsAsync. One parameter
        // per board, never the ids written into the text; ix_submissions_board (board_id, state)
        // serves the filter.
        public async Task<IReadOnlyList<SubmissionRecord>> GetQueueForBoardsAsync(
            IReadOnlyCollection<string> boardIds,
            int limit,
            CancellationToken cancellationToken = default)
        {
            var records = new List<SubmissionRecord>();

            if (boardIds is null || boardIds.Count == 0)
                return records;

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            List<string> names = [];
            int index = 0;

            foreach (string boardId in boardIds.Distinct(StringComparer.Ordinal))
            {
                string name = FormattableString.Invariant($"@b{index++}");
                names.Add(name);
                command.Parameters.AddWithValue(name, boardId);
            }

            command.CommandText = $"""
                SELECT id, board_id, account_id, contact_email, upload_token_hash, base_revision,
                       state, summary, format_version, created_utc, expires_utc, decided_utc,
                       decision_comment, touches_shared_files
                FROM submissions
                WHERE state IN (@state, @approved) AND board_id IN ({string.Join(", ", names)})
                ORDER BY id ASC LIMIT @limit;
                """;

            command.Parameters.AddWithValue("@state", SubmissionState.Pending);
            command.Parameters.AddWithValue("@approved", SubmissionState.Approved);
            command.Parameters.AddWithValue("@limit", Math.Clamp(limit, 1, 200));

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                records.Add(MySqlSubmissionStore.ReadSubmission(reader));

            return records;
        }
    }
}
