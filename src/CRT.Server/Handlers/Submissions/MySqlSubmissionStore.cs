using System.Globalization;
using System.Text.Json;
using CRT.Server.Configuration;
using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The real ISubmissionStore, against MariaDB. An untested I/O boundary, the same as
    // MySqlAccountStore and for the same reason: there are no decisions in this file. Every rule
    // lives in SubmissionFlows and SubmissionValidator, both pure and both fully tested.
    //
    // The three rules from MySqlAccountStore's header apply unchanged here: every value reaches
    // SQL through a parameter, DATETIME(3) holds UTC, and a connection is taken per operation from
    // the pool.
    //
    // ONE THING SPECIFIC TO THIS STORE: the manifest's rows are stored as JSON in one column - see
    // 0002_submission_files.sql for why eleven tables would be worse. That means the shape is
    // validated by SubmissionValidator before it is written and by the deserialiser when it is
    // read, never by the database.
    // ###########################################################################################
    public sealed class MySqlSubmissionStore : ISubmissionStore
    {
        // Matched to the app's own JSON conventions so a payload written by one version reads back
        // in another. Property names are case-insensitive on read for the same reason.
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        private readonly string thisConnectionString;

        public MySqlSubmissionStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        // ###########################################################################################
        // Creates the submission and its file rows in ONE TRANSACTION.
        //
        // The transaction matters here in a way it did not for accounts: a submission row without
        // its files is a submission that can never be finalised and never be collected, because
        // nothing would report it incomplete. Both halves land or neither does.
        // ###########################################################################################
        public async Task<long> CreateAsync(NewSubmission submission, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            // ###########################################################################################
            // THE SYSTEM ROW HAS TO EXIST FIRST, or the foreign key below refuses the insert.
            //
            // submissions.system_id is NOT NULL with a foreign key to systems(system_id), and until
            // this was added NOTHING ever wrote to `systems` - so the table was empty and every
            // real submission would have failed. It was invisible because every test uses an
            // in-memory fake store; see 0004's header.
            //
            // INSERT IGNORE rather than "check then insert": two contributors submitting for the
            // same new system at the same moment would both see it absent and both insert, and the
            // second would fail on the primary key. Letting the database settle it is the only
            // version without a race.
            //
            // A system that already exists is left completely alone - origin in particular. A
            // system records where it CAME FROM, so one that shipped with CRT stays 'shipped'
            // however many contributions it later receives; an ON DUPLICATE KEY UPDATE here would
            // rewrite that on every submission.
            //
            // current_revision stays NULL: this system has not been published yet. Publishing is
            // the maintainer's act and is what fills it in.
            // ###########################################################################################
            await using (MySqlCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT IGNORE INTO systems
                        (system_id, manufacturer, hardware, board, origin, is_accepting, created_utc)
                    VALUES (@systemId, @manufacturer, @hardware, @board, @origin, 1, @created);
                    """;

                command.Parameters.AddWithValue("@systemId", submission.SystemId);
                command.Parameters.AddWithValue("@manufacturer", submission.Manufacturer);
                command.Parameters.AddWithValue("@hardware", submission.Hardware);
                command.Parameters.AddWithValue("@board", submission.Board);

                // A system that first appears through this pipeline is 'contributed' by
                // definition - a shipped one was already in the tree before any submission.
                command.Parameters.AddWithValue("@origin", SystemDescriptorRules.SystemOrigin.Contributed);
                command.Parameters.AddWithValue("@created", submission.CreatedUtc.UtcDateTime);

                await command.ExecuteNonQueryAsync(cancellationToken);
            }

            long submissionId;

            await using (MySqlCommand command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO submissions
                        (system_id, account_id, contact_email, created_ip, upload_token_hash,
                         base_revision, state, format_version, summary, created_utc, expires_utc)
                    VALUES (@systemId, @accountId, @contactEmail, @createdIp, @uploadTokenHash,
                            @baseRevision, @state, @formatVersion, @summary, @created, @expires);
                    SELECT LAST_INSERT_ID();
                    """;

                command.Parameters.AddWithValue("@systemId", submission.SystemId);
                command.Parameters.AddWithValue("@accountId", (object?)submission.AccountId ?? DBNull.Value);
                command.Parameters.AddWithValue("@contactEmail", (object?)submission.ContactEmail ?? DBNull.Value);
                command.Parameters.AddWithValue("@createdIp", (object?)submission.CreatedIp ?? DBNull.Value);
                command.Parameters.AddWithValue("@uploadTokenHash", submission.UploadTokenHash);
                command.Parameters.AddWithValue("@baseRevision", submission.BaseRevision);
                command.Parameters.AddWithValue("@state", SubmissionState.Uploading);
                command.Parameters.AddWithValue("@formatVersion", submission.FormatVersion);
                command.Parameters.AddWithValue("@summary", submission.Summary);
                command.Parameters.AddWithValue("@created", submission.CreatedUtc.UtcDateTime);
                command.Parameters.AddWithValue("@expires", submission.ExpiresUtc.UtcDateTime);

                object? id = await command.ExecuteScalarAsync(cancellationToken);
                submissionId = Convert.ToInt64(id, CultureInfo.InvariantCulture);
            }

            foreach (SubmissionFile file in submission.Files)
            {
                await using MySqlCommand command = connection.CreateCommand();

                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO submission_files (submission_id, path, sha256, size_bytes)
                    VALUES (@submissionId, @path, @hash, @size);
                    """;

                command.Parameters.AddWithValue("@submissionId", submissionId);
                command.Parameters.AddWithValue("@path", file.Path);
                command.Parameters.AddWithValue("@hash", file.Sha256);
                command.Parameters.AddWithValue("@size", file.SizeBytes);

                await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);

            return submissionId;
        }

        public async Task<SubmissionRecord?> FindAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT id, system_id, account_id, contact_email, upload_token_hash, base_revision,
                       state, summary, format_version, created_utc, expires_utc, decided_utc,
                       decision_comment
                FROM submissions WHERE id = @id LIMIT 1;
                """;

            command.Parameters.AddWithValue("@id", submissionId);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            if (!await reader.ReadAsync(cancellationToken))
                return null;

            return MySqlSubmissionStore.ReadSubmission(reader);
        }

        public async Task<IReadOnlyList<SubmissionFileRecord>> GetFilesAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT id, path, sha256, size_bytes, is_uploaded
                FROM submission_files WHERE submission_id = @id ORDER BY id;
                """;

            command.Parameters.AddWithValue("@id", submissionId);

            var files = new List<SubmissionFileRecord>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                files.Add(new SubmissionFileRecord(
                    reader.GetInt64(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    reader.GetInt64(3),
                    reader.GetBoolean(4)));
            }

            return files;
        }

        // Marks every file carrying this hash - one blob can appear at several paths, and
        // uploading it once must satisfy all of them.
        public Task MarkUploadedAsync(long submissionId, string sha256, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE submission_files SET is_uploaded = 1
                WHERE submission_id = @submissionId AND sha256 = @hash;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@submissionId", submissionId);
                    command.Parameters.AddWithValue("@hash", sha256);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // Records a system's published revision and content hash (Phase 5, task 6).
        //
        // An UPDATE rather than an upsert: the `systems` row is inserted by CreateAsync when the
        // first submission for that system arrives, so by publish time it always exists. An
        // INSERT ... ON DUPLICATE KEY here could create a row with no manufacturer/hardware/board
        // and no origin, which the NOT NULL columns would refuse anyway - failing loudly at the
        // wrong layer instead of revealing that the caller published a system nobody registered.
        //
        // origin is deliberately UNTOUCHED. It records where a system CAME FROM and is set once,
        // so a contributed system stays contributed however many times it is later revised -
        // including by the maintainer.
        //
        // *** THERE IS NO `updated_utc` COLUMN ON `systems`, AND THIS SETS NONE. *** A first
        // version of this method wrote one; the schema has only created_utc. It would have thrown
        // against MariaDB on the very first publish while every test passed, because the tests run
        // against an in-memory fake - which is precisely the failure migration 0004 exists to
        // correct, three of them at once, found by reading the schema rather than by any test.
        // When the publish time is wanted, system.json carries PublishedUtc, and the audit table
        // records the act.
        // ###########################################################################################
        public Task SetSystemPublishedAsync(
            string systemId,
            string revision,
            string contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE systems
                   SET current_revision = @revision,
                       content_hash = @hash
                 WHERE system_id = @systemId;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@revision", revision);
                    command.Parameters.AddWithValue("@hash", contentHash);
                    command.Parameters.AddWithValue("@systemId", systemId);
                },
                cancellationToken);
        }

        public Task SetStateAsync(long submissionId, string state, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            // decided_utc is set for a terminal state only. An 'uploading' row has not been
            // decided, and stamping it would make the audit trail claim a decision nobody made.
            bool isTerminal = state is not SubmissionState.Uploading;

            return this.ExecuteAsync(
                isTerminal
                    ? "UPDATE submissions SET state = @state, decided_utc = @when WHERE id = @id;"
                    : "UPDATE submissions SET state = @state WHERE id = @id;",
                command =>
                {
                    command.Parameters.AddWithValue("@state", state);
                    command.Parameters.AddWithValue("@id", submissionId);

                    if (isTerminal)
                        command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // Records a reviewer's decision - the state, who made it, when, and why.
        //
        // *** THE COLUMNS WERE CHECKED AGAINST 0001_initial.sql, one by one. *** `submissions` has
        // state, decided_utc, decided_by and decision_comment, and no updated_utc - the fourth
        // schema-versus-code disagreement in this project was writing exactly such a column on a
        // table that does not have one, which every test passed because they run against the
        // in-memory fake. Enumerate the real columns; do not infer them.
        //
        // A NULL comment is stored as NULL rather than an empty string, so "no comment was given"
        // and "an empty comment was given" stay distinguishable in the audit trail.
        // ###########################################################################################
        public Task SetDecisionAsync(
            long submissionId,
            string state,
            long decidedByAccountId,
            string? comment,
            DateTimeOffset decidedUtc,
            CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE submissions
                   SET state = @state,
                       decided_utc = @when,
                       decided_by = @by,
                       decision_comment = @comment
                 WHERE id = @id;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@state", state);
                    command.Parameters.AddWithValue("@when", decidedUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@by", decidedByAccountId);
                    command.Parameters.AddWithValue(
                        "@comment",
                        string.IsNullOrWhiteSpace(comment) ? DBNull.Value : comment);
                    command.Parameters.AddWithValue("@id", submissionId);
                },
                cancellationToken);
        }

        public Task SavePayloadAsync(long submissionId, SubmissionManifest manifest, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            string rowsJson = JsonSerializer.Serialize(manifest.Rows, MySqlSubmissionStore.JsonOptions);
            string renamesJson = JsonSerializer.Serialize(manifest.Renames, MySqlSubmissionStore.JsonOptions);

            // REPLACE rather than INSERT so a retried create does not fail on the primary key.
            return this.ExecuteAsync(
                """
                REPLACE INTO submission_payloads
                    (submission_id, format_version, rows_json, renames_json, created_utc)
                VALUES (@id, @formatVersion, @rows, @renames, @created);
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@id", submissionId);
                    command.Parameters.AddWithValue("@formatVersion", manifest.FormatVersion);
                    command.Parameters.AddWithValue("@rows", rowsJson);
                    command.Parameters.AddWithValue("@renames", renamesJson);
                    command.Parameters.AddWithValue("@created", DateTime.UtcNow);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // Reads a submission back as a manifest.
        //
        // The FILES come from submission_files rather than from the stored JSON, because that table
        // is the one carrying upload state - reconstructing from the payload would lose it.
        // ###########################################################################################
        public async Task<SubmissionManifest?> LoadPayloadAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            SubmissionRecord? submission = await this.FindAsync(submissionId, cancellationToken);

            if (submission is null)
                return null;

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            // The three name parts come from the `systems` row, joined via the submission's own
            // system_id - see the comment on the manifest below for why they have to be here at
            // all. A LEFT JOIN, so a payload whose system row is somehow absent still loads rather
            // than vanishing; the validator then reports the missing names, which is the truth.
            command.CommandText = """
                SELECT p.format_version, p.rows_json, p.renames_json,
                       COALESCE(y.manufacturer, ''), COALESCE(y.hardware, ''), COALESCE(y.board, '')
                FROM submission_payloads p
                JOIN submissions s ON s.id = p.submission_id
                LEFT JOIN systems y ON y.system_id = s.system_id
                WHERE p.submission_id = @id
                LIMIT 1;
                """;

            command.Parameters.AddWithValue("@id", submissionId);

            SubmissionRows? rows;
            List<SubmissionRename>? renames;
            int formatVersion;
            string manufacturer;
            string hardware;
            string board;

            await using (MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken))
            {
                if (!await reader.ReadAsync(cancellationToken))
                    return null;

                formatVersion = reader.GetInt32(0);

                rows = JsonSerializer.Deserialize<SubmissionRows>(
                    reader.GetString(1), MySqlSubmissionStore.JsonOptions);

                renames = reader.IsDBNull(2)
                    ? []
                    : JsonSerializer.Deserialize<List<SubmissionRename>>(
                        reader.GetString(2), MySqlSubmissionStore.JsonOptions);

                manufacturer = reader.GetString(3);
                hardware = reader.GetString(4);
                board = reader.GetString(5);
            }

            if (rows is null)
                return null;

            IReadOnlyList<SubmissionFileRecord> files = await this.GetFilesAsync(submissionId, cancellationToken);

            return new SubmissionManifest
            {
                FormatVersion = formatVersion,
                SystemId = submission.SystemId,

                // ###########################################################################################
                // *** THE THREE NAME PARTS MUST BE RESTORED, OR FINALISE REJECTS EVERY SUBMISSION. ***
                //
                // Validation runs TWICE - once at create against the manifest as sent, and again at
                // finalise against the manifest as RELOADED from here. This method used to leave
                // Manufacturer/Hardware/Board unset, so the reloaded manifest carried three empty
                // strings and the second pass answered "does not name the hardware", "does not name
                // the board", and - because BuildSystemId("","","") cannot equal the stored id -
                // "system identifier does not match".
                //
                // Every submission that reached finalise was therefore rejected, whatever the
                // client sent, and the findings blamed the CLIENT for it ("a fault in the
                // submitting application"), which is where two days of looking went.
                //
                // They come from the `systems` row rather than the payload JSON because that is
                // where they live: the schema stores them as their own columns precisely "so the
                // review app can list by manufacturer without parsing" (0001_initial.sql).
                // ###########################################################################################
                Manufacturer = manufacturer,
                Hardware = hardware,
                Board = board,

                BaseRevision = submission.BaseRevision,
                Summary = submission.Summary ?? string.Empty,
                CreatedUtc = submission.CreatedUtc,
                Rows = rows,
                Renames = renames ?? [],
                Files = files
                    .Select(file => new SubmissionFile
                    {
                        Path = file.Path,
                        Sha256 = file.Sha256,
                        SizeBytes = file.SizeBytes
                    })
                    .ToList()
            };
        }

        // ###########################################################################################
        // Replaces a submission's findings.
        //
        // Replaces rather than appends: findings are recomputed at finalise time against the
        // complete set of files, and keeping the earlier partial set would show the contributor
        // problems that have since been resolved.
        // ###########################################################################################
        public async Task SaveFindingsAsync(long submissionId, IReadOnlyList<ValidationFinding> findings, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(findings);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            await using (MySqlCommand clear = connection.CreateCommand())
            {
                clear.Transaction = transaction;
                clear.CommandText = "DELETE FROM submission_findings WHERE submission_id = @id;";
                clear.Parameters.AddWithValue("@id", submissionId);

                await clear.ExecuteNonQueryAsync(cancellationToken);
            }

            foreach (ValidationFinding finding in findings)
            {
                await using MySqlCommand command = connection.CreateCommand();

                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO submission_findings (submission_id, severity, code, subject, message)
                    VALUES (@id, @severity, @code, @subject, @message);
                    """;

                command.Parameters.AddWithValue("@id", submissionId);
                command.Parameters.AddWithValue(
                    "@severity", finding.Severity == ValidationSeverity.Error ? "error" : "warning");
                command.Parameters.AddWithValue("@code", finding.Code);
                command.Parameters.AddWithValue("@subject", (object?)finding.Subject ?? DBNull.Value);
                command.Parameters.AddWithValue("@message", finding.Message);

                await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
        }

        public async Task<IReadOnlyList<ValidationFinding>> GetFindingsAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT severity, code, subject, message
                FROM submission_findings WHERE submission_id = @id ORDER BY id;
                """;

            command.Parameters.AddWithValue("@id", submissionId);

            var findings = new List<ValidationFinding>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                findings.Add(new ValidationFinding
                {
                    Severity = string.Equals(reader.GetString(0), "error", StringComparison.Ordinal)
                        ? ValidationSeverity.Error
                        : ValidationSeverity.Warning,
                    Code = reader.GetString(1),
                    Subject = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                    Message = reader.GetString(3)
                });
            }

            return findings;
        }

        public async Task<IReadOnlyList<long>> GetExpiredUploadsAsync(DateTimeOffset now, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            // Bounded, so one sweep cannot lock a huge number of rows at once.
            command.CommandText = """
                SELECT id FROM submissions
                WHERE state = @state AND expires_utc IS NOT NULL AND expires_utc <= @now
                LIMIT 500;
                """;

            command.Parameters.AddWithValue("@state", SubmissionState.Uploading);
            command.Parameters.AddWithValue("@now", now.UtcDateTime);

            var ids = new List<long>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                ids.Add(reader.GetInt64(0));

            return ids;
        }

        public async Task<IReadOnlyList<SubmissionRecord>> GetForAccountAsync(long accountId, int limit, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT id, system_id, account_id, contact_email, upload_token_hash, base_revision,
                       state, summary, format_version, created_utc, expires_utc, decided_utc,
                       decision_comment
                FROM submissions WHERE account_id = @accountId
                ORDER BY id DESC LIMIT @limit;
                """;

            command.Parameters.AddWithValue("@accountId", accountId);
            command.Parameters.AddWithValue("@limit", Math.Clamp(limit, 1, 200));

            var records = new List<SubmissionRecord>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                records.Add(MySqlSubmissionStore.ReadSubmission(reader));

            return records;
        }

        // ###########################################################################################
        // The review queue (Phase 5, task 2).
        //
        // ORDER BY id ASC - oldest first - which is the opposite of GetForAccountAsync above and
        // deliberately so. That one answers "what did I send lately", where newest-first is right.
        // This one is a work queue, and the oldest submission is the one that has been waiting
        // longest; newest-first would let a steady trickle of new contributions keep burying the
        // one that has been waiting a month.
        //
        // The index ix_submissions_state (state, created_utc) covers this filter.
        // ###########################################################################################
        public async Task<IReadOnlyList<SubmissionRecord>> GetQueueAsync(int limit, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT id, system_id, account_id, contact_email, upload_token_hash, base_revision,
                       state, summary, format_version, created_utc, expires_utc, decided_utc,
                       decision_comment
                FROM submissions WHERE state = @state
                ORDER BY id ASC LIMIT @limit;
                """;

            command.Parameters.AddWithValue("@state", SubmissionState.Pending);
            command.Parameters.AddWithValue("@limit", Math.Clamp(limit, 1, 200));

            var records = new List<SubmissionRecord>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                records.Add(MySqlSubmissionStore.ReadSubmission(reader));

            return records;
        }

        // -----------------------------------------------------------------------------------
        // Plumbing. Same shape as MySqlAccountStore's.
        // -----------------------------------------------------------------------------------

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            return connection;
        }

        private async Task ExecuteAsync(
            string sql,
            Action<MySqlCommand> addParameters,
            CancellationToken cancellationToken)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = sql;
            addParameters(command);

            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        private static SubmissionRecord ReadSubmission(MySqlDataReader reader)
        {
            return new SubmissionRecord(
                reader.GetInt64(0),
                reader.GetString(1),
                reader.IsDBNull(2) ? null : reader.GetInt64(2),
                reader.IsDBNull(3) ? null : reader.GetString(3),
                reader.IsDBNull(4) ? null : reader.GetString(4),
                reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                reader.GetString(6),
                reader.IsDBNull(7) ? null : reader.GetString(7),
                reader.GetInt32(8),
                MySqlSubmissionStore.ReadUtc(reader, 9)!.Value,
                MySqlSubmissionStore.ReadUtc(reader, 10),
                MySqlSubmissionStore.ReadUtc(reader, 11),

                // Column 12, added LAST in every SELECT so that no existing ordinal moved - the
                // mapper reads by index and a column inserted in the middle would silently shift
                // every field after it, which compiles and reports the wrong data.
                reader.IsDBNull(12) ? null : reader.GetString(12));
        }

        // DATETIME carries no timezone, so a value read back is Unspecified. Tagging it Utc is
        // what makes every comparison correct - an Unspecified DateTime converted to
        // DateTimeOffset would be interpreted in the SERVER'S LOCAL ZONE, silently shifting every
        // expiry by the UTC offset.
        private static DateTimeOffset? ReadUtc(MySqlDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal))
                return null;

            return new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(ordinal), DateTimeKind.Utc));
        }
    }
}
