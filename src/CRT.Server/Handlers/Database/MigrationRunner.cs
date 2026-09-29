using MySqlConnector;

namespace CRT.Server.Handlers.Database
{
    // ###########################################################################################
    // Applies the migrations MigrationPlan decided on. This is the I/O half and is DELIBERATELY
    // DUMB: it reads files, asks MigrationPlan what to do, and executes. Every rule about ordering,
    // gaps, checksums and merge collisions lives in MigrationPlan, where it is unit tested with no
    // database. If a rule ends up here it is a rule no test can reach.
    //
    // WHY MIGRATIONS RUN AT STARTUP rather than from a separate command. This service is deployed
    // by hand by the project owner, and a separate step is a step that gets forgotten - the failure
    // mode being a service running against a schema it does not match, which surfaces as confusing
    // runtime errors rather than as a clear refusal. Running at startup means the deployment
    // either works or does not start, which is the same bargain ServerOptions makes.
    //
    // EACH MIGRATION RUNS IN ITS OWN TRANSACTION, and its schema_migrations row is written inside
    // that transaction. So a migration either fully applied and is recorded, or did neither.
    //
    // *** MariaDB DDL IS NOT TRANSACTIONAL. *** CREATE TABLE commits implicitly, so a migration
    // that fails halfway leaves the tables it already created. The transaction still buys the one
    // thing that matters most - the recorded row cannot claim a migration ran when it did not -
    // but it does NOT make a failed migration atomic. Write each migration so that it is safe to
    // inspect and fix by hand afterwards, and keep them small.
    // ###########################################################################################
    public static class MigrationRunner
    {
        // The table recording what has been applied. Created here rather than in 0001, because
        // something has to exist before the first migration can be recorded.
        private const string SchemaMigrationsTable = "schema_migrations";

        // ###########################################################################################
        // Reads the migration files from a directory, in whatever order the filesystem returns them
        // - MigrationPlan sorts. A file whose name does not parse is a FAILURE, not something to
        // skip quietly: a migration nobody notices is not running is the whole problem this class
        // exists to prevent.
        // ###########################################################################################
        public static IReadOnlyList<MigrationScript> ReadMigrations(
            string directory,
            out IReadOnlyList<string> failures)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(directory);

            var problems = new List<string>();
            var scripts = new List<MigrationScript>();

            if (!Directory.Exists(directory))
            {
                problems.Add(
                    $"The migrations directory [{directory}] does not exist. It ships beside the " +
                    "binaries; a missing one means the publish was incomplete.");

                failures = problems;
                return Array.Empty<MigrationScript>();
            }

            foreach (string path in Directory.GetFiles(directory, "*.sql"))
            {
                string fileName = Path.GetFileName(path);

                if (!MigrationPlan.TryParseNumber(fileName, out int number))
                {
                    problems.Add(
                        $"Migration file [{fileName}] is not named <number>_<description>.sql " +
                        "(for example 0001_initial.sql), so its position in the sequence is " +
                        "undefined. Rename it or remove it.");

                    continue;
                }

                string sql = File.ReadAllText(path);

                scripts.Add(new MigrationScript(number, fileName, sql, MigrationPlan.ComputeChecksum(sql)));
            }

            failures = problems;
            return scripts;
        }

        // ###########################################################################################
        // Brings the database up to date, or throws with every reason at once.
        //
        // Throwing is correct here for the same reason it is correct in Program.Main's configuration
        // check: the process exits non-zero, systemd reports the unit failed, and the journal carries
        // the reasons. An outage is noticed at once and harms nobody; a service running against a
        // schema it does not match is a silent wrong answer.
        // ###########################################################################################
        public static async Task<int> MigrateAsync(
            string connectionString,
            string migrationsDirectory,
            ILogger logger,
            CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(connectionString);
            ArgumentNullException.ThrowIfNull(logger);

            IReadOnlyList<MigrationScript> available =
                MigrationRunner.ReadMigrations(migrationsDirectory, out IReadOnlyList<string> readFailures);

            await using var connection = new MySqlConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            await MigrationRunner.EnsureSchemaMigrationsTableAsync(connection, cancellationToken);

            IReadOnlyList<AppliedMigration> applied =
                await MigrationRunner.ReadAppliedAsync(connection, cancellationToken);

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            var allFailures = readFailures.Concat(plan.Failures).ToList();

            if (allFailures.Count > 0)
            {
                foreach (string failure in allFailures)
                    logger.LogCritical("Migration error: {Failure}", failure);

                throw new MigrationFailedException(
                    $"The database schema cannot be brought up to date: {allFailures.Count} " +
                    $"error(s). See the journal for details " +
                    $"(journalctl -u crt-server -n 20 --no-pager -p warning). " +
                    $"First: {allFailures[0]}");
            }

            if (plan.ToApply.Count == 0)
            {
                logger.LogInformation(
                    "Database schema is up to date ({Count} migration(s) already applied).",
                    applied.Count);

                return 0;
            }

            foreach (MigrationScript script in plan.ToApply)
            {
                logger.LogInformation("Applying migration {Number:D4} ({FileName})...", script.Number, script.FileName);

                await MigrationRunner.ApplyOneAsync(connection, script, cancellationToken);

                logger.LogInformation("Applied migration {Number:D4}.", script.Number);
            }

            logger.LogInformation("Applied {Count} migration(s).", plan.ToApply.Count);

            return plan.ToApply.Count;
        }

        // ###########################################################################################
        // The bootstrap table. IF NOT EXISTS rather than a migration, because this is what records
        // that migrations ran - it cannot record its own creation.
        // ###########################################################################################
        private static async Task EnsureSchemaMigrationsTableAsync(
            MySqlConnection connection,
            CancellationToken cancellationToken)
        {
            await using var command = connection.CreateCommand();

            command.CommandText = $"""
                CREATE TABLE IF NOT EXISTS {MigrationRunner.SchemaMigrationsTable} (
                    number       INT           NOT NULL,
                    file_name    VARCHAR(255)  NOT NULL,
                    checksum     CHAR(64)      NOT NULL,
                    applied_utc  DATETIME(3)   NOT NULL,
                    PRIMARY KEY (number)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
                """;

            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        private static async Task<IReadOnlyList<AppliedMigration>> ReadAppliedAsync(
            MySqlConnection connection,
            CancellationToken cancellationToken)
        {
            var rows = new List<AppliedMigration>();

            await using var command = connection.CreateCommand();
            command.CommandText =
                $"SELECT number, file_name, checksum, applied_utc FROM {MigrationRunner.SchemaMigrationsTable} ORDER BY number;";

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                rows.Add(new AppliedMigration(
                    reader.GetInt32(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    DateTime.SpecifyKind(reader.GetDateTime(3), DateTimeKind.Utc)));
            }

            return rows;
        }

        // ###########################################################################################
        // Applies one migration and records it, in one transaction.
        //
        // The SQL file is executed as a SINGLE command text. MySqlConnector allows multiple
        // statements in one command, which is what lets a migration file read as ordinary SQL
        // rather than being split on semicolons by a parser of our own - and a naive split would
        // break the moment a migration contains a semicolon inside a string literal or a trigger
        // body.
        //
        // See the class header: DDL commits implicitly in MariaDB, so this transaction protects
        // the RECORD, not the schema change itself.
        // ###########################################################################################
        private static async Task ApplyOneAsync(
            MySqlConnection connection,
            MigrationScript script,
            CancellationToken cancellationToken)
        {
            await using MySqlTransaction transaction =
                await connection.BeginTransactionAsync(cancellationToken);

            try
            {
                await using (MySqlCommand migrationCommand = connection.CreateCommand())
                {
                    migrationCommand.Transaction = transaction;
                    migrationCommand.CommandText = script.Sql;

                    await migrationCommand.ExecuteNonQueryAsync(cancellationToken);
                }

                await using (MySqlCommand recordCommand = connection.CreateCommand())
                {
                    recordCommand.Transaction = transaction;
                    recordCommand.CommandText = $"""
                        INSERT INTO {MigrationRunner.SchemaMigrationsTable}
                            (number, file_name, checksum, applied_utc)
                        VALUES (@number, @fileName, @checksum, @appliedUtc);
                        """;

                    recordCommand.Parameters.AddWithValue("@number", script.Number);
                    recordCommand.Parameters.AddWithValue("@fileName", script.FileName);
                    recordCommand.Parameters.AddWithValue("@checksum", script.Checksum);
                    recordCommand.Parameters.AddWithValue("@appliedUtc", DateTime.UtcNow);

                    await recordCommand.ExecuteNonQueryAsync(cancellationToken);
                }

                await transaction.CommitAsync(cancellationToken);
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync(cancellationToken);

                throw new MigrationFailedException(
                    $"Migration {script.Number:D4} ({script.FileName}) failed: {ex.Message}. " +
                    "Note that MariaDB commits DDL implicitly, so any tables this migration " +
                    "already created still exist and must be inspected by hand before retrying.",
                    ex);
            }
        }
    }
}
