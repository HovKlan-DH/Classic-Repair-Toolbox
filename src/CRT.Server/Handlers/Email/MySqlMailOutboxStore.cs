using System.Globalization;
using CRT.Server.Configuration;
using MySqlConnector;

namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // The real IMailOutboxStore, against migration 0022's mail_outbox. An untested I/O boundary like
    // every MySql* store: no decisions here - MailOutboxFlow decides, and is tested against a fake.
    // Every value reaches SQL through a parameter; DATETIME(3) holds UTC; a connection per call.
    // ###########################################################################################
    public sealed class MySqlMailOutboxStore : IMailOutboxStore
    {
        private readonly string thisConnectionString;

        public MySqlMailOutboxStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public async Task AddAsync(OutboxMail mail, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(mail);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                INSERT INTO mail_outbox
                    (to_address, subject, body_text, body_html, reply_to, created_utc, attempts, next_attempt_utc, last_error)
                VALUES (@to, @subject, @text, @html, @replyTo, @created, @attempts, @next, @error);
                """;

            command.Parameters.AddWithValue("@to", MySqlMailOutboxStore.Cut(mail.ToAddress, 320));
            command.Parameters.AddWithValue("@subject", MySqlMailOutboxStore.Cut(mail.Subject, 500));
            command.Parameters.AddWithValue("@text", (object?)mail.Body ?? DBNull.Value);
            command.Parameters.AddWithValue("@html", (object?)mail.HtmlBody ?? DBNull.Value);
            command.Parameters.AddWithValue("@replyTo", mail.ReplyToAddress is null ? DBNull.Value : MySqlMailOutboxStore.Cut(mail.ReplyToAddress, 320));
            command.Parameters.AddWithValue("@created", mail.CreatedUtc.UtcDateTime);
            command.Parameters.AddWithValue("@attempts", Math.Max(1, mail.Attempts));
            command.Parameters.AddWithValue("@next", mail.NextAttemptUtc.UtcDateTime);
            command.Parameters.AddWithValue("@error", mail.LastError is null ? DBNull.Value : MySqlMailOutboxStore.Cut(mail.LastError, 500));

            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task<IReadOnlyList<OutboxMail>> GetDueAsync(DateTimeOffset nowUtc, int limit, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            // A mail kept without its words is never due - there is nothing to send.
            command.CommandText = """
                SELECT id, to_address, subject, body_text, body_html, reply_to, created_utc, attempts, next_attempt_utc, last_error
                FROM mail_outbox
                WHERE next_attempt_utc <= @now AND body_text IS NOT NULL
                ORDER BY id ASC LIMIT @limit;
                """;

            command.Parameters.AddWithValue("@now", nowUtc.UtcDateTime);
            command.Parameters.AddWithValue("@limit", Math.Clamp(limit, 1, 500));

            var mails = new List<OutboxMail>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                mails.Add(new OutboxMail(
                    reader.GetInt64(0),
                    reader.GetString(1),
                    reader.GetString(2),
                    reader.IsDBNull(3) ? null : reader.GetString(3),
                    reader.IsDBNull(4) ? null : reader.GetString(4),
                    reader.IsDBNull(5) ? null : reader.GetString(5),
                    MySqlMailOutboxStore.Utc(reader.GetDateTime(6)),
                    Convert.ToInt32(reader.GetValue(7), CultureInfo.InvariantCulture),
                    MySqlMailOutboxStore.Utc(reader.GetDateTime(8)),
                    reader.IsDBNull(9) ? null : reader.GetString(9)));
            }

            return mails;
        }

        public async Task RemoveAsync(long id, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = "DELETE FROM mail_outbox WHERE id = @id;";
            command.Parameters.AddWithValue("@id", id);

            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task RecordFailureAsync(long id, int attempts, DateTimeOffset nextAttemptUtc, string? error, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = "UPDATE mail_outbox SET attempts = @attempts, next_attempt_utc = @next, last_error = @error WHERE id = @id;";
            command.Parameters.AddWithValue("@attempts", attempts);
            command.Parameters.AddWithValue("@next", nextAttemptUtc.UtcDateTime);
            command.Parameters.AddWithValue("@error", error is null ? DBNull.Value : MySqlMailOutboxStore.Cut(error, 500));
            command.Parameters.AddWithValue("@id", id);

            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task<int> RemoveCreatedBeforeAsync(DateTimeOffset createdBefore, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = "DELETE FROM mail_outbox WHERE created_utc < @before;";
            command.Parameters.AddWithValue("@before", createdBefore.UtcDateTime);

            return await command.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task<int> CountAsync(CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = "SELECT COUNT(*) FROM mail_outbox;";

            return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken), CultureInfo.InvariantCulture);
        }

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);
            return connection;
        }

        // DATETIME carries no zone; it holds UTC (MySqlAccountStore's header).
        private static DateTimeOffset Utc(DateTime value) => new(DateTime.SpecifyKind(value, DateTimeKind.Utc));

        private static string Cut(string value, int length) => value.Length <= length ? value : value[..length];
    }
}
