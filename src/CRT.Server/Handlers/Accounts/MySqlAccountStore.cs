using CRT.Server.Configuration;
using MySqlConnector;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // The real IAccountStore, against MariaDB.
    //
    // AN UNTESTED I/O BOUNDARY, by the same rule as SmtpEmailSender and as ScopeScpiClient in
    // CRT.App. There are no decisions in this file - every rule about expiry, rotation, reuse and
    // enumeration lives in AccountFlows, which is fully tested against FakeAccountStore. What is
    // left here is SQL, and testing it would mean testing MariaDB.
    //
    // THREE RULES THIS FILE FOLLOWS THROUGHOUT:
    //
    // 1. EVERY value reaching SQL goes through a PARAMETER. Never string concatenation, not even
    //    for an integer id, not even for a value this service generated itself. The moment one
    //    call is written differently, it becomes the one a reviewer has to reason about.
    //
    // 2. DATETIME(3) columns hold UTC (see 0001_initial.sql). Values go in as UtcDateTime and come
    //    back tagged Utc, so nothing in the rest of the service has to remember.
    //
    // 3. A connection per operation, from the pool. MySqlConnector pools by connection string, so
    //    this is cheap, and it avoids a long-lived connection being held across an await by a
    //    request that then fails.
    // ###########################################################################################
    public sealed class MySqlAccountStore : IAccountStore
    {
        private readonly string thisConnectionString;

        public MySqlAccountStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        // -----------------------------------------------------------------------------------
        // Accounts.
        // -----------------------------------------------------------------------------------

        private const string AccountColumns =
            "id, email, email_normalised, password_hash, display_name, is_verified, " +
            "is_administrator, is_reviewer, is_locked, created_utc, last_login_utc";

        public async Task<AccountRecord?> FindByNormalisedEmailAsync(string normalisedEmail, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                $"SELECT {MySqlAccountStore.AccountColumns} FROM accounts WHERE email_normalised = @email LIMIT 1;";
            command.Parameters.AddWithValue("@email", normalisedEmail);

            return await MySqlAccountStore.ReadAccountAsync(command, cancellationToken);
        }

        public async Task<AccountRecord?> FindByIdAsync(long accountId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText =
                $"SELECT {MySqlAccountStore.AccountColumns} FROM accounts WHERE id = @id LIMIT 1;";
            command.Parameters.AddWithValue("@id", accountId);

            return await MySqlAccountStore.ReadAccountAsync(command, cancellationToken);
        }

        public async Task<long> CreateAccountAsync(NewAccount account, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                INSERT INTO accounts (email, email_normalised, password_hash, display_name, created_utc)
                VALUES (@email, @normalised, @hash, @displayName, @created);
                SELECT LAST_INSERT_ID();
                """;

            command.Parameters.AddWithValue("@email", account.Email);
            command.Parameters.AddWithValue("@normalised", account.NormalisedEmail);
            command.Parameters.AddWithValue("@hash", account.PasswordHash);
            command.Parameters.AddWithValue("@displayName", account.DisplayName);
            command.Parameters.AddWithValue("@created", account.CreatedUtc.UtcDateTime);

            object? id = await command.ExecuteScalarAsync(cancellationToken);

            return Convert.ToInt64(id, System.Globalization.CultureInfo.InvariantCulture);
        }

        public Task SetPasswordHashAsync(long accountId, string passwordHash, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                "UPDATE accounts SET password_hash = @hash WHERE id = @id;",
                command =>
                {
                    command.Parameters.AddWithValue("@hash", passwordHash);
                    command.Parameters.AddWithValue("@id", accountId);
                },
                cancellationToken);
        }

        public Task SetVerifiedAsync(long accountId, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                "UPDATE accounts SET is_verified = 1 WHERE id = @id;",
                command => command.Parameters.AddWithValue("@id", accountId),
                cancellationToken);
        }

        public Task SetLastLoginAsync(long accountId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                "UPDATE accounts SET last_login_utc = @when WHERE id = @id;",
                command =>
                {
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@id", accountId);
                },
                cancellationToken);
        }

        // -----------------------------------------------------------------------------------
        // Tokens.
        // -----------------------------------------------------------------------------------

        public Task CreateTokenAsync(NewAccountToken token, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                INSERT INTO account_tokens (account_id, purpose, token_hash, created_utc, expires_utc)
                VALUES (@accountId, @purpose, @hash, @created, @expires);
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@accountId", token.AccountId);
                    command.Parameters.AddWithValue("@purpose", token.Purpose);
                    command.Parameters.AddWithValue("@hash", token.TokenHash);
                    command.Parameters.AddWithValue("@created", token.CreatedUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@expires", token.ExpiresUtc.UtcDateTime);
                },
                cancellationToken);
        }

        public async Task<AccountTokenRecord?> FindTokenByHashAsync(string tokenHash, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT id, account_id, purpose, created_utc, expires_utc, consumed_utc
                FROM account_tokens WHERE token_hash = @hash LIMIT 1;
                """;

            command.Parameters.AddWithValue("@hash", tokenHash);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            if (!await reader.ReadAsync(cancellationToken))
                return null;

            return new AccountTokenRecord(
                reader.GetInt64(0),
                reader.GetInt64(1),
                reader.GetString(2),
                MySqlAccountStore.ReadUtc(reader, 3)!.Value,
                MySqlAccountStore.ReadUtc(reader, 4)!.Value,
                MySqlAccountStore.ReadUtc(reader, 5));
        }

        public Task ConsumeTokenAsync(long tokenId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            // "AND consumed_utc IS NULL" so a concurrent second redemption of the same link
            // cannot overwrite the first one's timestamp.
            return this.ExecuteAsync(
                "UPDATE account_tokens SET consumed_utc = @when WHERE id = @id AND consumed_utc IS NULL;",
                command =>
                {
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@id", tokenId);
                },
                cancellationToken);
        }

        public Task ConsumeOutstandingTokensAsync(long accountId, string purpose, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE account_tokens SET consumed_utc = @when
                WHERE account_id = @accountId AND purpose = @purpose AND consumed_utc IS NULL;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@accountId", accountId);
                    command.Parameters.AddWithValue("@purpose", purpose);
                },
                cancellationToken);
        }

        // -----------------------------------------------------------------------------------
        // Sessions.
        // -----------------------------------------------------------------------------------

        public async Task<long> CreateSessionAsync(NewSession session, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                INSERT INTO sessions (account_id, refresh_token_hash, created_utc, expires_utc, user_agent, created_ip)
                VALUES (@accountId, @hash, @created, @expires, @userAgent, @ip);
                SELECT LAST_INSERT_ID();
                """;

            command.Parameters.AddWithValue("@accountId", session.AccountId);
            command.Parameters.AddWithValue("@hash", session.RefreshTokenHash);
            command.Parameters.AddWithValue("@created", session.CreatedUtc.UtcDateTime);
            command.Parameters.AddWithValue("@expires", session.ExpiresUtc.UtcDateTime);
            command.Parameters.AddWithValue("@userAgent", (object?)session.UserAgent ?? DBNull.Value);
            command.Parameters.AddWithValue("@ip", (object?)session.CreatedIp ?? DBNull.Value);

            object? id = await command.ExecuteScalarAsync(cancellationToken);

            return Convert.ToInt64(id, System.Globalization.CultureInfo.InvariantCulture);
        }

        public async Task<SessionRecord?> FindSessionByHashAsync(string refreshTokenHash, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT id, account_id, created_utc, expires_utc, last_used_utc,
                       revoked_utc, revoked_reason, replaced_by_id
                FROM sessions WHERE refresh_token_hash = @hash LIMIT 1;
                """;

            command.Parameters.AddWithValue("@hash", refreshTokenHash);

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            if (!await reader.ReadAsync(cancellationToken))
                return null;

            return new SessionRecord(
                reader.GetInt64(0),
                reader.GetInt64(1),
                MySqlAccountStore.ReadUtc(reader, 2)!.Value,
                MySqlAccountStore.ReadUtc(reader, 3)!.Value,
                MySqlAccountStore.ReadUtc(reader, 4),
                MySqlAccountStore.ReadUtc(reader, 5),
                reader.IsDBNull(6) ? null : reader.GetString(6),
                reader.IsDBNull(7) ? null : reader.GetInt64(7));
        }

        public Task RotateSessionAsync(long sessionId, long replacedBySessionId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            // "AND replaced_by_id IS NULL" makes the rotation atomic against a concurrent refresh
            // with the same token: only one can win, and the loser's chain is then detected as
            // reuse on its next presentation - which is exactly the intended behaviour.
            return this.ExecuteAsync(
                """
                UPDATE sessions SET replaced_by_id = @replacedBy, last_used_utc = @when
                WHERE id = @id AND replaced_by_id IS NULL;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@replacedBy", replacedBySessionId);
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@id", sessionId);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // SLIDING EXPIRY - push a live session's expiry forward, leaving its token untouched.
        //
        // *** THE WHERE CLAUSE IS THE SECURITY, NOT THE CALLER. *** Every condition the
        // authenticated path checks is repeated here as part of the UPDATE, so the row can only
        // move if it is still genuinely usable at the moment the write lands:
        //
        //   - `revoked_utc IS NULL`    - a signed-out or force-revoked session must stay dead.
        //     Without this, a sign-out racing an in-flight request could be undone by the
        //     extension that follows it, and the session would outlive the revocation.
        //   - `replaced_by_id IS NULL` - a rotated session is spent; extending one would hand a
        //     superseded token a fresh lease.
        //   - `expires_utc > @when`    - an ALREADY EXPIRED session is never resurrected. This is
        //     the one that matters most: without it, a stale token found in an old backup could
        //     be extended back to life indefinitely.
        //
        // So a TOCTOU gap between the caller's check and this write cannot grant anything - the
        // database re-decides atomically. Zero rows affected is a correct, silent outcome; the
        // session simply keeps the expiry it had.
        //
        // `last_used_utc` is written alongside, which is what finally makes that column mean
        // something: until now only rotation touched it, so a session used daily for a month
        // still read as never used.
        // ###########################################################################################
        public Task ExtendSessionAsync(long sessionId, DateTimeOffset newExpiresUtc, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE sessions SET expires_utc = @expires, last_used_utc = @when
                WHERE id = @id
                  AND revoked_utc IS NULL
                  AND replaced_by_id IS NULL
                  AND expires_utc > @when;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@expires", newExpiresUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@id", sessionId);
                },
                cancellationToken);
        }

        public Task RevokeSessionAsync(long sessionId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE sessions SET revoked_utc = @when, revoked_reason = @reason
                WHERE id = @id AND revoked_utc IS NULL;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@reason", reason);
                    command.Parameters.AddWithValue("@id", sessionId);
                },
                cancellationToken);
        }

        public Task RevokeAllSessionsAsync(long accountId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                UPDATE sessions SET revoked_utc = @when, revoked_reason = @reason
                WHERE account_id = @accountId AND revoked_utc IS NULL;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                    command.Parameters.AddWithValue("@reason", reason);
                    command.Parameters.AddWithValue("@accountId", accountId);
                },
                cancellationToken);
        }

        // -----------------------------------------------------------------------------------
        // Rate limiting.
        //
        // Failures are recorded in the audit table rather than in a table of their own: they are
        // exactly the kind of security event the audit trail is for, and a separate table would
        // need its own cleanup. The action name is what makes them findable.
        // -----------------------------------------------------------------------------------

        private const string AuthFailureAction = "login.failed";

        public async Task<IReadOnlyList<DateTimeOffset>> GetRecentAuthFailuresAsync(
            string? normalisedEmail,
            string? ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            // The subject holds the normalised address and the detail holds the IP, so one table
            // answers both buckets. Both are compared with a parameter, never interpolated.
            command.CommandText = """
                SELECT at_utc FROM audit
                WHERE action = @action AND at_utc > @since
                  AND (@subject IS NULL OR subject = @subject)
                  AND (@detail IS NULL OR detail = @detail)
                LIMIT 500;
                """;

            command.Parameters.AddWithValue("@action", MySqlAccountStore.AuthFailureAction);
            command.Parameters.AddWithValue("@since", since.UtcDateTime);
            command.Parameters.AddWithValue("@subject", (object?)normalisedEmail ?? DBNull.Value);
            command.Parameters.AddWithValue("@detail", (object?)ipAddress ?? DBNull.Value);

            var times = new List<DateTimeOffset>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                times.Add(MySqlAccountStore.ReadUtc(reader, 0)!.Value);

            return times;
        }

        public Task RecordAuthFailureAsync(string? normalisedEmail, string? ipAddress, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            return this.WriteAuditAsync(
                new AuditEntry(
                    null,
                    normalisedEmail ?? "unknown",
                    MySqlAccountStore.AuthFailureAction,
                    normalisedEmail,
                    ipAddress,
                    whenUtc),
                cancellationToken);
        }

        public Task ClearAuthFailuresAsync(string normalisedEmail, CancellationToken cancellationToken = default)
        {
            // The audit trail is APPEND-ONLY, so a successful login does not delete rows - it
            // renames the action, which keeps the history while taking those rows out of the
            // limiter's window. Deleting them would destroy the record that the attempts happened,
            // which is precisely what an attacker would want.
            return this.ExecuteAsync(
                """
                UPDATE audit SET action = 'login.failed.cleared'
                WHERE action = @action AND subject = @subject;
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@action", MySqlAccountStore.AuthFailureAction);
                    command.Parameters.AddWithValue("@subject", normalisedEmail);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // The mail-request bucket. Stored in the same `audit` table as everything else, under its
        // own action, with the IP in `subject` rather than in `detail`.
        //
        // SUBJECT, NOT DETAIL, AND THAT IS DELIBERATE: `subject` is VARCHAR(255) and carries the
        // ix_audit_subject index, while `detail` is TEXT and is indexed by nothing. This lookup
        // runs on every registration and every password-reset request, so it has to be an indexed
        // one - the auth-failure query gets away with its unindexed `detail` comparison only
        // because it is already narrowed by an indexed action-and-time range.
        // ###########################################################################################
        private const string MailRequestAction = "mail.requested";

        public async Task<IReadOnlyList<DateTimeOffset>> GetRecentMailRequestsAsync(
            string ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = """
                SELECT at_utc FROM audit
                WHERE action = @action AND subject = @subject AND at_utc > @since
                LIMIT 500;
                """;

            command.Parameters.AddWithValue("@action", MySqlAccountStore.MailRequestAction);
            command.Parameters.AddWithValue("@subject", ipAddress);
            command.Parameters.AddWithValue("@since", since.UtcDateTime);

            var times = new List<DateTimeOffset>();

            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                times.Add(MySqlAccountStore.ReadUtc(reader, 0)!.Value);

            return times;
        }

        public Task RecordMailRequestAsync(string ipAddress, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            // The ADDRESS THE MAIL WOULD GO TO IS NOT RECORDED. It is the victim's address in the
            // abuse case, and writing it here would build a log of which addresses a stranger
            // probed - the very list an enumeration attack is trying to assemble.
            return this.WriteAuditAsync(
                new AuditEntry(
                    null,
                    ipAddress,
                    MySqlAccountStore.MailRequestAction,
                    ipAddress,
                    null,
                    whenUtc),
                cancellationToken);
        }

        public Task WriteAuditAsync(AuditEntry entry, CancellationToken cancellationToken = default)
        {
            return this.ExecuteAsync(
                """
                INSERT INTO audit (actor_account_id, actor_label, action, subject, detail, at_utc)
                VALUES (@actorId, @actorLabel, @action, @subject, @detail, @at);
                """,
                command =>
                {
                    command.Parameters.AddWithValue("@actorId", (object?)entry.ActorAccountId ?? DBNull.Value);
                    command.Parameters.AddWithValue("@actorLabel", entry.ActorLabel);
                    command.Parameters.AddWithValue("@action", entry.Action);
                    command.Parameters.AddWithValue("@subject", (object?)entry.Subject ?? DBNull.Value);
                    command.Parameters.AddWithValue("@detail", (object?)entry.Detail ?? DBNull.Value);
                    command.Parameters.AddWithValue("@at", entry.AtUtc.UtcDateTime);
                },
                cancellationToken);
        }

        // -----------------------------------------------------------------------------------
        // Plumbing.
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

        private static async Task<AccountRecord?> ReadAccountAsync(MySqlCommand command, CancellationToken cancellationToken)
        {
            await using MySqlDataReader reader = await command.ExecuteReaderAsync(cancellationToken);

            if (!await reader.ReadAsync(cancellationToken))
                return null;

            return new AccountRecord(
                reader.GetInt64(0),
                reader.GetString(1),
                reader.GetString(2),
                reader.GetString(3),
                reader.GetString(4),
                reader.GetBoolean(5),
                reader.GetBoolean(6),
                reader.GetBoolean(7),
                reader.GetBoolean(8),
                MySqlAccountStore.ReadUtc(reader, 9)!.Value,
                MySqlAccountStore.ReadUtc(reader, 10));
        }

        // DATETIME has no timezone of its own, so a value read back is Unspecified. Tagging it Utc
        // is what makes every comparison in AccountFlows correct - an Unspecified DateTime
        // converted to DateTimeOffset would be interpreted in the SERVER'S LOCAL ZONE, silently
        // shifting every expiry by the UTC offset.
        private static DateTimeOffset? ReadUtc(MySqlDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal))
                return null;

            return new DateTimeOffset(
                DateTime.SpecifyKind(reader.GetDateTime(ordinal), DateTimeKind.Utc));
        }
    }
}
