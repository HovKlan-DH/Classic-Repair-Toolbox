using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // An in-memory IAccountStore, so the account flows can be tested with no database.
    //
    // This is the direct equivalent of MockMiniproRunner in CRT.App: the interface is the seam,
    // the real implementation is an untested I/O boundary, and this stands in under test. Without
    // it, "registering with an existing address sends the already-registered mail rather than
    // creating a second account" could not be tested at all.
    //
    // IT BEHAVES LIKE THE REAL THING WHERE THAT MATTERS. In particular it looks accounts up by
    // NORMALISED email only, so a test cannot accidentally pass because the fake was more
    // forgiving than MariaDB's unique index.
    //
    // The public collections let a test assert on what was written - which sessions were revoked,
    // what audit rows exist - rather than only on the returned outcome.
    // ###########################################################################################
    public sealed class FakeAccountStore : IAccountStore
    {
        private long thisNextAccountId = 1;
        private long thisNextTokenId = 1;
        private long thisNextSessionId = 1;

        public Dictionary<long, AccountRecord> Accounts { get; } = [];

        public Dictionary<long, AccountTokenRecord> Tokens { get; } = [];

        // The plaintext hash that was stored for each token, so a test can look a token up the
        // way the flow does.
        public Dictionary<long, string> TokenHashes { get; } = [];

        public Dictionary<long, SessionRecord> Sessions { get; } = [];

        public Dictionary<long, string> SessionHashes { get; } = [];

        public List<AuditEntry> Audit { get; } = [];

        public List<(string? Email, string? Ip, DateTimeOffset When)> AuthFailures { get; } = [];

        // The maintainer pools: (system id, account id), exactly the `maintainers` table's key. A
        // test puts an account in a pool by adding to this directly.
        public HashSet<(string SystemId, long AccountId)> Maintainers { get; } = [];

        // -----------------------------------------------------------------------------------
        // Accounts.
        // -----------------------------------------------------------------------------------

        public Task<AccountRecord?> FindByNormalisedEmailAsync(string normalisedEmail, CancellationToken cancellationToken = default)
        {
            AccountRecord? found = this.Accounts.Values
                .FirstOrDefault(account => string.Equals(account.NormalisedEmail, normalisedEmail, StringComparison.Ordinal));

            return Task.FromResult(found);
        }

        public Task<AccountRecord?> FindByIdAsync(long accountId, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(this.Accounts.TryGetValue(accountId, out AccountRecord? account) ? account : null);
        }

        public Task<long> CreateAccountAsync(NewAccount account, CancellationToken cancellationToken = default)
        {
            long id = this.thisNextAccountId++;

            this.Accounts[id] = new AccountRecord(
                id,
                account.Email,
                account.NormalisedEmail,
                account.PasswordHash,
                account.DisplayName,
                IsVerified: false,
                IsAdministrator: false,
                IsLocked: false,
                account.CreatedUtc,
                LastLoginUtc: null);

            return Task.FromResult(id);
        }

        public Task SetPasswordHashAsync(long accountId, string passwordHash, CancellationToken cancellationToken = default)
        {
            this.Accounts[accountId] = this.Accounts[accountId] with { PasswordHash = passwordHash };
            return Task.CompletedTask;
        }

        public Task SetVerifiedAsync(long accountId, CancellationToken cancellationToken = default)
        {
            this.Accounts[accountId] = this.Accounts[accountId] with { IsVerified = true };
            return Task.CompletedTask;
        }

        public Task SetLastLoginAsync(long accountId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.Accounts[accountId] = this.Accounts[accountId] with { LastLoginUtc = whenUtc };
            return Task.CompletedTask;
        }

        // -----------------------------------------------------------------------------------
        // Tokens.
        // -----------------------------------------------------------------------------------

        public Task CreateTokenAsync(NewAccountToken token, CancellationToken cancellationToken = default)
        {
            long id = this.thisNextTokenId++;

            this.Tokens[id] = new AccountTokenRecord(
                id, token.AccountId, token.Purpose, token.CreatedUtc, token.ExpiresUtc, ConsumedUtc: null);

            this.TokenHashes[id] = token.TokenHash;

            return Task.CompletedTask;
        }

        public Task<AccountTokenRecord?> FindTokenByHashAsync(string tokenHash, CancellationToken cancellationToken = default)
        {
            foreach ((long id, string hash) in this.TokenHashes)
            {
                if (string.Equals(hash, tokenHash, StringComparison.Ordinal))
                    return Task.FromResult<AccountTokenRecord?>(this.Tokens[id]);
            }

            return Task.FromResult<AccountTokenRecord?>(null);
        }

        public Task ConsumeTokenAsync(long tokenId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.Tokens[tokenId] = this.Tokens[tokenId] with { ConsumedUtc = whenUtc };
            return Task.CompletedTask;
        }

        public Task ConsumeOutstandingTokensAsync(long accountId, string purpose, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            foreach (long id in this.Tokens.Keys.ToList())
            {
                AccountTokenRecord token = this.Tokens[id];

                if (token.AccountId == accountId &&
                    string.Equals(token.Purpose, purpose, StringComparison.Ordinal) &&
                    token.ConsumedUtc is null)
                {
                    this.Tokens[id] = token with { ConsumedUtc = whenUtc };
                }
            }

            return Task.CompletedTask;
        }

        // -----------------------------------------------------------------------------------
        // Sessions.
        // -----------------------------------------------------------------------------------

        public Task<long> CreateSessionAsync(NewSession session, CancellationToken cancellationToken = default)
        {
            long id = this.thisNextSessionId++;

            this.Sessions[id] = new SessionRecord(
                id, session.AccountId, session.CreatedUtc, session.ExpiresUtc,
                LastUsedUtc: null, RevokedUtc: null, RevokedReason: null, ReplacedById: null);

            this.SessionHashes[id] = session.RefreshTokenHash;

            return Task.FromResult(id);
        }

        public Task<SessionRecord?> FindSessionByHashAsync(string refreshTokenHash, CancellationToken cancellationToken = default)
        {
            foreach ((long id, string hash) in this.SessionHashes)
            {
                if (string.Equals(hash, refreshTokenHash, StringComparison.Ordinal))
                    return Task.FromResult<SessionRecord?>(this.Sessions[id]);
            }

            return Task.FromResult<SessionRecord?>(null);
        }

        public Task RotateSessionAsync(long sessionId, long replacedBySessionId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.Sessions[sessionId] = this.Sessions[sessionId] with
            {
                ReplacedById = replacedBySessionId,
                LastUsedUtc = whenUtc
            };

            return Task.CompletedTask;
        }

        public Task RevokeSessionAsync(long sessionId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.Sessions[sessionId] = this.Sessions[sessionId] with
            {
                RevokedUtc = whenUtc,
                RevokedReason = reason
            };

            return Task.CompletedTask;
        }

        // ###########################################################################################
        // Sliding expiry.
        //
        // *** THE GUARDS MIRROR MySqlAccountStore.ExtendSessionAsync's WHERE CLAUSE, and that is
        // the whole value of this method. *** A permissive fake would happily extend a revoked,
        // rotated or expired session and certify a caller the real database silently refuses -
        // so a test proving "a signed-out session cannot be revived" would prove nothing at all.
        // If that SQL gains or loses a condition, change this in the same sitting.
        // ###########################################################################################
        public Task ExtendSessionAsync(long sessionId, DateTimeOffset newExpiresUtc, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            if (!this.Sessions.TryGetValue(sessionId, out SessionRecord? session))
                return Task.CompletedTask;

            if (session.RevokedUtc is not null || session.ReplacedById is not null || session.ExpiresUtc <= whenUtc)
                return Task.CompletedTask;

            this.Sessions[sessionId] = session with
            {
                ExpiresUtc = newExpiresUtc,
                LastUsedUtc = whenUtc
            };

            this.ExtensionCount++;

            return Task.CompletedTask;
        }

        // How many times an extension actually landed. Lets a test assert that a read-only screen
        // is NOT writing a row on every single request, which is the thing the half-lifetime
        // threshold exists to prevent.
        public int ExtensionCount { get; private set; }

        public Task RevokeAllSessionsAsync(long accountId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            foreach (long id in this.Sessions.Keys.ToList())
            {
                SessionRecord session = this.Sessions[id];

                if (session.AccountId == accountId && session.RevokedUtc is null)
                    this.Sessions[id] = session with { RevokedUtc = whenUtc, RevokedReason = reason };
            }

            return Task.CompletedTask;
        }

        // -----------------------------------------------------------------------------------
        // Rate limiting and audit.
        // -----------------------------------------------------------------------------------

        public Task<IReadOnlyList<DateTimeOffset>> GetRecentAuthFailuresAsync(
            string? normalisedEmail,
            string? ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<DateTimeOffset> matches = this.AuthFailures
                .Where(failure => failure.When > since)
                .Where(failure => normalisedEmail is null || string.Equals(failure.Email, normalisedEmail, StringComparison.Ordinal))
                .Where(failure => ipAddress is null || string.Equals(failure.Ip, ipAddress, StringComparison.Ordinal))
                .Select(failure => failure.When)
                .ToList();

            return Task.FromResult(matches);
        }

        public Task RecordAuthFailureAsync(string? normalisedEmail, string? ipAddress, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.AuthFailures.Add((normalisedEmail, ipAddress, whenUtc));
            return Task.CompletedTask;
        }

        public Task ClearAuthFailuresAsync(string normalisedEmail, CancellationToken cancellationToken = default)
        {
            this.AuthFailures.RemoveAll(failure => string.Equals(failure.Email, normalisedEmail, StringComparison.Ordinal));
            return Task.CompletedTask;
        }

        // The mail-sending budget, per IP. Kept as its own list for the same reason the real store
        // keys it on its own action: it counts successful requests, not failures.
        public List<(string Ip, DateTimeOffset When)> MailRequests { get; } = [];

        public Task<IReadOnlyList<DateTimeOffset>> GetRecentMailRequestsAsync(
            string ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<DateTimeOffset> matches = this.MailRequests
                .Where(request => request.When > since)
                .Where(request => string.Equals(request.Ip, ipAddress, StringComparison.Ordinal))
                .Select(request => request.When)
                .ToList();

            return Task.FromResult(matches);
        }

        public Task RecordMailRequestAsync(string ipAddress, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.MailRequests.Add((ipAddress, whenUtc));
            return Task.CompletedTask;
        }

        // -----------------------------------------------------------------------------------
        // Maintainer pools (Phase 6 roles).
        // -----------------------------------------------------------------------------------

        public Task<IReadOnlySet<string>> GetReviewedSystemIdsAsync(long accountId, CancellationToken cancellationToken = default)
        {
            IReadOnlySet<string> ids = this.Maintainers
                .Where(row => row.AccountId == accountId)
                .Select(row => row.SystemId)
                .ToHashSet(StringComparer.Ordinal);

            return Task.FromResult(ids);
        }

        public Task<IReadOnlyList<MaintainerRecord>> GetMaintainersOfSystemAsync(string systemId, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(this.MaintainerRows(row => string.Equals(row.SystemId, systemId, StringComparison.Ordinal)));
        }

        public Task<IReadOnlyList<MaintainerRecord>> ListMaintainersAsync(CancellationToken cancellationToken = default)
        {
            return Task.FromResult(this.MaintainerRows(_ => true));
        }

        // Joined to the account the way the real query is, and an orphan pair (no such account)
        // is dropped the way an inner join drops it.
        private IReadOnlyList<MaintainerRecord> MaintainerRows(Func<(string SystemId, long AccountId), bool> where)
        {
            return this.Maintainers
                .Where(where)
                .Where(row => this.Accounts.ContainsKey(row.AccountId))
                .Select(row => new MaintainerRecord(
                    row.SystemId,
                    row.AccountId,
                    this.Accounts[row.AccountId].DisplayName,
                    this.Accounts[row.AccountId].Email,
                    this.Accounts[row.AccountId].IsAdministrator,
                    this.Accounts[row.AccountId].IsVerified,
                    this.Accounts[row.AccountId].IsLocked))
                .OrderBy(row => row.SystemId, StringComparer.Ordinal)
                .ThenBy(row => row.DisplayName, StringComparer.Ordinal)
                .ThenBy(row => row.AccountId)
                .ToList();
        }

        public Task AddMaintainerAsync(string systemId, long accountId, long grantedByAccountId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            this.Maintainers.Add((systemId, accountId));
            return Task.CompletedTask;
        }

        public Task RemoveMaintainerAsync(string systemId, long accountId, CancellationToken cancellationToken = default)
        {
            this.Maintainers.Remove((systemId, accountId));
            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<AccountRecord>> ListAccountsAsync(int limit, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<AccountRecord> accounts = this.Accounts.Values
                .OrderBy(account => account.DisplayName, StringComparer.Ordinal)
                .ThenBy(account => account.Id)
                .Take(Math.Clamp(limit, 1, 1000))
                .ToList();

            return Task.FromResult(accounts);
        }

        public Task<IReadOnlyList<AccountRecord>> GetAdministratorsAsync(CancellationToken cancellationToken = default)
        {
            IReadOnlyList<AccountRecord> admins = this.Accounts.Values
                .Where(account => account.IsAdministrator)
                .OrderBy(account => account.Id)
                .ToList();

            return Task.FromResult(admins);
        }

        // ---------------------------------------------------------------------------------------
        // Maintainer invitations (2026-09-27). TokenHash is kept beside each record, as the tokens'
        // hashes are, so a test can prove the plaintext code was never stored.
        // ---------------------------------------------------------------------------------------

        private long thisNextInvitationId = 1;

        public Dictionary<long, MaintainerInvitationRecord> Invitations { get; } = [];

        public Dictionary<long, string> InvitationHashes { get; } = [];

        public Task<long> CreateInvitationAsync(NewMaintainerInvitation invitation, CancellationToken cancellationToken = default)
        {
            long id = this.thisNextInvitationId++;

            this.Invitations[id] = new MaintainerInvitationRecord(
                id,
                invitation.SystemId,
                invitation.Email,
                invitation.NormalisedEmail,
                invitation.InvitedByAccountId,
                invitation.CreatedUtc,
                invitation.ExpiresUtc,
                AcceptedUtc: null,
                WithdrawnUtc: null);

            this.InvitationHashes[id] = invitation.TokenHash;

            return Task.FromResult(id);
        }

        public Task<MaintainerInvitationRecord?> FindInvitationByHashAsync(string tokenHash, CancellationToken cancellationToken = default)
        {
            long id = this.InvitationHashes.FirstOrDefault(pair => pair.Value == tokenHash).Key;
            return Task.FromResult(this.Invitations.GetValueOrDefault(id));
        }

        public Task<MaintainerInvitationRecord?> FindInvitationByIdAsync(long invitationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Invitations.GetValueOrDefault(invitationId));

        public Task<IReadOnlyList<MaintainerInvitationRecord>> ListOpenInvitationsAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<MaintainerInvitationRecord> open = this.Invitations.Values
                .Where(invitation => invitation.IsOpenAt(nowUtc))
                .OrderBy(invitation => invitation.SystemId, StringComparer.Ordinal)
                .ThenBy(invitation => invitation.CreatedUtc)
                .ThenBy(invitation => invitation.Id)
                .ToList();

            return Task.FromResult(open);
        }

        public Task WithdrawInvitationAsync(long invitationId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            if (this.Invitations.TryGetValue(invitationId, out MaintainerInvitationRecord? invitation) &&
                invitation.AcceptedUtc is null && invitation.WithdrawnUtc is null)
            {
                this.Invitations[invitationId] = invitation with { WithdrawnUtc = whenUtc };
            }

            return Task.CompletedTask;
        }

        public Task<long?> AcceptInvitationsAsync(NewAccount account, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            if (this.Accounts.Values.Any(existing => existing.NormalisedEmail == account.NormalisedEmail))
                return Task.FromResult<long?>(null);

            long id = this.thisNextAccountId++;

            this.Accounts[id] = new AccountRecord(
                id, account.Email, account.NormalisedEmail, account.PasswordHash, account.DisplayName,
                IsVerified: true, IsAdministrator: false, IsLocked: false, account.CreatedUtc, LastLoginUtc: null);

            foreach (MaintainerInvitationRecord invitation in this.Invitations.Values
                .Where(invitation => invitation.NormalisedEmail == account.NormalisedEmail && invitation.IsOpenAt(whenUtc))
                .ToList())
            {
                this.Maintainers.Add((invitation.SystemId, id));
                this.Invitations[invitation.Id] = invitation with { AcceptedUtc = whenUtc };
            }

            return Task.FromResult<long?>(id);
        }

        public Task<IReadOnlyList<AuditEntry>> GetAuditForSubjectsAsync(IReadOnlyCollection<string> subjects, int limit, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<AuditEntry> found = this.Audit
                .Select((entry, index) => (entry, index))
                .Where(item => item.entry.Subject is not null && subjects.Contains(item.entry.Subject))
                .OrderByDescending(item => item.entry.AtUtc)
                .ThenByDescending(item => item.index)
                .Take(Math.Clamp(limit, 1, 1000))
                .Select(item => item.entry)
                .ToList();

            return Task.FromResult(found);
        }

        public Task WriteAuditAsync(AuditEntry entry, CancellationToken cancellationToken = default)
        {
            this.Audit.Add(entry);
            return Task.CompletedTask;
        }
    }
}
