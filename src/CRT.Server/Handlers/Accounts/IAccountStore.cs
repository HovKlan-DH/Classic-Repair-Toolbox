namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // The "talk to the accounts database" seam. Exactly the same idea as IMiniproRunner and
    // IScopeClient in CRT.App: the flow logic depends on this interface, a MySQL implementation
    // sits behind it as an untested I/O boundary, and an in-memory fake stands in under test.
    //
    // WHY THIS EXISTS. CLAUDE.md test rule 6 forbids a test that needs a network call, and a
    // database is one. Without this seam, "registering with an existing address sends the
    // already-registered mail rather than creating a second account" could not be tested at all -
    // and that is precisely the logic most worth pinning down, because getting it wrong either
    // leaks who has an account or lets one person register twice.
    //
    // WHAT BELONGS HERE AND WHAT DOES NOT. These methods are deliberately dumb: find a row, insert
    // a row, mark a row consumed. Not one of them decides anything. Every rule - whether a token
    // has expired, whether a password is acceptable, which mail to send, whether a session may be
    // refreshed - lives in a pure class that takes this interface as an argument. If a decision
    // ends up inside an implementation of this interface, it is a decision no test can reach.
    //
    // NULLABLE RETURNS RATHER THAN EXCEPTIONS. "No account with that address" is the ordinary
    // case on every registration, not an exceptional one.
    // ###########################################################################################
    public interface IAccountStore
    {
        // -----------------------------------------------------------------------------------
        // Accounts.
        // -----------------------------------------------------------------------------------

        // Looked up by the NORMALISED address - see AccountRules.NormaliseEmail. Passing a raw
        // address here would miss an account registered with different capitalisation.
        Task<AccountRecord?> FindByNormalisedEmailAsync(string normalisedEmail, CancellationToken cancellationToken = default);

        Task<AccountRecord?> FindByIdAsync(long accountId, CancellationToken cancellationToken = default);

        // Several accounts in ONE query, keyed by id - for a screen that names many people (a
        // board's detail), which would otherwise ask once per person (code review, 2026-10-09). An
        // id with no account is simply absent.
        Task<IReadOnlyDictionary<long, AccountRecord>> FindByIdsAsync(IReadOnlyCollection<long> accountIds, CancellationToken cancellationToken = default);

        // Returns the new account's id. The caller has already validated everything; this inserts.
        Task<long> CreateAccountAsync(NewAccount account, CancellationToken cancellationToken = default);

        Task SetPasswordHashAsync(long accountId, string passwordHash, CancellationToken cancellationToken = default);

        Task SetVerifiedAsync(long accountId, CancellationToken cancellationToken = default);

        Task SetLastLoginAsync(long accountId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // The maintainer's own changes (2026-10-03) - see AccountSelfServiceFlows. The name is
        // already validated and trimmed.
        Task SetDisplayNameAsync(long accountId, string displayName, CancellationToken cancellationToken = default);

        // False - and nothing written - when another account holds the normalised address: the
        // unique index decides, so two accounts racing for one address cannot both win.
        Task<bool> SetEmailAsync(long accountId, string email, string normalisedEmail, CancellationToken cancellationToken = default);

        // -----------------------------------------------------------------------------------
        // One-shot tokens (verification, password reset).
        // -----------------------------------------------------------------------------------

        Task CreateTokenAsync(NewAccountToken token, CancellationToken cancellationToken = default);

        // Found by HASH, never by the token itself - only the hash is stored. See SecureToken.
        Task<AccountTokenRecord?> FindTokenByHashAsync(string tokenHash, CancellationToken cancellationToken = default);

        Task ConsumeTokenAsync(long tokenId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // Invalidates every outstanding token of one purpose for one account. Used when a password
        // is reset: any other reset link already in flight must stop working, or an attacker who
        // triggered one earlier still holds a way in.
        Task ConsumeOutstandingTokensAsync(long accountId, string purpose, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // -----------------------------------------------------------------------------------
        // Sessions (refresh tokens).
        // -----------------------------------------------------------------------------------

        Task<long> CreateSessionAsync(NewSession session, CancellationToken cancellationToken = default);

        Task<SessionRecord?> FindSessionByHashAsync(string refreshTokenHash, CancellationToken cancellationToken = default);

        // Marks a session rotated and points it at its successor - see SessionRecord.ReplacedById
        // for why that chain matters.
        Task RotateSessionAsync(long sessionId, long replacedBySessionId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        Task RevokeSessionAsync(long sessionId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // Pushes a live session's expiry forward WITHOUT issuing a new token - sliding expiry, so
        // a desktop maintainer signs in once rather than on a schedule. Deliberately NOT rotation:
        // see SessionExtensionRules' header for why rotating a file-backed token can lock an
        // account out of every session after nothing worse than a power cut.
        Task ExtendSessionAsync(long sessionId, DateTimeOffset newExpiresUtc, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // Revokes every live session for an account. Used on reuse detection and on password
        // change - both cases where every outstanding credential must stop working at once.
        Task RevokeAllSessionsAsync(long accountId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // Revokes every live session for an account EXCEPT the one making the request - a password
        // or address changed from the Maintainer tab signs out CRT on every other computer, and
        // leaves the one the maintainer is sitting at signed in.
        Task RevokeOtherSessionsAsync(long accountId, long keepSessionId, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // -----------------------------------------------------------------------------------
        // Maintainer pools (Phase 6 roles, 2026-09-25). One row per (board, account) in the
        // `maintainers` table; being in a pool is the whole of what makes an account a maintainer.
        // Read on EVERY review request (ReviewEndpoints.AuthoriseAsync), which is what makes
        // removal take effect on the next call rather than at next login.
        // -----------------------------------------------------------------------------------

        // The board ids this account reviews. Empty for an ordinary account, and for an
        // administrator too - the administrator is in every pool by definition and never needs
        // rows.
        Task<IReadOnlySet<string>> GetReviewedBoardIdsAsync(long accountId, CancellationToken cancellationToken = default);

        // Who reviews this board, with the display names system.json mirrors and the addresses a
        // new-submission mail goes to.
        Task<IReadOnlyList<MaintainerRecord>> GetMaintainersOfBoardAsync(string boardId, CancellationToken cancellationToken = default);

        // Every pool row there is, for the administrator's overview. Small by nature - a handful
        // of people across a few dozen boards.
        Task<IReadOnlyList<MaintainerRecord>> ListMaintainersAsync(CancellationToken cancellationToken = default);

        // Idempotent: adding somebody already in the pool changes nothing.
        Task AddMaintainerAsync(string boardId, long accountId, long grantedByAccountId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // Idempotent: removing somebody not in the pool changes nothing.
        Task RemoveMaintainerAsync(string boardId, long accountId, CancellationToken cancellationToken = default);

        // Every account, for the administrator to pick a maintainer from. Bounded, because a table
        // that has somehow grown large is a sign of abuse to look into, not a list to page.
        Task<IReadOnlyList<AccountRecord>> ListAccountsAsync(int limit, CancellationToken cancellationToken = default);

        // Where a submission goes when no maintainer is assigned to its board, or it changes
        // shared files.
        Task<IReadOnlyList<AccountRecord>> GetAdministratorsAsync(CancellationToken cancellationToken = default);

        // -----------------------------------------------------------------------------------
        // Rate limiting and audit.
        // -----------------------------------------------------------------------------------

        // The times of recent failed attempts, for AuthRateLimitPolicy. Either key may be null,
        // meaning "do not filter on this".
        Task<IReadOnlyList<DateTimeOffset>> GetRecentAuthFailuresAsync(
            string? normalisedEmail,
            string? ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default);

        Task RecordAuthFailureAsync(string? normalisedEmail, string? ipAddress, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        Task ClearAuthFailuresAsync(string normalisedEmail, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // The times of recent MAIL-SENDING requests from one address, for
        // AuthRateLimitPolicy.CheckMailRequest.
        //
        // SEPARATE FROM THE AUTH-FAILURE BUCKET, because it counts a different thing. An auth
        // failure is a FAILED attempt; here a SUCCESS is the thing being abused - registration and
        // "forgot my password" both send mail to an address the caller names, so an unlimited rate
        // is a way to use this service to flood somebody else's inbox, and every one of those
        // requests succeeds.
        //
        // Counted per IP only. There is deliberately no per-address bucket: the address belongs to
        // the VICTIM rather than to the caller, so limiting on it would let anyone lock a specific
        // person out of their own password-reset by spending the budget on their behalf.
        // ###########################################################################################
        Task<IReadOnlyList<DateTimeOffset>> GetRecentMailRequestsAsync(
            string ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default);

        Task RecordMailRequestAsync(string ipAddress, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        Task WriteAuditAsync(AuditEntry entry, CancellationToken cancellationToken = default);

        // The audit rows naming any of `subjects`, newest first - a board's history on the Boards
        // screen (2026-09-27): its id, and "#{id}" for each of its submissions.
        Task<IReadOnlyList<AuditEntry>> GetAuditForSubjectsAsync(IReadOnlyCollection<string> subjects, int limit, CancellationToken cancellationToken = default);

        // ---------------------------------------------------------------------------------------
        // Maintainer invitations (2026-09-27, migration 0012) - see MaintainerInvitationFlows.
        // ---------------------------------------------------------------------------------------

        Task<long> CreateInvitationAsync(NewMaintainerInvitation invitation, CancellationToken cancellationToken = default);

        Task<MaintainerInvitationRecord?> FindInvitationByHashAsync(string tokenHash, CancellationToken cancellationToken = default);

        Task<MaintainerInvitationRecord?> FindInvitationByIdAsync(long invitationId, CancellationToken cancellationToken = default);

        // Every invitation still OPEN at `nowUtc` - not accepted, not withdrawn, not expired.
        Task<IReadOnlyList<MaintainerInvitationRecord>> ListOpenInvitationsAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default);

        Task WithdrawInvitationAsync(long invitationId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Accepting: in ONE transaction, creates the account VERIFIED (the code proved the mailbox),
        // puts it in the pool of every board its address has an open invitation to, and marks
        // those invitations accepted. Null - and nothing written - when the address has an account
        // already (taken between the flow's check and this write).
        // ###########################################################################################
        Task<long?> AcceptInvitationsAsync(NewAccount account, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);
    }

    public sealed record NewMaintainerInvitation(
        string BoardId,
        string Email,
        string NormalisedEmail,
        string TokenHash,
        long InvitedByAccountId,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc);

    public sealed record MaintainerInvitationRecord(
        long Id,
        string BoardId,
        string Email,
        string NormalisedEmail,
        long? InvitedByAccountId,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc,
        DateTimeOffset? AcceptedUtc,
        DateTimeOffset? WithdrawnUtc)
    {
        public bool IsOpenAt(DateTimeOffset nowUtc) =>
            this.AcceptedUtc is null && this.WithdrawnUtc is null && this.ExpiresUtc > nowUtc;
    }

    // ###########################################################################################
    // An account as stored. PasswordHash is the full PHC string - see PasswordHashEncoding.
    // ###########################################################################################
    public sealed record AccountRecord(
        long Id,
        string Email,
        string NormalisedEmail,
        string PasswordHash,
        string DisplayName,
        bool IsVerified,
        bool IsAdministrator,
        bool IsLocked,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? LastLoginUtc);

    // ###########################################################################################
    // One pool row, joined to the account it names. BoardId is present so a whole-table listing
    // can be grouped by board without a second lookup.
    //
    // The account's three flags travel with it (code review, 2026-09-25): a pool row can outlive
    // the account's fitness to review - the account made administrator by hand, locked, or never
    // verified - and "does this board have a maintainer who can approve?" has to see that
    // (ReviewAuthority.CanGiveMaintainerApproval).
    // ###########################################################################################
    public sealed record MaintainerRecord(
        string BoardId,
        long AccountId,
        string DisplayName,
        string Email,
        bool IsAdministrator = false,
        bool IsVerified = true,
        bool IsLocked = false);

    public sealed record NewAccount(
        string Email,
        string NormalisedEmail,
        string PasswordHash,
        string DisplayName,
        DateTimeOffset CreatedUtc);

    // ###########################################################################################
    // The purposes a one-shot token can have. Strings rather than an enum because they are
    // stored as text and read in a log line; the constants stop them being mistyped.
    //
    // EmailChange (2026-10-03, migration 0016): the code mailed to a maintainer's NEW address. It
    // is the only purpose that carries a PendingEmail - the address the code is for.
    // ###########################################################################################
    public static class TokenPurpose
    {
        public const string EmailVerification = "email_verification";
        public const string PasswordReset = "password_reset";
        public const string EmailChange = "email_change";
    }

    public sealed record NewAccountToken(
        long AccountId,
        string Purpose,
        string TokenHash,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc,
        string? PendingEmail = null,
        string? PendingEmailNormalised = null);

    public sealed record AccountTokenRecord(
        long Id,
        long AccountId,
        string Purpose,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc,
        DateTimeOffset? ConsumedUtc,
        string? PendingEmail = null,
        string? PendingEmailNormalised = null);

    public sealed record NewSession(
        long AccountId,
        string RefreshTokenHash,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc,
        string? UserAgent,
        string? CreatedIp);

    // ###########################################################################################
    // A session as stored.
    //
    // ReplacedById is what makes REUSE DETECTION possible. Each refresh issues a new session and
    // points the old one at it. A refresh token that has already been rotated being presented
    // again means two parties hold it - the legitimate client and a thief - and the correct
    // response is to revoke the whole chain rather than guess which is which.
    // ###########################################################################################
    public sealed record SessionRecord(
        long Id,
        long AccountId,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc,
        DateTimeOffset? LastUsedUtc,
        DateTimeOffset? RevokedUtc,
        string? RevokedReason,
        long? ReplacedById);

    // ###########################################################################################
    // One audit row. ActorLabel carries a text copy of the identity because the account may later
    // be deleted, and an audit trail that loses its actor attributes nothing.
    // ###########################################################################################
    public sealed record AuditEntry(
        long? ActorAccountId,
        string ActorLabel,
        string Action,
        string? Subject,
        string? Detail,
        DateTimeOffset AtUtc);
}
