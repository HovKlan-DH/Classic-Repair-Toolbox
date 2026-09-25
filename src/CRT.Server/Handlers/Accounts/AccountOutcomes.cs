namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // What the flows in AccountFlows answer with. These are the vocabulary the endpoints map onto
    // status codes, and they are deliberately COARSE where revealing more would leak.
    //
    // Note what is NOT here: there is no "no such account", no "wrong password" and no "account
    // locked". A login either succeeds, is rate limited, or fails - because distinguishing the
    // failures turns the one endpoint everybody can reach into an account enumeration oracle.
    // ###########################################################################################

    // ###########################################################################################
    // Registration. Accepted is returned BOTH for a new account and for an address that already
    // has one - see AccountFlows.RegisterAsync. Only Invalid is distinguishable, because that is
    // about the input rather than about who exists.
    // ###########################################################################################
    public sealed record RegistrationOutcome(bool IsAccepted, IReadOnlyList<string> Failures)
    {
        public static RegistrationOutcome Accepted() =>
            new(true, Array.Empty<string>());

        public static RegistrationOutcome Invalid(IReadOnlyList<string> failures) =>
            new(false, failures);
    }

    // ###########################################################################################
    // Redeeming a one-shot token (verification or password reset).
    //
    // These ARE distinguished, unlike login's, because the token is the only subject - none of
    // them reveals whether a particular account exists - and because the user needs different
    // things: an expired link means ask for a new one, an invalid one means check what you pasted.
    // ###########################################################################################
    public enum TokenRedemptionOutcome
    {
        Success,
        Invalid,
        Expired,
        AlreadyUsed,

        // Only from CompletePasswordResetAsync: the token was fine, the new password was not.
        PasswordRejected
    }

    // ###########################################################################################
    // A newly issued session. The RefreshToken here is the PLAINTEXT, which exists only in this
    // object on its way to the caller - the store holds nothing but its hash.
    // ###########################################################################################
    public sealed record IssuedSession(long SessionId, string RefreshToken, DateTimeOffset ExpiresUtc);

    // ###########################################################################################
    // Login. Three states, and the failure carries no reason on purpose.
    // ###########################################################################################
    public sealed record LoginOutcome(
        bool IsSuccess,
        bool IsRateLimited,
        TimeSpan RetryAfter,
        AccountRecord? Account,
        IssuedSession? Session)
    {
        public static LoginOutcome Succeeded(AccountRecord account, IssuedSession session) =>
            new(true, false, TimeSpan.Zero, account, session);

        public static LoginOutcome Failed() =>
            new(false, false, TimeSpan.Zero, null, null);

        public static LoginOutcome RateLimited(TimeSpan retryAfter) =>
            new(false, true, retryAfter, null, null);
    }

    // ###########################################################################################
    // Refresh. A failure is undifferentiated for the same reason login's is - and note that reuse
    // detection answers Failed() too, so a thief presenting a stolen token is told nothing about
    // having just triggered a revocation.
    // ###########################################################################################
    public sealed record RefreshOutcome(bool IsSuccess, AccountRecord? Account, IssuedSession? Session)
    {
        public static RefreshOutcome Succeeded(AccountRecord account, IssuedSession session) =>
            new(true, account, session);

        public static RefreshOutcome Failed() =>
            new(false, null, null);
    }

    // ###########################################################################################
    // Request shapes. Plain records rather than anything bound directly from JSON, so the
    // endpoint does the reading and the flow takes values - the thin-rim rule from the csproj.
    // ###########################################################################################
    public sealed record RegistrationRequest(
        string? Email,
        string? Password,
        string? DisplayName,
        string? IpAddress);

    public sealed record LoginRequest(
        string? Email,
        string? Password,
        string? UserAgent,
        string? IpAddress);
}
