using System.Security.Cryptography;
using System.Text;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // Generates the opaque tokens this service hands out - email verification links, password
    // reset links and refresh tokens - and hashes them for storage.
    //
    // *** ONLY THE HASH IS EVER STORED. *** The plaintext token exists exactly twice: in the mail
    // or HTTP response that carries it to its owner, and in memory for the instant it takes to
    // look it up. A database read - a backup left on a laptop, a SQL injection in some future
    // endpoint, a support person browsing tables - must not yield working password-reset links or
    // live sessions. This is the same reasoning as never storing a password, applied to the
    // credentials that stand in for one.
    //
    // WHY SHA-256 HERE, WHEN PASSWORDS NEED ARGON2. A password is low-entropy and chosen by a
    // human, so it must be made expensive to guess. These tokens are 256 bits from a CSPRNG -
    // there is nothing to guess, and no amount of hardware brute-forces a random 256-bit value.
    // Argon2 here would cost a memory-hard hash on every single authenticated request and buy
    // nothing at all.
    //
    // TOKENS ARE OPAQUE AND DATABASE-BACKED, NOT JWTs. Forced by Phase 6's definition of done:
    // "removing a maintainer takes effect immediately, proven by a test using an already-issued
    // token". No stateless token can satisfy that - being valid without a lookup is its whole
    // point. The cost is one indexed single-row SELECT per authenticated request, which is
    // nothing at this audience size.
    // ###########################################################################################
    public static class SecureToken
    {
        // 32 bytes = 256 bits. Base64url-encodes to 43 characters, which fits comfortably in a
        // URL and in a mail client's line wrapping.
        public const int TokenBytes = 32;

        // SHA-256 hex is always 64 characters, which is why the database columns are CHAR(64).
        public const int HashLength = 64;

        // ###########################################################################################
        // A new token, as a URL-safe string.
        //
        // BASE64URL, not plain base64: these travel in verification and reset links, and '+' and
        // '/' are both unsafe in a URL path or query - '+' silently becomes a space when a form
        // decoder gets hold of it, which would produce a token that looks right and never matches.
        // Padding is stripped for the same reason ('=' needs escaping).
        //
        // RandomNumberGenerator, never System.Random: Random is a predictable PRNG, and a
        // predictable password-reset token is a way into every account on the service.
        // ###########################################################################################
        public static string Create()
        {
            byte[] bytes = RandomNumberGenerator.GetBytes(SecureToken.TokenBytes);

            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        // ###########################################################################################
        // The value stored in the database and looked up on presentation.
        //
        // Lowercase hex so a lookup cannot miss on casing. The database column is CHAR(64) with a
        // case-insensitive collation, but relying on the collation for correctness would mean the
        // code breaks if the column is ever recreated differently.
        // ###########################################################################################
        public static string Hash(string token)
        {
            ArgumentException.ThrowIfNullOrEmpty(token);

            byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(token));

            return Convert.ToHexStringLower(hash);
        }

        // ###########################################################################################
        // Compares two token hashes in fixed time.
        //
        // Ordinarily a token is found by looking its hash up in an indexed column, which is a
        // database operation and not a comparison we control. This exists for the paths that
        // compare in memory - and it is the right default, because an early-exit string comparison
        // on a secret leaks through timing how many leading characters matched.
        // ###########################################################################################
        public static bool HashesEqual(string? left, string? right)
        {
            if (left is null || right is null)
                return false;

            if (left.Length != right.Length)
                return false;

            return CryptographicOperations.FixedTimeEquals(
                Encoding.UTF8.GetBytes(left),
                Encoding.UTF8.GetBytes(right));
        }
    }
}
