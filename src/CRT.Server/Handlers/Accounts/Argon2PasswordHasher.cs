using System.Security.Cryptography;
using System.Text;
using CRT.Server.Configuration;
using Konscious.Security.Cryptography;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // Hashes and verifies passwords with Argon2id.
    //
    // ARGON2ID, not bcrypt and never a bare SHA. The strategy document permits either Argon2id or
    // bcrypt; Argon2id is chosen because it is memory-hard, which is what makes a GPU or ASIC
    // attack expensive rather than merely slow. A bare SHA-256 of a password is not a password
    // hash at all - a commodity GPU tries billions per second.
    //
    // The parameters come from configuration (ServerOptions.Argon2*) and default to RFC 9106's
    // second recommended profile: 64 MiB, 3 iterations, parallelism 2. ServerOptionsValidator
    // refuses to start below a floor, so they cannot be quietly weakened to the point of being
    // decorative.
    //
    // *** THE RATE LIMITER MUST SIT IN FRONT OF THIS, NOT BEHIND IT. *** Each call allocates
    // roughly MemoryKib and holds it for the duration. At 64 MiB and parallelism 2 that is ~128
    // MiB transient per concurrent login, so unlimited login attempts are a memory-exhaustion
    // vector on a small box regardless of whether the passwords are right. See
    // AuthRateLimitPolicy, which is applied before the handler runs.
    //
    // VERIFICATION USES THE STORED PARAMETERS, never the configured ones - see
    // PasswordHashEncoding's header. That is what makes raising the cost safe.
    // ###########################################################################################
    public sealed class Argon2PasswordHasher
    {
        // 16 bytes is the RFC 9106 recommendation for a salt, and is what every mainstream
        // implementation uses. Longer buys nothing; shorter starts to make rainbow tables
        // conceivable across a large user base.
        private const int SaltBytes = 16;

        // 32 bytes of output. Matches the security level of the rest of the stack and is what the
        // PHC test vectors use.
        private const int HashBytes = 32;

        private readonly ServerOptions thisOptions;

        public Argon2PasswordHasher(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            this.thisOptions = options;
        }

        // ###########################################################################################
        // Hashes a new password with the CURRENT configured parameters, and a fresh random salt.
        //
        // The salt is per-password and never reused: two people choosing the same password must not
        // produce the same stored hash, or a single cracked hash reveals every account sharing it.
        // ###########################################################################################
        public string Hash(string password)
        {
            ArgumentException.ThrowIfNullOrEmpty(password);

            byte[] salt = RandomNumberGenerator.GetBytes(Argon2PasswordHasher.SaltBytes);

            byte[] hash = Argon2PasswordHasher.ComputeHash(
                password,
                salt,
                this.thisOptions.Argon2MemoryKib,
                this.thisOptions.Argon2Iterations,
                this.thisOptions.Argon2Parallelism);

            return PasswordHashEncoding.Encode(
                this.thisOptions.Argon2MemoryKib,
                this.thisOptions.Argon2Iterations,
                this.thisOptions.Argon2Parallelism,
                salt,
                hash);
        }

        // ###########################################################################################
        // Verifies a password against a stored hash.
        //
        // A MALFORMED STORED HASH FAILS THE LOGIN rather than throwing. The caller is an
        // authentication endpoint; an unparseable hash means that account cannot be logged into,
        // which is the safe answer, and the alternative - a 500 - tells an attacker that this
        // particular account is interesting.
        //
        // NeedsRehash is reported alongside so the caller can upgrade the stored hash at the one
        // moment the plaintext is legitimately in hand. It is only meaningful when the result is
        // true.
        // ###########################################################################################
        public PasswordVerificationResult Verify(string password, string? storedHash)
        {
            if (string.IsNullOrEmpty(password))
                return new PasswordVerificationResult(false, false);

            if (!PasswordHashEncoding.TryDecode(storedHash, out DecodedPasswordHash decoded))
                return new PasswordVerificationResult(false, false);

            byte[] computed = Argon2PasswordHasher.ComputeHash(
                password,
                decoded.Salt,
                decoded.MemoryKib,
                decoded.Iterations,
                decoded.Parallelism);

            // FIXED-TIME comparison. A byte-by-byte early-exit comparison leaks, through timing,
            // how many leading bytes were correct - which turns finding a collision into a
            // per-byte search rather than a search of the whole space.
            bool matches = CryptographicOperations.FixedTimeEquals(computed, decoded.Hash);

            if (!matches)
                return new PasswordVerificationResult(false, false);

            bool needsRehash = PasswordHashEncoding.NeedsRehash(
                decoded,
                this.thisOptions.Argon2MemoryKib,
                this.thisOptions.Argon2Iterations,
                this.thisOptions.Argon2Parallelism);

            return new PasswordVerificationResult(true, needsRehash);
        }

        private static byte[] ComputeHash(
            string password,
            byte[] salt,
            int memoryKib,
            int iterations,
            int parallelism)
        {
            using var argon2 = new Argon2id(Encoding.UTF8.GetBytes(password))
            {
                Salt = salt,
                MemorySize = memoryKib,
                Iterations = iterations,
                DegreeOfParallelism = parallelism
            };

            return argon2.GetBytes(Argon2PasswordHasher.HashBytes);
        }
    }

    // ###########################################################################################
    // NeedsRehash is only meaningful when IsValid is true - there is nothing to upgrade when the
    // password was wrong.
    // ###########################################################################################
    public readonly record struct PasswordVerificationResult(bool IsValid, bool NeedsRehash);
}
