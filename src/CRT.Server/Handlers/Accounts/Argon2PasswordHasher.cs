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
    // *** AND NO MORE COMPUTATIONS RUN AT ONCE THAN THE MACHINE HAS CORES FOR (code review,
    // 2026-10-11). *** The limiter counts per address and per IP, and every unknown address now
    // costs a full computation (VerifyWithoutAccountAsync) - so many addresses from many IPs could
    // still run any number at once. They wait for a turn instead (ConcurrentComputations), without
    // holding a thread; more at once would not finish any sooner, only take more memory.
    //
    // VERIFICATION USES THE STORED PARAMETERS, never the configured ones - see
    // PasswordHashEncoding's header. That is what makes raising the cost safe - and why a wrong
    // password is padded to the configured cost (PaddingMemoryKib).
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

        // The salt the computations whose result is thrown away hash with. Any salt costs the same;
        // this one is random so the work cannot be skipped by anything that recognises it.
        private static readonly byte[] NoAccountSalt = RandomNumberGenerator.GetBytes(Argon2PasswordHasher.SaltBytes);

        private readonly ServerOptions thisOptions;

        // A turn to compute - see the header.
        private readonly SemaphoreSlim thisTurns;

        // What this hasher has computed. Read by tests only: it is how AccountFlowTests proves every
        // failed login costs the same work, which a wall-clock test could not show without being
        // flaky. Counters, never a list - the hasher lives as long as the service.
        private readonly object thisCountsLock = new();
        private long thisComputations;
        private double thisWork;
        private int thisRunning;
        private int thisMostRunning;

        internal long Computations
        {
            get
            {
                lock (this.thisCountsLock)
                    return this.thisComputations;
            }
        }

        // The cost of every computation so far, added up (Argon2PasswordHasher.Cost).
        internal double Work
        {
            get
            {
                lock (this.thisCountsLock)
                    return this.thisWork;
            }
        }

        // The most computations that ever ran at the same moment.
        internal int MostRunning
        {
            get
            {
                lock (this.thisCountsLock)
                    return this.thisMostRunning;
            }
        }

        public Argon2PasswordHasher(ServerOptions options)
            : this(options, Argon2PasswordHasher.ConcurrentComputations(Environment.ProcessorCount, options?.Argon2Parallelism ?? 1))
        {
        }

        // `concurrentComputations` is for tests - the machine's cores otherwise.
        internal Argon2PasswordHasher(ServerOptions options, int concurrentComputations)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentOutOfRangeException.ThrowIfLessThan(concurrentComputations, 1);

            this.thisOptions = options;
            this.thisTurns = new SemaphoreSlim(concurrentComputations, concurrentComputations);
        }

        // ###########################################################################################
        // How many computations may run at once: as many as there are cores for each one's lanes,
        // and at least one. Each computation runs `parallelism` lanes side by side, so on a machine
        // with fewer cores than that it already has them all.
        // ###########################################################################################
        internal static int ConcurrentComputations(int processorCount, int parallelism) =>
            Math.Max(1, processorCount / Math.Max(1, parallelism));

        // ###########################################################################################
        // Hashes a new password with the CURRENT configured parameters, and a fresh random salt.
        //
        // The salt is per-password and never reused: two people choosing the same password must not
        // produce the same stored hash, or a single cracked hash reveals every account sharing it.
        // ###########################################################################################
        public Task<string> HashAsync(string password, CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrEmpty(password);

            return this.InTurnAsync(
                () =>
                {
                    byte[] salt = RandomNumberGenerator.GetBytes(Argon2PasswordHasher.SaltBytes);

                    byte[] hash = this.ComputeHash(
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
                },
                cancellationToken);
        }

        // ###########################################################################################
        // Verifies a password against a stored hash.
        //
        // A MALFORMED STORED HASH FAILS THE LOGIN rather than throwing. The caller is an
        // authentication endpoint; an unparseable hash means that account cannot be logged into,
        // which is the safe answer, and the alternative - a 500 - tells an attacker that this
        // particular account is interesting. For the same reason it costs what a real verification
        // costs (VerifyWithoutAccountAsync): failing fast would single the account out by its timing.
        //
        // A WRONG PASSWORD COSTS THE CONFIGURED PARAMETERS' WORK, whatever the hash was made with -
        // see PaddingMemoryKib.
        //
        // NeedsRehash is reported alongside so the caller can upgrade the stored hash at the one
        // moment the plaintext is legitimately in hand. It is only meaningful when the result is
        // true.
        // ###########################################################################################
        public Task<PasswordVerificationResult> VerifyAsync(string password, string? storedHash, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(password))
                return Task.FromResult(new PasswordVerificationResult(false, false));

            if (!PasswordHashEncoding.TryDecode(storedHash, out DecodedPasswordHash decoded))
                return this.VerifyWithoutAccountAsync(password, cancellationToken);

            return this.InTurnAsync(() => this.Verify(password, decoded), cancellationToken);
        }

        // ###########################################################################################
        // A verification that always fails, at the cost of a real one - for a login with no hash to
        // check: an unknown address, a locked account, a stored hash that cannot be read (security
        // review, 2026-10-11).
        //
        // *** A FAILED LOGIN MUST TAKE AS LONG AS ANY OTHER. *** The ANSWER was already the same
        // whatever the login found (AccountFlows' property 2), but only an account with a readable
        // hash ran Argon2 - hundreds of milliseconds at the configured cost - so the response TIME
        // told an anonymous caller which addresses have an account, one request per address. This
        // runs a computation with the CONFIGURED parameters and throws the result away; a wrong
        // password for a hash made with other parameters is brought to the same cost (Verify).
        //
        // An empty password costs nothing here, exactly as in VerifyAsync - and LoginAsync refuses
        // one before it looks anybody up, so that answer says nothing about an account.
        // ###########################################################################################
        public Task<PasswordVerificationResult> VerifyWithoutAccountAsync(string? password, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(password))
                return Task.FromResult(new PasswordVerificationResult(false, false));

            return this.InTurnAsync(
                () =>
                {
                    _ = this.ComputeHash(
                        password,
                        Argon2PasswordHasher.NoAccountSalt,
                        this.thisOptions.Argon2MemoryKib,
                        this.thisOptions.Argon2Iterations,
                        this.thisOptions.Argon2Parallelism);

                    return new PasswordVerificationResult(false, false);
                },
                cancellationToken);
        }

        // ###########################################################################################
        // *** A WRONG PASSWORD COSTS WHAT THE CONFIGURED PARAMETERS COST, whatever its hash was made
        // with (code review, 2026-10-11). *** A hash keeps the parameters it was made with until
        // its account next signs in, while an unknown address costs the configured ones - so after
        // the cost was RAISED, a wrong password for an account not yet rehashed answered sooner
        // than an unknown address, and the time told which addresses have accounts again.
        //
        // So a wrong password is followed by one more computation, thrown away, making up the
        // difference: Argon2's time is close to proportional to memory x passes over its lanes
        // (Cost), so this is the memory that costs the difference at the configured passes and
        // lanes. When only the memory was raised - the usual way to raise it - the two together
        // also take the configured memory, as a configured computation does.
        //
        // Close, not exact, and only UPWARD. 0 when there is nothing to make up, or less than
        // Argon2's smallest memory (8 KiB per lane). A hash made at a HIGHER cost than configured -
        // the cost was lowered - still answers later than an unknown address until its account
        // signs in and is rehashed; the no-account path cannot know which cost to match.
        // ###########################################################################################
        internal static int PaddingMemoryKib(DecodedPasswordHash stored, int memoryKib, int iterations, int parallelism)
        {
            double missing =
                Argon2PasswordHasher.Cost(memoryKib, iterations, parallelism) -
                Argon2PasswordHasher.Cost(stored.MemoryKib, stored.Iterations, stored.Parallelism);

            double padding = Math.Round(missing * parallelism / iterations);

            return padding >= 8.0 * parallelism ? (int)Math.Min(padding, int.MaxValue) : 0;
        }

        // How long a computation takes, near enough: memory x passes, spread over its lanes.
        internal static double Cost(int memoryKib, int iterations, int parallelism) =>
            (double)memoryKib * iterations / Math.Max(1, parallelism);

        private PasswordVerificationResult Verify(string password, DecodedPasswordHash decoded)
        {
            byte[] computed = this.ComputeHash(
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
            {
                int padding = Argon2PasswordHasher.PaddingMemoryKib(
                    decoded,
                    this.thisOptions.Argon2MemoryKib,
                    this.thisOptions.Argon2Iterations,
                    this.thisOptions.Argon2Parallelism);

                if (padding > 0)
                {
                    _ = this.ComputeHash(
                        password,
                        Argon2PasswordHasher.NoAccountSalt,
                        padding,
                        this.thisOptions.Argon2Iterations,
                        this.thisOptions.Argon2Parallelism);
                }

                return new PasswordVerificationResult(false, false);
            }

            bool needsRehash = PasswordHashEncoding.NeedsRehash(
                decoded,
                this.thisOptions.Argon2MemoryKib,
                this.thisOptions.Argon2Iterations,
                this.thisOptions.Argon2Parallelism);

            return new PasswordVerificationResult(true, needsRehash);
        }

        // ###########################################################################################
        // Waits for a turn, without holding a thread, then computes on this one. A verification and
        // its padding share one turn, so both kinds of failed login wait the same way.
        // ###########################################################################################
        private async Task<T> InTurnAsync<T>(Func<T> computation, CancellationToken cancellationToken)
        {
            await this.thisTurns.WaitAsync(cancellationToken).ConfigureAwait(false);

            try
            {
                return computation();
            }
            finally
            {
                this.thisTurns.Release();
            }
        }

        private byte[] ComputeHash(
            string password,
            byte[] salt,
            int memoryKib,
            int iterations,
            int parallelism)
        {
            lock (this.thisCountsLock)
            {
                this.thisComputations++;
                this.thisWork += Argon2PasswordHasher.Cost(memoryKib, iterations, parallelism);
                this.thisRunning++;
                this.thisMostRunning = Math.Max(this.thisMostRunning, this.thisRunning);
            }

            try
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
            finally
            {
                lock (this.thisCountsLock)
                    this.thisRunning--;
            }
        }
    }

    // ###########################################################################################
    // NeedsRehash is only meaningful when IsValid is true - there is nothing to upgrade when the
    // password was wrong.
    // ###########################################################################################
    public readonly record struct PasswordVerificationResult(bool IsValid, bool NeedsRehash);
}
