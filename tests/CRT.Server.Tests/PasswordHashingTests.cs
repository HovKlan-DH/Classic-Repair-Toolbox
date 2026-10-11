using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers password hashing, the stored format, and the rehash-on-login rule.
    //
    // THESE TESTS USE DELIBERATELY WEAK ARGON2 PARAMETERS (8 MiB, 1 iteration). The real defaults
    // are 64 MiB and 3 iterations, which would make a suite that hashes a dozen times take several
    // seconds and allocate over a gigabyte in total. What is under test here is the FORMAT and the
    // DECISIONS around it, not Argon2 itself - the algorithm is the library's problem, and its
    // strength is a configuration value the validator enforces a floor on.
    //
    // The one exception is DefaultOptions, used where the test is specifically about the
    // configured cost being honoured.
    // ###########################################################################################
    public class PasswordHashingTests
    {
        // Fast enough for a test suite, structurally identical to production.
        private static ServerOptions FastOptions()
        {
            return new ServerOptions
            {
                Argon2MemoryKib = 8192,
                Argon2Iterations = 1,
                Argon2Parallelism = 1
            };
        }

        // -----------------------------------------------------------------------------------
        // Hashing and verification.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_password_verifies_against_its_own_hash()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = await hasher.HashAsync("correct horse battery staple");

            Assert.True((await hasher.VerifyAsync("correct horse battery staple", stored)).IsValid);
        }

        [Fact]
        public async Task A_wrong_password_does_not_verify()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = await hasher.HashAsync("correct horse battery staple");

            Assert.False((await hasher.VerifyAsync("correct horse battery stapler", stored)).IsValid);
        }

        [Fact]
        public async Task Verification_is_case_sensitive()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = await hasher.HashAsync("MyPassphrase123");

            Assert.False((await hasher.VerifyAsync("mypassphrase123", stored)).IsValid);
        }

        [Fact]
        public async Task The_same_password_hashed_twice_gives_different_stored_values()
        {
            // The salt is per-password and random. Without this, two people choosing the same
            // password produce identical stored hashes, so one cracked hash reveals every account
            // sharing it - and the stored values themselves leak who shares a password.
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string first = await hasher.HashAsync("the same password");
            string second = await hasher.HashAsync("the same password");

            Assert.NotEqual(first, second);

            // Both must still verify - different salts, same password.
            Assert.True((await hasher.VerifyAsync("the same password", first)).IsValid);
            Assert.True((await hasher.VerifyAsync("the same password", second)).IsValid);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not a hash at all")]
        [InlineData("$argon2id$v=19$m=8192,t=1")]                    // truncated
        [InlineData("$argon2i$v=19$m=8192,t=1,p=1$c2FsdA$aGFzaA")]   // wrong variant
        [InlineData("$argon2id$v=16$m=8192,t=1,p=1$c2FsdA$aGFzaA")]  // old Argon2 version
        [InlineData("$argon2id$v=19$m=0,t=1,p=1$c2FsdA$aGFzaA")]     // zero cost
        public async Task A_malformed_stored_hash_fails_the_login_rather_than_throwing(string? stored)
        {
            // The caller is an authentication endpoint. An unparseable stored hash must mean "this
            // login fails", not a 500 - which would tell an attacker that this particular account
            // is interesting.
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            PasswordVerificationResult result = await hasher.VerifyAsync("any password", stored);

            Assert.False(result.IsValid);
            Assert.False(result.NeedsRehash);

            // ...and costs what a real verification costs (security review, 2026-10-11): failing
            // at once would single the account out by how quickly the login answered.
            Assert.Equal(1, hasher.Computations);
        }

        // ###########################################################################################
        // *** A VERIFICATION WITH NO ACCOUNT FAILS - AT THE COST OF A REAL ONE (security review,
        // 2026-10-11). *** LoginAsync calls it for an unknown address and a locked account, so the
        // time a failed login takes says nothing about which one it was.
        // ###########################################################################################
        [Fact]
        public async Task A_verification_with_no_account_fails_after_one_computation()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            PasswordVerificationResult result = await hasher.VerifyWithoutAccountAsync("correct horse battery staple");

            Assert.False(result.IsValid);
            Assert.False(result.NeedsRehash);
            Assert.Equal(1, hasher.Computations);
        }

        // Exactly as Verify: an empty password is refused before any work, and LoginAsync refuses
        // one before it looks anybody up, so that answer says nothing about an account.
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public async Task A_verification_with_no_account_and_no_password_costs_nothing(string? password)
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            Assert.False((await hasher.VerifyWithoutAccountAsync(password)).IsValid);
            Assert.Equal(0, hasher.Computations);
        }

        // -----------------------------------------------------------------------------------
        // The stored format. This is what makes a cost increase safe.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_stored_hash_carries_the_parameters_it_was_made_with()
        {
            var options = PasswordHashingTests.FastOptions();
            var hasher = new Argon2PasswordHasher(options);

            string stored = await hasher.HashAsync("a password");

            Assert.True(PasswordHashEncoding.TryDecode(stored, out DecodedPasswordHash decoded));
            Assert.Equal(options.Argon2MemoryKib, decoded.MemoryKib);
            Assert.Equal(options.Argon2Iterations, decoded.Iterations);
            Assert.Equal(options.Argon2Parallelism, decoded.Parallelism);
        }

        [Fact]
        public async Task The_stored_hash_is_a_standard_PHC_string()
        {
            // Interoperability is deliberate: a hash written here should be readable by any other
            // Argon2 implementation, so a future move off this library does not strand every
            // stored password.
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = await hasher.HashAsync("a password");

            Assert.StartsWith("$argon2id$v=19$m=", stored, StringComparison.Ordinal);

            string[] parts = stored.Split('$');
            Assert.Equal(6, parts.Length);

            // PHC omits base64 padding - on the SALT and HASH segments specifically. The '=' in
            // "v=19" and "m=65536,t=3,p=2" is part of the format and must stay.
            Assert.DoesNotContain('=', parts[4]);
            Assert.DoesNotContain('=', parts[5]);
        }

        [Fact]
        public async Task Raising_the_cost_does_not_invalidate_existing_passwords()
        {
            // THE property the whole format exists for. A hash made with the old parameters must
            // still verify after the configuration is raised - verification reads the parameters
            // from the stored string, never from configuration.
            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            string storedUnderWeak = await new Argon2PasswordHasher(weak).HashAsync("a password");

            var strong = new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 };
            PasswordVerificationResult result = await new Argon2PasswordHasher(strong).VerifyAsync("a password", storedUnderWeak);

            Assert.True(result.IsValid);

            // ...and the login path is told to upgrade it, at the one moment the plaintext is
            // legitimately available.
            Assert.True(result.NeedsRehash);
        }

        [Fact]
        public async Task A_hash_at_the_current_cost_is_not_rehashed()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = await hasher.HashAsync("a password");

            Assert.False((await hasher.VerifyAsync("a password", stored)).NeedsRehash);
        }

        [Fact]
        public async Task A_hash_STRONGER_than_the_current_cost_is_not_downgraded()
        {
            // Rehashing on any difference rather than on being BELOW policy would silently weaken
            // a password whenever the configuration was lowered - the opposite of the intent.
            var strong = new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 };
            string storedUnderStrong = await new Argon2PasswordHasher(strong).HashAsync("a password");

            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            PasswordVerificationResult result = await new Argon2PasswordHasher(weak).VerifyAsync("a password", storedUnderStrong);

            Assert.True(result.IsValid);
            Assert.False(result.NeedsRehash);
        }

        [Fact]
        public void Any_single_parameter_being_below_policy_triggers_a_rehash()
        {
            // Each of the three is checked independently - an OR, not an AND. A hash with plenty
            // of memory but one iteration is still below policy.
            var decoded = new DecodedPasswordHash(65536, 1, 2, [1], [2]);

            Assert.True(PasswordHashEncoding.NeedsRehash(decoded, 65536, 3, 2));
        }

        // -----------------------------------------------------------------------------------
        // A wrong password costs the configured work, whatever its hash was made with.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** A WRONG PASSWORD ON A CHEAPER HASH IS PADDED TO THE CONFIGURED COST (code review,
        // 2026-10-11). *** After the cost was raised, an account not yet rehashed answered a wrong
        // password sooner than an unknown address (VerifyWithoutAccountAsync), and the time said
        // which addresses have accounts. Counted as work (memory x passes over lanes), not timed.
        // ###########################################################################################
        [Fact]
        public async Task A_wrong_password_on_a_hash_made_at_a_lower_cost_does_the_configured_work()
        {
            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            string storedUnderWeak = await new Argon2PasswordHasher(weak).HashAsync("a password");

            var strong = new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 };
            var verifying = new Argon2PasswordHasher(strong);
            var noAccount = new Argon2PasswordHasher(strong);

            Assert.False((await verifying.VerifyAsync("not the password", storedUnderWeak)).IsValid);
            await noAccount.VerifyWithoutAccountAsync("not the password");

            Assert.Equal(Argon2PasswordHasher.Cost(16384, 2, 1), noAccount.Work);
            Assert.Equal(noAccount.Work, verifying.Work);
            Assert.Equal(2, verifying.Computations);
        }

        // The right password is not padded: it is about to be rehashed at the configured cost anyway.
        [Fact]
        public async Task The_right_password_on_a_hash_made_at_a_lower_cost_is_not_padded()
        {
            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            string storedUnderWeak = await new Argon2PasswordHasher(weak).HashAsync("a password");

            var verifying = new Argon2PasswordHasher(new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 });

            Assert.True((await verifying.VerifyAsync("a password", storedUnderWeak)).IsValid);
            Assert.Equal(1, verifying.Computations);
        }

        // A hash at the configured cost, or above it, already costs at least as much: nothing to add.
        [Fact]
        public async Task A_wrong_password_on_a_hash_at_or_above_the_configured_cost_is_not_padded()
        {
            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            var strong = new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 };

            var atCost = new Argon2PasswordHasher(weak);
            Assert.False((await atCost.VerifyAsync("not the password", await new Argon2PasswordHasher(weak).HashAsync("a password"))).IsValid);
            Assert.Equal(1, atCost.Computations);

            var belowStored = new Argon2PasswordHasher(weak);
            Assert.False((await belowStored.VerifyAsync("not the password", await new Argon2PasswordHasher(strong).HashAsync("a password"))).IsValid);
            Assert.Equal(1, belowStored.Computations);
        }

        // ###########################################################################################
        // The padding is the memory that costs the difference at the configured passes and lanes -
        // and when only the memory was raised, the two computations together take the configured
        // memory, as a configured computation does.
        // ###########################################################################################
        [Theory]
        [InlineData(32768, 3, 2, 65536, 3, 2, 32768)]   // memory doubled: the other half
        [InlineData(65536, 1, 2, 65536, 3, 2, 43691)]   // passes tripled: two thirds more at three passes
        [InlineData(65536, 3, 1, 65536, 3, 2, 0)]       // one lane at the same memory already costs more
        [InlineData(65536, 3, 2, 65536, 3, 2, 0)]       // the configured cost itself
        [InlineData(65532, 3, 2, 65536, 3, 2, 0)]       // less than Argon2's smallest memory short
        public void The_padding_makes_up_the_cost_at_the_configured_passes_and_lanes(
            int storedMemoryKib, int storedIterations, int storedParallelism,
            int memoryKib, int iterations, int parallelism,
            int expected)
        {
            var stored = new DecodedPasswordHash(storedMemoryKib, storedIterations, storedParallelism, [1], [2]);

            Assert.Equal(expected, Argon2PasswordHasher.PaddingMemoryKib(stored, memoryKib, iterations, parallelism));
        }

        // -----------------------------------------------------------------------------------
        // No more computations at once than the machine has cores for.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** COMPUTATIONS WAIT FOR A TURN (code review, 2026-10-11). *** Every unknown address
        // costs a computation of ~128 MiB, and the limiter counts per address and per IP - so
        // logins from many addresses could run any number at once. More than the cores can run
        // would finish no sooner, only take more memory.
        // ###########################################################################################
        [Fact]
        public async Task No_more_computations_run_at_once_than_the_hasher_has_turns_for()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions(), concurrentComputations: 1);

            await Task.WhenAll(Enumerable.Range(0, 6).Select(index =>
                Task.Run(() => hasher.VerifyWithoutAccountAsync($"password {index}"))));

            Assert.Equal(6, hasher.Computations);
            Assert.Equal(1, hasher.MostRunning);
        }

        [Theory]
        [InlineData(8, 2, 4)]
        [InlineData(4, 2, 2)]
        [InlineData(2, 2, 1)]
        [InlineData(1, 2, 1)]   // fewer cores than lanes: one at a time, never none
        [InlineData(6, 1, 6)]
        public void A_turn_for_every_computations_lanes_and_at_least_one(int processorCount, int parallelism, int expected)
        {
            Assert.Equal(expected, Argon2PasswordHasher.ConcurrentComputations(processorCount, parallelism));
        }

        // -----------------------------------------------------------------------------------
        // Encoding round-trips.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_encoded_hash_round_trips_through_decode()
        {
            byte[] salt = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
            byte[] hash = new byte[32];
            Array.Fill(hash, (byte)0xAB);

            string encoded = PasswordHashEncoding.Encode(65536, 3, 2, salt, hash);

            Assert.True(PasswordHashEncoding.TryDecode(encoded, out DecodedPasswordHash decoded));
            Assert.Equal(65536, decoded.MemoryKib);
            Assert.Equal(3, decoded.Iterations);
            Assert.Equal(2, decoded.Parallelism);
            Assert.Equal(salt, decoded.Salt);
            Assert.Equal(hash, decoded.Hash);
        }

        [Fact]
        public void Salt_lengths_that_need_different_base64_padding_all_round_trip()
        {
            // PHC strips base64 padding, so decoding has to restore it. A length whose remainder
            // differs is where an off-by-one in that restoration shows up - and it would corrupt
            // the salt, making every affected password unverifiable.
            for (int length = 1; length <= 24; length++)
            {
                byte[] salt = new byte[length];
                Array.Fill(salt, (byte)length);

                string encoded = PasswordHashEncoding.Encode(8192, 1, 1, salt, [9, 9, 9, 9]);

                Assert.True(PasswordHashEncoding.TryDecode(encoded, out DecodedPasswordHash decoded));
                Assert.Equal(salt, decoded.Salt);
            }
        }
    }
}
