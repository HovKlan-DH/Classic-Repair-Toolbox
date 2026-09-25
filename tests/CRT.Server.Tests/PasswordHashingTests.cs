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
        public void A_password_verifies_against_its_own_hash()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = hasher.Hash("correct horse battery staple");

            Assert.True(hasher.Verify("correct horse battery staple", stored).IsValid);
        }

        [Fact]
        public void A_wrong_password_does_not_verify()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = hasher.Hash("correct horse battery staple");

            Assert.False(hasher.Verify("correct horse battery stapler", stored).IsValid);
        }

        [Fact]
        public void Verification_is_case_sensitive()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = hasher.Hash("MyPassphrase123");

            Assert.False(hasher.Verify("mypassphrase123", stored).IsValid);
        }

        [Fact]
        public void The_same_password_hashed_twice_gives_different_stored_values()
        {
            // The salt is per-password and random. Without this, two people choosing the same
            // password produce identical stored hashes, so one cracked hash reveals every account
            // sharing it - and the stored values themselves leak who shares a password.
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string first = hasher.Hash("the same password");
            string second = hasher.Hash("the same password");

            Assert.NotEqual(first, second);

            // Both must still verify - different salts, same password.
            Assert.True(hasher.Verify("the same password", first).IsValid);
            Assert.True(hasher.Verify("the same password", second).IsValid);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not a hash at all")]
        [InlineData("$argon2id$v=19$m=8192,t=1")]                    // truncated
        [InlineData("$argon2i$v=19$m=8192,t=1,p=1$c2FsdA$aGFzaA")]   // wrong variant
        [InlineData("$argon2id$v=16$m=8192,t=1,p=1$c2FsdA$aGFzaA")]  // old Argon2 version
        [InlineData("$argon2id$v=19$m=0,t=1,p=1$c2FsdA$aGFzaA")]     // zero cost
        public void A_malformed_stored_hash_fails_the_login_rather_than_throwing(string? stored)
        {
            // The caller is an authentication endpoint. An unparseable stored hash must mean "this
            // login fails", not a 500 - which would tell an attacker that this particular account
            // is interesting.
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            PasswordVerificationResult result = hasher.Verify("any password", stored);

            Assert.False(result.IsValid);
            Assert.False(result.NeedsRehash);
        }

        // -----------------------------------------------------------------------------------
        // The stored format. This is what makes a cost increase safe.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_stored_hash_carries_the_parameters_it_was_made_with()
        {
            var options = PasswordHashingTests.FastOptions();
            var hasher = new Argon2PasswordHasher(options);

            string stored = hasher.Hash("a password");

            Assert.True(PasswordHashEncoding.TryDecode(stored, out DecodedPasswordHash decoded));
            Assert.Equal(options.Argon2MemoryKib, decoded.MemoryKib);
            Assert.Equal(options.Argon2Iterations, decoded.Iterations);
            Assert.Equal(options.Argon2Parallelism, decoded.Parallelism);
        }

        [Fact]
        public void The_stored_hash_is_a_standard_PHC_string()
        {
            // Interoperability is deliberate: a hash written here should be readable by any other
            // Argon2 implementation, so a future move off this library does not strand every
            // stored password.
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = hasher.Hash("a password");

            Assert.StartsWith("$argon2id$v=19$m=", stored, StringComparison.Ordinal);

            string[] parts = stored.Split('$');
            Assert.Equal(6, parts.Length);

            // PHC omits base64 padding - on the SALT and HASH segments specifically. The '=' in
            // "v=19" and "m=65536,t=3,p=2" is part of the format and must stay.
            Assert.DoesNotContain('=', parts[4]);
            Assert.DoesNotContain('=', parts[5]);
        }

        [Fact]
        public void Raising_the_cost_does_not_invalidate_existing_passwords()
        {
            // THE property the whole format exists for. A hash made with the old parameters must
            // still verify after the configuration is raised - verification reads the parameters
            // from the stored string, never from configuration.
            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            string storedUnderWeak = new Argon2PasswordHasher(weak).Hash("a password");

            var strong = new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 };
            PasswordVerificationResult result = new Argon2PasswordHasher(strong).Verify("a password", storedUnderWeak);

            Assert.True(result.IsValid);

            // ...and the login path is told to upgrade it, at the one moment the plaintext is
            // legitimately available.
            Assert.True(result.NeedsRehash);
        }

        [Fact]
        public void A_hash_at_the_current_cost_is_not_rehashed()
        {
            var hasher = new Argon2PasswordHasher(PasswordHashingTests.FastOptions());

            string stored = hasher.Hash("a password");

            Assert.False(hasher.Verify("a password", stored).NeedsRehash);
        }

        [Fact]
        public void A_hash_STRONGER_than_the_current_cost_is_not_downgraded()
        {
            // Rehashing on any difference rather than on being BELOW policy would silently weaken
            // a password whenever the configuration was lowered - the opposite of the intent.
            var strong = new ServerOptions { Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1 };
            string storedUnderStrong = new Argon2PasswordHasher(strong).Hash("a password");

            var weak = new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 };
            PasswordVerificationResult result = new Argon2PasswordHasher(weak).Verify("a password", storedUnderStrong);

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
