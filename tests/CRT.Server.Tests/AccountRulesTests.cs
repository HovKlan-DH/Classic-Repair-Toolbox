using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers email normalisation, address plausibility, and the password/display-name policy.
    //
    // The normalisation tests are the consequential ones: that value IS the unique index in the
    // accounts table, so a rule change here silently changes who can register.
    // ###########################################################################################
    public class AccountRulesTests
    {
        // -----------------------------------------------------------------------------------
        // Email normalisation.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("dennis@example.com", "dennis@example.com")]
        [InlineData("Dennis@Example.com", "dennis@example.com")]
        [InlineData("DENNIS@EXAMPLE.COM", "dennis@example.com")]
        [InlineData("  dennis@example.com  ", "dennis@example.com")]
        public void Addresses_differing_only_by_case_or_spacing_normalise_to_one_value(
            string input, string expected)
        {
            // This is what stops one person registering twice by capitalising a letter, and what
            // makes "an address gets exactly ONE verification mail" true - the property the whole
            // anti-enumeration design rests on.
            Assert.Equal(expected, AccountRules.NormaliseEmail(input));
        }

        [Fact]
        public void Normalisation_uses_invariant_culture()
        {
            // Turkish lowercases 'I' to a dotless 'i', so a culture-sensitive lowercase would
            // normalise the same address differently depending on the server's locale - and this
            // value is a database key. Guarding it explicitly because the failure would only
            // appear on a box with a Turkish locale, which nobody would think to test.
            System.Globalization.CultureInfo original = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

                Assert.Equal("ilkay@example.com", AccountRules.NormaliseEmail("ILKAY@example.com"));
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }

        // -----------------------------------------------------------------------------------
        // Address plausibility. Deliberately narrow - see AccountRules' header.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("dennis@example.com")]
        [InlineData("first.last@example.co.uk")]
        [InlineData("user+tag@example.com")]
        [InlineData("a@b.dk")]
        public void An_ordinary_address_is_accepted(string email)
        {
            Assert.True(AccountRules.IsPlausibleEmail(email));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("no-at-sign")]
        [InlineData("two@at@signs.com")]
        [InlineData("@example.com")]          // nothing before the @
        [InlineData("dennis@")]               // nothing after
        [InlineData("dennis@localhost")]      // no dot: mail would never leave the box
        [InlineData("dennis@example.")]       // trailing dot
        [InlineData("dennis@.example.com")]   // leading dot in the domain
        [InlineData("dennis@ex..ample.com")]  // consecutive dots
        [InlineData("Dennis <d@example.com>")] // a pasted display-name form
        public void A_malformed_address_is_rejected(string? email)
        {
            Assert.False(AccountRules.IsPlausibleEmail(email));
        }

        [Fact]
        public void An_over_length_address_is_rejected()
        {
            string local = new('a', 320);

            Assert.False(AccountRules.IsPlausibleEmail($"{local}@example.com"));
        }

        // -----------------------------------------------------------------------------------
        // Password policy.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_ordinary_passphrase_is_accepted()
        {
            Assert.Empty(AccountRules.ValidatePassword("correct horse battery staple"));
        }

        [Fact]
        public void A_short_password_is_rejected_and_the_message_suggests_a_passphrase()
        {
            IReadOnlyList<string> failures = AccountRules.ValidatePassword("short");

            string failure = Assert.Single(failures);
            Assert.Contains("at least 12 characters", failure);

            // The message has to help rather than scold - this audience is hobbyists at a bench,
            // not employees under a corporate policy.
            Assert.Contains("words together", failure);
        }

        [Fact]
        public void A_password_of_exactly_the_minimum_length_is_accepted()
        {
            // Boundary: the rule is "at least", so twelve must pass.
            Assert.Empty(AccountRules.ValidatePassword(new string('x', 12)));
        }

        [Fact]
        public void A_whitespace_only_password_is_rejected_even_when_long_enough()
        {
            // Passes a naive length check and is almost certainly a paste accident.
            IReadOnlyList<string> failures = AccountRules.ValidatePassword(new string(' ', 20));

            Assert.Contains(failures, failure => failure.Contains("only spaces"));
        }

        [Fact]
        public void An_absurdly_long_password_is_rejected()
        {
            // Unbounded input reaches the hasher, which allocates per call.
            IReadOnlyList<string> failures = AccountRules.ValidatePassword(new string('x', 500));

            Assert.Contains(failures, failure => failure.Contains("at most 256"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void A_missing_password_is_rejected(string? password)
        {
            Assert.Contains(AccountRules.ValidatePassword(password), failure => failure.Contains("required"));
        }

        [Fact]
        public void No_composition_rule_is_enforced()
        {
            // Deliberately pinned. Required character classes are withdrawn NIST guidance and
            // reliably produce weaker passwords - "Password1!" and a sticky note. If someone later
            // adds a "must contain a digit" rule, this test is the conversation.
            Assert.Empty(AccountRules.ValidatePassword("all lowercase letters here"));
        }

        // -----------------------------------------------------------------------------------
        // Display name.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_ordinary_display_name_is_accepted()
        {
            Assert.Empty(AccountRules.ValidateDisplayName("Dennis"));
        }

        [Fact]
        public void A_non_latin_display_name_is_accepted()
        {
            // Contributors are worldwide, and the column is utf8mb4 precisely so this works.
            Assert.Empty(AccountRules.ValidateDisplayName("Ð¯ÐµÐ¼Ð¾Ð½Ñ‚"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void A_missing_display_name_is_rejected(string? displayName)
        {
            Assert.Contains(
                AccountRules.ValidateDisplayName(displayName),
                failure => failure.Contains("required"));
        }

        [Fact]
        public void A_display_name_with_control_characters_is_rejected()
        {
            // A newline in a name would let it break a log line or a mail header.
            Assert.Contains(
                AccountRules.ValidateDisplayName("Dennis\nBCC: someone@example.com"),
                failure => failure.Contains("control characters"));
        }
    }
}
