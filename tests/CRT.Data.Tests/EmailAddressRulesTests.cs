using System;
using System.Globalization;
using System.Threading;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers the email rule shared by the app and the server.
    //
    // IT LIVES IN CRT.DATA BECAUSE BOTH SIDES NEED IT: the app validates the address typed into
    // the Submit dialog before enabling the button, and the server validates contributors' contact
    // addresses. Two copies would drift, and the drift would show as the app accepting something
    // the server refuses - which reads to a user as the server being broken.
    //
    // AccountRulesTests in CRT.Server.Tests still exercises the same rule through
    // AccountRules.IsPlausibleEmail, which now delegates here. That duplication is deliberate: it
    // is what proves the delegation still works, and it would catch the server-side wrapper being
    // changed to do something of its own.
    // ###########################################################################################
    public class EmailAddressRulesTests
    {
        [Theory]
        [InlineData("dennis@example.com")]
        [InlineData("first.last@example.co.uk")]
        [InlineData("user+tag@example.com")]
        [InlineData("a@b.dk")]
        public void An_ordinary_address_is_accepted(string email)
        {
            Assert.True(EmailAddressRules.IsPlausible(email));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("no-at-sign")]
        [InlineData("two@at@signs.com")]
        [InlineData("@example.com")]
        [InlineData("dennis@")]
        [InlineData("dennis@localhost")]
        [InlineData("dennis@example.")]
        [InlineData("dennis@.example.com")]
        [InlineData("dennis@ex..ample.com")]
        [InlineData("Dennis <d@example.com>")]
        public void A_malformed_address_is_rejected(string? email)
        {
            Assert.False(EmailAddressRules.IsPlausible(email));
        }

        [Fact]
        public void An_over_length_address_is_rejected()
        {
            Assert.False(EmailAddressRules.IsPlausible(new string('a', 320) + "@example.com"));
        }

        [Theory]
        [InlineData("dennis@example.com", "dennis@example.com")]
        [InlineData("Dennis@Example.com", "dennis@example.com")]
        [InlineData("  DENNIS@EXAMPLE.COM  ", "dennis@example.com")]
        public void Normalisation_collapses_case_and_spacing(string input, string expected)
        {
            Assert.Equal(expected, EmailAddressRules.Normalise(input));
        }

        [Fact]
        public void Normalisation_uses_invariant_culture()
        {
            // Turkish lowercases 'I' to a dotless 'i'. On the server this value is a database key,
            // so a culture-sensitive lowercase would make the same address normalise differently
            // depending on the machine's locale - a failure that only appears on a box nobody
            // thought to test.
            CultureInfo original = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("tr-TR");

                Assert.Equal("ilkay@example.com", EmailAddressRules.Normalise("ILKAY@example.com"));
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = original;
            }
        }

        [Fact]
        public void A_plausible_address_survives_normalisation()
        {
            // The two rules have to agree: normalising an accepted address must not produce one
            // the same rule would then reject, or a value could pass validation and fail on its
            // way into storage.
            foreach (string email in new[] { "Dennis@Example.com", "  a@b.dk  ", "USER+TAG@EXAMPLE.CO.UK" })
            {
                Assert.True(EmailAddressRules.IsPlausible(email));
                Assert.True(EmailAddressRules.IsPlausible(EmailAddressRules.Normalise(email)));
            }
        }
    }
}
