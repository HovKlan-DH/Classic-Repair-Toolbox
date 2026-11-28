using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // AccountPages - the account web pages: "I forgot my password", the administrator's sign-in and a verification link mailed
    // before server 17.0.0 (nobody makes an account here since - owner decision, 2026-11-28). The security of a page served
    // to anybody is the point: every word a person typed is encoded, nothing runs or loads from anywhere, and the one style
    // block is the one the Content-Security-Policy's hash allows.
    // ###########################################################################################
    public class AccountPagesTests
    {
        private const string Base = "https://classic-repair-toolbox.dk";

        // Every page this class makes, for the checks every page must pass.
        private static IEnumerable<string> EveryPage()
        {
            yield return AccountPages.ForgotPassword(AccountPagesTests.Base);
            yield return AccountPages.ForgotPassword(AccountPagesTests.Base, "dennis@example.com", codeAsked: true, resetErrors: ["That code is not valid."]);
            yield return AccountPages.PasswordSet(AccountPagesTests.Base);
            yield return AccountPages.FromAnotherSite(AccountPagesTests.Base);
            yield return AccountPages.SignIn(AccountPagesTests.Base, "dennis@example.com", "/collections/k7mq3x9vbn2p", [AccountPages.SignInRefused]);
            yield return AccountPages.SignedIn(AccountPagesTests.Base, "Dennis");
            yield return AccountPages.SignedOut(AccountPagesTests.Base);

            foreach (TokenRedemptionOutcome outcome in Enum.GetValues<TokenRedemptionOutcome>())
                yield return AccountPages.Verified(AccountPagesTests.Base, outcome).Html;
        }

        [Theory]
        [InlineData("https://classic-repair-toolbox.dk/api", "https://classic-repair-toolbox.dk")]
        [InlineData("https://classic-repair-toolbox.dk/api/", "https://classic-repair-toolbox.dk")]
        [InlineData("https://classic-repair-toolbox.dk/crt-test/api", "https://classic-repair-toolbox.dk/crt-test")]
        [InlineData("https://classic-repair-toolbox.dk", "https://classic-repair-toolbox.dk")]
        public void The_pages_stand_where_the_API_does_less_its_api(string apiBase, string expected)
        {
            Assert.Equal(expected, AccountPages.BaseUrl(apiBase));
            Assert.Equal(expected + "/account/sign-in", AccountPages.Link(AccountPages.BaseUrl(apiBase), "sign-in"));
        }

        // ###########################################################################################
        // *** WHAT A PERSON TYPED IS SHOWN, NEVER RUN. *** An address, a name and the server's sentences come back into the
        // page they were typed on - each HTML-encoded, so a "<script>" is those characters and a quote closes no attribute.
        // ###########################################################################################
        [Fact]
        public void What_was_typed_comes_back_encoded()
        {
            string page = AccountPages.SignIn(AccountPagesTests.Base, "x\"onfocus=\"alert(1)@example.com", "/collections/", ["<img src=x>"]);

            Assert.DoesNotContain("\"onfocus=\"", page, StringComparison.Ordinal);
            Assert.Contains("&quot;onfocus=&quot;", page, StringComparison.Ordinal);
            Assert.DoesNotContain("<img", page, StringComparison.Ordinal);

            string signedIn = AccountPages.SignedIn(AccountPagesTests.Base, "<script>alert(1)</script>");
            Assert.DoesNotContain("<script>alert", signedIn, StringComparison.Ordinal);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", signedIn, StringComparison.Ordinal);

            string forgot = AccountPages.ForgotPassword(AccountPagesTests.Base, "<i>@example.com", codeAsked: true);
            Assert.DoesNotContain("<i>@", forgot, StringComparison.Ordinal);
        }

        // No script, no picture, no font, nothing from another site - the policy allows none, and no page asks for one.
        [Fact]
        public void No_page_runs_or_loads_anything()
        {
            foreach (string page in AccountPagesTests.EveryPage())
            {
                Assert.DoesNotContain("<script", page, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain(" src=", page, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("<link", page, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain(" style=", page, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotMatch(new Regex(@"<[^>]*\son[a-z]+\s*=", RegexOptions.IgnoreCase), page);
            }
        }

        // The one style block is exactly what the policy's hash allows - a change to the style that left the hash behind would
        // leave every page unstyled in a browser, and no test of the text would see it.
        [Fact]
        public void The_style_block_is_the_one_the_security_policy_allows()
        {
            foreach (string page in AccountPagesTests.EveryPage())
            {
                string style = Assert.Single(Regex.Matches(page, "<style>(.*?)</style>", RegexOptions.Singleline)).Groups[1].Value;
                string hash = "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(style)));

                Assert.Equal(AccountPages.StyleHash, hash);
            }

            Assert.Contains($"style-src '{AccountPages.StyleHash}'", AccountPages.ContentSecurityPolicy, StringComparison.Ordinal);
            Assert.StartsWith("default-src 'none';", AccountPages.ContentSecurityPolicy, StringComparison.Ordinal);
            Assert.Contains("frame-ancestors 'none'", AccountPages.ContentSecurityPolicy, StringComparison.Ordinal);
        }

        [Fact]
        public void Every_page_says_it_is_Classic_Repair_Toolbox_and_asks_to_be_kept_by_no_search_engine()
        {
            foreach (string page in AccountPagesTests.EveryPage())
            {
                Assert.Contains("<header><a href=\"https://classic-repair-toolbox.dk/\">Classic Repair Toolbox</a></header>", page, StringComparison.Ordinal);
                Assert.Contains("<meta name=\"robots\" content=\"noindex\">", page, StringComparison.Ordinal);
            }
        }

        // ###########################################################################################
        // *** NO PAGE OFFERS TO MAKE AN ACCOUNT *** (owner decision, 2026-11-28: "I do not want the possibility for everyone to
        // create an account") - an account is a maintainer's, made by an invitation.
        // ###########################################################################################
        [Fact]
        public void No_page_offers_to_make_an_account()
        {
            foreach (string page in AccountPagesTests.EveryPage())
            {
                Assert.DoesNotContain("/account/register", page, StringComparison.Ordinal);
                Assert.DoesNotContain("Create an account", page, StringComparison.OrdinalIgnoreCase);
            }
        }

        // The forms post to this service's own pages, under the base - the test instance's too.
        [Fact]
        public void The_forms_post_to_the_pages_of_their_own_instance()
        {
            const string test = "https://classic-repair-toolbox.dk/crt-test";

            Assert.Contains("action=\"https://classic-repair-toolbox.dk/crt-test/account/forgot-password\"", AccountPages.ForgotPassword(test), StringComparison.Ordinal);
            Assert.Contains("action=\"https://classic-repair-toolbox.dk/crt-test/account/reset-password\"", AccountPages.ForgotPassword(test), StringComparison.Ordinal);
            Assert.Contains("action=\"https://classic-repair-toolbox.dk/crt-test/account/sign-in\"", AccountPages.SignIn(test), StringComparison.Ordinal);
        }

        // The neutral answer stays neutral: a code asked for says nothing of whether the address has an account (AccountFlows'
        // header).
        [Fact]
        public void Asking_for_a_code_says_nothing_of_whether_the_address_has_an_account() =>
            Assert.Contains("If dennis@example.com has an account, a code is on its way to it.", AccountPages.ForgotPassword(AccountPagesTests.Base, "dennis@example.com", codeAsked: true), StringComparison.Ordinal);

        // ###########################################################################################
        // *** A FORM ANOTHER WEBSITE POSTS IS REFUSED *** (code review, 2026-11-25, Security). Asking for a code mails somebody,
        // the budget counted per IP address - a hidden form on a busy page would have spent every visitor's. Sec-Fetch-Site says
        // it in every current browser; an older one's Origin is read instead; a request with neither is no browser, which could
        // post directly anyway.
        // ###########################################################################################
        [Theory]
        [InlineData("same-origin", null, false)]
        [InlineData("same-site", null, false)]
        [InlineData("none", null, false)]
        [InlineData("SAME-ORIGIN", null, false)]
        [InlineData("cross-site", null, true)]
        [InlineData("cross-site", "https://classic-repair-toolbox.dk", true)]
        [InlineData("something-new", null, true)]
        [InlineData(null, "https://classic-repair-toolbox.dk", false)]
        [InlineData(null, "https://www.classic-repair-toolbox.dk", false)]
        [InlineData(null, "https://evil.example", true)]
        [InlineData(null, "https://classic-repair-toolbox.dk.evil.example", true)]
        [InlineData(null, "https://evilclassic-repair-toolbox.dk", true)]
        [InlineData(null, "http://classic-repair-toolbox.dk", true)]
        [InlineData(null, "null", true)]
        [InlineData(null, null, false)]
        [InlineData("", "", false)]
        public void A_form_from_another_website_is_told_apart_from_the_sites_own(string? fetchSite, string? origin, bool fromAnotherSite)
        {
            Assert.Equal(fromAnotherSite, AccountPages.IsFromAnotherSite(fetchSite, origin, "https://classic-repair-toolbox.dk/crt-test"));
        }

        // The page saying so names the pages here to send the form from - the instance's own.
        [Fact]
        public void The_refusal_names_the_pages_of_its_own_instance()
        {
            string page = AccountPages.FromAnotherSite("https://classic-repair-toolbox.dk/crt-test");

            Assert.Contains("href=\"https://classic-repair-toolbox.dk/crt-test/account/sign-in\"", page, StringComparison.Ordinal);
            Assert.Contains("href=\"https://classic-repair-toolbox.dk/crt-test/account/forgot-password\"", page, StringComparison.Ordinal);
        }

        // "no-referrer" makes a browser send "Origin: null" with the site's OWN forms, which the check above refuses - and
        // "strict-origin" still lets no reset code or verify link's token leave in a Referer, only the site's address.
        [Fact]
        public void The_pages_send_only_the_sites_address_as_referrer_so_their_own_forms_keep_their_origin()
        {
            Assert.Equal("strict-origin", AccountPages.ReferrerPolicy);
        }

        // How long a code lasts is the server's own lifetime, said in CRT.Data's words - never a number typed beside it.
        [Fact]
        public void A_code_asked_for_says_the_lifetime_the_server_gives_it()
        {
            Assert.Equal(AccountFlows.PasswordResetTokenLifetime, TimeSpan.FromHours(AccountWording.PasswordResetCodeHours));
            Assert.Contains(
                System.Net.WebUtility.HtmlEncode(AccountWording.ResetCodeWorks),
                AccountPages.ForgotPassword(AccountPagesTests.Base, "dennis@example.com", codeAsked: true),
                StringComparison.Ordinal);
            Assert.Contains($"{AccountWording.PasswordResetCodeHours} hours", AccountWording.ResetCodeWorks, StringComparison.Ordinal);
        }

        // A link mailed before server 17.0.0: a page per outcome, the failures told apart as the token is the only subject - an
        // expired link's way on is a password set with a mailed code, which confirms the address too.
        [Fact]
        public void A_verification_link_is_answered_with_what_to_do_next()
        {
            Assert.Equal(200, AccountPages.Verified(AccountPagesTests.Base, TokenRedemptionOutcome.Success).Status);
            Assert.Contains("Your address is confirmed", AccountPages.Verified(AccountPagesTests.Base, TokenRedemptionOutcome.Success).Html, StringComparison.Ordinal);
            Assert.Equal(200, AccountPages.Verified(AccountPagesTests.Base, TokenRedemptionOutcome.AlreadyUsed).Status);

            (string expired, int expiredStatus) = AccountPages.Verified(AccountPagesTests.Base, TokenRedemptionOutcome.Expired);
            Assert.Equal(400, expiredStatus);
            Assert.Contains("href=\"https://classic-repair-toolbox.dk/account/forgot-password\"", expired, StringComparison.Ordinal);

            Assert.Equal(400, AccountPages.Verified(AccountPagesTests.Base, TokenRedemptionOutcome.Invalid).Status);
        }

        // ###########################################################################################
        // Signing in on the website - the administrator's alone since server 17.0.0, for "All collections" - and back to the
        // page after.
        // ###########################################################################################
        [Fact]
        public void Sign_in_posts_here_carries_the_page_to_go_back_to_and_says_whose_it_is()
        {
            string page = AccountPages.SignIn(AccountPagesTests.Base, returnTo: "/collections/k7mq3x9vbn2p");

            Assert.Contains("<form method=\"post\" action=\"https://classic-repair-toolbox.dk/account/sign-in\">", page, StringComparison.Ordinal);
            Assert.Contains("<input type=\"hidden\" name=\"return\" value=\"/collections/k7mq3x9vbn2p\">", page, StringComparison.Ordinal);
            Assert.Contains("autocomplete=\"current-password\"", page, StringComparison.Ordinal);
            Assert.Contains("For the site's administrator", page, StringComparison.Ordinal);
            Assert.Contains("href=\"https://classic-repair-toolbox.dk/account/forgot-password\"", page, StringComparison.Ordinal);
        }

        // A refused sign-in says no reason - whether the address has an account is nobody's business (AccountFlows' header);
        // right ones of an account that is no administrator's are told the web's sign-in is the administrator's, and where a
        // maintainer signs in.
        [Fact]
        public void A_refused_sign_in_says_no_reason_and_a_maintainer_is_sent_to_CRT()
        {
            Assert.Equal("The address or the password is not right.", AccountPages.SignInRefused);
            Assert.Contains("administrator alone", AccountPages.SignInAdministratorsOnly, StringComparison.Ordinal);
            Assert.Contains("\"Maintainer\" tab", AccountPages.SignInAdministratorsOnly, StringComparison.Ordinal);
        }
    }
}
