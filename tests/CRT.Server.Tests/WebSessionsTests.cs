using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Signed in on the website (WebSessions): a cookie of each instance's own, the token every signed-in form carries, and
    // where a sign-in may send a person back to - a collection's page of this site, and nowhere else.
    // ###########################################################################################
    public sealed class WebSessionsTests
    {
        private const string Live = "https://classic-repair-toolbox.dk";

        private const string Test = "https://classic-repair-toolbox.dk/crt-test";

        // The test instance's pages live under the live site's path too - its cookie must not be the live one's.
        [Fact]
        public void Each_instance_has_a_cookie_of_its_own()
        {
            Assert.Equal(("crt_session", "/"), (WebSessions.CookieName(Live), WebSessions.CookiePath(Live)));
            Assert.Equal(("crt_session_crt_test", "/crt-test/"), (WebSessions.CookieName(Test), WebSessions.CookiePath(Test)));
        }

        [Fact]
        public void A_form_s_token_is_its_session_s_and_no_other()
        {
            string token = WebSessions.Csrf("session-one");

            Assert.True(WebSessions.CsrfMatches("session-one", token));
            Assert.False(WebSessions.CsrfMatches("session-two", token));
            Assert.False(WebSessions.CsrfMatches("session-one", null));
            Assert.False(WebSessions.CsrfMatches("session-one", ""));
            Assert.False(WebSessions.CsrfMatches("session-one", token.ToUpperInvariant()));
        }

        [Theory]
        [InlineData(Live, "/collections/k7mq3x9vbn2p", "/collections/k7mq3x9vbn2p")]
        [InlineData(Live, "/collections/k7mq3x9vbn2p#comments-U7-12", "/collections/k7mq3x9vbn2p#comments-U7-12")]
        [InlineData(Test, "/crt-test/collections/k7mq3x9vbn2p", "/crt-test/collections/k7mq3x9vbn2p")]
        public void A_sign_in_goes_back_to_a_collection_s_page_of_this_site(string site, string asked, string expected) =>
            Assert.Equal(expected, WebSessions.ReturnPath(asked, site));

        // An open redirect is a phishing link wearing this site's address.
        [Theory]
        [InlineData(Live, "https://evil.example/collections/x")]
        [InlineData(Live, "//evil.example/collections/x")]
        [InlineData(Live, "/account/register")]
        [InlineData(Live, "/collections/../account/sign-out")]
        [InlineData(Live, "/collections/")]
        [InlineData(Live, "/collections/x?y=z")]
        [InlineData(Test, "/collections/k7mq3x9vbn2p")]
        [InlineData(Live, null)]
        public void Anywhere_else_is_no_place_to_go(string site, string? asked) =>
            Assert.Null(WebSessions.ReturnPath(asked, site));
    }
}
