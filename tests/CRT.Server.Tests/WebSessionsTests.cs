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

        // The test instance's pages live under the live site's path too - its cookie must not be the live one's, NOR the live
        // one sent to it: the live cookie goes only where it is read, never to "/" (code review, 2026-11-27, Security).
        [Fact]
        public void Each_instance_has_a_cookie_of_its_own_sent_only_where_it_is_read()
        {
            Assert.Equal("crt_session", WebSessions.CookieName(Live));
            Assert.Equal(["/collections/", "/account/sign-out", "/account/sign-in"], WebSessions.CookiePaths(Live));
            Assert.Equal("crt_session_crt_test", WebSessions.CookieName(Test));
            Assert.Equal(["/crt-test/collections/", "/crt-test/account/sign-out", "/crt-test/account/sign-in"], WebSessions.CookiePaths(Test));

            // No path of the live cookie is a prefix of the test instance's pages.
            Assert.DoesNotContain(WebSessions.CookiePaths(Live), path => "/crt-test/collections/k7mq3x9vbn2p".StartsWith(path, StringComparison.Ordinal));
        }

        [Fact]
        public void A_sign_in_is_kept_on_every_path_and_the_earlier_whole_site_cookie_taken_away()
        {
            var context = new Microsoft.AspNetCore.Http.DefaultHttpContext();

            WebSessions.Keep(context.Response, Live, "token-1", new DateTimeOffset(2026, 12, 27, 0, 0, 0, TimeSpan.Zero));

            string[] cookies = context.Response.Headers.SetCookie.ToArray()!;

            Assert.Contains(cookies, cookie => cookie.StartsWith("crt_session=token-1;", StringComparison.Ordinal) && cookie.Contains("path=/collections/", StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.StartsWith("crt_session=token-1;", StringComparison.Ordinal) && cookie.Contains("path=/account/sign-out", StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.StartsWith("crt_session=token-1;", StringComparison.Ordinal) && cookie.Contains("path=/account/sign-in", StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.StartsWith("crt_session=;", StringComparison.Ordinal) && cookie.Contains("path=/;", StringComparison.Ordinal));
            Assert.All(cookies, cookie => Assert.Contains("httponly", cookie, StringComparison.OrdinalIgnoreCase));

            // Whatever the page turns out to be, nobody between keeps the token (code review, 2026-11-27, Security).
            Assert.Equal("private, no-store", context.Response.Headers.CacheControl.ToString());
        }

        [Fact]
        public void Signing_out_takes_the_cookie_away_on_every_path_it_ever_had()
        {
            var context = new Microsoft.AspNetCore.Http.DefaultHttpContext();

            WebSessions.Forget(context.Response, Test);

            string[] cookies = context.Response.Headers.SetCookie.ToArray()!;

            Assert.Equal(4, cookies.Length);
            Assert.All(cookies, cookie => Assert.StartsWith("crt_session_crt_test=;", cookie, StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.Contains("path=/crt-test/;", StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.Contains("path=/crt-test/collections/", StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.Contains("path=/crt-test/account/sign-out", StringComparison.Ordinal));
            Assert.Contains(cookies, cookie => cookie.Contains("path=/crt-test/account/sign-in", StringComparison.Ordinal));
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
        // A long anchor - "CN2_Keyboard" pin "PB7" - still goes back (code review, 2026-11-27).
        [InlineData(Live, "/collections/k7mq3x9vbn2p#comments-CN2-Keyboard-PB7", "/collections/k7mq3x9vbn2p#comments-CN2-Keyboard-PB7")]
        [InlineData(Test, "/crt-test/collections/k7mq3x9vbn2p", "/crt-test/collections/k7mq3x9vbn2p")]
        // "Your collections" (step 4).
        [InlineData(Live, "/collections/", "/collections/")]
        [InlineData(Test, "/crt-test/collections/", "/crt-test/collections/")]
        public void A_sign_in_goes_back_to_a_collection_s_page_of_this_site(string site, string asked, string expected) =>
            Assert.Equal(expected, WebSessions.ReturnPath(asked, site));

        // The longest anchor a page can make - a label and a pin of the longest a capture names - still goes back.
        [Fact]
        public void The_longest_anchor_a_page_makes_still_goes_back()
        {
            var image = new global::Handlers.DataHandling.CollectionImageEntry(
                new string('a', 64), "1" + new string('L', 99), new string('P', 100), null, null, null, null, null, null, null, DateTimeOffset.UtcNow);
            string anchor = global::Handlers.DataHandling.CollectionRules.AnchorsOf([image, image])[1];
            string asked = $"/collections/k7mq3x9vbn2p#comments-{anchor}";

            Assert.Equal(asked, WebSessions.ReturnPath(asked, Live));
        }

        // An open redirect is a phishing link wearing this site's address.
        [Theory]
        [InlineData(Live, "https://evil.example/collections/x")]
        [InlineData(Live, "//evil.example/collections/x")]
        [InlineData(Live, "/account/register")]
        [InlineData(Live, "/collections/../account/sign-out")]
        [InlineData(Live, "/collections")]
        [InlineData(Live, "/collections/x?y=z")]
        [InlineData(Live, "/collections//evil.example")]
        [InlineData(Test, "/collections/k7mq3x9vbn2p")]
        [InlineData(Live, "/collections/k7mq3x9vbn2p/delete")]
        [InlineData(Live, "/collections/k7mq3x9vbn2p#a/b")]
        [InlineData(Live, null)]
        public void Anywhere_else_is_no_place_to_go(string site, string? asked) =>
            Assert.Null(WebSessions.ReturnPath(asked, site));

        // A browser signed in already ends the session it held as it signs in again - the previous person's on a shared
        // computer - and the new one is kept (code review, 2026-11-27, Security: the old one stayed live for 30 days).
        [Fact]
        public async Task Signing_in_again_ends_the_session_the_browser_held()
        {
            DateTimeOffset now = new(2026, 11, 27, 12, 0, 0, TimeSpan.Zero);
            var options = new CRT.Server.Configuration.ServerOptions
            {
                Argon2MemoryKib = 8192,
                Argon2Iterations = 1,
                Argon2Parallelism = 1,
                PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
                MailFromAddress = "noreply@classic-repair-toolbox.dk",
                RefreshTokenDays = 30
            };
            var hasher = new Argon2PasswordHasher(options);
            var store = new CRT.Server.Tests.Fakes.FakeAccountStore();

            await AccountFlows.RegisterAsync(new RegistrationRequest("anna@example.com", "correct horse battery staple", "Anna", null), store, new CRT.Server.Tests.Fakes.FakeEmailSender(), hasher, options, now);
            store.ConfirmAddress("anna@example.com");

            async Task<string> SignInAsync() =>
                (await AccountFlows.LoginAsync(new LoginRequest("anna@example.com", "correct horse battery staple", null, null), store, hasher, options, now)).Session!.RefreshToken;

            string earlier = await SignInAsync();
            string kept = await SignInAsync();

            await WebSessions.EndEarlierAsync(earlier, kept, store, now, CancellationToken.None);
            await WebSessions.EndEarlierAsync(kept, kept, store, now, CancellationToken.None);
            await WebSessions.EndEarlierAsync(null, kept, store, now, CancellationToken.None);

            Assert.Null(await AccountFlows.AuthenticateAsync(earlier, store, now));
            Assert.NotNull(await AccountFlows.AuthenticateAsync(kept, store, now));
        }
    }
}
