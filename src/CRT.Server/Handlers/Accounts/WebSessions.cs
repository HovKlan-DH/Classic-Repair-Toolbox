using System.Security.Cryptography;
using System.Text;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // SIGNED IN ON THE WEBSITE (step 3, 2026-11-26) - to comment on a shared collection's page. The sign-in is the very session
    // CRT gets (AccountFlows.LoginAsync; 30 days, sliding as it is used - owner decision, 2026-11-26: "30 days, like CRT"),
    // kept by the browser in a cookie:
    //
    //   - HttpOnly (no script reads it - the pages run none anyway), Secure (HTTPS only), SameSite=Lax (another site's form
    //     posting here does not carry it).
    //   - Named PER INSTANCE (CookieName), and SENT ONLY WHERE IT IS READ (CookiePaths): the collections' pages, the sign-out
    //     and the sign-in - never "/" (code review, 2026-11-27, Security: the live cookie, a bearer token with every right of
    //     the account, went to every page of the site, the test instance's "/crt-test/" - every push to main - among them). So
    //     it is three cookies of one name and token, one per path; one of the "/" path an earlier build set is taken away as a
    //     new sign-in is kept and at the sign-out.
    //   - Written again at every page signed in, so the browser keeps it as the session slides (Keep).
    //
    // *** EVERY SIGNED-IN FORM CARRIES A TOKEN MADE FROM THE SESSION (Csrf) *** - besides the cookie's SameSite and the
    // refusal of a form from another site (AccountPages.IsFromAnotherSite), so a forged post fails three ways. Nobody but the
    // browser holding the session can make it: it is a checksum of the session's own secret.
    // ###########################################################################################
    public static class WebSessions
    {
        // The site's path the pages live under ("" live, "/crt-test" for the test instance).
        public static string PathOf(string siteBaseUrl) =>
            Uri.TryCreate(siteBaseUrl, UriKind.Absolute, out Uri? uri) ? uri.AbsolutePath.TrimEnd('/') : string.Empty;

        public static string CookieName(string siteBaseUrl)
        {
            string path = WebSessions.PathOf(siteBaseUrl).Trim('/');

            return path.Length == 0
                ? "crt_session"
                : "crt_session_" + string.Concat(path.Select(c => char.IsAsciiLetterOrDigit(c) ? c : '_'));
        }

        public const string SignOutPath = "/account/sign-out";

        public const string SignInPath = "/account/sign-in";

        // Where the cookie is sent: the collections' pages, the sign-out, and the sign-in - which ends the session a browser
        // signed in already held (code review, 2026-11-27, Security: replaced in the browser only, it stayed live on the
        // server for 30 days). Nothing else of the site reads it.
        public static IReadOnlyList<string> CookiePaths(string siteBaseUrl) =>
            [
                WebSessions.PathOf(siteBaseUrl) + global::Handlers.DataHandling.CollectionRoutes.PagePath + "/",
                WebSessions.PathOf(siteBaseUrl) + WebSessions.SignOutPath,
                WebSessions.PathOf(siteBaseUrl) + WebSessions.SignInPath
            ];

        // ###########################################################################################
        // A BROWSER SIGNED IN ALREADY - the same account or another - ENDS THE SESSION IT HELD as it signs in again (code review,
        // 2026-11-27, Security: replaced only in the browser, it stayed a live bearer token on the server for up to 30 days, on a
        // shared computer the previous person's). `earlier` is the cookie the sign-in came with (TokenOf), `kept` the new one.
        // ###########################################################################################
        public static async Task EndEarlierAsync(string? earlier, string kept, IAccountStore store, DateTimeOffset now, CancellationToken cancellationToken)
        {
            if (earlier is not null && !string.Equals(earlier, kept, StringComparison.Ordinal))
                await AccountFlows.LogoutAsync(earlier, store, now, cancellationToken);
        }

        // The path the cookie had until 2026-11-27 - the site's whole - only ever taken away.
        public static string EarlierCookiePath(string siteBaseUrl) => WebSessions.PathOf(siteBaseUrl) + "/";

        public static string? TokenOf(HttpRequest request, string siteBaseUrl) =>
            request.Cookies.TryGetValue(WebSessions.CookieName(siteBaseUrl), out string? token) && !string.IsNullOrWhiteSpace(token)
                ? token
                : null;

        // *** THE EARLIER COOKIE IS TAKEN AWAY FIRST. *** ASP.NET's delete also drops every Set-Cookie of the name written before
        // it whose text holds "path=<its path>" - and "path=/" is in "path=/collections/" - so taken away last, it took the
        // cookies just kept with it.
        // A response carrying the cookie is kept by nobody between the browser and the service - whatever page it turns out to
        // be (code review, 2026-11-27, Security); a page that sets its own Cache-Control keeps this one when it writes a cookie.
        public static void Keep(HttpResponse response, string siteBaseUrl, string token, DateTimeOffset expiresUtc)
        {
            response.Cookies.Delete(WebSessions.CookieName(siteBaseUrl), WebSessions.Options(WebSessions.EarlierCookiePath(siteBaseUrl), null));

            foreach (string path in WebSessions.CookiePaths(siteBaseUrl))
                response.Cookies.Append(WebSessions.CookieName(siteBaseUrl), token, WebSessions.Options(path, expiresUtc));

            response.Headers.CacheControl = "private, no-store";
        }

        public static void Forget(HttpResponse response, string siteBaseUrl)
        {
            foreach (string path in WebSessions.CookiePaths(siteBaseUrl).Prepend(WebSessions.EarlierCookiePath(siteBaseUrl)))
                response.Cookies.Delete(WebSessions.CookieName(siteBaseUrl), WebSessions.Options(path, null));

            response.Headers.CacheControl = "private, no-store";
        }

        private static CookieOptions Options(string path, DateTimeOffset? expiresUtc) =>
            new()
            {
                HttpOnly = true,
                Secure = true,
                SameSite = SameSiteMode.Lax,
                Path = path,
                Expires = expiresUtc,
                IsEssential = true
            };

        // The token a signed-in form carries: a checksum of the session's secret, so only its browser can make it.
        public static string Csrf(string sessionToken) =>
            Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes("crt-web-form|" + sessionToken)))[..32];

        public static bool CsrfMatches(string sessionToken, string? presented) =>
            presented is { Length: 32 } &&
            CryptographicOperations.FixedTimeEquals(Encoding.ASCII.GetBytes(WebSessions.Csrf(sessionToken)), Encoding.ASCII.GetBytes(presented));

        // ###########################################################################################
        // Where to go after signing in or out: a collection's page of THIS site, or "Your collections" (step 4), as asked -
        // anything else (another site, a path elsewhere, a scheme) is no place to go, and the account page is. So no link can
        // send a person on from here to somewhere of its choosing. A place on the page ("#comments-CN2-Keyboard-PB7") is
        // letters, digits and "-", as long as an anchor is (CollectionRules.AnchorsOf; code review, 2026-11-27: 40 characters
        // in all cut a long one off, and the reader was not sent back).
        // ###########################################################################################
        // As long as the longest anchor a page makes: "comments-", "capture-", a label and a pin of the longest a capture may name
        // them, the "-" between and a "-300" making it unique - with room (code review, 2026-11-27: a fixed 200 was shorter).
        public const int ReturnAnchorAtMost = (2 * global::Handlers.DataHandling.CollectionRules.MaximumNameLength) + 40;

        public static string? ReturnPath(string? asked, string siteBaseUrl)
        {
            string prefix = WebSessions.PathOf(siteBaseUrl) + global::Handlers.DataHandling.CollectionRoutes.PagePath + "/";

            if (string.IsNullOrEmpty(asked) || !asked.StartsWith(prefix, StringComparison.Ordinal))
                return null;

            string rest = asked[prefix.Length..];
            int hash = rest.IndexOf('#', StringComparison.Ordinal);
            string page = hash < 0 ? rest : rest[..hash];
            string anchor = hash < 0 ? string.Empty : rest[(hash + 1)..];

            bool pageAllowed = page.Length == 0 || global::Handlers.DataHandling.CollectionRules.IsId(page);
            bool anchorAllowed = anchor.Length <= WebSessions.ReturnAnchorAtMost && anchor.All(c => char.IsAsciiLetterOrDigit(c) || c == '-');

            return pageAllowed && anchorAllowed ? asked : null;
        }
    }
}
