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
    //   - Named and scoped PER INSTANCE (CookieName, CookiePath): the test instance's pages live under "/crt-test/", and a
    //     cookie of the live site's path would be sent to them too.
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

        public static string CookiePath(string siteBaseUrl) => WebSessions.PathOf(siteBaseUrl) + "/";

        public static string? TokenOf(HttpRequest request, string siteBaseUrl) =>
            request.Cookies.TryGetValue(WebSessions.CookieName(siteBaseUrl), out string? token) && !string.IsNullOrWhiteSpace(token)
                ? token
                : null;

        public static void Keep(HttpResponse response, string siteBaseUrl, string token, DateTimeOffset expiresUtc) =>
            response.Cookies.Append(WebSessions.CookieName(siteBaseUrl), token, WebSessions.Options(siteBaseUrl, expiresUtc));

        public static void Forget(HttpResponse response, string siteBaseUrl) =>
            response.Cookies.Delete(WebSessions.CookieName(siteBaseUrl), WebSessions.Options(siteBaseUrl, null));

        private static CookieOptions Options(string siteBaseUrl, DateTimeOffset? expiresUtc) =>
            new()
            {
                HttpOnly = true,
                Secure = true,
                SameSite = SameSiteMode.Lax,
                Path = WebSessions.CookiePath(siteBaseUrl),
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
        // Where to go after signing in or out: a collection's page of THIS site, as asked - anything else (another site, a
        // path elsewhere, a scheme) is no place to go, and the account page is. So no link can send a person on from here to
        // somewhere of its choosing.
        // ###########################################################################################
        public static string? ReturnPath(string? asked, string siteBaseUrl)
        {
            string prefix = WebSessions.PathOf(siteBaseUrl) + global::Handlers.DataHandling.CollectionRoutes.PagePath + "/";

            if (string.IsNullOrEmpty(asked) || !asked.StartsWith(prefix, StringComparison.Ordinal))
                return null;

            string rest = asked[prefix.Length..];

            return rest.Length is > 0 and <= 40 && rest.All(c => char.IsAsciiLetterOrDigit(c) || c is '#' or '-')
                ? asked
                : null;
        }
    }
}
