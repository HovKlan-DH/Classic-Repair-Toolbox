using System.Net;
using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // THE ACCOUNT WEB PAGES (2026-11-25). *** NOBODY MAKES AN ACCOUNT HERE SINCE SERVER 17.0.0 *** (owner decision, 2026-11-28:
    // "Entirely remove the new "Account" possibility ... I do not want the possibility for everyone to create an account" - an
    // account is a maintainer's, made by invitation). What is left: "I forgot my password", the web's sign-in - the
    // ADMINISTRATOR'S, for the list of every shared collection - and the link a verification mail sent before then carried.
    // Served by this service under "/account/" (AccountPageEndpoints, Apache forwarding the path as it forwards "/api/" -
    // INSTALLING.md); the link in a verification mail, GET /api/accounts/verify, answers the same page as "/account/verify".
    //
    // PURE: strings in, a page out. Every word a person typed is HTML-ENCODED here and nowhere else (Encode) - a page never
    // builds markup out of what it was sent, so a name like "<script>" is shown as those characters. No script at all, no
    // picture, no font and nothing from another site; the one style block is allowed by its own hash in the page's
    // Content-Security-Policy (StyleHash), so not even a style can be injected. A plain form, posting to this service.
    //
    // What a page says of a refusal is the flows' own sentences (AccountFlows, AccountRules, DisplayNames), so the window, the
    // JSON answers and these pages refuse in one set of words. The neutral answer stays neutral: "a code is on its way" for
    // any address (AccountFlows' header - no page here says whether an address has an account).
    // ###########################################################################################
    public static class AccountPages
    {
        // The pages' path under the site, and the site's base for them: the public API's address less its "/api" - so the
        // live service's pages are "https://classic-repair-toolbox.dk/account/" and the test instance's
        // "https://classic-repair-toolbox.dk/crt-test/account/", with no setting of their own.
        public const string Path = "/account";

        public static string BaseUrl(string? publicApiBaseUrl)
        {
            string url = (publicApiBaseUrl ?? string.Empty).Trim().TrimEnd('/');

            return url.EndsWith("/api", StringComparison.OrdinalIgnoreCase) ? url[..^4] : url;
        }

        public static string Link(string baseUrl, string page) => $"{baseUrl}{AccountPages.Path}/{page}".TrimEnd('/');

        // ###########################################################################################
        // The link in a verification mail, opened. The token is the only subject, so the failures are told apart (AccountFlows'
        // header): an expired link needs a new mail, a broken one needs the whole link.
        // ###########################################################################################
        public static (string Html, int Status) Verified(string baseUrl, TokenRedemptionOutcome outcome) => outcome switch
        {
            TokenRedemptionOutcome.Success => (AccountPages.Layout(baseUrl, "Your address is confirmed",
                $"<p>Thank you. You can sign in now - in Classic Repair Toolbox, through {Encode(AccountWording.WhereInCrt)}.</p>"), 200),

            TokenRedemptionOutcome.AlreadyUsed => (AccountPages.Layout(baseUrl, "Your address is confirmed",
                "<p>That link has already been used, so there is nothing more to do. You can close this page.</p>"), 200),

            // The way on is a new password set with a mailed code, which confirms the address too (AccountFlows
            // .CompletePasswordResetAsync) - here, or in CRT's account window. An address not confirmed cannot sign in (2026-11-25,
            // AccountFlows.LoginAsync), so "sign in and press "Send the mail again"", what this page said until then, is none.
            TokenRedemptionOutcome.Expired => (AccountPages.Layout(baseUrl, "That link has expired",
                $"<p>Set a password with a code mailed to you on <a href=\"{Encode(AccountPages.Link(baseUrl, "forgot-password"))}\">" +
                $"{Encode(AccountWording.ForgotPasswordButton)}</a> - that confirms the address too. It is the same in Classic Repair " +
                $"Toolbox, through {Encode(AccountWording.WhereInCrt)}.</p>"), 400),

            _ => (AccountPages.Layout(baseUrl, "That link is not valid",
                "<p>Check that the whole link was opened - a mail program can break a long link over two lines.</p>"), 400)
        };

        // ###########################################################################################
        // "I forgot my password" - ask for a code, then set a password with it, on one page: the code comes by mail, and the
        // page it is typed into should be the one already open. `codeAsked`: a code was just asked for, said in the neutral
        // words (any address). `resetErrors`: what was wrong with the code or the new password.
        // ###########################################################################################
        public static string ForgotPassword(string baseUrl, string? email = null, bool codeAsked = false, IReadOnlyList<string>? resetErrors = null) =>
            AccountPages.Layout(
                baseUrl,
                AccountWording.ForgotPasswordButton,
                "<h2>1. Ask for a code</h2>" +
                (codeAsked
                    ? $"<div class=\"notice\" role=\"status\"><p>If {Encode(email)} has an account, a code is on its way to it. {Encode(AccountWording.ResetCodeWorks)}</p></div>"
                    : string.Empty) +
                $"<form method=\"post\" action=\"{Encode(AccountPages.Link(baseUrl, "forgot-password"))}\">" +
                "<label for=\"email\">Email address</label>" +
                $"<input type=\"email\" id=\"email\" name=\"email\" maxlength=\"{EmailAddressRules.MaximumLength}\" autocomplete=\"email\" required value=\"{Encode(email)}\">" +
                "<button type=\"submit\">Send me a code</button>" +
                "</form>" +
                "<h2>2. Set a new password</h2>" +
                AccountPages.Errors(resetErrors) +
                $"<form method=\"post\" action=\"{Encode(AccountPages.Link(baseUrl, "reset-password"))}\">" +
                "<label for=\"code\">Code from the mail</label>" +
                "<input type=\"text\" id=\"code\" name=\"code\" autocomplete=\"one-time-code\" required>" +
                "<label for=\"newPassword\">New password</label>" +
                "<input type=\"password\" id=\"newPassword\" name=\"newPassword\" autocomplete=\"new-password\" required>" +
                $"<p class=\"muted\">{Encode(AccountWording.PasswordLine)}</p>" +
                "<button type=\"submit\">Set password</button>" +
                "</form>");

        // ###########################################################################################
        // "Sign in" on the website (step 3, 2026-11-26) - THE ADMINISTRATOR'S ALONE since server 17.0.0 (owner decision, 2026-11-28:
        // "I (as web admin) still want to be able to see all collections, so I can control it looks correct"; nobody comments any
        // more), for the list of every shared collection and the delete on a collection's page. `returnTo`: the page to go back to
        // (WebSessions.ReturnPath decided it is one), carried in the form. A wrong address or password is said without saying why,
        // as every sign-in is (AccountFlows' header); right ones of an account that is no administrator's are refused in words,
        // and no session is kept for it (AccountPageEndpoints.SignInAsync).
        // ###########################################################################################
        public const string SignInRefused = "The address or the password is not right.";

        public const string SignInSlowDown = "Too many tries - wait a few minutes, then try again.";

        public const string SignInAdministratorsOnly =
            "Signing in here is for the site's administrator alone - the list of every shared collection is theirs. A maintainer signs " +
            "in in Classic Repair Toolbox, on the \"Maintainer\" tab.";

        public static string SignIn(string baseUrl, string? email = null, string? returnTo = null, IReadOnlyList<string>? errors = null) =>
            AccountPages.Layout(
                baseUrl,
                AccountWording.SignInButton,
                AccountPages.Errors(errors) +
                "<p>For the site's administrator: the list of every shared oscilloscope collection, and taking one down.</p>" +
                $"<form method=\"post\" action=\"{Encode(AccountPages.Link(baseUrl, "sign-in"))}\">" +
                (returnTo is null ? string.Empty : $"<input type=\"hidden\" name=\"return\" value=\"{Encode(returnTo)}\">") +
                "<label for=\"email\">Email address</label>" +
                $"<input type=\"email\" id=\"email\" name=\"email\" maxlength=\"{EmailAddressRules.MaximumLength}\" autocomplete=\"email\" required value=\"{Encode(email)}\">" +
                "<label for=\"password\">Password</label>" +
                "<input type=\"password\" id=\"password\" name=\"password\" autocomplete=\"current-password\" required>" +
                $"<button type=\"submit\">{Encode(AccountWording.SignInButton)}</button>" +
                "</form>" +
                $"<p class=\"muted\"><a href=\"{Encode(AccountPages.Link(baseUrl, "forgot-password"))}\">{Encode(AccountWording.ForgotPasswordButton)}</a></p>");

        public static string SignedOut(string baseUrl) =>
            AccountPages.Layout(baseUrl, "You are signed out", "<p>This browser is no longer signed in.</p>");

        public static string SignedIn(string baseUrl, string displayName) =>
            AccountPages.Layout(
                baseUrl,
                "You are signed in",
                $"<p>Signed in as <b>{Encode(displayName)}</b>. See <a href=\"{Encode(Collections.CollectionPages.ListUrl(baseUrl))}\">" +
                "all collections</a>, or open a collection's link to take it down.</p>");

        public static string PasswordSet(string baseUrl) =>
            AccountPages.Layout(
                baseUrl,
                "Your password is set",
                $"<p>Sign in with it - in Classic Repair Toolbox, through {Encode(AccountWording.WhereInCrt)}.</p>");

        // A form posted from another website (IsFromAnotherSite) - nothing done, and the way to do it here said.
        public static string FromAnotherSite(string baseUrl) =>
            AccountPages.Layout(
                baseUrl,
                "That form came from another website",
                // Every form of the site is refused so (code review, 2026-11-27: this page spoke of making an account to all of
                // them).
                "<p>Nothing was done. Open the page here - " +
                $"<a href=\"{Encode(AccountPages.Link(baseUrl, "sign-in"))}\">the sign-in</a>, " +
                $"<a href=\"{Encode(AccountPages.Link(baseUrl, "forgot-password"))}\">{Encode(AccountWording.ForgotPasswordButton)}</a> " +
                "or the collection's own page - and send the form from it.</p>");

        // ###########################################################################################
        // *** A FORM ANOTHER WEBSITE POSTS HERE IS REFUSED *** (code review, 2026-11-25, Security). Asking for a code MAILS
        // somebody, and the mail budget is counted per IP address (AccountFlows.CheckMailBudgetAsync) - so a hidden form on a
        // busy page, posting here from every visitor's browser, would spend each visitor's budget and flood an inbox, the limit
        // never tripping. The JSON routes take application/json, which no form can send.
        //
        //   Sec-Fetch-Site  every current browser sends it and no page can set it: "same-origin" or "same-site" (a page of
        //                   this site), or "none" (typed or bookmarked) is ours; "cross-site" - and anything else - is not.
        //   Origin          what an older browser sends instead with a POST: this site's own scheme and host, or one of its
        //                   subdomains, is ours; another, or "null" (a sandboxed frame), is not. The pages' Referrer-Policy
        //                   is "strict-origin" for this (ReferrerPolicy) - "no-referrer" makes a browser send "null" for
        //                   the site's own forms.
        //   neither         no browser - a script, which could post directly anyway, so nothing is gained by refusing it.
        // ###########################################################################################
        public static bool IsFromAnotherSite(string? fetchSite, string? origin, string baseUrl)
        {
            if (!string.IsNullOrWhiteSpace(fetchSite))
            {
                string site = fetchSite.Trim();

                return !(site.Equals("same-origin", StringComparison.OrdinalIgnoreCase) ||
                         site.Equals("same-site", StringComparison.OrdinalIgnoreCase) ||
                         site.Equals("none", StringComparison.OrdinalIgnoreCase));
            }

            if (string.IsNullOrWhiteSpace(origin))
                return false;

            if (!Uri.TryCreate(origin.Trim(), UriKind.Absolute, out Uri? from) ||
                !Uri.TryCreate(baseUrl, UriKind.Absolute, out Uri? ours))
                return true;

            if (!string.Equals(from.Scheme, ours.Scheme, StringComparison.OrdinalIgnoreCase))
                return true;

            return !(string.Equals(from.Host, ours.Host, StringComparison.OrdinalIgnoreCase) ||
                     from.Host.EndsWith("." + ours.Host, StringComparison.OrdinalIgnoreCase));
        }

        // The pages' Referrer-Policy: a link sends no more than this site's address, so a verify link's token never leaves in
        // a Referer - and a form's POST keeps its Origin, which IsFromAnotherSite reads ("no-referrer" would send "null").
        public const string ReferrerPolicy = "strict-origin";

        // ###########################################################################################
        // The frame every page stands in: the name of the site, the page, and where CRT is downloaded.
        // ###########################################################################################
        private static string Layout(string baseUrl, string title, string body)
        {
            string site = AccountPages.SiteRoot(baseUrl);

            return "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\">" +
                "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
                "<meta name=\"robots\" content=\"noindex\">" +
                $"<title>{Encode(title)} - Classic Repair Toolbox</title>" +
                $"<style>{AccountPages.Css}</style></head><body>" +
                $"<header><a href=\"{Encode(site)}\">Classic Repair Toolbox</a></header>" +
                $"<main><h1>{Encode(title)}</h1>{body}</main>" +
                "<footer><p>Classic Repair Toolbox helps hobbyists repair vintage computers. " +
                $"<a href=\"{Encode(EmailTemplates.CrtDownloadUrl)}\">Download it</a> - it is free.</p></footer>" +
                "</body></html>";
        }

        private static string Errors(IReadOnlyList<string>? errors) =>
            errors is { Count: > 0 }
                ? "<div class=\"errors\" role=\"alert\">" + string.Concat(errors.Select(error => $"<p>{Encode(error)}</p>")) + "</div>"
                : string.Empty;

        private static string SiteRoot(string baseUrl) =>
            Uri.TryCreate(baseUrl, UriKind.Absolute, out Uri? uri) ? uri.GetLeftPart(UriPartial.Authority) + "/" : "/";

        // Every word a person typed, and every sentence, before it goes into the page.
        private static string Encode(string? text) => WebUtility.HtmlEncode(text ?? string.Empty);

        // ###########################################################################################
        // The pages' one style block - light, and dark where the reader's system is dark. Its hash goes into the
        // Content-Security-Policy (AccountPageEndpoints), which then allows no other style, script or picture at all.
        // ###########################################################################################
        public const string Css =
            ":root{color-scheme:light dark;--bg:#f4f5f7;--card:#ffffff;--fg:#1d1f23;--muted:#555b66;--edge:#d5d8de;--link:#2a78d6;--go:#497f54;--fail:#bc5454}" +
            "@media (prefers-color-scheme:dark){:root{--bg:#1b1d21;--card:#25282d;--fg:#e6e8eb;--muted:#a9afb8;--edge:#3a3e45;--link:#7fb2f0;--go:#497f54;--fail:#e08a8a}}" +
            "body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.5 system-ui,-apple-system,\"Segoe UI\",Roboto,sans-serif}" +
            "header{padding:14px 20px;border-bottom:1px solid var(--edge);background:var(--card)}" +
            "header a{color:var(--fg);font-weight:600;text-decoration:none}" +
            "main{max-width:520px;margin:28px auto;padding:24px;background:var(--card);border:1px solid var(--edge);border-radius:8px}" +
            "h1{font-size:1.4em;margin:0 0 12px}h2{font-size:1.1em;margin:28px 0 8px}p{margin:0 0 12px}" +
            ".muted{color:var(--muted);font-size:.9em;margin-top:4px}" +
            "label{display:block;font-weight:600;margin:16px 0 4px}" +
            "input[type=text],input[type=email],input[type=password]{box-sizing:border-box;width:100%;padding:8px 10px;font:inherit;color:var(--fg);background:var(--bg);border:1px solid var(--edge);border-radius:4px}" +
            "button{margin-top:12px;padding:8px 16px;font:inherit;color:#ffffff;background:var(--go);border:0;border-radius:4px;cursor:pointer}" +
            "a{color:var(--link)}" +
            ".errors{border:1px solid var(--fail);color:var(--fail);border-radius:4px;padding:8px 12px;margin:0 0 12px}.errors p{margin:0}" +
            ".notice{border:1px solid var(--go);border-radius:4px;padding:8px 12px;margin:0 0 12px}.notice p{margin:0}" +
            "footer{max-width:520px;margin:0 auto 28px;padding:0 24px;color:var(--muted);font-size:.85em}";

        public static readonly string StyleHash =
            "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(AccountPages.Css)));

        // The headers every page is sent with: nothing but the style above may run or load, no other site may frame it, its
        // links send no address (a verify link's token must not leave in a Referer), and nothing keeps a copy.
        public static readonly string ContentSecurityPolicy =
            $"default-src 'none'; style-src '{AccountPages.StyleHash}'; form-action 'self'; frame-ancestors 'none'; base-uri 'none'";
    }
}
