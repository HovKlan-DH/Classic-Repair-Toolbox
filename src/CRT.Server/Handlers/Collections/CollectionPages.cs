using System.Net;
using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // A SHARED COLLECTION'S PAGE (owner requests, 2026-11-25): what a forum reader opens from a link. Built by the server, so
    // a forum's link preview (which runs no script) gets its title and first picture (the og: tags).
    //
    //   - IT NAMES NOBODY: the board, the owner's title and words, the collection's id, the date it was shared and the date it
    //     is deleted (owner decision: "For the name of the collection I do not think any user identification should be shown
    //     at all"; "The collection page should also show date of creation and date of deletion").
    //   - IT SAYS WHERE IT COMES FROM: Classic Repair Toolbox, with a link to download it.
    //   - EACH CAPTURE BESIDE CRT'S OWN BASELINE FOR THAT PIN, when the board has one - the baseline is in the published data,
    //     linked from there, never uploaded - and their settings in a table, a value that differs from the baseline's in bold,
    //     so a wrong oscilloscope setup shows before anybody has to ask.
    //   - Every capture has its own anchor ("#U7-12"), so a forum reply can point at one pin.
    //
    // PURE: a collection in, a page out. Every word anybody typed is HTML-ENCODED here (Encode); the page carries no script,
    // and its one style block is allowed by its hash (StyleHash). Search engines are asked to keep nothing (noindex) - a
    // collection is reached through its link alone, and nothing lists them.
    // ###########################################################################################
    public static class CollectionPages
    {
        // The collection as a person names it: its title, else its board.
        public static string NameOf(CollectionRecord collection) =>
            collection.Title ?? $"Oscilloscope captures of {collection.HardwareName} {collection.BoardName}";

        public static string BoardLine(CollectionRecord collection) =>
            $"{collection.HardwareName} - {collection.BoardName}" + (collection.Region is { } region ? $" ({region})" : string.Empty);

        public static string ImageUrl(string siteBaseUrl, string id, string sha256) =>
            $"{CollectionRules.PageUrl(siteBaseUrl, id)}/{sha256}.png";

        // ###########################################################################################
        // The page of a collection shown. `baselineUrl` gives the address of a capture's baseline picture in the published data,
        // or null when it is not there.
        // ###########################################################################################
        public static string Page(string siteBaseUrl, CollectionRecord collection, Func<CollectionBaseline, string?> baselineUrl, CollectionPageView? view = null)
        {
            view ??= CollectionPageView.SignedOut;

            string name = CollectionPages.NameOf(collection);
            IReadOnlyList<string> anchors = CollectionRules.AnchorsOf(collection.Images.Select(image => image.Entry).ToList());
            var body = new StringBuilder();

            body.Append($"<p class=\"board\">{Encode(CollectionPages.BoardLine(collection))}</p>");
            body.Append("<p class=\"facts\">")
                .Append($"Collection <code>{Encode(collection.Id)}</code>")
                .Append($" | Shared {Encode(CollectionRules.DateText(collection.CreatedUtc))}")
                .Append($" | Deleted on {Encode(CollectionRules.DateText(collection.DeleteAtUtc))}")
                .Append("</p>");

            // Who is signed in, and the way out - the form carries its token (WebSessions.Csrf).
            if (view.Viewer is { } viewer)
            {
                body.Append("<form class=\"signedin\" method=\"post\" action=\"")
                    .Append(Encode(Accounts.AccountPages.Link(Accounts.AccountPages.BaseUrl(siteBaseUrl), "sign-out")))
                    .Append("\">")
                    .Append($"Signed in as <b>{Encode(viewer.DisplayName)}</b> ")
                    .Append(CollectionPages.Hidden("csrf", view.Csrf))
                    .Append(CollectionPages.Hidden("return", CollectionPages.PagePath(siteBaseUrl, collection.Id)))
                    .Append("<button type=\"submit\" class=\"link\">Sign out</button></form>");
            }

            if (collection.Description is { } description)
                body.Append($"<p class=\"text\">{Encode(description)}</p>");

            if (collection.ForumLinks.Count > 0)
            {
                body.Append("<p class=\"discussion\">Discussion: ");
                body.Append(string.Join(" | ", collection.ForumLinks.Select(link =>
                    $"<a href=\"{Encode(link)}\" rel=\"nofollow ugc noopener noreferrer\">{Encode(CollectionPages.LinkText(link))}</a>")));
                body.Append("</p>");
            }

            body.Append("<nav class=\"contents\"><h2>Captures</h2><ol>");

            for (int i = 0; i < collection.Images.Count; i++)
            {
                CollectionImageEntry image = collection.Images[i].Entry;

                body.Append($"<li><a href=\"#{anchors[i]}\">{Encode(CollectionRules.WhereOf(image))}</a>");

                if (image.SignalName is { } signal)
                    body.Append($" {Encode(signal)}");

                if (image.LooksWrong)
                    body.Append(" <span class=\"flag\">Looks wrong</span>");

                int comments = view.Comments.Count(comment => comment.ImageSha256 == image.Sha256 && !comment.IsHidden);

                if (comments > 0)
                    body.Append($" <span class=\"muted\">({comments} comment{(comments == 1 ? string.Empty : "s")})</span>");

                body.Append("</li>");
            }

            body.Append("</ol></nav>");

            for (int i = 0; i < collection.Images.Count; i++)
            {
                body.Append(CollectionPages.CaptureSection(
                    siteBaseUrl,
                    collection,
                    collection.Images[i],
                    anchors[i],
                    previous: i > 0 ? anchors[i - 1] : null,
                    next: i < collection.Images.Count - 1 ? anchors[i + 1] : null,
                    baselineUrl,
                    view));
            }

            string description2 =
                $"{collection.Images.Count} oscilloscope capture{(collection.Images.Count == 1 ? string.Empty : "s")} of " +
                $"{CollectionPages.BoardLine(collection)}, shared with Classic Repair Toolbox.";

            var preview = new StringBuilder()
                .Append("<meta property=\"og:type\" content=\"website\">")
                .Append("<meta property=\"og:site_name\" content=\"Classic Repair Toolbox\">")
                .Append($"<meta property=\"og:title\" content=\"{Encode(name)}\">")
                .Append($"<meta property=\"og:description\" content=\"{Encode(description2)}\">")
                .Append($"<meta property=\"og:url\" content=\"{Encode(CollectionRules.PageUrl(siteBaseUrl, collection.Id))}\">");

            if (collection.Images.Count > 0)
            {
                CollectionImageRecord first = collection.Images[0];

                preview
                    .Append($"<meta property=\"og:image\" content=\"{Encode(CollectionPages.ImageUrl(siteBaseUrl, collection.Id, first.Entry.Sha256))}\">")
                    .Append($"<meta property=\"og:image:width\" content=\"{first.Width}\">")
                    .Append($"<meta property=\"og:image:height\" content=\"{first.Height}\">")
                    .Append("<meta name=\"twitter:card\" content=\"summary_large_image\">");
            }

            return CollectionPages.Layout(siteBaseUrl, name, body.ToString(), preview.ToString());
        }

        // A collection deleted - when, and why in a few plain words. Its link stays answerable, so a forum reader years later
        // learns the page went on purpose.
        public static string Deleted(string siteBaseUrl, CollectionRecord collection)
        {
            string why = collection.DeletedReason switch
            {
                CollectionDeletedReasons.Expired => "Its time was up - it was kept for as long as the person who shared it chose.",
                CollectionDeletedReasons.Owner => "The person who shared it deleted it.",
                _ => "It was removed."
            };

            return CollectionPages.Layout(
                siteBaseUrl,
                "This collection was deleted",
                $"<p>Collection <code>{Encode(collection.Id)}</code> was deleted on " +
                $"{Encode(CollectionRules.DateText(collection.DeletedUtc ?? collection.DeleteAtUtc))}. {Encode(why)}</p>");
        }

        public static string NotFound(string siteBaseUrl) =>
            CollectionPages.Layout(
                siteBaseUrl,
                "There is no such collection",
                "<p>Check that the whole link was copied - a collection's address ends in its 12-character id.</p>");

        public static string NotSetUp(string siteBaseUrl) =>
            CollectionPages.Layout(siteBaseUrl, "Collections are not available", $"<p>{Encode(CollectionFlows.NotSetUp)}</p>");

        // ###########################################################################################
        // One capture: the picture beside its baseline, their settings, what it should read, and the owner's note.
        // ###########################################################################################
        private static string CaptureSection(
            string siteBaseUrl,
            CollectionRecord collection,
            CollectionImageRecord capture,
            string anchor,
            string? previous,
            string? next,
            Func<CollectionBaseline, string?> baselineUrl,
            CollectionPageView view)
        {
            CollectionImageEntry image = capture.Entry;
            CollectionBaseline? baseline = image.Baseline;
            string where = CollectionRules.WhereOf(image);
            string? baselinePicture = baseline is { Sha256: not null } ? baselineUrl(baseline) : null;
            var html = new StringBuilder();

            html.Append($"<section class=\"capture\" id=\"{anchor}\">");
            html.Append($"<h2>{Encode(where)}");

            if (image.SignalName is { } signal)
                html.Append($" <span class=\"signal\">{Encode(signal)}</span>");

            if (image.LooksWrong)
                html.Append(" <span class=\"flag\">Looks wrong</span>");

            html.Append("</h2><div class=\"pair\">");
            html.Append("<figure>")
                .Append($"<img src=\"{Encode(CollectionPages.ImageUrl(siteBaseUrl, collection.Id, image.Sha256))}\" width=\"{capture.Width}\" height=\"{capture.Height}\" alt=\"{Encode($"Oscilloscope capture of {where}")}\">")
                .Append("<figcaption>This capture</figcaption></figure>");

            if (baselinePicture is not null)
            {
                html.Append("<figure>")
                    .Append($"<img src=\"{Encode(baselinePicture)}\" alt=\"{Encode($"Classic Repair Toolbox's baseline of {where}")}\">")
                    .Append("<figcaption>Classic Repair Toolbox's baseline</figcaption></figure>");
            }

            html.Append("</div>");

            bool withBaseline = baseline is not null;

            html.Append("<table class=\"settings\"><thead><tr><th></th><th>This capture</th>");

            if (withBaseline)
                html.Append("<th>Baseline</th>");

            html.Append("</tr></thead><tbody>");
            html.Append(CollectionPages.SettingRow("T/DIV", image.TimePerDivision, baseline?.TimePerDivision, withBaseline));
            html.Append(CollectionPages.SettingRow("V/DIV", image.VoltsPerDivision, baseline?.VoltsPerDivision, withBaseline));
            html.Append(CollectionPages.SettingRow("Trigger level", image.TriggerLevel, baseline?.TriggerLevel, withBaseline));
            html.Append("</tbody></table>");

            if (image.ExpectedReading is { } reading)
                html.Append($"<p><b>Expected reading:</b> {Encode(reading)}</p>");

            if (image.Caption is { } caption)
                html.Append($"<p class=\"text\">{Encode(caption)}</p>");

            html.Append(CollectionPages.CommentsOf(siteBaseUrl, collection, image.Sha256, anchor, view));

            html.Append("<p class=\"steps\">");
            html.Append(previous is null ? "Previous" : $"<a href=\"#{previous}\">Previous</a>");
            html.Append(" | <a href=\"#top\">Top</a> | ");
            html.Append(next is null ? "Next" : $"<a href=\"#{next}\">Next</a>");
            html.Append("</p></section>");

            return html.ToString();
        }

        // ###########################################################################################
        // A CAPTURE'S COMMENTS (step 3): each with its writer's name - "Owner" by the collection's owner's, "Administrator" by an
        // administrator's - when, and its text as typed; its writer's "Edit" and "Delete", the owner's "Hide" and "Show again";
        // then the form to comment, or the way to sign in. No script: editing or deleting one is a page of its own state
        // (?edit=, ?delete=), every change a form posting to this service with the session's token.
        // ###########################################################################################
        private static string CommentsOf(string siteBaseUrl, CollectionRecord collection, string sha256, string anchor, CollectionPageView view)
        {
            List<CollectionCommentRecord> comments = view.Comments.Where(comment => comment.ImageSha256 == sha256).ToList();
            string page = CollectionRules.PageUrl(siteBaseUrl, collection.Id);
            CollectionViewer? viewer = view.Viewer;
            bool moderates = viewer is not null && (viewer.Id == collection.AccountId || viewer.IsAdministrator);
            var html = new StringBuilder();

            html.Append($"<div class=\"comments\" id=\"comments-{anchor}\">");
            html.Append($"<h3>Comments{(comments.Count > 0 ? $" ({comments.Count})" : string.Empty)}</h3>");

            foreach (CollectionCommentRecord comment in comments)
            {
                bool mine = viewer is not null && viewer.Id == comment.AccountId;

                html.Append($"<article class=\"comment{(comment.IsHidden ? " hidden" : string.Empty)}\" id=\"comment-{comment.Id}\">");
                html.Append($"<p class=\"who\"><b>{Encode(comment.AuthorName)}</b>");

                if (comment.AccountId == collection.AccountId)
                    html.Append($" <span class=\"badge\">{CollectionComments.OwnerBadge}</span>");

                if (comment.AuthorIsAdministrator)
                    html.Append($" <span class=\"badge\">{CollectionComments.AdministratorBadge}</span>");

                html.Append($" <span class=\"muted\">{Encode(CollectionComments.WhenText(comment.CreatedUtc))}{(comment.EditedUtc is null ? string.Empty : " - edited")}</span>");

                if (comment.IsHidden)
                    html.Append(" <span class=\"flag\">Hidden - only you see it</span>");

                html.Append("</p>");

                if (mine && view.EditingId == comment.Id)
                {
                    html.Append($"<form method=\"post\" action=\"{Encode($"{page}/comments/{comment.Id}/edit")}\">")
                        .Append(CollectionPages.Hidden("csrf", view.Csrf))
                        .Append(CollectionPages.TextArea(view.ErrorOn == comment.Id.ToString(System.Globalization.CultureInfo.InvariantCulture) ? view.Typed : comment.Body))
                        .Append(CollectionPages.ErrorsFor(view, comment.Id.ToString(System.Globalization.CultureInfo.InvariantCulture)))
                        .Append("<p class=\"actions\"><button type=\"submit\">Save</button> ")
                        .Append($"<a href=\"{Encode($"{page}#comment-{comment.Id}")}\">Cancel</a></p></form>");
                }
                else
                {
                    html.Append($"<p class=\"text\">{Encode(comment.Body)}</p>");
                }

                if ((mine || viewer is { IsAdministrator: true }) && view.DeletingId == comment.Id)
                {
                    html.Append($"<form class=\"actions\" method=\"post\" action=\"{Encode($"{page}/comments/{comment.Id}/delete")}\">")
                        .Append("Delete this comment? ")
                        .Append(CollectionPages.Hidden("csrf", view.Csrf))
                        .Append("<button type=\"submit\" class=\"danger\">Delete comment</button> ")
                        .Append($"<a href=\"{Encode($"{page}#comment-{comment.Id}")}\">Keep it</a></form>");
                }
                else if (view.EditingId != comment.Id)
                {
                    var actions = new List<string>();

                    if (mine)
                        actions.Add($"<a href=\"{Encode($"{page}?edit={comment.Id}#comment-{comment.Id}")}\">Edit</a>");

                    if (mine || viewer is { IsAdministrator: true })
                        actions.Add($"<a href=\"{Encode($"{page}?delete={comment.Id}#comment-{comment.Id}")}\">Delete</a>");

                    if (moderates && !mine)
                    {
                        string verb = comment.IsHidden ? "show" : "hide";

                        actions.Add(
                            $"<form class=\"inline\" method=\"post\" action=\"{Encode($"{page}/comments/{comment.Id}/{verb}")}\">" +
                            CollectionPages.Hidden("csrf", view.Csrf) +
                            $"<button type=\"submit\" class=\"link\">{(comment.IsHidden ? "Show again" : "Hide")}</button></form>");
                    }

                    if (actions.Count > 0)
                        html.Append($"<p class=\"actions\">{string.Join(" | ", actions)}</p>");
                }

                html.Append("</article>");
            }

            if (viewer is not null)
            {
                html.Append($"<form method=\"post\" action=\"{Encode($"{page}/comments")}\">")
                    .Append(CollectionPages.Hidden("csrf", view.Csrf))
                    .Append(CollectionPages.Hidden("image", sha256))
                    .Append($"<label for=\"comment-on-{anchor}\">Your comment</label>")
                    .Append(CollectionPages.TextArea(view.ErrorOn == sha256 ? view.Typed : null, $"comment-on-{anchor}"))
                    .Append(CollectionPages.ErrorsFor(view, sha256))
                    .Append("<p class=\"actions\"><button type=\"submit\">Comment</button></p></form>");
            }
            else
            {
                string signIn = Accounts.AccountPages.Link(Accounts.AccountPages.BaseUrl(siteBaseUrl), "sign-in") +
                                "?return=" + Uri.EscapeDataString($"{CollectionPages.PagePath(siteBaseUrl, collection.Id)}#comments-{anchor}");
                string register = Accounts.AccountPages.Link(Accounts.AccountPages.BaseUrl(siteBaseUrl), "register");

                html.Append($"<p class=\"muted\"><a href=\"{Encode(signIn)}\">Sign in</a> to comment - ")
                    .Append($"<a href=\"{Encode(register)}\">make an account</a> if you have none.</p>");
            }

            html.Append("</div>");
            return html.ToString();
        }

        private static string TextArea(string? text, string? id = null) =>
            $"<textarea name=\"body\"{(id is null ? string.Empty : $" id=\"{id}\"")} rows=\"4\" maxlength=\"{CollectionComments.MaximumLength}\" required>{Encode(text)}</textarea>";

        private static string ErrorsFor(CollectionPageView view, string where) =>
            view.ErrorOn == where && view.Errors.Count > 0
                ? "<div class=\"errors\" role=\"alert\">" + string.Concat(view.Errors.Select(error => $"<p>{Encode(error)}</p>")) + "</div>"
                : string.Empty;

        private static string Hidden(string name, string? value) => $"<input type=\"hidden\" name=\"{name}\" value=\"{Encode(value)}\">";

        // The page's path on the site, as a sign-in returns to it ("/crt-test/collections/<id>" on the test instance).
        public static string PagePath(string siteBaseUrl, string id) =>
            $"{Accounts.WebSessions.PathOf(siteBaseUrl)}{CollectionRoutes.PagePath}/{id}";

        // A setting of the capture beside the baseline's - in bold when the two differ (spaces and capitals aside).
        private static string SettingRow(string name, string? value, string? baselineValue, bool withBaseline)
        {
            bool differs = withBaseline && value is not null && baselineValue is not null &&
                           !string.Equals(CollectionPages.Comparable(value), CollectionPages.Comparable(baselineValue), StringComparison.Ordinal);

            string shown = value is null ? "<span class=\"muted\">not read</span>" : differs ? $"<b>{Encode(value)}</b>" : Encode(value);

            return $"<tr><th>{Encode(name)}</th><td>{shown}</td>" +
                   (withBaseline ? $"<td>{(baselineValue is null ? "<span class=\"muted\">not given</span>" : Encode(baselineValue))}</td>" : string.Empty) +
                   "</tr>";
        }

        private static string Comparable(string value) => string.Concat(value.Where(c => !char.IsWhiteSpace(c))).ToLowerInvariant();

        // A forum link as shown: its host and path, no scheme.
        private static string LinkText(string link) =>
            Uri.TryCreate(link, UriKind.Absolute, out Uri? uri) ? (uri.Host + uri.PathAndQuery).TrimEnd('/') : link;

        private static string Layout(string siteBaseUrl, string title, string body, string head = "")
        {
            string site = Uri.TryCreate(siteBaseUrl, UriKind.Absolute, out Uri? uri) ? uri.GetLeftPart(UriPartial.Authority) + "/" : "/";

            return "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\">" +
                "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
                "<meta name=\"robots\" content=\"noindex, nofollow\">" +
                $"<title>{Encode(title)} - Classic Repair Toolbox</title>" +
                head +
                $"<style>{CollectionPages.Css}</style></head><body id=\"top\">" +
                $"<header><a href=\"{Encode(site)}\">Classic Repair Toolbox</a><span>Oscilloscope collection</span></header>" +
                $"<main><h1>{Encode(title)}</h1>{body}</main>" +
                "<footer><p>Shared with Classic Repair Toolbox, which helps hobbyists repair vintage computers. " +
                $"<a href=\"{Encode(EmailTemplates.CrtDownloadUrl)}\">Download it</a> - it is free.</p>" +
                $"<p>Something on this page that should not be here? Mail <a href=\"mailto:{Encode(EmailTemplates.ContactAddress)}\">" +
                $"{Encode(EmailTemplates.ContactAddress)}</a> with its address.</p></footer>" +
                "</body></html>";
        }

        private static string Encode(string? text) => WebUtility.HtmlEncode(text ?? string.Empty);

        // ###########################################################################################
        // The page's one style block, light and dark by the reader's system - the account pages' colors. Its hash goes into the
        // Content-Security-Policy, which allows no other style and no script at all.
        // ###########################################################################################
        public const string Css =
            ":root{color-scheme:light dark;--bg:#f4f5f7;--card:#ffffff;--fg:#1d1f23;--muted:#555b66;--edge:#d5d8de;--link:#2a78d6;--flag:#bc5454;--diff:#fff3cd;--go:#497f54;--fail:#bc5454}" +
            "@media (prefers-color-scheme:dark){:root{--bg:#1b1d21;--card:#25282d;--fg:#e6e8eb;--muted:#a9afb8;--edge:#3a3e45;--link:#7fb2f0;--flag:#e08a8a;--diff:#4a4020;--go:#497f54;--fail:#bc5454}}" +
            "body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.5 system-ui,-apple-system,\"Segoe UI\",Roboto,sans-serif}" +
            "header{display:flex;justify-content:space-between;gap:12px;padding:14px 20px;border-bottom:1px solid var(--edge);background:var(--card)}" +
            "header a{color:var(--fg);font-weight:600;text-decoration:none}header span{color:var(--muted)}" +
            "main{max-width:1100px;margin:24px auto;padding:24px;background:var(--card);border:1px solid var(--edge);border-radius:8px}" +
            "h1{font-size:1.5em;margin:0 0 6px}h2{font-size:1.15em;margin:0 0 10px}p{margin:0 0 12px}" +
            ".board{font-size:1.1em;margin-bottom:4px}.facts{color:var(--muted);font-size:.9em}" +
            ".text{white-space:pre-wrap;overflow-wrap:anywhere}.discussion{overflow-wrap:anywhere}a{color:var(--link)}code{font-size:.95em}" +
            ".contents ol{margin:0 0 8px;padding-left:24px;columns:2 260px}.contents h2{margin-top:20px}" +
            ".capture{border-top:1px solid var(--edge);padding-top:20px;margin-top:24px}" +
            ".signal{color:var(--muted);font-weight:normal}" +
            ".flag{display:inline-block;font-size:.75em;font-weight:600;color:var(--flag);border:1px solid var(--flag);border-radius:10px;padding:0 8px;vertical-align:middle}" +
            ".pair{display:flex;flex-wrap:wrap;gap:16px;margin-bottom:12px}" +
            "figure{margin:0;flex:1 1 400px;max-width:100%}figure img{display:block;max-width:100%;height:auto;border:1px solid var(--edge)}" +
            "figcaption{color:var(--muted);font-size:.9em;margin-top:4px}" +
            ".settings{border-collapse:collapse;margin-bottom:12px}.settings th,.settings td{text-align:left;padding:4px 14px 4px 0}" +
            ".settings thead th{color:var(--muted);font-weight:600;font-size:.9em}.settings b{background:var(--diff);padding:0 4px;border-radius:3px}" +
            ".muted{color:var(--muted)}.steps{font-size:.9em;color:var(--muted)}" +
            ".signedin{font-size:.9em;color:var(--muted);margin:0 0 12px}" +
            ".comments{margin:8px 0 12px;padding:12px 16px;border:1px solid var(--edge);border-radius:6px;background:var(--bg)}" +
            ".comments h3{font-size:1em;margin:0 0 8px}.comment{border-top:1px solid var(--edge);padding:8px 0}.comment:first-of-type{border-top:0}" +
            ".comment.hidden{opacity:.65}.who{margin:0 0 4px}.comment .text{margin:0 0 4px}.actions{font-size:.9em;margin:4px 0 0}" +
            ".badge{display:inline-block;font-size:.75em;font-weight:600;color:var(--card);background:var(--link);border-radius:10px;padding:0 8px;vertical-align:middle}" +
            "label{display:block;font-weight:600;margin:8px 0 4px}" +
            "textarea{box-sizing:border-box;width:100%;padding:8px 10px;font:inherit;color:var(--fg);background:var(--card);border:1px solid var(--edge);border-radius:4px}" +
            "button{padding:6px 14px;font:inherit;color:#ffffff;background:var(--go);border:0;border-radius:4px;cursor:pointer}" +
            "button.danger{background:var(--fail)}button.link{padding:0;color:var(--link);background:none;text-decoration:underline}" +
            "form.inline{display:inline}" +
            ".errors{border:1px solid var(--fail);color:var(--fail);border-radius:4px;padding:8px 12px;margin:8px 0}.errors p{margin:0}" +
            "footer{max-width:1100px;margin:0 auto 28px;padding:0 24px;color:var(--muted);font-size:.85em}";

        public static readonly string StyleHash =
            "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Css)));

        // Nothing but the style above may run, pictures only from this site or where the published data is served, and forms
        // post only to this site (the comments and signing out).
        public static string ContentSecurityPolicy(IEnumerable<string> imageOrigins)
        {
            string origins = string.Join(" ", imageOrigins.Distinct(StringComparer.OrdinalIgnoreCase));

            return $"default-src 'none'; style-src '{CollectionPages.StyleHash}'; img-src 'self'{(origins.Length > 0 ? " " + origins : string.Empty)}; " +
                   "form-action 'self'; frame-ancestors 'none'; base-uri 'none'";
        }
    }

    // ###########################################################################################
    // WHO READS THE PAGE, AND WHAT OF THE COMMENTS THEY SEE (step 3). Signed out: the comments shown to everybody, and the way to
    // sign in. Signed in: a form under each capture, the token every form carries (WebSessions.Csrf), and what the reader is
    // doing - a comment being edited (?edit=) or asked about before it is deleted (?delete=), or a refusal with what was typed.
    // ###########################################################################################
    public sealed record CollectionPageView
    {
        public static readonly CollectionPageView SignedOut = new();

        public CollectionViewer? Viewer { get; init; }

        public string? Csrf { get; init; }

        // As shown to this reader (CollectionCommentFlows.ShownTo).
        public IReadOnlyList<CollectionCommentRecord> Comments { get; init; } = [];

        public long? EditingId { get; init; }

        public long? DeletingId { get; init; }

        // What a refusal is about - a capture's checksum (a new comment) or a comment's id (an edit) - its sentences, and the
        // text that was typed, put back in its box.
        public string? ErrorOn { get; init; }

        public IReadOnlyList<string> Errors { get; init; } = [];

        public string? Typed { get; init; }
    }

    // The reader signed in - never the account's address or anything else of it.
    public sealed record CollectionViewer(long Id, string DisplayName, bool IsAdministrator);
}
