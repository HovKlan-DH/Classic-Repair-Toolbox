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
    //   - IT SAYS WHERE IT COMES FROM: Classic Repair Toolbox, with a link to its project page.
    //   - EACH CAPTURE SEEN LARGE TELLS OF ITS COMPONENT as CRT's component info window did - its facts, its pins, and CRT's
    //     drawing of its pinout with the pin marked (InfoOf).
    //   - Every capture has its own anchor ("#U7-12") and its own address seen large ("#view.U7-12"), so a forum reply can
    //     point at one pin.
    //
    // PURE: a collection in, a page out. Every word anybody typed is HTML-ENCODED here (Encode); the page's one style block and
    // its one small script - the viewer's keys, fixed text with nothing typed in it - are allowed by their hashes (StyleHash,
    // ScriptHash), and nothing else may run. Search engines are asked to keep nothing (noindex) - a
    // collection is reached through its link alone; the one list of them, "Your collections" (List, step 4), shows a reader
    // only their own and those they commented on, and every one to the administrator alone.
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
        // THE PAGE OF A COLLECTION SHOWN.
        //
        // *** THE LOOK IS THE OWNER'S OSCILLOSCOPE NETWORK CAPTURE GALLERY'S (owner request, 2026-11-27: "find heavy inspiration in
        // my other project ... a nice dark view with nice UI functionalities"). *** A dark page under a frosted bar - CRT on the
        // left, the board and the red "Will be deleted" badge in the middle, "Download all" on the right - then the captures as a
        // grid of cards, each opening the VIEWER (Viewers). The viewer is the page's own CSS (:target), so it works with no
        // script at all; the one small script (Script, allowed by its hash) only adds the keys, the wheel and "Copy". Under the
        // grid each capture's section: the picture, its settings, the note and the comments. (CRT's own baseline beside each
        // went on 2026-11-27 - owner decision: "that is not relevant".)
        // ###########################################################################################
        public static string Page(string siteBaseUrl, CollectionRecord collection, CollectionPageView? view = null)
        {
            view ??= CollectionPageView.SignedOut;

            string name = CollectionPages.NameOf(collection);
            IReadOnlyList<string> anchors = CollectionRules.AnchorsOf(collection.Images.Select(image => image.Entry).ToList());
            var body = new StringBuilder();

            body.Append("<section class=\"panel intro\">");
            body.Append($"<h1>{Encode(name)}</h1>");
            body.Append("<p class=\"facts\">")
                .Append($"Collection <code>{Encode(collection.Id)}</code>")
                .Append($" | Shared {Encode(CollectionRules.DateText(collection.CreatedUtc))}")
                .Append($" | {CollectionPages.Counted(collection.Images.Count, "capture")}")
                .Append("</p>");

            if (view.Viewer is { } viewer)
                body.Append(CollectionPages.SignedInLine(siteBaseUrl, viewer, view.Csrf, CollectionPages.PagePath(siteBaseUrl, collection.Id), onList: false));

            if (view.PageErrors.Count > 0)
                body.Append("<div class=\"errors\" role=\"alert\">" + string.Concat(view.PageErrors.Select(error => $"<p>{Encode(error)}</p>")) + "</div>");

            // Shown to the administrator alone - and, sharing stopped, to its owner (CollectionFlows.FindShownToAsync).
            if (view.Offline != CollectionOffline.No)
            {
                body.Append("<div class=\"errors\" role=\"status\"><p>")
                    .Append(Encode(CollectionPages.OfflineLine(view.Offline, ownerReads: view.Viewer is { } reader && reader.Id == collection.AccountId)))
                    .Append("</p></div>");
            }

            if (view.Viewer is { IsAdministrator: true })
                body.Append(CollectionPages.RemoveCollection(siteBaseUrl, collection, view));

            if (collection.Description is { } description)
                body.Append($"<p class=\"text\">{Encode(description)}</p>");

            if (collection.ForumLinks.Count > 0)
            {
                body.Append("<p class=\"discussion\">Discussion: ");
                body.Append(string.Join(" | ", collection.ForumLinks.Select(link =>
                    $"<a href=\"{Encode(link)}\" rel=\"nofollow ugc noopener noreferrer\">{Encode(CollectionPages.LinkText(link))}</a>")));
                body.Append("</p>");
            }

            body.Append("</section>");

            body.Append(CollectionPages.Gallery(siteBaseUrl, collection, anchors, view));

            for (int i = 0; i < collection.Images.Count; i++)
            {
                body.Append(CollectionPages.CaptureSection(
                    siteBaseUrl,
                    collection,
                    collection.Images[i],
                    anchors[i],
                    previous: i > 0 ? anchors[i - 1] : null,
                    next: i < collection.Images.Count - 1 ? anchors[i + 1] : null,
                    view));
            }

            body.Append(CollectionPages.Viewers(siteBaseUrl, collection, anchors));

            string middle =
                $"<div class=\"meta-line1\">{Encode(CollectionPages.BoardLine(collection))}</div>" +
                $"<div class=\"meta-line2\"><span class=\"badge-red\">Will be deleted <b>{Encode(CollectionRules.DateText(collection.DeleteAtUtc))}</b></span></div>";

            string actions = collection.Images.Count == 0
                ? string.Empty
                : $"<a class=\"btn\" href=\"{Encode(CollectionPages.DownloadUrl(siteBaseUrl, collection.Id))}\">" +
                  $"Download {CollectionPages.Counted(collection.Images.Count, "capture")}</a>";

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

            return CollectionPages.Layout(name, body.ToString(), preview.ToString(), middle: middle, actions: actions, panelled: false);
        }

        // "Download all" - every capture in one ZIP (CollectionPageEndpoints.DownloadAsync).
        public static string DownloadUrl(string siteBaseUrl, string id) => $"{CollectionRules.PageUrl(siteBaseUrl, id)}/download";

        // ###########################################################################################
        // THE GRID OF CAPTURES (the gallery's cards): each picture opens the viewer, its name goes to its section below. Three or
        // fewer get larger cards, one a single wide card, as the gallery's few do.
        // ###########################################################################################
        private static string Gallery(string siteBaseUrl, CollectionRecord collection, IReadOnlyList<string> anchors, CollectionPageView view)
        {
            if (collection.Images.Count == 0)
                return "<section class=\"panel\"><p class=\"muted\">No captures.</p></section>";

            string grid = collection.Images.Count switch
            {
                1 => "grid grid-few grid-single",
                <= 3 => "grid grid-few",
                _ => "grid"
            };

            var html = new StringBuilder($"<div class=\"{grid}\">");

            for (int i = 0; i < collection.Images.Count; i++)
            {
                CollectionImageEntry image = collection.Images[i].Entry;
                string where = CollectionRules.WhereOf(image);

                // The very comments the capture's own section lists for this reader - the owner's and the administrator's count
                // hidden ones too, as that section does (code review, 2026-11-27: the two counts differed).
                int comments = view.Comments.Count(comment => comment.ImageSha256 == image.Sha256);

                html.Append("<figure class=\"card\">")
                    .Append($"<a class=\"thumb\" href=\"#{CollectionPages.ViewerId(anchors[i])}\" aria-label=\"{Encode($"Open {where} large")}\">")
                    .Append($"<img class=\"thumb-img\" src=\"{Encode(CollectionPages.ImageUrl(siteBaseUrl, collection.Id, image.Sha256))}\" ")
                    .Append($"width=\"{collection.Images[i].Width}\" height=\"{collection.Images[i].Height}\" loading=\"lazy\" alt=\"{Encode($"Oscilloscope capture of {where}")}\">")
                    .Append("</a><figcaption class=\"caption\">")
                    .Append($"<a class=\"cap-title\" href=\"#{anchors[i]}\">{Encode(where)}</a>")
                    .Append("<div class=\"cap-sub\">");

                var sub = new List<string>();

                if (image.SignalName is { } signal)
                    sub.Add(Encode(signal));

                if (comments > 0)
                    sub.Add(CollectionPages.Counted(comments, "comment"));

                html.Append(sub.Count == 0 ? "&nbsp;" : string.Join(" | ", sub))
                    .Append("</div></figcaption></figure>");
            }

            return html.Append("</div>").ToString();
        }

        // ###########################################################################################
        // THE VIEWER - one per capture - opened by its id in the address (#view.U7-12), so a link to it opens the page with the
        // capture large; closed by any other ("#viewer.closed", no element's, so the page does not move). The ids hold a ".",
        // which a capture's anchor never does (letters, digits and "-" - CollectionRules.AnchorsOf), so none can be taken for
        // another. Previous and next go round, as the gallery's did - by LARGE AREAS of the picture (owner requests, 2026-11-27:
        // "click in a big part of the right-side for right-navigation and also big part of the left-side ... Same also for
        // up/down"; asked, "Top and bottom areas too"): its left and its top go back, its right and its bottom on, the middle
        // nothing - and by the arrow keys and the wheel (Script). Beside the picture what CRT's component info window said of its
        // component (InfoOf), and under it "Link", which offers the viewer's own address ready for Lemon64, Forum64 or Reddit.
        // ###########################################################################################
        public const string ViewerClosed = "viewer.closed";

        public static string ViewerId(string anchor) => $"view.{anchor}";

        // A capture's own address - the page with it open in the viewer.
        public static string ViewerUrl(string siteBaseUrl, string id, string anchor) =>
            $"{CollectionRules.PageUrl(siteBaseUrl, id)}#{CollectionPages.ViewerId(anchor)}";

        // A capture saved from the viewer, or in "Download all"'s ZIP: its place on the page and its anchor - "001-U7-12.png".
        public static string CaptureFileName(int index, string anchor) =>
            $"{(index + 1).ToString("D3", System.Globalization.CultureInfo.InvariantCulture)}-{anchor}.png";

        private static string Viewers(string siteBaseUrl, CollectionRecord collection, IReadOnlyList<string> anchors)
        {
            var html = new StringBuilder();
            int count = collection.Images.Count;

            for (int i = 0; i < count; i++)
            {
                CollectionImageEntry image = collection.Images[i].Entry;
                string where = CollectionRules.WhereOf(image) + (image.SignalName is { } signal ? $" - {signal}" : string.Empty);
                string picture = CollectionPages.ImageUrl(siteBaseUrl, collection.Id, image.Sha256);
                string? previous = count > 1 ? CollectionPages.ViewerId(anchors[(i - 1 + count) % count]) : null;
                string? next = count > 1 ? CollectionPages.ViewerId(anchors[(i + 1) % count]) : null;
                string info = CollectionPages.InfoOf(siteBaseUrl, collection, image);

                html.Append($"<div class=\"modal\" id=\"{Encode(CollectionPages.ViewerId(anchors[i]))}\" role=\"dialog\" aria-modal=\"true\" aria-label=\"{Encode(where)}\">")
                    .Append($"<a class=\"backdrop\" href=\"#{CollectionPages.ViewerClosed}\" aria-label=\"Close\"></a>")
                    .Append("<div class=\"dialog\">")
                    .Append($"<a class=\"close\" href=\"#{CollectionPages.ViewerClosed}\" title=\"Close (Esc)\">&#215;</a>")
                    .Append(info.Length == 0 ? "<div class=\"vbody wide\">" : "<div class=\"vbody\">")
                    .Append("<div class=\"viewer\">")
                    .Append($"<div class=\"stage\"><img src=\"{Encode(picture)}\" loading=\"lazy\" alt=\"{Encode($"Oscilloscope capture of {CollectionRules.WhereOf(image)}")}\"></div>");

                if (previous is not null && next is not null)
                {
                    html.Append($"<a class=\"zone zone-left prev\" href=\"#{Encode(previous)}\" title=\"Previous (left arrow)\"><span>&#8249;</span></a>")
                        .Append($"<a class=\"zone zone-right next\" href=\"#{Encode(next)}\" title=\"Next (right arrow)\"><span>&#8250;</span></a>")
                        .Append($"<a class=\"zone zone-top prev\" href=\"#{Encode(previous)}\" title=\"Previous (up arrow)\" tabindex=\"-1\"></a>")
                        .Append($"<a class=\"zone zone-bottom next\" href=\"#{Encode(next)}\" title=\"Next (down arrow)\" tabindex=\"-1\"></a>");
                }

                html.Append("</div>")
                    .Append(info)
                    .Append("</div><div class=\"footer\"><div class=\"title\">")
                    .Append($"<div class=\"m-name\">{Encode(where)}</div>")
                    .Append($"<div class=\"m-time\">{Encode(CollectionPages.SettingsText(image.TimePerDivision, image.VoltsPerDivision, image.TriggerLevel))}</div></div>")
                    .Append("<div class=\"actions\">")
                    .Append(CollectionPages.LinkPanel(siteBaseUrl, collection, image, anchors[i]))
                    .Append($"<a class=\"btn\" href=\"{Encode(picture)}\" download=\"{Encode(CollectionPages.CaptureFileName(i, anchors[i]))}\">Download</a>")
                    .Append($"<span class=\"counter\">{i + 1} / {count}</span>")
                    .Append("</div></div></div></div>");
            }

            return html.ToString();
        }

        // ###########################################################################################
        // "LINK" (owner request, 2026-11-27: "Every image (when viewing it large) should have a link, and clicking that link will
        // expand an option to copy a link to either Forum64, Lemon64 or Reddit. The link should then navigate to this exact image,
        // and show it large"): a <details> - opened with no script - holding the viewer's address and a "Copy" for each forum's
        // way of writing a link (Script copies the button's data-copy). The words are the capture's place and the board.
        // ###########################################################################################
        public static string CaptureLinkText(CollectionRecord collection, CollectionImageEntry image) =>
            CollectionRules.WhereOf(image) + (image.SignalName is { } signal ? $" - {signal}" : string.Empty) +
            $" ({collection.HardwareName} {collection.BoardName})";

        // Lemon64's phpBB, Forum64's WoltLab and Reddit's Markdown each write a link their own way.
        public static string ForLemon64(string url, string text) => $"[url={url}]{text}[/url]";

        public static string ForForum64(string url, string text) => $"[url='{url}']{text}[/url]";

        public static string ForReddit(string url, string text) =>
            $"[{text.Replace("[", "(", StringComparison.Ordinal).Replace("]", ")", StringComparison.Ordinal)}]({url})";

        private static string LinkPanel(string siteBaseUrl, CollectionRecord collection, CollectionImageEntry image, string anchor)
        {
            string url = CollectionPages.ViewerUrl(siteBaseUrl, collection.Id, anchor);
            string text = CollectionPages.CaptureLinkText(collection, image);

            return "<details class=\"share\"><summary class=\"btn\">Link</summary><div class=\"share-panel\">" +
                   "<label>A link to this capture, shown large:</label>" +
                   $"<input class=\"share-url\" type=\"text\" readonly value=\"{Encode(url)}\" aria-label=\"The capture's link\">" +
                   $"<button type=\"button\" class=\"btn\" data-copy=\"{Encode(CollectionPages.ForLemon64(url, text))}\">Copy for Lemon64</button>" +
                   $"<button type=\"button\" class=\"btn\" data-copy=\"{Encode(CollectionPages.ForForum64(url, text))}\">Copy for Forum64</button>" +
                   $"<button type=\"button\" class=\"btn\" data-copy=\"{Encode(CollectionPages.ForReddit(url, text))}\">Copy for Reddit</button>" +
                   "</div></details>";
        }

        // ###########################################################################################
        // WHAT CRT'S COMPONENT INFO WINDOW SAID OF THE CAPTURE'S COMPONENT (owner request, 2026-11-27: "show something like the
        // component info window from the CRT application ... to get some context for this specific component and pin"; asked,
        // "Also the drawn pinout"): its title, category line, description and purpose, CRT's own drawing of its pinout with the
        // capture's pin marked (a picture CRT uploaded - CollectionImageEntry.PinoutSha256), and its pins with this capture's
        // highlighted. Every word as CRT sent it, encoded. Empty when CRT sent nothing of it.
        // ###########################################################################################
        private static string InfoOf(string siteBaseUrl, CollectionRecord collection, CollectionImageEntry image)
        {
            CollectionComponent? component = collection.ComponentOf(image);

            if (component is null && image.PinoutSha256 is null)
                return string.Empty;

            var html = new StringBuilder("<aside class=\"info\">");

            html.Append($"<h3>{Encode(component?.Title ?? image.ComponentLabel)}</h3>");

            if (component?.CategoryLine is { } category)
                html.Append($"<p class=\"muted\">{Encode(category)}</p>");

            if (component?.Description is { } description)
                html.Append($"<p>{Encode(description)}</p>");

            if (component?.Purpose is { } purpose)
                html.Append($"<p class=\"text\">{Encode(purpose)}</p>");

            if (image.PinoutSha256 is { } pinout)
            {
                html.Append($"<img class=\"pinout\" src=\"{Encode(CollectionPages.ImageUrl(siteBaseUrl, collection.Id, pinout))}\" loading=\"lazy\" ")
                    .Append($"alt=\"{Encode($"The pinout of {image.ComponentLabel}" + (image.Pin is { } marked ? $", pin {marked} marked" : string.Empty))}\">");
            }

            IReadOnlyList<CollectionPin> pins = component?.Pins ?? [];

            if (pins.Count > 0)
            {
                bool nets = pins.Any(pin => pin.Net is not null);
                bool readings = pins.Any(pin => pin.Reading is not null);

                html.Append("<table class=\"pins\"><thead><tr><th>Pin</th><th>Name</th><th>Direction</th>")
                    .Append(nets ? "<th>Net</th>" : string.Empty)
                    .Append(readings ? "<th>Expected reading</th>" : string.Empty)
                    .Append("</tr></thead><tbody>");

                foreach (CollectionPin pin in pins)
                {
                    bool here = image.Pin is { } pinOfCapture && string.Equals(pin.Pin.Trim(), pinOfCapture.Trim(), StringComparison.OrdinalIgnoreCase);

                    html.Append(here ? "<tr class=\"here\">" : "<tr>")
                        .Append($"<td>{Encode(pin.Pin)}</td><td>{Encode(pin.Name)}</td><td>{Encode(pin.Direction)}</td>")
                        .Append(nets ? $"<td>{Encode(pin.Net)}</td>" : string.Empty)
                        .Append(readings ? $"<td class=\"reading\">{Encode(pin.Reading)}</td>" : string.Empty)
                        .Append("</tr>");
                }

                html.Append("</tbody></table>");
            }

            return html.Append("</aside>").ToString();
        }

        // "T/DIV 50us | V/DIV 2V | Trigger level 1V" - what was read back, the rest left out.
        private static string SettingsText(string? timePerDivision, string? voltsPerDivision, string? triggerLevel)
        {
            var parts = new List<string>();

            if (timePerDivision is not null)
                parts.Add($"T/DIV {timePerDivision}");

            if (voltsPerDivision is not null)
                parts.Add($"V/DIV {voltsPerDivision}");

            if (triggerLevel is not null)
                parts.Add($"Trigger level {triggerLevel}");

            return string.Join(" | ", parts);
        }

        // ###########################################################################################
        // *** A COLLECTION DELETED KEEPS ITS PAGE FOR GOOD (owner request, 2026-11-27: "replaced by a standard web page, explaining
        // the deletion period has passed etc. ... and do keep this page forever"). *** Its row stays in the database - only its
        // words, captures and comments go (ICollectionStore.MarkDeletedAsync) - so the link a forum post holds answers years
        // later: what the page was, when and why it went, that it went on purpose, and what CRT is. It names nobody, as the
        // page did not.
        // ###########################################################################################
        public const string DeletedTitle = "This collection has been deleted";

        public const string DeletedWhatItWas =
            "This page showed an oscilloscope collection: captures of a vintage computer's signals, taken with Classic Repair " +
            "Toolbox and shared so that others could help read them.";

        public const string DeletedForGood =
            "This page stays here for good, so anybody following a link to it knows the collection went on purpose - and if you came " +
            "from a forum post, the discussion there may still say what was found.";

        public static string DeletedWhy(CollectionRecord collection)
        {
            string when = CollectionRules.DateText(collection.DeletedUtc ?? collection.DeleteAtUtc);

            return collection.DeletedReason switch
            {
                CollectionDeletedReasons.Expired =>
                    $"It was deleted on {when}, as the time it was shared for had passed. Whoever shares a collection chooses how long " +
                    "it stays online - a month, six months or a year from the last time it was shared - and when that time is up, " +
                    "its captures and comments go.",
                CollectionDeletedReasons.Owner =>
                    $"Its owner deleted it on {when}, and its captures and comments went with it.",
                CollectionDeletedReasons.Administrator =>
                    $"It was taken down by the site's administrator on {when}, and its captures and comments went with it.",
                _ => $"It was deleted on {when}."
            };
        }

        public static string Deleted(CollectionRecord collection) =>
            CollectionPages.Layout(
                CollectionPages.DeletedTitle,
                $"<p class=\"lead\">{Encode(CollectionPages.DeletedWhatItWas)}</p>" +
                $"<p>{Encode(CollectionPages.DeletedWhy(collection))}</p>" +
                $"<p>{Encode(CollectionPages.DeletedForGood)}</p>" +
                $"<p class=\"facts\">Collection <code>{Encode(collection.Id)}</code></p>" +
                CollectionPages.WhatCrtIs(),
                middle: "<div class=\"meta-line2\"><span class=\"badge-red badge-large\">Deleted " +
                        $"<b>{Encode(CollectionRules.DateText(collection.DeletedUtc ?? collection.DeleteAtUtc))}</b></span></div>");

        // ###########################################################################################
        // ITS OWNER STOPPED SHARING IT (migration 0028) - what everybody but its owner and the administrator sees: nothing of it,
        // and that the page may come back under this same link.
        // ###########################################################################################
        public const string SharingStoppedTitle = "This collection is not shared at the moment";

        public const string SharingStoppedText =
            "Its owner has stopped sharing it for now, so its captures and comments are not shown. If they share it again, it comes " +
            "back under this same link - so try it again later.";

        public static string SharingStopped(CollectionRecord collection) =>
            CollectionPages.Layout(
                CollectionPages.SharingStoppedTitle,
                $"<p class=\"lead\">{Encode(CollectionPages.SharingStoppedText)}</p>" +
                $"<p class=\"facts\">Collection <code>{Encode(collection.Id)}</code></p>" +
                CollectionPages.WhatCrtIs(),
                middle: "<div class=\"meta-line2\"><span class=\"badge-grey badge-large\">Not shared at the moment</span></div>");

        // What Classic Repair Toolbox is, for a reader who knows nothing of it.
        private static string WhatCrtIs() =>
            "<h2>What is Classic Repair Toolbox?</h2>" +
            "<p>A free program that helps hobbyists repair vintage computers - schematics, component data and oscilloscope " +
            "baselines for Commodore, Amstrad and ZX Spectrum boards, and the captures of your own oscilloscope beside them. " +
            $"<a href=\"{Encode(EmailTemplates.CrtDownloadUrl)}\">Download it</a>.</p>";

        // A collection offline for everybody else, as the administrator sees it - nobody can comment on it, and its comments are
        // still hidden or deleted here - or, its sharing stopped, as its owner and the administrator see it.
        public const string OfflineOwnerLocked =
            "Offline: its owner's account is locked, so only you see this page. Nobody can comment on it until the account is unlocked.";

        public const string OfflineOwnerGone =
            "Offline: its owner has no account any more, so only you see this page. Nobody can comment on it.";

        public const string OfflineSharingStoppedOwner =
            "Not shared: you stopped sharing it, so only you and the administrator see this page. \"Share\" in CRT brings it back " +
            "under this same link.";

        public const string OfflineSharingStoppedAdministrator =
            "Not shared: its owner stopped sharing it, so only its owner and you see this page. Nobody can comment on it meanwhile.";

        private static string OfflineLine(CollectionOffline offline, bool ownerReads) => offline switch
        {
            CollectionOffline.OwnerGone => CollectionPages.OfflineOwnerGone,
            CollectionOffline.SharingStopped => ownerReads ? CollectionPages.OfflineSharingStoppedOwner : CollectionPages.OfflineSharingStoppedAdministrator,
            _ => CollectionPages.OfflineOwnerLocked
        };

        public static string NotFound() =>
            CollectionPages.Layout(
                "There is no such collection",
                "<p>Check that the whole link was copied - a collection's address ends in its 12-character id.</p>");

        public static string NotSetUp() =>
            CollectionPages.Layout("Collections are not available", $"<p>{Encode(CollectionFlows.NotSetUp)}</p>");

        // ###########################################################################################
        // "YOUR COLLECTIONS" (step 4, 2026-11-26; CollectionListFlows says what each list holds): the reader's own and those they
        // commented on, and for the administrator every collection with whose it is and the newest comments, each a link to
        // the comment on its page - where it is hidden or deleted. The one page naming owners, so the administrator's alone.
        // ###########################################################################################
        public const string ListTitle = "Your collections";

        public const string NoneSharedYet =
            "None yet - captures are collected in CRT's component info window, under \"Collect captures into\", and shared from there.";

        public const string OwnerLockedFlag = "Owner locked - offline";

        public const string OwnerGoneFlag = "No owner account - offline";

        public const string NotSharedFlag = "Not shared - offline";

        // ###########################################################################################
        // `locking` is the account the administrator asked to lock (?lock=<id>) - asked first, as a page of its own state - and
        // `problem` what a lock or unlock refused said.
        // ###########################################################################################
        public static string List(string siteBaseUrl, CollectionViewer viewer, string? csrf, CollectionListView list, LockedAccount? locking = null, string? problem = null)
        {
            var body = new StringBuilder();
            bool administrator = viewer.IsAdministrator;

            body.Append(CollectionPages.SignedInLine(siteBaseUrl, viewer, csrf, CollectionPages.ListPath(siteBaseUrl), onList: true));

            if (administrator && problem is not null)
                body.Append($"<div class=\"errors\" role=\"alert\"><p>{Encode(problem)}</p></div>");

            if (administrator && locking is not null)
                body.Append(CollectionPages.LockQuestion(siteBaseUrl, csrf, locking));

            body.Append($"<h2>Shared by you ({list.Own.Count})</h2>");
            body.Append(list.Own.Count == 0
                ? $"<p class=\"muted\">{Encode(CollectionPages.NoneSharedYet)}</p>"
                : CollectionPages.Rows(siteBaseUrl, list.Own.Select(collection => new ListedRow(collection))));

            body.Append($"<h2>You commented on ({list.CommentedOn.Count})</h2>");
            body.Append(list.CommentedOn.Count == 0
                ? "<p class=\"muted\">None yet.</p>"
                : CollectionPages.Rows(siteBaseUrl, list.CommentedOn.Select(collection => new ListedRow(collection))));

            if (administrator && list.All is { } all)
            {
                body.Append($"<h2>Every collection ({all.Count}) <span class=\"badge\">{CollectionComments.AdministratorBadge}</span></h2>");

                if (all.Count >= CollectionListFlows.AllShownAtMost)
                    body.Append($"<p class=\"muted\">The newest {CollectionListFlows.AllShownAtMost} are listed.</p>");

                body.Append(all.Count == 0
                    ? "<p class=\"muted\">None.</p>"
                    : CollectionPages.Rows(siteBaseUrl, all.Select(listed => new ListedRow(
                        listed.Collection,
                        CollectionPages.OwnerText(listed),
                        listed.OwnerGone ? CollectionPages.OwnerGoneFlag : listed.OwnerLocked ? CollectionPages.OwnerLockedFlag : null,
                        listed.IsOffline || listed.OwnerIsAdministrator ? null : CollectionPages.LockLink(siteBaseUrl, listed.Collection.AccountId)))));
            }

            if (administrator && list.NewestComments is { } comments)
            {
                Dictionary<string, CollectionRecord> named = list.Own
                    .Concat(list.CommentedOn)
                    .Concat(list.All?.Select(listed => listed.Collection) ?? [])
                    .DistinctBy(collection => collection.Id)
                    .ToDictionary(collection => collection.Id, StringComparer.Ordinal);

                body.Append($"<h2>Newest comments ({comments.Count}) <span class=\"badge\">{CollectionComments.AdministratorBadge}</span></h2>");
                body.Append(comments.Count == 0 ? "<p class=\"muted\">None.</p>" : CollectionPages.NewestComments(siteBaseUrl, comments, named));
            }

            if (administrator && list.Locked is { } locked)
            {
                body.Append($"<h2 id=\"locked\">Locked accounts ({locked.Count}) <span class=\"badge\">{CollectionComments.AdministratorBadge}</span></h2>");

                if (locked.Count == 0)
                {
                    body.Append("<p class=\"muted\">None.</p>");
                }
                else
                {
                    body.Append("<ul class=\"collections\">");

                    foreach (LockedAccount account in locked)
                    {
                        body.Append($"<li><b>{Encode(account.Name)}</b> <span class=\"muted\">({Encode(account.Email)})</span> ")
                            .Append($"<form class=\"inline\" method=\"post\" action=\"{Encode(CollectionPages.AccountActionUrl(siteBaseUrl, account.Id, "unlock"))}\">")
                            .Append(CollectionPages.Hidden("csrf", csrf))
                            .Append("<button type=\"submit\" class=\"link\">Unlock</button></form></li>");
                    }

                    body.Append("</ul>");
                }
            }

            return CollectionPages.Layout(CollectionPages.ListTitle, $"<div class=\"list\">{body}</div>", kind: CollectionPages.ListKind);
        }

        // ###########################################################################################
        // LOCKING AN ACCOUNT (owner decision, 2026-11-27; AccountLockFlows): the administrator's, asked first, saying what it does.
        // ###########################################################################################
        public const string LockQuestionText =
            "It can no longer sign in, its collections go offline and its comments are hidden from everybody but you. " +
            "\"Unlock\" under \"Locked accounts\" puts them back; it then has to sign in again.";

        public const string AccountLockedFlag = "Account locked";

        // A hidden comment is shown to the collection's owner AND the administrator (CollectionCommentFlows.ShownTo) - said so,
        // as "only you" misled each about the other (code review, 2026-11-27).
        public const string HiddenFlag = "Hidden - only the owner and the administrator see it";

        private static string LockQuestion(string siteBaseUrl, string? csrf, LockedAccount account) =>
            $"<form class=\"admin\" method=\"post\" action=\"{Encode(CollectionPages.AccountActionUrl(siteBaseUrl, account.Id, "lock"))}\">" +
            $"<p>Lock the account of <b>{Encode(account.Name)}</b> ({Encode(account.Email)})? {Encode(CollectionPages.LockQuestionText)}</p>" +
            CollectionPages.Hidden("csrf", csrf) +
            "<p class=\"actions\"><button type=\"submit\" class=\"danger\">Lock account</button> " +
            $"<a href=\"{Encode(CollectionPages.ListUrl(siteBaseUrl))}\">Keep it</a></p></form>";

        // The link asking to lock an account, the question then at the top of the list.
        private static string LockLink(string siteBaseUrl, long accountId) =>
            $"<a href=\"{Encode($"{CollectionPages.ListUrl(siteBaseUrl)}?lock={accountId}")}\">Lock account</a>";

        // Where a lock or unlock posts: /collections/accounts/<id>/lock.
        public static string AccountActionUrl(string siteBaseUrl, long accountId, string action) =>
            $"{CollectionPages.ListUrl(siteBaseUrl)}accounts/{accountId}/{action}";

        private sealed record ListedRow(CollectionRecord Collection, string? Owner = null, string? Flag = null, string? LockLink = null);

        private const string ListKind = "Oscilloscope collections";

        // The list asked for signed out: what it is, and the way in back to it.
        public static string ListSignedOut(string siteBaseUrl)
        {
            string signIn = Accounts.AccountPages.Link(Accounts.AccountPages.BaseUrl(siteBaseUrl), "sign-in") +
                            "?return=" + Uri.EscapeDataString(CollectionPages.ListPath(siteBaseUrl));

            return CollectionPages.Layout(
                CollectionPages.ListTitle,
                $"<p><a href=\"{Encode(signIn)}\">Sign in</a> to see the collections you shared and those you commented on.</p>",
                kind: CollectionPages.ListKind);
        }

        // The list's address, and its path on the site as a sign-in returns to it ("/crt-test/collections/" on the test instance).
        public static string ListUrl(string siteBaseUrl) => $"{siteBaseUrl.TrimEnd('/')}{CollectionRoutes.PagePath}/";

        public static string ListPath(string siteBaseUrl) => $"{Accounts.WebSessions.PathOf(siteBaseUrl)}{CollectionRoutes.PagePath}/";

        private static string Rows(string siteBaseUrl, IEnumerable<ListedRow> rows)
        {
            var html = new StringBuilder("<ul class=\"collections\">");

            foreach ((CollectionRecord collection, string? owner, string? flag, string? lockLink) in rows)
            {
                html.Append($"<li><a href=\"{Encode(CollectionRules.PageUrl(siteBaseUrl, collection.Id))}\">{Encode(CollectionPages.NameOf(collection))}</a>");

                if (flag is not null)
                    html.Append($" <span class=\"flag\">{Encode(flag)}</span>");

                // Its owner stopped sharing it - offline to everybody but its owner and the administrator (migration 0028).
                if (collection.IsSharingStopped)
                    html.Append($" <span class=\"flag\">{CollectionPages.NotSharedFlag}</span>");

                html.Append("<br><span class=\"muted\">")
                    .Append(Encode(CollectionPages.BoardLine(collection)))
                    .Append($" | {CollectionPages.Counted(collection.ImageCount, "capture")} | {CollectionPages.Counted(collection.CommentCount, "comment")}")
                    .Append($" | Shared {Encode(CollectionRules.DateText(collection.CreatedUtc))}")
                    .Append($" | Deleted on {Encode(CollectionRules.DateText(collection.DeleteAtUtc))}");

                if (owner is not null)
                    html.Append($" | By {Encode(owner)}");

                if (lockLink is not null)
                    html.Append($" | {lockLink}");

                html.Append("</span></li>");
            }

            return html.Append("</ul>").ToString();
        }

        private static string OwnerText(CollectionListed listed) =>
            listed.OwnerEmail is { } email ? $"{listed.OwnerName} ({email})" : listed.OwnerName;

        private static string NewestComments(string siteBaseUrl, IReadOnlyList<CollectionCommentRecord> comments, IReadOnlyDictionary<string, CollectionRecord> named)
        {
            var html = new StringBuilder();

            foreach (CollectionCommentRecord comment in comments)
            {
                string on = named.TryGetValue(comment.CollectionId, out CollectionRecord? collection) ? CollectionPages.NameOf(collection) : comment.CollectionId;

                html.Append($"<article class=\"comment{(comment.IsHidden ? " hidden" : string.Empty)}\">")
                    .Append($"<p class=\"who\"><b>{Encode(comment.AuthorName)}</b>")
                    .Append($" <span class=\"muted\">{Encode(CollectionComments.WhenText(comment.CreatedUtc))}{(comment.EditedUtc is null ? string.Empty : " - edited")}</span>");

                if (comment.IsHidden)
                    html.Append(" <span class=\"flag\">Hidden</span>");

                if (comment.AuthorIsLocked)
                    html.Append($" <span class=\"flag\">{CollectionPages.AccountLockedFlag}</span>");

                html.Append("</p>")
                    .Append($"<p class=\"text\">{Encode(CollectionPages.Excerpt(comment.Body, 300))}</p>")
                    .Append($"<p class=\"actions\">On <a href=\"{Encode($"{CollectionRules.PageUrl(siteBaseUrl, comment.CollectionId)}#comment-{comment.Id}")}\">{Encode(on)}</a>");

                if (!comment.AuthorIsLocked && !comment.AuthorIsAdministrator)
                    html.Append($" | {CollectionPages.LockLink(siteBaseUrl, comment.AccountId)}");

                html.Append("</p></article>");
            }

            return html.ToString();
        }

        private static string Counted(int count, string what) => $"{count} {what}{(count == 1 ? string.Empty : "s")}";

        // At most `length` characters, never half of one written as two (a surrogate pair).
        private static string Excerpt(string text, int length)
        {
            if (text.Length <= length)
                return text;

            int cut = char.IsHighSurrogate(text[length - 1]) ? length - 1 : length;

            return text[..cut].TrimEnd() + "...";
        }

        // Who is signed in, the way to the list (but on the list), and the way out - the form carries its token (WebSessions.Csrf)
        // and where to come back to.
        private static string SignedInLine(string siteBaseUrl, CollectionViewer viewer, string? csrf, string returnPath, bool onList) =>
            "<form class=\"signedin\" method=\"post\" action=\"" +
            Encode(Accounts.AccountPages.Link(Accounts.AccountPages.BaseUrl(siteBaseUrl), "sign-out")) + "\">" +
            $"Signed in as <b>{Encode(viewer.DisplayName)}</b> | " +
            (onList ? string.Empty : $"<a href=\"{Encode(CollectionPages.ListUrl(siteBaseUrl))}\">{CollectionPages.ListTitle}</a> | ") +
            CollectionPages.Hidden("csrf", csrf) +
            CollectionPages.Hidden("return", returnPath) +
            "<button type=\"submit\" class=\"link\">Sign out</button></form>";

        // ###########################################################################################
        // THE ADMINISTRATOR TAKES A COLLECTION DOWN (step 4) - asked first, as a page of its own state (?remove=1), then a form
        // posting to this service. Its page then says it was removed; its captures and comments go. The owner deletes their own
        // in CRT, which keeps its window in step with the server - so the web offers it to the administrator alone.
        // ###########################################################################################
        public const string RemoveQuestion =
            "Delete this collection for everybody? Its page will then say it was removed, and its captures and comments go. Its owner is not mailed.";

        private static string RemoveCollection(string siteBaseUrl, CollectionRecord collection, CollectionPageView view)
        {
            string page = CollectionRules.PageUrl(siteBaseUrl, collection.Id);

            if (!view.RemovingCollection)
            {
                return $"<p class=\"admin\"><span class=\"badge\">{CollectionComments.AdministratorBadge}</span> " +
                       $"<a href=\"{Encode(page + "?remove=1")}\">Delete this collection</a></p>";
            }

            return $"<form class=\"admin\" method=\"post\" action=\"{Encode(page + "/delete")}\">" +
                   $"<p>{Encode(CollectionPages.RemoveQuestion)}</p>" +
                   CollectionPages.Hidden("csrf", view.Csrf) +
                   "<p class=\"actions\"><button type=\"submit\" class=\"danger\">Delete collection</button> " +
                   $"<a href=\"{Encode(page)}\">Keep it</a></p></form>";
        }

        // ###########################################################################################
        // One capture: the picture, its settings, what it should read, and the owner's note.
        // ###########################################################################################
        private static string CaptureSection(
            string siteBaseUrl,
            CollectionRecord collection,
            CollectionImageRecord capture,
            string anchor,
            string? previous,
            string? next,
            CollectionPageView view)
        {
            CollectionImageEntry image = capture.Entry;
            string where = CollectionRules.WhereOf(image);
            var html = new StringBuilder();

            html.Append($"<section class=\"panel capture\" id=\"{anchor}\">");
            html.Append($"<h2>{Encode(where)}");

            if (image.SignalName is { } signal)
                html.Append($" <span class=\"signal\">{Encode(signal)}</span>");

            // The picture opens the viewer (Viewers), as the grid's cards do.
            html.Append("</h2><div class=\"pair\">");
            html.Append("<figure>")
                .Append($"<a class=\"framed\" href=\"#{CollectionPages.ViewerId(anchor)}\">")
                .Append($"<img src=\"{Encode(CollectionPages.ImageUrl(siteBaseUrl, collection.Id, image.Sha256))}\" width=\"{capture.Width}\" height=\"{capture.Height}\" loading=\"lazy\" alt=\"{Encode($"Oscilloscope capture of {where}")}\">")
                .Append("</a></figure>");

            html.Append("</div>");

            html.Append("<table class=\"settings\"><tbody>");
            html.Append(CollectionPages.SettingRow("T/DIV", image.TimePerDivision));
            html.Append(CollectionPages.SettingRow("V/DIV", image.VoltsPerDivision));
            html.Append(CollectionPages.SettingRow("Trigger level", image.TriggerLevel));
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
                    html.Append($" <span class=\"flag\">{CollectionPages.HiddenFlag}</span>");

                // Shown to the administrator alone (CollectionCommentFlows.ShownTo).
                if (comment.AuthorIsLocked)
                    html.Append($" <span class=\"flag\">{CollectionPages.AccountLockedFlag} - only you see it</span>");

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

                    // An offline collection's comments are not edited - CollectionCommentFlows.EditAsync asks FindShownAsync.
                    if (mine && view.Offline == CollectionOffline.No)
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

            // Nobody comments on a collection offline (CollectionCommentFlows.AddAsync), so the administrator, the one reader of
            // its page, gets no form to be refused by.
            if (viewer is not null && view.Offline == CollectionOffline.No)
            {
                html.Append($"<form method=\"post\" action=\"{Encode($"{page}/comments")}\">")
                    .Append(CollectionPages.Hidden("csrf", view.Csrf))
                    .Append(CollectionPages.Hidden("image", sha256))
                    .Append($"<label for=\"comment-on-{anchor}\">Your comment</label>")
                    .Append(CollectionPages.TextArea(view.ErrorOn == sha256 ? view.Typed : null, $"comment-on-{anchor}"))
                    .Append(CollectionPages.ErrorsFor(view, sha256))
                    .Append("<p class=\"actions\"><button type=\"submit\">Comment</button></p></form>");
            }
            else if (viewer is null)
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

        // A setting of the capture, as read back from the oscilloscope.
        private static string SettingRow(string name, string? value) =>
            $"<tr><th>{Encode(name)}</th><td>{(value is null ? "<span class=\"muted\">not read</span>" : Encode(value))}</td></tr>";

        // A forum link as shown: its host and path, no scheme.
        private static string LinkText(string link) =>
            Uri.TryCreate(link, UriKind.Absolute, out Uri? uri) ? (uri.Host + uri.PathAndQuery).TrimEnd('/') : link;

        // ###########################################################################################
        // Every page of collections: the frosted bar - CRT on the left, `middle` (HTML) in the middle, `actions` (HTML) on the
        // right - then the page. `panelled`: the title and `body` in one panel; else `body` is the page's own panels.
        // ###########################################################################################
        private static string Layout(
            string title,
            string body,
            string head = "",
            string kind = "Oscilloscope collection",
            string middle = "",
            string actions = "",
            bool panelled = true)
        {
            return "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\">" +
                "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
                "<meta name=\"robots\" content=\"noindex, nofollow\">" +
                $"<title>{Encode(title)} - Classic Repair Toolbox</title>" +
                head +
                $"<style>{CollectionPages.Css}</style></head><body id=\"top\">" +
                "<header class=\"topbar\">" +
                // CRT's project page (owner request, 2026-11-27), not the site's front page.
                $"<div class=\"brand\"><a class=\"brand-title\" href=\"{Encode(EmailTemplates.CrtProjectUrl)}\">Classic Repair Toolbox</a>" +
                $"<div class=\"brand-sub\">{Encode(kind)}</div></div>" +
                $"<div class=\"meta\">{middle}</div>" +
                $"<div class=\"header-actions\">{actions}</div>" +
                "</header>" +
                "<main class=\"container\">" +
                (panelled ? $"<section class=\"panel\"><h1>{Encode(title)}</h1>{body}</section>" : body) +
                "</main>" +
                "<footer><p>Shared with Classic Repair Toolbox, which helps hobbyists repair vintage computers. " +
                $"<a href=\"{Encode(EmailTemplates.CrtDownloadUrl)}\">Download it</a> - it is free.</p>" +
                $"<p>Something on this page that should not be here? Mail <a href=\"mailto:{Encode(EmailTemplates.ContactAddress)}\">" +
                $"{Encode(EmailTemplates.ContactAddress)}</a> with its address.</p></footer>" +
                $"<script>{CollectionPages.Script}</script>" +
                "</body></html>";
        }

        private static string Encode(string? text) => WebUtility.HtmlEncode(text ?? string.Empty);

        // ###########################################################################################
        // The page's one style block - the owner's capture gallery's look (Page says it): a dark slate page lit from its top
        // left, a frosted bar, light rounded cards that lift under the pointer, the viewer. Its hash goes into the
        // Content-Security-Policy, which allows no other style. A phone gets the bar stacked and not fixed.
        // ###########################################################################################
        public const string Css =
            ":root{--bg0:#eef3f9;--bg1:#303841;--card:#e3e6ea;--card2:#f4f6f8;--text:#0f172a;--muted:#475467;" +
            "--border:rgba(15,23,42,.12);--border2:rgba(15,23,42,.18);--link:#075985;--red:#bc5454;--go:#497f54;--diff:#fff3cd;" +
            "--shadow:0 12px 30px rgba(15,23,42,.18);--radius:14px;--topbar-h:74px}" +
            "*{box-sizing:border-box}html{scroll-padding-top:calc(var(--topbar-h) + 16px)}" +
            "body{margin:0;min-height:100vh;color:var(--text);font:16px/1.5 ui-sans-serif,system-ui,-apple-system,\"Segoe UI\",Roboto,Arial,sans-serif;" +
            "background:radial-gradient(900px 480px at 20% 0%,#dfeaf9 0%,var(--bg1) 55%),linear-gradient(180deg,var(--bg1),var(--bg0));background-attachment:fixed}" +
            "body:has(.modal:target){overflow:hidden}a{color:var(--link)}" +
            ".topbar{position:fixed;top:0;left:0;right:0;z-index:50;height:var(--topbar-h);padding:0 20px;display:grid;grid-template-columns:1fr auto 1fr;" +
            "align-items:center;gap:14px;background:rgba(255,255,255,.72);border-bottom:1px solid var(--border);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px)}" +
            ".brand{grid-column:1;justify-self:start;display:flex;flex-direction:column;gap:2px;min-width:0}" +
            ".brand-title{font-size:20px;font-weight:750;letter-spacing:.2px;line-height:1.15;white-space:nowrap;color:var(--text);text-decoration:none}" +
            ".brand-title:hover{text-decoration:underline}.brand-sub{font-size:14px;color:var(--muted);line-height:1.15;white-space:nowrap}" +
            ".meta{grid-column:2;justify-self:center;display:flex;flex-direction:column;align-items:center;gap:4px;text-align:center;white-space:nowrap;font-size:12px;color:var(--muted)}" +
            ".meta-line1{font-size:18px;color:rgba(15,23,42,.78)}.meta-line2{font-size:14px}.header-actions{grid-column:3;justify-self:end;display:flex;align-items:center;gap:10px}" +
            ".badge-red,.badge-grey{display:inline-block;padding:.35em .65em;font-size:.85em;line-height:1.1;border-radius:.375rem;color:#fff;" +
            "border:1px solid rgba(0,0,0,.08);box-shadow:0 1px 2px rgba(0,0,0,.1)}.badge-red{background:var(--red)}.badge-grey{background:#5b6472}" +
            ".badge-large{font-size:1.3em;padding:.45em .8em}" +
            ".btn{display:inline-flex;align-items:center;justify-content:center;text-decoration:none;border:1px solid rgba(2,132,199,.25);" +
            "background:rgba(2,132,199,.12);color:var(--link);padding:8px 12px;border-radius:10px;font-weight:650;white-space:nowrap}" +
            ".btn:hover{border-color:rgba(2,132,199,.38);background:rgba(2,132,199,.16)}" +
            ".container{max-width:1200px;margin:0 auto;padding:calc(var(--topbar-h) + 18px) 18px 10px}" +
            ".panel{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);padding:18px 22px;margin:0 0 18px}" +
            "h1{font-size:1.5em;margin:0 0 6px;line-height:1.25}h2{font-size:1.15em;margin:0 0 10px}h3{font-size:1em;margin:0 0 8px}p{margin:0 0 12px}" +
            ".lead{font-size:1.1em}.facts{color:var(--muted);font-size:.9em}.muted{color:var(--muted)}" +
            "code{font-size:.95em;background:rgba(2,132,199,.08);border:1px solid rgba(2,132,199,.18);padding:1px 6px;border-radius:999px;color:var(--link)}" +
            ".text{white-space:pre-wrap;overflow-wrap:anywhere}.discussion{overflow-wrap:anywhere}" +
            ".grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:16px;margin:0 0 18px}" +
            ".grid.grid-few{grid-template-columns:repeat(auto-fit,minmax(320px,1fr))}" +
            ".grid.grid-single{justify-content:center}.grid.grid-single>.card{grid-column:1/-1;max-width:640px;width:100%;justify-self:center}" +
            ".card{margin:0;background:var(--card);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow);" +
            "transition:transform 140ms ease,box-shadow 140ms ease,border-color 140ms ease}" +
            ".card:hover{transform:translateY(-2px);box-shadow:0 16px 38px rgba(15,23,42,.26);border-color:var(--border2)}" +
            ".thumb{display:block;padding:12px 12px 0}.framed{display:block}" +
            ".thumb-img,.framed img{display:block;width:100%;height:auto;aspect-ratio:5/3;object-fit:contain;background:#0b1220;border-radius:12px;" +
            "border:1px solid rgba(15,23,42,.18);box-shadow:inset 0 0 0 1px rgba(255,255,255,.35),0 8px 18px rgba(15,23,42,.18)}" +
            ".caption{padding:10px 12px 14px;font-size:13px;color:var(--text);border-top:1px solid var(--border);margin-top:12px}" +
            ".cap-title{display:block;font-weight:650;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--text);text-decoration:none}" +
            ".cap-title:hover{text-decoration:underline}" +
            ".cap-sub{display:block;margin-top:4px;font-size:12px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}" +
            ".signal{color:var(--muted);font-weight:normal}" +
            ".flag{display:inline-block;font-size:.75em;font-weight:600;color:var(--red);border:1px solid var(--red);border-radius:10px;padding:0 8px;vertical-align:middle}" +
            ".badge{display:inline-block;font-size:.75em;font-weight:600;color:#fff;background:var(--link);border-radius:10px;padding:0 8px;vertical-align:middle}" +
            "figure{margin:0}.pair{display:flex;flex-wrap:wrap;gap:16px;margin-bottom:12px}.pair figure{flex:1 1 400px;max-width:100%}" +
            ".settings{border-collapse:collapse;margin-bottom:12px}.settings th,.settings td{text-align:left;padding:4px 14px 4px 0}" +
            ".steps{font-size:.9em;color:var(--muted);margin:0}.signedin{font-size:.9em;color:var(--muted);margin:0 0 12px}" +
            ".comments{margin:8px 0 12px;padding:12px 16px;border:1px solid var(--border);border-radius:10px;background:var(--card2)}" +
            ".comment{border-top:1px solid var(--border);padding:8px 0}.comment:first-of-type{border-top:0}" +
            ".comment.hidden{opacity:.65}.who{margin:0 0 4px}.comment .text{margin:0 0 4px}.actions{font-size:.9em;margin:4px 0 0}" +
            "label{display:block;font-weight:600;margin:8px 0 4px}" +
            "textarea{width:100%;padding:8px 10px;font:inherit;color:var(--text);background:#fff;border:1px solid var(--border2);border-radius:8px}" +
            "button{padding:6px 14px;font:inherit;color:#fff;background:var(--go);border:0;border-radius:8px;cursor:pointer}" +
            "button.danger{background:var(--red)}button.link{padding:0;color:var(--link);background:none;text-decoration:underline;border-radius:0}" +
            "form.inline{display:inline}" +
            ".errors{border:1px solid var(--red);color:#8f2f2f;background:rgba(188,84,84,.08);border-radius:8px;padding:8px 12px;margin:8px 0}.errors p{margin:0}" +
            ".list h2{margin-top:24px}.collections{list-style:none;padding:0;margin:0 0 12px}" +
            ".collections li{padding:8px 0;border-top:1px solid var(--border);overflow-wrap:anywhere}.collections li:first-child{border-top:0}" +
            ".admin{margin:0 0 12px;padding:8px 12px;border:1px dashed var(--border2);border-radius:8px}.admin p{margin:0 0 6px}" +
            "footer{max-width:1164px;width:calc(100% - 36px);margin:0 auto 28px;padding:10px 22px;color:var(--muted);font-size:.85em;" +
            "background:rgba(255,255,255,.72);border:1px solid var(--border);border-radius:var(--radius)}footer p{margin:4px 0}" +
            ".modal{display:none;position:fixed;inset:0;z-index:2000}.modal:target{display:block}" +
            ".backdrop{position:absolute;inset:0;background:rgba(15,23,42,.6)}" +
            ".dialog{position:relative;width:85vw;max-width:1600px;height:92vh;margin:4vh auto;display:flex;flex-direction:column;" +
            "background:rgba(255,255,255,.94);border:1px solid var(--border);border-radius:16px;box-shadow:0 26px 70px rgba(15,23,42,.36);overflow:hidden}" +
            ".close{position:absolute;top:10px;right:10px;z-index:1;width:40px;height:40px;display:grid;place-items:center;border-radius:12px;" +
            "border:1px solid var(--border);background:rgba(255,255,255,.9);color:var(--text);font-size:22px;line-height:1;text-decoration:none}" +
            ".close:hover{border-color:var(--border2)}" +
            ".vbody{display:grid;grid-template-columns:minmax(0,1fr) 380px;grid-template-rows:minmax(0,1fr);flex:1 1 auto;min-height:0}" +
            ".vbody.wide{grid-template-columns:minmax(0,1fr)}" +
            ".viewer{position:relative;display:grid;grid-template:minmax(0,1fr)/minmax(0,1fr);min-height:0;min-width:0}" +
            ".zone{position:absolute;z-index:1;display:flex;align-items:center;color:rgba(15,23,42,.7);text-decoration:none;transition:background 120ms ease}" +
            ".zone span{display:grid;place-items:center;width:44px;height:44px;padding-bottom:4px;border-radius:50%;font-size:34px;line-height:1;" +
            "color:var(--text);background:rgba(255,255,255,.82);border:1px solid var(--border2);box-shadow:0 4px 14px rgba(0,0,0,.35)}" +
            ".zone:hover span{background:#fff}" +
            ".zone-left{left:0;top:0;bottom:0;width:30%;justify-content:flex-start;padding-left:16px}" +
            ".zone-right{right:0;top:0;bottom:0;width:30%;justify-content:flex-end;padding-right:16px}" +
            ".zone-top{left:30%;right:30%;top:0;height:28%}.zone-bottom{left:30%;right:30%;bottom:0;height:28%}" +
            ".zone-left:hover{background:linear-gradient(90deg,rgba(15,23,42,.14),transparent)}" +
            ".zone-right:hover{background:linear-gradient(270deg,rgba(15,23,42,.14),transparent)}" +
            ".zone-top:hover{background:linear-gradient(180deg,rgba(15,23,42,.12),transparent)}" +
            ".zone-bottom:hover{background:linear-gradient(0deg,rgba(15,23,42,.12),transparent)}" +
            ".info{overflow:auto;padding:16px 18px;border-left:1px solid var(--border);background:var(--card2);font-size:14px}" +
            ".info h3{font-size:1.05em;margin:0 0 4px;padding-right:46px}.info p{margin:0 0 8px}" +
            ".info .pinout{display:block;max-width:100%;height:auto;margin:10px auto;border-radius:8px;border:1px solid var(--border)}" +
            ".pins{border-collapse:collapse;width:100%;font-size:12px;margin-top:6px}" +
            ".pins th,.pins td{text-align:left;padding:3px 6px;border-bottom:1px solid var(--border);vertical-align:top}" +
            ".pins th{color:var(--muted);font-weight:600}.pins tr.here td{background:var(--diff);font-weight:700}.pins .reading{white-space:pre-line}" +
            ".stage{display:grid;grid-template:minmax(0,1fr)/minmax(0,1fr);place-items:center;padding:14px;min-height:0;min-width:0;overflow:hidden;" +
            "background:radial-gradient(900px 600px at 30% 20%,rgba(14,165,233,.14),transparent 60%),linear-gradient(180deg,rgba(2,6,23,.06),rgba(2,6,23,.02))}" +
            ".stage img{width:100%;height:100%;object-fit:contain;border-radius:12px;background:#0b1220;border:1px solid rgba(15,23,42,.22);" +
            "box-shadow:inset 0 0 0 1px rgba(255,255,255,.25),0 18px 45px rgba(15,23,42,.24)}" +
            ".modal .footer{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:14px;padding:12px 14px;border-top:1px solid var(--border);color:var(--muted);font-size:13px}" +
            ".modal .title{grid-column:2;justify-self:center;text-align:center;display:flex;flex-direction:column;gap:2px;min-width:0}" +
            ".m-name{color:var(--text);font-weight:750;font-size:14px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}" +
            ".m-time{color:var(--muted);font-size:12px;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}" +
            ".modal .actions{grid-column:3;justify-self:end;display:flex;align-items:center;gap:10px;margin:0;font-size:13px}" +
            "button.btn{font:inherit;font-weight:650;cursor:pointer}" +
            ".share{position:relative}.share>summary{list-style:none;cursor:pointer}.share>summary::-webkit-details-marker{display:none}" +
            ".share-panel{position:absolute;right:0;bottom:calc(100% + 8px);z-index:3;width:330px;display:grid;gap:6px;padding:12px;" +
            "background:#fff;border:1px solid var(--border2);border-radius:12px;box-shadow:var(--shadow)}" +
            ".share-panel label{margin:0;font-size:12px}" +
            ".share-url{width:100%;padding:6px 8px;font:12px ui-monospace,Consolas,monospace;color:var(--text);border:1px solid var(--border2);border-radius:6px}" +
            "@media (max-width:720px){.topbar{position:static;height:auto;grid-template-columns:1fr;justify-items:center;padding:10px 14px;gap:8px}" +
            ".brand,.meta,.header-actions{grid-column:1;justify-self:center;align-items:center;text-align:center}.meta{white-space:normal}" +
            ".container{padding:14px 10px 6px}html{scroll-padding-top:12px}.grid.grid-few{grid-template-columns:1fr}.panel{padding:14px}" +
            "footer{width:calc(100% - 20px)}.dialog{width:96vw;height:96vh;margin:2vh auto}.zone span{width:34px;height:34px;font-size:26px}" +
            ".vbody{grid-template-columns:minmax(0,1fr);grid-template-rows:minmax(0,55%) minmax(0,45%)}.vbody.wide{grid-template-rows:minmax(0,1fr)}" +
            ".info{border-left:0;border-top:1px solid var(--border)}" +
            ".modal .footer{grid-template-columns:1fr;justify-items:center}.modal .title,.modal .actions{grid-column:1;justify-self:center}" +
            ".share-panel{right:auto;left:50%;transform:translateX(-50%);width:min(330px,90vw)}}";

        public static readonly string StyleHash =
            "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Css)));

        // ###########################################################################################
        // THE PAGE'S ONE SCRIPT - what the viewer cannot do in CSS (owner requests, 2026-11-27): Esc closes it; Left and Up go to
        // the previous capture, Right and Down to the next, and so does the WHEEL - one step per turn, at most one every 350 ms,
        // never over the component's facts, which scroll; and "Copy for ..." puts its data-copy on the clipboard, saying
        // "Copied". Nothing else: the viewer itself is CSS (:target) and works without it. Fixed text, nothing anybody typed in
        // it - the words copied are the button's attribute, encoded on the page - and allowed by its hash alone (ScriptHash).
        // The address is REPLACED, so the keys and the wheel leave no trail of captures for "Back" to walk through.
        // ###########################################################################################
        public const string Script =
            "(function(){var last=0;" +
            "function open(){return document.querySelector(\".modal:target\");}" +
            "function go(m,s){var l=m.querySelector(s);if(!l)return false;location.replace(l.getAttribute(\"href\"));return true;}" +
            "document.addEventListener(\"keydown\",function(e){var m=open();if(!m||e.altKey||e.ctrlKey||e.metaKey)return;" +
            "if(e.target.closest&&e.target.closest(\"input,textarea\")&&e.key!==\"Escape\")return;" +
            "var s=e.key===\"Escape\"?\".close\":(e.key===\"ArrowRight\"||e.key===\"ArrowDown\")?\".next\":(e.key===\"ArrowLeft\"||e.key===\"ArrowUp\")?\".prev\":null;" +
            "if(s&&go(m,s))e.preventDefault();});" +
            "document.addEventListener(\"wheel\",function(e){var m=open();if(!m||e.ctrlKey||e.deltaY===0)return;" +
            "if(e.target.closest&&e.target.closest(\".info,.share-panel\"))return;e.preventDefault();" +
            "var now=Date.now();if(now-last<350)return;last=now;go(m,e.deltaY>0?\".next\":\".prev\");},{passive:false});" +
            "document.addEventListener(\"click\",function(e){var b=e.target.closest&&e.target.closest(\"[data-copy]\");" +
            "if(!b||!navigator.clipboard)return;var w=b.textContent;navigator.clipboard.writeText(b.getAttribute(\"data-copy\"))" +
            ".then(function(){b.textContent=\"Copied\";setTimeout(function(){b.textContent=w;},1500);});});})();";

        public static readonly string ScriptHash =
            "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Script)));

        // Nothing but the style and the script above may run, pictures only from this site (the captures and their pinout
        // drawings - nothing of the published data since the baseline went, server 16.0.0), and forms post only to this site
        // (the comments and signing out). The script copies to the clipboard, which asks for nothing more.
        public static string ContentSecurityPolicy() =>
            $"default-src 'none'; style-src '{CollectionPages.StyleHash}'; script-src '{CollectionPages.ScriptHash}'; " +
            "img-src 'self'; form-action 'self'; frame-ancestors 'none'; base-uri 'none'";
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

        // The administrator asked to delete the collection (?remove=1) - the question shown, with its form.
        public bool RemovingCollection { get; init; }

        // Offline for everybody else - shown to the administrator alone (CollectionFlows.FindShownToAsync), saying so.
        public CollectionOffline Offline { get; init; }

        // Why a form was refused, when its own place on the page cannot say it - a hide, a "Show again", a delete, or an edit of
        // a comment gone meanwhile (code review, 2026-11-27: the page came back with an error status and no words).
        public IReadOnlyList<string> PageErrors { get; init; } = [];

        // What a refusal is about - a capture's checksum (a new comment) or a comment's id (an edit) - its sentences, and the
        // text that was typed, put back in its box.
        public string? ErrorOn { get; init; }

        public IReadOnlyList<string> Errors { get; init; } = [];

        public string? Typed { get; init; }
    }

    // The reader signed in - never the account's address or anything else of it.
    public sealed record CollectionViewer(long Id, string DisplayName, bool IsAdministrator);
}
