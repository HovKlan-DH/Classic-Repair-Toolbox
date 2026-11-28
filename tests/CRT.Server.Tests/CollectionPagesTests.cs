using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // A shared collection's page (CollectionPages): it names nobody, says where it comes from and when it goes, shows each
    // capture with its settings and, seen large, what CRT's component info window said of its component, and carries nothing
    // anybody typed as markup.
    // ###########################################################################################
    public sealed class CollectionPagesTests
    {
        private static readonly DateTimeOffset Created = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private const string Site = "https://classic-repair-toolbox.dk";

        private static readonly string ShaA = new('a', 64);

        private static readonly string ShaB = new('b', 64);

        private static readonly string PinoutSha = new('d', 64);

        private static CollectionImageRecord Capture(string sha, string label, string pin, string? caption = null, string? pinout = null) =>
            new(new CollectionImageEntry(sha, label, pin, "/RAS", "Pulsing", "100us", "2V", "1V", caption, Created, pinout), 800, 480);

        private static CollectionRecord Collection(params CollectionImageRecord[] images) =>
            new("k7mq3x9vbn2p", new string('f', 64), "Commodore 64", "250407", "PAL", "No picture after changing the PLA", "It went black.\nThen it smelled.",
                ["https://www.lemon64.com/forum/viewtopic.php?t=1"], Created, Created, Created.AddYears(1), null, null, images, images.Length);

        private static string Page(CollectionRecord collection, CollectionPageView? view = null) =>
            CollectionPages.Page(Site, collection, view);

        [Fact]
        public void The_page_says_the_board_the_id_and_when_it_was_shared_and_goes()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.Contains("<h1>No picture after changing the PLA</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<code>k7mq3x9vbn2p</code>", html, StringComparison.Ordinal);
            Assert.Contains("Shared 2026-November-26", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"badge-red\">Will be deleted <b>2027-November-26</b></span>", html, StringComparison.Ordinal);
            Assert.Contains("Shared with Classic Repair Toolbox", html, StringComparison.Ordinal);
        }

        // Owner request, 2026-11-28: "I want to have the board hardware and model shown in the "top user panel", where the collection
        // headline is (move it there). Also the oscilloscope brand and model should be shown there" - the bar over the page keeps
        // only the date it goes.
        [Fact]
        public void The_board_and_the_oscilloscope_stand_under_the_title()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")) with { Oscilloscopes = ["Rigol DS1054Z", "Siglent <SDS>"] });

            Assert.Contains(
                "<h1>No picture after changing the PLA</h1><p class=\"kit\"><span>Board: <b>Commodore 64 - 250407 (PAL)</b></span>" +
                "<span>Oscilloscope: <b>Rigol DS1054Z, Siglent &lt;SDS&gt;</b></span></p>",
                html,
                StringComparison.Ordinal);

            string meta = html[html.IndexOf("<div class=\"meta\">", StringComparison.Ordinal)..html.IndexOf("<div class=\"header-actions\">", StringComparison.Ordinal)];
            Assert.DoesNotContain("Commodore 64", meta, StringComparison.Ordinal);

            // No oscilloscope known (captures from before CRT kept it): the board alone.
            Assert.Contains(
                "<p class=\"kit\"><span>Board: <b>Commodore 64 - 250407 (PAL)</b></span></p>",
                Page(Collection(Capture(ShaA, "U7", "12"))),
                StringComparison.Ordinal);
        }

        // Owner request, 2026-11-28: "For each image I want it to show the component display name as its headline - e.g. like this:
        // "<b>U53</b> | <b>Dynamic RAM</b> | <b>4164</b>"" - on its card, its section and its viewer; its label alone when CRT told
        // nothing of its component. The pin and signal under it.
        [Fact]
        public void A_capture_s_headline_is_its_component_each_part_bold()
        {
            CollectionComponent component = new("U53", "U53 | Dynamic <RAM> | 4164", null, null, null, null);

            string html = Page(Collection(Capture(ShaA, "U53", "2"), Capture(ShaB, "U8", "3")) with { Components = [component] });

            Assert.Contains("<a class=\"cap-title\" href=\"#U53-2\"><b>U53</b> | <b>Dynamic &lt;RAM&gt;</b> | <b>4164</b></a>", html, StringComparison.Ordinal);
            Assert.Contains("<h2><b>U53</b> | <b>Dynamic &lt;RAM&gt;</b> | <b>4164</b> <span class=\"signal\">Pin 2 - /RAS</span></h2>", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"m-name\"><b>U53</b> | <b>Dynamic &lt;RAM&gt;</b> | <b>4164</b></div>", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"cap-sub\">Pin 2 - /RAS</div>", html, StringComparison.Ordinal);

            Assert.Contains("<a class=\"cap-title\" href=\"#U8-3\"><b>U8</b></a>", html, StringComparison.Ordinal);
        }

        // Owner request, 2026-11-28: "For the front page of a collection, then it should not show any large images - only the
        // thumbnails": every card a thumbnail however few, and each capture's section its picture as a thumbnail beside its
        // settings - large only in the viewer.
        [Fact]
        public void The_page_itself_shows_its_captures_as_thumbnails()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.Contains("<div class=\"grid\">", html, StringComparison.Ordinal);
            Assert.DoesNotContain("grid-few", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"capture-body\"><a class=\"framed capture-thumb\" href=\"#view.U7-12\"", html, StringComparison.Ordinal);
            Assert.Contains(".capture-thumb{flex:0 0 240px;", CollectionPages.Css, StringComparison.Ordinal);
        }

        // Owner decision, 2026-11-25: "For the name of the collection I do not think any user identification should be shown" -
        // and since 2026-11-28 nobody is behind a collection but a key, whose hash is never shown.
        [Fact]
        public void The_page_names_nobody()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.DoesNotContain("shared by", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(">Owner<", html, StringComparison.Ordinal);
            Assert.DoesNotContain(new string('f', 64), html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_collection_with_no_title_is_named_by_its_board() =>
            Assert.Contains("<h1>Oscilloscope captures of Commodore 64 250407</h1>", Page(Collection(Capture(ShaA, "U7", "12")) with { Title = null }), StringComparison.Ordinal);

        // The administrator's delete refused says why at the top of the page (code review, 2026-11-27: the page came back with an
        // error status and no words).
        [Fact]
        public void A_refusal_with_no_form_to_say_it_in_is_said_at_the_top()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView { PageErrors = ["That collection was deleted <now>."] });

            Assert.Contains("<div class=\"errors\" role=\"alert\"><p>That collection was deleted &lt;now&gt;.</p></div>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("role=\"alert\"", Page(Collection(Capture(ShaA, "U7", "12"))), StringComparison.Ordinal);
        }

        // *** NOTHING ANYBODY TYPED IS EVER READ AS MARKUP. ***
        [Fact]
        public void Every_word_the_owner_typed_is_shown_as_text()
        {
            CollectionRecord collection = Collection(Capture(ShaA, "<b>U7</b>", "12", caption: "<script>alert(1)</script>")) with
            {
                Title = "<img src=x onerror=alert(1)>",
                Description = "\"quoted\" & <i>tagged</i>"
            };

            string html = Page(collection);

            Assert.DoesNotContain("<script>alert", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<img src=x", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<i>tagged", html, StringComparison.Ordinal);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html, StringComparison.Ordinal);
            Assert.Contains("&quot;quoted&quot; &amp; &lt;i&gt;tagged&lt;/i&gt;", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_forum_link_carries_nothing_of_the_page()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.Contains(
                "<a href=\"https://www.lemon64.com/forum/viewtopic.php?t=1\" rel=\"nofollow ugc noopener noreferrer\">www.lemon64.com/forum/viewtopic.php?t=1</a>",
                html,
                StringComparison.Ordinal);
        }

        [Fact]
        public void Each_capture_has_an_anchor_of_its_own_and_steps_to_the_next()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12"), Capture(ShaB, "U7", "12"), Capture(new string('d', 64), "R 5", "1")));

            Assert.Contains("id=\"U7-12\"", html, StringComparison.Ordinal);
            Assert.Contains("id=\"U7-12-2\"", html, StringComparison.Ordinal);
            Assert.Contains("id=\"R-5-1\"", html, StringComparison.Ordinal);
            Assert.Contains("<a href=\"#U7-12-2\">Next</a>", html, StringComparison.Ordinal);
            Assert.Contains("<a href=\"#U7-12\">Previous</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_capture_s_picture_is_the_collection_s_own_address()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.Contains($"src=\"{Site}/collections/k7mq3x9vbn2p/{ShaA}.png\" width=\"800\" height=\"480\"", html, StringComparison.Ordinal);
        }

        // Owner decision, 2026-11-27: the baseline "is not relevant" - a capture shows its own settings alone.
        [Fact]
        public void A_capture_shows_its_own_settings_and_no_baseline()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12") with { Entry = Capture(ShaA, "U7", "12").Entry with { TriggerLevel = null } }));

            Assert.Contains("<th>T/DIV</th><td>100us</td></tr>", html, StringComparison.Ordinal);
            Assert.Contains("<th>Trigger level</th><td><span class=\"muted\">not read</span></td></tr>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("aseline", html, StringComparison.Ordinal);
        }

        // ---- The viewer and "Download all" (owner request, 2026-11-27: the owner's capture gallery's look) -------------------

        // Each card's picture and each section's open the capture's viewer; its previous and next go round, it says where in the
        // list it is, and "Download" saves it under its place and anchor - with no script at all (:target).
        [Fact]
        public void Each_capture_opens_in_a_viewer_whose_previous_and_next_go_round()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12"), Capture(ShaB, "U8", "3")));

            // Its card in the grid, its section's picture, and the other capture's viewer's four areas.
            Assert.Equal(6, html.Split("href=\"#view.U7-12\"").Length - 1);
            Assert.Contains("<div class=\"modal\" id=\"view.U7-12\"", html, StringComparison.Ordinal);

            // Owner request, 2026-11-27: a big part of the left and the top goes back, of the right and the bottom on.
            Assert.Contains("<a class=\"zone zone-left prev\" href=\"#view.U8-3\"", html, StringComparison.Ordinal);
            Assert.Contains("<a class=\"zone zone-right next\" href=\"#view.U8-3\"", html, StringComparison.Ordinal);
            Assert.Contains("<a class=\"zone zone-top prev\" href=\"#view.U8-3\"", html, StringComparison.Ordinal);
            Assert.Contains("<a class=\"zone zone-bottom next\" href=\"#view.U8-3\"", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"counter\">1 / 2</span>", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"counter\">2 / 2</span>", html, StringComparison.Ordinal);
            Assert.Contains($"href=\"{Site}/collections/k7mq3x9vbn2p/{ShaA}.png\" download=\"001-U7-12.png\"", html, StringComparison.Ordinal);
            Assert.Contains($"href=\"#{CollectionPages.ViewerClosed}\"", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"m-time\">Pin 12 - /RAS | T/DIV 100us | V/DIV 2V | Trigger level 1V</div>", html, StringComparison.Ordinal);
        }

        // One capture: nothing to step to.
        [Fact]
        public void A_single_capture_s_viewer_has_no_previous_or_next()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.DoesNotContain("class=\"zone ", html, StringComparison.Ordinal);
        }

        // The viewers' ids hold a ".", which no capture's anchor does - so no capture can take a viewer's place.
        [Fact]
        public void A_viewer_s_id_is_never_a_capture_s_anchor()
        {
            IReadOnlyList<string> anchors = CollectionRules.AnchorsOf(
                [Capture(ShaA, "view", "U7-12").Entry, Capture(ShaB, "viewer", "closed").Entry]);

            Assert.DoesNotContain(anchors, anchor => anchor.Contains('.', StringComparison.Ordinal));
            Assert.DoesNotContain(CollectionPages.ViewerClosed, anchors);
        }

        // Owner request, 2026-11-27: "show something like the component info window ... for this specific component and pin" - the
        // facts, CRT's drawing of the pinout with the pin marked, and the pins with this capture's highlighted.
        [Fact]
        public void A_capture_seen_large_tells_of_its_component_its_pinout_and_its_pin()
        {
            CollectionComponent component = new(
                "U7", "U7 | VIC-II | 6569", "IC | 6569 | DIP-40", "Video chip <b>", "Makes the picture",
                [new CollectionPin("11", "AEC", "Output", null, null), new CollectionPin("12", "/RAS", "Output", "RAS", "Pulsing\nLow [64K]")]);

            string html = Page(Collection(Capture(ShaA, "U7", "12", pinout: PinoutSha)) with { Components = [component] });

            Assert.Contains("<aside class=\"info\"><h3>U7 | VIC-II | 6569</h3><p class=\"muted\">IC | 6569 | DIP-40</p><p>Video chip &lt;b&gt;</p>", html, StringComparison.Ordinal);
            Assert.Contains($"<img class=\"pinout\" src=\"{Site}/collections/k7mq3x9vbn2p/{PinoutSha}.png\"", html, StringComparison.Ordinal);
            Assert.Contains("alt=\"The pinout of U7, pin 12 marked\"", html, StringComparison.Ordinal);
            Assert.Contains("<tr class=\"here\"><td>12</td><td>/RAS</td><td>Output</td><td>RAS</td><td class=\"reading\">Pulsing\nLow [64K]</td></tr>", html, StringComparison.Ordinal);
            Assert.Contains("<tr><td>11</td><td>AEC</td><td>Output</td><td></td><td class=\"reading\"></td></tr>", html, StringComparison.Ordinal);
        }

        // A capture of a component CRT told nothing of fills the viewer alone; a column no pin has is left out.
        [Fact]
        public void A_capture_whose_component_was_not_told_of_fills_the_viewer()
        {
            string bare = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.DoesNotContain("<aside", bare, StringComparison.Ordinal);
            Assert.Contains("<div class=\"vbody wide\">", bare, StringComparison.Ordinal);

            string noNets = Page(Collection(Capture(ShaA, "U7", "12")) with
            {
                Components = [new CollectionComponent("u7", null, null, null, null, [new CollectionPin("12", "/RAS", "Output", null, null)])]
            });

            Assert.Contains("<thead><tr><th>Pin</th><th>Name</th><th>Direction</th></tr></thead>", noNets, StringComparison.Ordinal);
        }

        // Owner request, 2026-11-27: "a link, and clicking that link will expand an option to copy a link to either Forum64,
        // Lemon64 or Reddit. The link should then navigate to this exact image, and show it large." By FORMAT since 2026-11-28
        // (owner: "OK to your proposal for the forum button"), each saying where it works - CRT.Data's one rule.
        [Fact]
        public void A_capture_seen_large_offers_its_own_link_in_each_format()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));
            string url = $"{Site}/collections/k7mq3x9vbn2p#view.U7-12";
            const string words = "U7, pin 12 - /RAS (Commodore 64 250407)";

            Assert.Contains("<details class=\"share\"><summary class=\"btn\">Link</summary>", html, StringComparison.Ordinal);
            Assert.Contains($"<input class=\"share-url\" type=\"text\" readonly value=\"{url}\"", html, StringComparison.Ordinal);
            Assert.Contains($"data-copy=\"[url={url}]{words}[/url]\">Copy as BBCode</button><span>Lemon64, Forum64,", html, StringComparison.Ordinal);
            Assert.Contains($"data-copy=\"[{words}]({url})\">Copy as Markdown</button><span>Reddit,", html, StringComparison.Ordinal);
            Assert.Contains($"data-copy=\"&lt;a href=&quot;{url}&quot;&gt;{words}&lt;/a&gt;\">Copy as HTML</button><span>Web pages,", html, StringComparison.Ordinal);
            Assert.Contains($"data-copy=\"{words}: {url}\">Copy plain link</button><span>Facebook groups,", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Copy for ", html, StringComparison.Ordinal);
        }

        // Owner request, 2026-11-27: the top left corner leads to CRT's project page, not the site's front page.
        [Fact]
        public void The_name_in_the_top_left_corner_leads_to_the_project_page() =>
            Assert.Contains(
                "<a class=\"brand-title\" href=\"https://github.com/HovKlan-DH/Classic-Repair-Toolbox\">Classic Repair Toolbox</a>",
                Page(Collection(Capture(ShaA, "U7", "12"))),
                StringComparison.Ordinal);

        [Fact]
        public void The_page_offers_every_capture_in_one_download()
        {
            Assert.Contains(
                $"<a class=\"btn\" href=\"{Site}/collections/k7mq3x9vbn2p/download\">Download 2 captures</a>",
                Page(Collection(Capture(ShaA, "U7", "12"), Capture(ShaB, "U8", "3"))),
                StringComparison.Ordinal);

            Assert.DoesNotContain("/download\"", Page(Collection()), StringComparison.Ordinal);
        }

        // A forum's link preview runs no script - the server's page carries the title and the first picture.
        [Fact]
        public void A_link_preview_gets_the_title_and_the_first_capture()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12"), Capture(ShaB, "U8", "3")));

            Assert.Contains("<meta property=\"og:title\" content=\"No picture after changing the PLA\">", html, StringComparison.Ordinal);
            Assert.Contains($"<meta property=\"og:image\" content=\"{Site}/collections/k7mq3x9vbn2p/{ShaA}.png\">", html, StringComparison.Ordinal);
            Assert.Contains("2 oscilloscope captures of Commodore 64 - 250407 (PAL), shared with Classic Repair Toolbox.", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Search_engines_are_asked_to_keep_nothing() =>
            Assert.Contains("<meta name=\"robots\" content=\"noindex, nofollow\">", Page(Collection(Capture(ShaA, "U7", "12"))), StringComparison.Ordinal);

        // Owner request, 2026-11-27: "a standard web page, explaining the deletion period has passed etc. ... what this was and
        // why it was deleted" - kept for good, naming nobody.
        [Theory]
        [InlineData(CollectionDeletedReasons.Expired, "It was deleted on 2027-January-26, as the time it was shared for had passed.")]
        [InlineData(CollectionDeletedReasons.Owner, "Whoever shared it deleted it on 2027-January-26")]
        [InlineData(CollectionDeletedReasons.Administrator, "It was taken down by the site&#39;s administrator on 2027-January-26")]
        public void A_deleted_collection_says_what_it_was_when_it_went_and_why(string reason, string why)
        {
            CollectionRecord gone = Collection() with { DeletedUtc = Created.AddMonths(2), DeletedReason = reason, Title = null };

            string html = CollectionPages.Deleted(gone);

            Assert.Contains($"<h1>{CollectionPages.DeletedTitle}</h1>", html, StringComparison.Ordinal);
            Assert.Contains(why, html, StringComparison.Ordinal);
            Assert.Contains("This page showed an oscilloscope collection", html, StringComparison.Ordinal);
            Assert.Contains("This page stays here for good", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"badge-red badge-large\">Deleted <b>2027-January-26</b></span>", html, StringComparison.Ordinal);
            Assert.Contains("What is Classic Repair Toolbox?", html, StringComparison.Ordinal);
            Assert.DoesNotContain("shared by", html, StringComparison.OrdinalIgnoreCase);
        }

        // Its owner stopped sharing it (migration 0028): to everybody else, only that - the page may come back under the link.
        [Fact]
        public void A_collection_its_owner_stopped_sharing_says_only_that_it_is_not_shared_at_the_moment()
        {
            string html = CollectionPages.SharingStopped(Collection(Capture(ShaA, "U7", "12")) with { UnsharedUtc = Created.AddDays(3) });

            Assert.Contains($"<h1>{CollectionPages.SharingStoppedTitle}</h1>", html, StringComparison.Ordinal);
            Assert.Contains("comes back under this same link", html, StringComparison.Ordinal);
            Assert.DoesNotContain(ShaA, html, StringComparison.Ordinal);
            Assert.DoesNotContain("No picture after changing the PLA", html, StringComparison.Ordinal);
            Assert.DoesNotContain("It went black", html, StringComparison.Ordinal);
        }

        // The administrator sees it still, told that nobody else does - so it can be checked, and deleted.
        [Fact]
        public void A_collection_whose_sharing_stopped_tells_the_administrator_that_nobody_else_sees_it()
        {
            CollectionRecord stopped = Collection(Capture(ShaA, "U7", "12")) with { UnsharedUtc = Created.AddDays(3) };

            string html = Page(stopped, view: AsAdministrator with { Offline = CollectionOffline.SharingStopped });

            Assert.Contains(System.Net.WebUtility.HtmlEncode(CollectionPages.OfflineSharingStopped), html, StringComparison.Ordinal);
            Assert.Contains(ShaA, html, StringComparison.Ordinal);
            Assert.Contains("?remove=1", html, StringComparison.Ordinal);

            // Shown to all, the page says nothing of it.
            Assert.DoesNotContain("Not shared:", Page(Collection(Capture(ShaA, "U7", "12")), view: AsAdministrator), StringComparison.Ordinal);
        }

        // The policy allows the page's one style and its one script by their hashes - either changed without its hash would be
        // blocked, so each is one with its hash.
        [Fact]
        public void The_policy_allows_the_page_s_own_style_and_script_and_nothing_else()
        {
            string policy = CollectionPages.ContentSecurityPolicy();
            string style = "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Css)));
            string script = "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Script)));

            Assert.Contains($"style-src '{style}'", policy, StringComparison.Ordinal);
            Assert.Contains($"script-src '{script}';", policy, StringComparison.Ordinal);

            // Pictures from this site alone - the captures and their pinout drawings; nothing of the published data since the
            // baseline went (2026-11-27).
            Assert.Contains("img-src 'self';", policy, StringComparison.Ordinal);
            Assert.Contains("default-src 'none'", policy, StringComparison.Ordinal);
            Assert.DoesNotContain("unsafe", policy, StringComparison.Ordinal);
        }

        // The script is the page's own fixed text, exactly as its hash allows it - nothing anybody typed reaches it.
        [Fact]
        public void The_page_carries_its_one_script_as_its_hash_allows_it()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12", caption: "</script><script>alert(1)</script>")));

            Assert.Equal(1, html.Split("<script>").Length - 1);
            Assert.Contains($"<script>{CollectionPages.Script}</script>", html, StringComparison.Ordinal);
        }

        // ---- No comments (owner decision, 2026-11-28: "do remove comment possibilities etc.") -------------------------------

        // Nobody writes on the page any more - signed in as the administrator neither: no comment form, no comment count.
        [Fact]
        public void The_page_takes_no_comment_from_anybody()
        {
            CollectionRecord collection = Collection(Capture(ShaA, "U7", "12"));

            foreach (string html in new[] { Page(collection), Page(collection, view: AsAdministrator) })
            {
                Assert.DoesNotContain("<textarea", html, StringComparison.Ordinal);
                Assert.DoesNotContain("omment", html, StringComparison.Ordinal);
                Assert.Contains("<div class=\"cap-sub\">Pin 12 - /RAS</div>", html, StringComparison.Ordinal);
            }
        }

        // The page names no way in for anybody: the website's sign-in is the administrator's alone.
        [Fact]
        public void Signed_out_the_page_offers_no_sign_in() =>
            Assert.DoesNotContain("sign-in", Page(Collection(Capture(ShaA, "U7", "12"))), StringComparison.Ordinal);

        [Fact]
        public void The_policy_lets_the_page_s_forms_post_to_this_site_alone() =>
            Assert.Contains("form-action 'self'", CollectionPages.ContentSecurityPolicy(), StringComparison.Ordinal);

        // ---- The administrator: "All collections" and the delete -----------------------------------------------------------

        private static readonly CollectionViewer Administrator = new(9, "Dennis@CRT", true);

        private static readonly CollectionViewer AMaintainer = new(7, "Bo", false);

        private const string Token = "0123456789abcdef0123456789abcdef";

        private static readonly CollectionPageView AsAdministrator = new() { Viewer = Administrator, Csrf = Token };

        private static CollectionRecord Listed(string id, string title, int captures) =>
            Collection() with { Id = id, Title = title, ImageCount = captures };

        [Fact]
        public void Signed_in_the_administrator_s_page_leads_to_all_collections_and_out()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: AsAdministrator);

            Assert.Contains($"<a href=\"{Site}/collections/\">{CollectionPages.ListTitle}</a>", html, StringComparison.Ordinal);
            Assert.Contains("Signed in as <b>Dennis@CRT</b>", html, StringComparison.Ordinal);
            Assert.Contains($"action=\"{Site}/account/sign-out\"", html, StringComparison.Ordinal);
            Assert.Contains($"<input type=\"hidden\" name=\"csrf\" value=\"{Token}\">", html, StringComparison.Ordinal);
        }

        // Whoever shared it deletes it in CRT, which keeps its window in step - the web offers it to the administrator alone.
        [Fact]
        public void Only_the_administrator_is_offered_to_delete_the_collection_and_is_asked_first()
        {
            CollectionRecord collection = Collection(Capture(ShaA, "U7", "12"));

            Assert.DoesNotContain("?remove=1", Page(collection), StringComparison.Ordinal);
            Assert.DoesNotContain("?remove=1", Page(collection, view: new CollectionPageView { Viewer = AMaintainer, Csrf = Token }), StringComparison.Ordinal);

            string offered = Page(collection, view: AsAdministrator);

            Assert.Contains($"href=\"{Site}/collections/k7mq3x9vbn2p?remove=1\">Delete this collection</a>", offered, StringComparison.Ordinal);
            Assert.DoesNotContain("/delete\"", offered, StringComparison.Ordinal);

            string asked = Page(collection, view: AsAdministrator with { RemovingCollection = true });

            Assert.Contains(CollectionPages.RemoveQuestion, asked, StringComparison.Ordinal);
            Assert.Contains($"<form class=\"admin\" method=\"post\" action=\"{Site}/collections/k7mq3x9vbn2p/delete\">", asked, StringComparison.Ordinal);
            Assert.Contains($"<input type=\"hidden\" name=\"csrf\" value=\"{Token}\">", asked, StringComparison.Ordinal);
        }

        // A viewer that is no administrator - a maintainer's session, say - is shown the page as anybody is.
        [Fact]
        public void Nobody_else_is_asked_to_delete_it_whatever_the_link_says()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView { Viewer = AMaintainer, Csrf = Token, RemovingCollection = true });

            Assert.DoesNotContain(CollectionPages.RemoveQuestion, html, StringComparison.Ordinal);
            Assert.DoesNotContain("Signed in as", html, StringComparison.Ordinal);
        }

        [Fact]
        public void All_collections_lists_every_one_with_its_board_captures_and_dates()
        {
            string html = CollectionPages.List(Site, Administrator, Token, [Listed("annafirst222", "My PLA", 1), Listed("bertfirst222", "Bert's VIC", 3)]);

            Assert.Contains($"<h1>{CollectionPages.ListTitle}</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h2>Every collection (2) <span class=\"badge\">Administrator</span></h2>", html, StringComparison.Ordinal);
            Assert.Contains($"<a href=\"{Site}/collections/annafirst222\">My PLA</a>", html, StringComparison.Ordinal);
            Assert.Contains("Commodore 64 - 250407 (PAL) | 1 capture | Shared 2026-November-26 | Deleted on 2027-November-26", html, StringComparison.Ordinal);
            Assert.Contains("Bert&#39;s VIC", html, StringComparison.Ordinal);
            Assert.Contains("3 captures", html, StringComparison.Ordinal);

            // Signed in, the way out comes back here; the list does not link to itself.
            Assert.Contains("<input type=\"hidden\" name=\"return\" value=\"/collections/\">", html, StringComparison.Ordinal);
            Assert.DoesNotContain($">{CollectionPages.ListTitle}</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void All_collections_marks_one_whose_sharing_its_owner_stopped()
        {
            string html = CollectionPages.List(Site, Administrator, Token,
                [Listed("annafirst222", "My PLA", 1) with { UnsharedUtc = Created.AddDays(1) }, Listed("annasecond22", "My VIC", 1)]);

            Assert.Equal(1, html.Split($"<span class=\"flag\">{CollectionPages.NotSharedFlag}</span>").Length - 1);
            Assert.Contains($"My PLA</a> <span class=\"flag\">{CollectionPages.NotSharedFlag}</span>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void An_empty_list_says_so() =>
            Assert.Contains("<h2>Every collection (0) <span class=\"badge\">Administrator</span></h2><p class=\"muted\">None.</p>",
                CollectionPages.List(Site, Administrator, Token, []), StringComparison.Ordinal);

        [Fact]
        public void Every_word_typed_is_shown_as_text_in_the_list()
        {
            string html = CollectionPages.List(Site, Administrator with { DisplayName = "<i>Dennis</i>" }, Token,
                [Listed("annafirst222", "<script>alert(1)</script>", 1) with { BoardName = "<b>250407</b>" }]);

            Assert.DoesNotContain("<script>alert", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<i>Dennis</i>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<b>250407</b>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Signed_out_the_list_says_it_is_the_administrator_s_and_how_to_sign_in_and_come_back_to_it()
        {
            string html = CollectionPages.ListSignedOut(Site);

            Assert.Contains("is the site's administrator's", html, StringComparison.Ordinal);
            Assert.Contains($"{Site}/account/sign-in?return=%2Fcollections%2F\">Sign in</a>", html, StringComparison.Ordinal);
            Assert.Contains("sign-in?return=%2Fcrt-test%2Fcollections%2F", CollectionPages.ListSignedOut(Site + "/crt-test"), StringComparison.Ordinal);
        }
    }
}
