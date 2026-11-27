using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Collections;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // A shared collection's page (CollectionPages): it names nobody, says where it comes from and when it goes, shows each
    // capture beside its baseline with the settings that differ in bold, and carries nothing anybody typed as markup.
    // ###########################################################################################
    public sealed class CollectionPagesTests : IDisposable
    {
        private static readonly DateTimeOffset Created = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private const string Site = "https://classic-repair-toolbox.dk";

        private static readonly string ShaA = new('a', 64);

        private static readonly string ShaB = new('b', 64);

        private static readonly string BaselineSha = new('c', 64);

        private readonly string thisTree = Path.Combine(Path.GetTempPath(), "crt-collection-tree-" + Guid.NewGuid().ToString("N"));

        public void Dispose()
        {
            if (Directory.Exists(this.thisTree))
                Directory.Delete(this.thisTree, recursive: true);
        }

        private static CollectionImageRecord Capture(string sha, string label, string pin, CollectionBaseline? baseline = null, string? caption = null) =>
            new(new CollectionImageEntry(sha, label, pin, "/RAS", "Pulsing", "100us", "2V", "1V", baseline, caption, Created), 800, 480);

        private static CollectionRecord Collection(params CollectionImageRecord[] images) =>
            new("k7mq3x9vbn2p", 42, "Commodore 64", "250407", "PAL", "No picture after changing the PLA", "It went black.\nThen it smelled.",
                ["https://www.lemon64.com/forum/viewtopic.php?t=1"], Created, Created, Created.AddYears(1), null, null, null, images, images.Length);

        private static string Page(CollectionRecord collection, Func<CollectionBaseline, string?>? baselineUrl = null, CollectionPageView? view = null) =>
            CollectionPages.Page(Site, collection, baselineUrl ?? (_ => null), view);

        [Fact]
        public void The_page_says_the_board_the_id_and_when_it_was_shared_and_goes()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.Contains("<h1>No picture after changing the PLA</h1>", html, StringComparison.Ordinal);
            Assert.Contains("Commodore 64 - 250407 (PAL)", html, StringComparison.Ordinal);
            Assert.Contains("<code>k7mq3x9vbn2p</code>", html, StringComparison.Ordinal);
            Assert.Contains("Shared 2026-November-26", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"badge-red\">Will be deleted <b>2027-November-26</b></span>", html, StringComparison.Ordinal);
            Assert.Contains("Shared with Classic Repair Toolbox", html, StringComparison.Ordinal);
        }

        // Owner decision, 2026-11-25: "For the name of the collection I do not think any user identification should be shown".
        [Fact]
        public void The_page_names_nobody()
        {
            // The page is made from the collection and its comments alone - no owner's account is handed to it - so this asks
            // that it says nothing of one: no "shared by", and no "Owner" mark while the owner has not commented.
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView { Comments = [Comment(5, 7, "Bo", "Check R7.")] });

            Assert.DoesNotContain("shared by", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(">Owner<", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_collection_with_no_title_is_named_by_its_board() =>
            Assert.Contains("<h1>Oscilloscope captures of Commodore 64 250407</h1>", Page(Collection(Capture(ShaA, "U7", "12")) with { Title = null }), StringComparison.Ordinal);

        // A hide, a delete or an edit refused where no form of it is drawn says why at the top of the page (code review,
        // 2026-11-27: the page came back with an error status and no words).
        [Fact]
        public void A_refusal_with_no_form_to_say_it_in_is_said_at_the_top()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView { PageErrors = ["There is no comment of yours with that id <now>."] });

            Assert.Contains("<div class=\"errors\" role=\"alert\"><p>There is no comment of yours with that id &lt;now&gt;.</p></div>", html, StringComparison.Ordinal);
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

        [Fact]
        public void The_baseline_stands_beside_the_capture_and_a_setting_that_differs_is_bold()
        {
            var baseline = new CollectionBaseline(BaselineSha, "Scope baseline/U7_12_RAS.png", "50us", "2 V", null);

            string html = Page(Collection(Capture(ShaA, "U7", "12", baseline)), _ => "https://example/baseline.png");

            Assert.Contains("<img src=\"https://example/baseline.png\"", html, StringComparison.Ordinal);
            Assert.Contains("<th>T/DIV</th><td><b>100us</b></td><td>50us</td>", html, StringComparison.Ordinal);

            // "2V" and "2 V" are the same setting.
            Assert.Contains("<th>V/DIV</th><td>2V</td><td>2 V</td>", html, StringComparison.Ordinal);
            Assert.Contains("<th>Trigger level</th><td>1V</td><td><span class=\"muted\">not given</span></td>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_capture_with_no_baseline_shows_its_own_settings_alone()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.Contains("<th>T/DIV</th><td>100us</td></tr>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<th>Baseline</th>", html, StringComparison.Ordinal);
        }

        // ---- The viewer and "Download all" (owner request, 2026-11-27: the owner's capture gallery's look) -------------------

        // Each card's picture and each section's open the capture's viewer; its previous and next go round, it says where in the
        // list it is, and "Download" saves it under its place and anchor - with no script at all (:target).
        [Fact]
        public void Each_capture_opens_in_a_viewer_whose_previous_and_next_go_round()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12"), Capture(ShaB, "U8", "3")));

            // Its card in the grid, its section's picture, and the other capture's viewer's previous and next.
            Assert.Equal(4, html.Split("href=\"#view.U7-12\"").Length - 1);
            Assert.Contains("<div class=\"modal\" id=\"view.U7-12\"", html, StringComparison.Ordinal);
            Assert.Contains("<a class=\"nav prev\" href=\"#view.U8-3\"", html, StringComparison.Ordinal);
            Assert.Contains("<a class=\"nav next\" href=\"#view.U8-3\"", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"counter\">1 / 2</span>", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"counter\">2 / 2</span>", html, StringComparison.Ordinal);
            Assert.Contains($"href=\"{Site}/collections/k7mq3x9vbn2p/{ShaA}.png\" download=\"001-U7-12.png\"", html, StringComparison.Ordinal);
            Assert.Contains($"href=\"#{CollectionPages.ViewerClosed}\"", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"m-time\">T/DIV 100us | V/DIV 2V | Trigger level 1V</div>", html, StringComparison.Ordinal);
        }

        // One capture: nothing to step to.
        [Fact]
        public void A_single_capture_s_viewer_has_no_previous_or_next()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")));

            Assert.DoesNotContain("class=\"nav ", html, StringComparison.Ordinal);
            Assert.Contains("class=\"grid grid-few grid-single\"", html, StringComparison.Ordinal);
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

        [Fact]
        public void A_capture_s_baseline_opens_in_a_viewer_of_its_own()
        {
            CollectionBaseline baseline = new(BaselineSha, "Scope baseline/U7_12.png", "50us", "2V", null);

            string html = Page(Collection(Capture(ShaA, "U7", "12", baseline)), baselineUrl: _ => "https://example/baseline.png");

            Assert.Contains("href=\"#base.U7-12\"", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"modal\" id=\"base.U7-12\"", html, StringComparison.Ordinal);
            Assert.Contains("U7, pin 12 - /RAS - Classic Repair Toolbox&#39;s baseline", html, StringComparison.Ordinal);
        }

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
        [InlineData(CollectionDeletedReasons.Owner, "Its owner deleted it on 2027-January-26")]
        [InlineData(CollectionDeletedReasons.Administrator, "It was taken down by the site&#39;s administrator on 2027-January-26")]
        public void A_deleted_collection_says_what_it_was_when_it_went_and_why(string reason, string why)
        {
            CollectionRecord gone = Collection() with { DeletedUtc = Created.AddMonths(2), DeletedReason = reason, Title = null };

            string html = CollectionPages.Deleted(Site, gone);

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
            string html = CollectionPages.SharingStopped(Site, Collection(Capture(ShaA, "U7", "12")) with { UnsharedUtc = Created.AddDays(3) });

            Assert.Contains($"<h1>{CollectionPages.SharingStoppedTitle}</h1>", html, StringComparison.Ordinal);
            Assert.Contains("comes back under this same link", html, StringComparison.Ordinal);
            Assert.DoesNotContain(ShaA, html, StringComparison.Ordinal);
            Assert.DoesNotContain("No picture after changing the PLA", html, StringComparison.Ordinal);
            Assert.DoesNotContain("It went black", html, StringComparison.Ordinal);
        }

        // Its owner and the administrator see it still, each told who else does - and nobody is offered a comment form.
        [Fact]
        public void A_collection_whose_sharing_stopped_tells_its_owner_and_the_administrator_who_sees_it()
        {
            CollectionRecord stopped = Collection(Capture(ShaA, "U7", "12")) with { UnsharedUtc = Created.AddDays(3) };

            string asOwner = Page(stopped, view: SignedIn(42) with { Offline = CollectionOffline.SharingStopped });
            string asAdministrator = Page(stopped, view: new CollectionPageView { Viewer = new CollectionViewer(1, "Dennis@CRT", true), Offline = CollectionOffline.SharingStopped });

            Assert.Contains(CollectionPages.OfflineSharingStoppedOwner.Replace("\"", "&quot;", StringComparison.Ordinal), asOwner, StringComparison.Ordinal);
            Assert.Contains(CollectionPages.OfflineSharingStoppedAdministrator, asAdministrator, StringComparison.Ordinal);
            Assert.DoesNotContain("<textarea", asOwner, StringComparison.Ordinal);
        }

        // The policy allows the page's one style and its one script by their hashes - either changed without its hash would be
        // blocked, so each is one with its hash.
        [Fact]
        public void The_policy_allows_the_page_s_own_style_and_script_and_nothing_else()
        {
            string policy = CollectionPages.ContentSecurityPolicy([Site]);
            string style = "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Css)));
            string script = "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Script)));

            Assert.Contains($"style-src '{style}'", policy, StringComparison.Ordinal);
            Assert.Contains($"script-src '{script}';", policy, StringComparison.Ordinal);
            Assert.Contains($"img-src 'self' {Site};", policy, StringComparison.Ordinal);
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

        // ---- Comments (step 3) ------------------------------------------------------------

        private static CollectionCommentRecord Comment(long id, long accountId, string name, string body, bool hidden = false, bool administrator = false, DateTimeOffset? edited = null) =>
            new(id, "k7mq3x9vbn2p", ShaA, accountId, name, administrator, body, Created.AddHours(2), edited, hidden ? Created.AddHours(3) : null);

        private static CollectionPageView SignedIn(long viewerId, params CollectionCommentRecord[] comments) =>
            new() { Viewer = new CollectionViewer(viewerId, viewerId == 42 ? "Anna" : "Bo", false), Csrf = "0123456789abcdef0123456789abcdef", Comments = comments };

        [Fact]
        public void Signed_out_a_capture_s_comments_show_and_the_way_to_sign_in_back_to_them()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView { Comments = [Comment(5, 7, "Bo", "Check R7.")] });

            Assert.Contains("<h3>Comments (1)</h3>", html, StringComparison.Ordinal);
            Assert.Contains("<b>Bo</b>", html, StringComparison.Ordinal);
            Assert.Contains("<p class=\"text\">Check R7.</p>", html, StringComparison.Ordinal);
            Assert.Contains($"{Site}/account/sign-in?return=%2Fcollections%2Fk7mq3x9vbn2p%23comments-U7-12", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<textarea", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"cap-sub\">/RAS | 1 comment</div>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Signed_in_each_capture_has_a_form_carrying_the_session_s_token_and_its_capture()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7));

            Assert.Contains($"<form method=\"post\" action=\"{Site}/collections/k7mq3x9vbn2p/comments\">", html, StringComparison.Ordinal);
            Assert.Contains("<input type=\"hidden\" name=\"csrf\" value=\"0123456789abcdef0123456789abcdef\">", html, StringComparison.Ordinal);
            Assert.Contains($"<input type=\"hidden\" name=\"image\" value=\"{ShaA}\">", html, StringComparison.Ordinal);
            Assert.Contains("Signed in as <b>Bo</b>", html, StringComparison.Ordinal);
            Assert.Contains($"action=\"{Site}/account/sign-out\"", html, StringComparison.Ordinal);
        }

        // The owner's comment says so; that alone ties a name to the collection, and only when the owner chooses to comment.
        [Fact]
        public void The_owner_s_and_an_administrator_s_comments_are_marked()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView
            {
                Comments = [Comment(5, 42, "Anna", "Thanks"), Comment(6, 9, "Dennis@CRT", "Moved", administrator: true)]
            });

            Assert.Contains("<b>Anna</b> <span class=\"badge\">Owner</span>", html, StringComparison.Ordinal);
            Assert.Contains("<b>Dennis@CRT</b> <span class=\"badge\">Administrator</span>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_comment_s_text_is_shown_as_typed()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView
            {
                Comments = [Comment(5, 7, "<i>Bo</i>", "<script>alert(1)</script> https://spam.example")]
            });

            Assert.DoesNotContain("<script>alert", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<i>Bo</i>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"https://spam.example", html, StringComparison.Ordinal);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt; https://spam.example", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Its_writer_may_edit_and_delete_it_and_the_owner_hide_it()
        {
            CollectionCommentRecord helpers = Comment(5, 7, "Bo", "Check R7.", edited: Created.AddHours(4));

            string asWriter = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7, helpers));
            string asOwner = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(42, helpers));

            Assert.Contains($"href=\"{Site}/collections/k7mq3x9vbn2p?edit=5#comment-5\">Edit</a>", asWriter, StringComparison.Ordinal);
            Assert.Contains($"href=\"{Site}/collections/k7mq3x9vbn2p?delete=5#comment-5\">Delete</a>", asWriter, StringComparison.Ordinal);
            Assert.DoesNotContain(">Hide</button>", asWriter, StringComparison.Ordinal);
            Assert.Contains(" - edited", asWriter, StringComparison.Ordinal);

            Assert.Contains($"action=\"{Site}/collections/k7mq3x9vbn2p/comments/5/hide\"", asOwner, StringComparison.Ordinal);
            Assert.DoesNotContain("?edit=5", asOwner, StringComparison.Ordinal);
        }

        [Fact]
        public void A_hidden_comment_shows_its_owner_that_only_they_see_it_and_the_way_back()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(42, Comment(5, 7, "Bo", "Spam", hidden: true)));

            Assert.Contains(CollectionPages.HiddenFlag, html, StringComparison.Ordinal);
            Assert.Contains($"action=\"{Site}/collections/k7mq3x9vbn2p/comments/5/show\"", html, StringComparison.Ordinal);
            Assert.Contains(">Show again</button>", html, StringComparison.Ordinal);

            // The grid of captures counts what the capture's section shows this reader - the owner sees the hidden one, so it
            // counts (code review, 2026-11-27: the two counts differed).
            Assert.Contains("<div class=\"cap-sub\">/RAS | 1 comment</div>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_comment_being_edited_is_a_form_and_one_being_deleted_asks_first()
        {
            CollectionCommentRecord helpers = Comment(5, 7, "Bo", "Check R7.");

            string editing = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7, helpers) with { EditingId = 5 });
            string deleting = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7, helpers) with { DeletingId = 5 });

            Assert.Contains($"action=\"{Site}/collections/k7mq3x9vbn2p/comments/5/edit\"", editing, StringComparison.Ordinal);
            Assert.Contains(">Check R7.</textarea>", editing, StringComparison.Ordinal);

            Assert.Contains("Delete this comment?", deleting, StringComparison.Ordinal);
            Assert.Contains($"action=\"{Site}/collections/k7mq3x9vbn2p/comments/5/delete\"", deleting, StringComparison.Ordinal);
        }

        [Fact]
        public void A_refused_comment_is_said_with_what_was_typed_put_back()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7) with
            {
                ErrorOn = ShaA,
                Errors = ["An account writes at most 100 comments a day - write the rest tomorrow."],
                Typed = "My long answer"
            });

            Assert.Contains(">My long answer</textarea>", html, StringComparison.Ordinal);
            Assert.Contains("An account writes at most 100 comments a day", html, StringComparison.Ordinal);
        }

        [Fact]
        public void The_policy_lets_the_page_s_forms_post_to_this_site_alone() =>
            Assert.Contains("form-action 'self'", CollectionPages.ContentSecurityPolicy([]), StringComparison.Ordinal);

        // ---- "Your collections" and the administrator's delete (step 4) ---------------------

        private static readonly CollectionViewer Bo = new(7, "Bo", false);

        private static readonly CollectionViewer Administrator = new(9, "Dennis@CRT", true);

        private const string Token = "0123456789abcdef0123456789abcdef";

        private static CollectionRecord Listed(string id, string title, int captures, int comments) =>
            Collection() with { Id = id, Title = title, ImageCount = captures, CommentCount = comments };

        [Fact]
        public void Signed_in_a_page_leads_to_your_collections()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7));

            Assert.Contains($"<a href=\"{Site}/collections/\">Your collections</a>", html, StringComparison.Ordinal);
        }

        // The owner deletes their own in CRT, which keeps its window in step - the web offers it to the administrator alone.
        [Fact]
        public void Only_the_administrator_is_offered_to_delete_the_collection_and_is_asked_first()
        {
            CollectionRecord collection = Collection(Capture(ShaA, "U7", "12"));
            var asAdministrator = new CollectionPageView { Viewer = Administrator, Csrf = Token };

            Assert.DoesNotContain("?remove=1", Page(collection, view: SignedIn(42)), StringComparison.Ordinal);
            Assert.DoesNotContain("?remove=1", Page(collection, view: SignedIn(7)), StringComparison.Ordinal);

            string offered = Page(collection, view: asAdministrator);

            Assert.Contains($"href=\"{Site}/collections/k7mq3x9vbn2p?remove=1\">Delete this collection</a>", offered, StringComparison.Ordinal);
            Assert.DoesNotContain("/delete\"", offered, StringComparison.Ordinal);

            string asked = Page(collection, view: asAdministrator with { RemovingCollection = true });

            Assert.Contains(CollectionPages.RemoveQuestion, asked, StringComparison.Ordinal);
            Assert.Contains($"<form class=\"admin\" method=\"post\" action=\"{Site}/collections/k7mq3x9vbn2p/delete\">", asked, StringComparison.Ordinal);
            Assert.Contains($"<input type=\"hidden\" name=\"csrf\" value=\"{Token}\">", asked, StringComparison.Ordinal);
        }

        // A collection asked to be deleted by somebody else is still only offered - the question is the administrator's.
        [Fact]
        public void Nobody_else_is_asked_to_delete_it_whatever_the_link_says() =>
            Assert.DoesNotContain(CollectionPages.RemoveQuestion, Page(Collection(Capture(ShaA, "U7", "12")), view: SignedIn(7) with { RemovingCollection = true }), StringComparison.Ordinal);

        [Fact]
        public void Your_collections_marks_one_whose_sharing_its_owner_stopped()
        {
            string html = CollectionPages.List(Site, Bo, Token, new CollectionListView(
                [Listed("annafirst222", "My PLA", 1, 0) with { UnsharedUtc = Created.AddDays(1) }, Listed("annasecond22", "My VIC", 1, 0)],
                [],
                null,
                null));

            Assert.Equal(1, html.Split($"<span class=\"flag\">{CollectionPages.NotSharedFlag}</span>").Length - 1);
            Assert.Contains($"My PLA</a> <span class=\"flag\">{CollectionPages.NotSharedFlag}</span>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Your_collections_lists_the_reader_s_own_and_those_they_commented_on()
        {
            string html = CollectionPages.List(Site, Bo, Token, new CollectionListView(
                [Listed("annafirst222", "My PLA", 1, 0)],
                [Listed("bertfirst222", "Bert's VIC", 3, 2)],
                null,
                null));

            Assert.Contains("<h1>Your collections</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h2>Shared by you (1)</h2>", html, StringComparison.Ordinal);
            Assert.Contains($"<a href=\"{Site}/collections/annafirst222\">My PLA</a>", html, StringComparison.Ordinal);
            Assert.Contains("Commodore 64 - 250407 (PAL) | 1 capture | 0 comments | Shared 2026-November-26 | Deleted on 2027-November-26", html, StringComparison.Ordinal);
            Assert.Contains("<h2>You commented on (1)</h2>", html, StringComparison.Ordinal);
            Assert.Contains("Bert&#39;s VIC", html, StringComparison.Ordinal);
            Assert.Contains("3 captures | 2 comments", html, StringComparison.Ordinal);

            // The administrator's parts, and anybody's address, are no part of it.
            Assert.DoesNotContain("Every collection", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Newest comments", html, StringComparison.Ordinal);
            Assert.DoesNotContain(" | By ", html, StringComparison.Ordinal);

            // Signed in, the way out comes back here; the list does not link to itself.
            Assert.Contains("<input type=\"hidden\" name=\"return\" value=\"/collections/\">", html, StringComparison.Ordinal);
            Assert.DoesNotContain(">Your collections</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void An_empty_list_says_where_collections_are_made()
        {
            string html = CollectionPages.List(Site, Bo, Token, new CollectionListView([], [], null, null));

            Assert.Contains("<h2>Shared by you (0)</h2>", html, StringComparison.Ordinal);
            Assert.Contains("under &quot;Collect captures into&quot;", html, StringComparison.Ordinal);
            Assert.Contains("<h2>You commented on (0)</h2>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void The_administrator_sees_every_collection_with_whose_it_is_and_the_newest_comments()
        {
            CollectionRecord anna = Listed("annafirst222", "Anna's PLA", 1, 1);

            string html = CollectionPages.List(Site, Administrator, Token, new CollectionListView(
                [],
                [],
                [new CollectionListed(anna, "Anna", "anna@example.com", false), new CollectionListed(Listed("cyfirst22222", "Cy's", 1, 0), "Cy", "cy@example.com", true)],
                [Comment(5, 7, "Bo", "Check R7.", hidden: true) with { CollectionId = "annafirst222" }]));

            Assert.Contains("<h2>Every collection (2) <span class=\"badge\">Administrator</span></h2>", html, StringComparison.Ordinal);
            Assert.Contains(" | By Anna (anna@example.com)", html, StringComparison.Ordinal);
            Assert.Contains($"Cy&#39;s</a> <span class=\"flag\">{CollectionPages.OwnerLockedFlag}</span>", html, StringComparison.Ordinal);

            Assert.Contains("<h2>Newest comments (1) <span class=\"badge\">Administrator</span></h2>", html, StringComparison.Ordinal);
            Assert.Contains("<b>Bo</b>", html, StringComparison.Ordinal);
            Assert.Contains("<span class=\"flag\">Hidden</span>", html, StringComparison.Ordinal);
            Assert.Contains($"On <a href=\"{Site}/collections/annafirst222#comment-5\">Anna&#39;s PLA</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Every_word_typed_is_shown_as_text_in_the_list()
        {
            string html = CollectionPages.List(Site, Administrator with { DisplayName = "<i>Dennis</i>" }, Token, new CollectionListView(
                [Listed("annafirst222", "<script>alert(1)</script>", 1, 0)],
                [],
                [new CollectionListed(Listed("annafirst222", "<script>alert(1)</script>", 1, 0), "<b>Anna</b>", "a@example.com", false)],
                [Comment(5, 7, "<i>Cy</i>", "<img src=x onerror=alert(1)>") with { CollectionId = "annafirst222" }],
                [new LockedAccount(8, "<u>Spam</u>", "<s>spam@example.com</s>")]),
                locking: new LockedAccount(7, "<em>Bo</em>", "bo@example.com"),
                problem: "<b>Refused</b>");

            Assert.DoesNotContain("<script>alert", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<i>Dennis</i>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<b>Anna</b>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<i>Cy</i>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<img src=x", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<u>Spam</u>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<s>spam", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<em>Bo</em>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<b>Refused</b>", html, StringComparison.Ordinal);
        }

        // ---- Locking an account from the list (AccountLockFlows; owner decision, 2026-11-27) ---------

        private static CollectionListView AdministratorsList(bool ownerLocked = false, bool ownerIsAdministrator = false, bool writerLocked = false) =>
            new(
                [],
                [],
                [new CollectionListed(Listed("annafirst222", "Anna's PLA", 1, 1) with { AccountId = 42 }, "Anna", "anna@example.com", ownerLocked, ownerIsAdministrator)],
                [Comment(5, 7, "Bo", "Buy cheap chips") with { CollectionId = "annafirst222", AuthorIsLocked = writerLocked }],
                []);

        [Fact]
        public void The_administrator_is_offered_to_lock_an_owner_and_a_writer()
        {
            string html = CollectionPages.List(Site, Administrator, Token, AdministratorsList());

            Assert.Contains($" | By Anna (anna@example.com) | <a href=\"{Site}/collections/?lock=42\">Lock account</a>", html, StringComparison.Ordinal);
            Assert.Contains($" | <a href=\"{Site}/collections/?lock=7\">Lock account</a>", html, StringComparison.Ordinal);
        }

        // An administrator's account is never locked from here, and one locked already says so instead.
        [Fact]
        public void No_lock_is_offered_for_an_administrator_or_an_account_locked_already()
        {
            Assert.DoesNotContain("?lock=42", CollectionPages.List(Site, Administrator, Token, AdministratorsList(ownerIsAdministrator: true)), StringComparison.Ordinal);
            Assert.DoesNotContain("?lock=42", CollectionPages.List(Site, Administrator, Token, AdministratorsList(ownerLocked: true)), StringComparison.Ordinal);

            string writerLocked = CollectionPages.List(Site, Administrator, Token, AdministratorsList(writerLocked: true));

            Assert.DoesNotContain("?lock=7", writerLocked, StringComparison.Ordinal);
            Assert.Contains($"<span class=\"flag\">{CollectionPages.AccountLockedFlag}</span>", writerLocked, StringComparison.Ordinal);

            CollectionListView administratorWrote = AdministratorsList() with
            {
                NewestComments = [Comment(6, 9, "Dennis@CRT", "Moved", administrator: true) with { CollectionId = "annafirst222" }]
            };

            Assert.DoesNotContain("?lock=9", CollectionPages.List(Site, Administrator, Token, administratorWrote), StringComparison.Ordinal);
        }

        // An owner whose account is gone was never locked - the flag says so, and no lock is offered (code review, 2026-11-27).
        [Fact]
        public void A_collection_whose_owner_has_no_account_is_flagged_as_such()
        {
            CollectionListView list = AdministratorsList() with
            {
                All = [new CollectionListed(Listed("annafirst222", "Anna's PLA", 1, 1) with { AccountId = 42 }, CollectionListFlows.NoOwner, null, OwnerLocked: false, OwnerGone: true)]
            };

            string html = CollectionPages.List(Site, Administrator, Token, list);

            Assert.Contains($"Anna&#39;s PLA</a> <span class=\"flag\">{CollectionPages.OwnerGoneFlag}</span>", html, StringComparison.Ordinal);
            Assert.DoesNotContain(CollectionPages.OwnerLockedFlag, html, StringComparison.Ordinal);
            Assert.DoesNotContain("?lock=42", html, StringComparison.Ordinal);
        }

        // The administrator opens a collection offline for everybody else (CollectionFlows.FindShownToAsync): told so, its
        // comments still hidden or deleted there - but no comment form or "Edit", which its offline state would refuse.
        [Fact]
        public void An_offline_collection_s_page_says_so_to_the_administrator_and_offers_no_comment_form()
        {
            CollectionPageView view = new()
            {
                Viewer = Administrator,
                Csrf = Token,
                Comments = [Comment(5, 9, "Dennis@CRT", "Mine", administrator: true), Comment(6, 7, "Bo", "Spam") with { AuthorIsLocked = true }],
                Offline = CollectionOffline.OwnerLocked
            };

            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: view);

            Assert.Contains(System.Net.WebUtility.HtmlEncode(CollectionPages.OfflineOwnerLocked), html, StringComparison.Ordinal);
            Assert.DoesNotContain("/comments\">", html, StringComparison.Ordinal);
            Assert.DoesNotContain("?edit=5", html, StringComparison.Ordinal);
            Assert.Contains("?delete=6", html, StringComparison.Ordinal);
            Assert.Contains("/comments/6/hide", html, StringComparison.Ordinal);
            Assert.Contains("?remove=1", html, StringComparison.Ordinal);

            string gone = Page(Collection(Capture(ShaA, "U7", "12")), view: view with { Offline = CollectionOffline.OwnerGone });

            Assert.Contains(System.Net.WebUtility.HtmlEncode(CollectionPages.OfflineOwnerGone), gone, StringComparison.Ordinal);

            // Shown to all, the page says nothing of it.
            string online = Page(Collection(Capture(ShaA, "U7", "12")), view: view with { Offline = CollectionOffline.No });

            Assert.DoesNotContain("Offline:", online, StringComparison.Ordinal);
            Assert.Contains("/comments\">", online, StringComparison.Ordinal);
        }

        [Fact]
        public void Locking_is_asked_first_saying_what_it_does()
        {
            string html = CollectionPages.List(Site, Administrator, Token, AdministratorsList(), locking: new LockedAccount(7, "Bo", "bo@example.com"));

            Assert.Contains($"<form class=\"admin\" method=\"post\" action=\"{Site}/collections/accounts/7/lock\">", html, StringComparison.Ordinal);
            Assert.Contains("Lock the account of <b>Bo</b> (bo@example.com)?", html, StringComparison.Ordinal);
            Assert.Contains("its comments are hidden from everybody but you", html, StringComparison.Ordinal);
            Assert.Contains($"<input type=\"hidden\" name=\"csrf\" value=\"{Token}\">", html, StringComparison.Ordinal);
            Assert.Contains(">Lock account</button>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void The_accounts_locked_are_listed_each_with_its_unlock()
        {
            string html = CollectionPages.List(Site, Administrator, Token, AdministratorsList() with { Locked = [new LockedAccount(8, "Spammer", "spam@example.com")] });

            Assert.Contains("<h2 id=\"locked\">Locked accounts (1) <span class=\"badge\">Administrator</span></h2>", html, StringComparison.Ordinal);
            Assert.Contains("<b>Spammer</b> <span class=\"muted\">(spam@example.com)</span>", html, StringComparison.Ordinal);
            Assert.Contains($"<form class=\"inline\" method=\"post\" action=\"{Site}/collections/accounts/8/unlock\">", html, StringComparison.Ordinal);
            Assert.Contains(">Unlock</button>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_refused_lock_is_said_at_the_top()
        {
            string html = CollectionPages.List(Site, Administrator, Token, AdministratorsList(), problem: AccountLockFlows.AdministratorsStay);

            Assert.Contains($"<div class=\"errors\" role=\"alert\"><p>{System.Net.WebUtility.HtmlEncode(AccountLockFlows.AdministratorsStay)}</p></div>", html, StringComparison.Ordinal);
        }

        // What only the administrator may see stays theirs, whatever a list handed to the page holds.
        [Fact]
        public void Nobody_else_sees_owners_comments_locks_or_questions()
        {
            string html = CollectionPages.List(
                Site,
                Bo,
                Token,
                AdministratorsList() with { Locked = [new LockedAccount(8, "Spammer", "spam@example.com")] },
                locking: new LockedAccount(7, "Bo", "bo@example.com"),
                problem: "Refused");

            Assert.DoesNotContain("Every collection", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Newest comments", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Locked accounts", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Lock the account", html, StringComparison.Ordinal);
            Assert.DoesNotContain("anna@example.com", html, StringComparison.Ordinal);
            Assert.DoesNotContain("Refused", html, StringComparison.Ordinal);
        }

        // On a collection's page the administrator - the one who sees a locked account's comments - is told why nobody else does.
        [Fact]
        public void A_locked_account_s_comment_is_marked_for_the_administrator()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12")), view: new CollectionPageView
            {
                Viewer = Administrator,
                Csrf = Token,
                Comments = [Comment(5, 7, "Bo", "Buy cheap chips") with { AuthorIsLocked = true }]
            });

            Assert.Contains($"<span class=\"flag\">{CollectionPages.AccountLockedFlag} - only you see it</span>", html, StringComparison.Ordinal);
        }

        // A long comment is cut in the list - never through a character written as two (an emoji), which would show as junk.
        [Fact]
        public void A_long_comment_is_cut_in_the_list_and_never_half_a_character()
        {
            string body = new string('a', 299) + "\U0001F600" + new string('b', 50);

            string html = CollectionPages.List(Site, Administrator, Token, new CollectionListView([], [], [], [Comment(5, 7, "Bo", body)]));

            Assert.Contains(new string('a', 299) + "...", html, StringComparison.Ordinal);
            Assert.DoesNotContain("\uD83D", html, StringComparison.Ordinal);
            Assert.DoesNotContain("bbb", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Signed_out_the_list_says_how_to_sign_in_and_come_back_to_it()
        {
            Assert.Contains($"{Site}/account/sign-in?return=%2Fcollections%2F\">Sign in</a>", CollectionPages.ListSignedOut(Site), StringComparison.Ordinal);
            Assert.Contains("sign-in?return=%2Fcrt-test%2Fcollections%2F", CollectionPages.ListSignedOut(Site + "/crt-test"), StringComparison.Ordinal);
        }

        // ---- Where a baseline's picture is served ------------------------------------------

        [Fact]
        public void A_baseline_is_linked_where_the_published_data_serves_it_only_when_it_is_there()
        {
            string stored = DataTreeLayout.StoredFilePath(BaselineSha, "Scope baseline/U7_12_RAS.png");

            var baselines = new CollectionBaselines(this.thisTree, "https://classic-repair-toolbox.dk/app-data-v4/Data/");
            var baseline = new CollectionBaseline(BaselineSha, "Scope baseline/U7_12_RAS.png", null, null, null);

            Assert.Null(baselines.UrlOf(baseline));

            Directory.CreateDirectory(Path.GetDirectoryName(Path.Combine(this.thisTree, stored))!);
            File.WriteAllBytes(Path.Combine(this.thisTree, stored), [1]);

            Assert.Equal($"https://classic-repair-toolbox.dk/app-data-v4/Data/{stored}", baselines.UrlOf(baseline));
            Assert.Equal("https://classic-repair-toolbox.dk", baselines.Origin);
            Assert.Null(baselines.UrlOf(baseline with { Sha256 = "../../x" }));
        }
    }
}
