using System.Security.Cryptography;
using System.Text;
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

        private static CollectionImageRecord Capture(string sha, string label, string pin, CollectionBaseline? baseline = null, bool looksWrong = false, string? caption = null) =>
            new(new CollectionImageEntry(sha, label, pin, "/RAS", "Pulsing", "100us", "2V", "1V", baseline, caption, looksWrong, Created), 800, 480);

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
            Assert.Contains("Deleted on 2027-November-26", html, StringComparison.Ordinal);
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

        [Fact]
        public void A_capture_the_owner_marked_says_it_looks_wrong_in_the_list_and_at_itself()
        {
            string html = Page(Collection(Capture(ShaA, "U7", "12", looksWrong: true)));

            Assert.Equal(2, html.Split("<span class=\"flag\">Looks wrong</span>").Length - 1);
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

        [Theory]
        [InlineData(CollectionDeletedReasons.Expired, "Its time was up")]
        [InlineData(CollectionDeletedReasons.Owner, "The person who shared it deleted it.")]
        [InlineData(CollectionDeletedReasons.Administrator, "It was removed.")]
        public void A_deleted_collection_says_when_it_went_and_why(string reason, string why)
        {
            CollectionRecord gone = Collection() with { DeletedUtc = Created.AddMonths(2), DeletedReason = reason, Title = null };

            string html = CollectionPages.Deleted(Site, gone);

            Assert.Contains("was deleted on 2027-January-26", html, StringComparison.Ordinal);
            Assert.Contains(why, html, StringComparison.Ordinal);
        }

        // The policy allows the page's one style by its hash - a style changed without it would be blocked, so the two are one.
        [Fact]
        public void The_policy_allows_the_page_s_own_style_and_no_script()
        {
            string policy = CollectionPages.ContentSecurityPolicy([Site]);
            string hash = "sha256-" + Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(CollectionPages.Css)));

            Assert.Contains($"style-src '{hash}'", policy, StringComparison.Ordinal);
            Assert.Contains($"img-src 'self' {Site};", policy, StringComparison.Ordinal);
            Assert.Contains("default-src 'none'", policy, StringComparison.Ordinal);
            Assert.DoesNotContain("script-src", policy, StringComparison.Ordinal);
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
            Assert.Contains("(1 comment)", html, StringComparison.Ordinal);
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

            Assert.Contains("Hidden - only you see it", html, StringComparison.Ordinal);
            Assert.Contains($"action=\"{Site}/collections/k7mq3x9vbn2p/comments/5/show\"", html, StringComparison.Ordinal);
            Assert.Contains(">Show again</button>", html, StringComparison.Ordinal);

            // Hidden comments are not counted in the list of captures.
            Assert.DoesNotContain("(1 comment)", html, StringComparison.Ordinal);
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
