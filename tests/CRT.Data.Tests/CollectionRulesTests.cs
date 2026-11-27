using Handlers.DataHandling;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // What a shared collection is (CollectionRules) - the rules CRT checks before sharing and the server as it takes one.
    // ###########################################################################################
    public class CollectionRulesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

        private static readonly string Sha = new('a', 64);

        internal static CollectionImageEntry Image(string sha, string label = "U7", string? pin = "12") =>
            new(sha, label, pin, "/RAS", "Pulsing", "50us", "2V", "1V", null, null, LooksWrong: false, CapturedUtc: null);

        internal static CollectionContentRequest Content(params CollectionImageEntry[] images) =>
            new("Commodore 64", "250407", "PAL", "No picture", "It went black.", [], CollectionRules.DeleteAt(Now, "1y"), images);

        [Fact]
        public void A_new_id_is_twelve_characters_none_of_which_looks_like_another()
        {
            for (int i = 0; i < 200; i++)
            {
                string id = CollectionRules.NewId();

                Assert.True(CollectionRules.IsId(id), id);
                Assert.DoesNotContain(id, c => c is '0' or 'o' or '1' or 'l' || char.IsUpper(c));
            }
        }

        [Theory]
        [InlineData("abcdefghijk")]
        [InlineData("abcdefghijkmn")]
        [InlineData("abcdefghijk0")]
        [InlineData("ABCDEFGHIJKM")]
        [InlineData("abcdefghij/x")]
        [InlineData(null)]
        public void Anything_else_is_no_id(string? id) => Assert.False(CollectionRules.IsId(id));

        [Theory]
        [InlineData("1m", 2026, 12, 26)]
        [InlineData("6m", 2027, 5, 26)]
        [InlineData("1y", 2027, 11, 26)]
        [InlineData("5y", 2031, 11, 26)]
        [InlineData("2y", 2027, 11, 26)]
        [InlineData(null, 2027, 11, 26)]
        public void Each_lifetime_counts_from_the_day_it_is_chosen_and_anything_else_is_a_year(string? key, int year, int month, int day) =>
            Assert.Equal(new DateTimeOffset(year, month, day, 12, 0, 0, TimeSpan.Zero), CollectionRules.DeleteAt(Now, key));

        // Owner decision, 2026-11-25: "no need to show the "2 years"".
        [Fact]
        public void The_choices_are_a_month_six_months_a_year_and_five_years_a_year_the_default()
        {
            Assert.Equal(["1 month", "6 months", "1 year", "5 years"], CollectionRules.Lifetimes.Select(lifetime => lifetime.Words));
            Assert.Equal("1 year", CollectionRules.LifetimeOf(CollectionRules.DefaultLifetimeKey).Words);
        }

        [Fact]
        public void A_whole_collection_as_it_should_be_has_no_problem() =>
            Assert.Empty(CollectionRules.ProblemsOf(Content(Image(Sha)), Now));

        [Fact]
        public void One_saying_no_board_is_refused_saying_so()
        {
            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(Content(Image(Sha)) with { BoardName = " " }, Now);

            Assert.Equal(["The collection does not say the board it is about."], problems);
        }

        [Fact]
        public void One_with_no_capture_is_refused() =>
            Assert.Contains("A collection holds at least one capture.", CollectionRules.ProblemsOf(Content(), Now));

        // "images": [null] in a body is said - the server answered 500 for it (code review, 2026-11-27).
        [Fact]
        public void An_empty_capture_is_refused_saying_so() =>
            Assert.Contains("A capture is empty.", CollectionRules.ProblemsOf(Content(Image(Sha), null!), Now));

        [Fact]
        public void One_with_more_captures_than_a_collection_holds_is_refused()
        {
            CollectionImageEntry[] images = Enumerable.Range(0, CollectionRules.MaximumImages + 1)
                .Select(i => Image(i.ToString("x64", System.Globalization.CultureInfo.InvariantCulture), pin: i.ToString(System.Globalization.CultureInfo.InvariantCulture)))
                .ToArray();

            Assert.Contains(CollectionRules.ProblemsOf(Content(images), Now), problem => problem.StartsWith("A collection holds at most 300 captures", StringComparison.Ordinal));
        }

        [Fact]
        public void The_same_capture_twice_is_refused_naming_where() =>
            Assert.Contains("U8, pin 3: the same capture is in the collection twice.", CollectionRules.ProblemsOf(Content(Image(Sha), Image(Sha, "U8", "3")), Now));

        [Fact]
        public void A_capture_not_named_by_a_checksum_is_refused() =>
            Assert.Contains("U7, pin 12: the capture's checksum is not one.", CollectionRules.ProblemsOf(Content(Image("../../etc/passwd")), Now));

        [Fact]
        public void A_baseline_not_named_by_a_checksum_is_refused()
        {
            CollectionImageEntry image = Image(Sha) with { Baseline = new CollectionBaseline("nope", "Scope baseline/a.png", null, null, null) };

            Assert.Contains("U7, pin 12: the baseline's checksum is not one.", CollectionRules.ProblemsOf(Content(image), Now));
        }

        [Fact]
        public void A_date_passed_and_one_past_five_years_are_refused()
        {
            Assert.Contains(
                "The date it is deleted on has passed - choose how long to keep it again.",
                CollectionRules.ProblemsOf(Content(Image(Sha)) with { DeleteAtUtc = Now.AddMinutes(-1) }, Now));

            Assert.Contains(
                "A collection is kept at most 5 years ahead.",
                CollectionRules.ProblemsOf(Content(Image(Sha)) with { DeleteAtUtc = Now.AddYears(5).AddDays(2) }, Now));

            // Five years to the day is the longest choice - with a day of room for a clock a little ahead.
            Assert.Empty(CollectionRules.ProblemsOf(Content(Image(Sha)) with { DeleteAtUtc = CollectionRules.DeleteAt(Now.AddHours(20), "5y") }, Now));
        }

        [Theory]
        [InlineData("https://www.lemon64.com/forum/viewtopic.php?t=12345", true)]
        [InlineData("http://www.forum64.de/index.php?thread/1/", true)]
        [InlineData("https://www.reddit.com/r/c64/comments/abc/", true)]
        [InlineData("www.lemon64.com/forum", false)]
        [InlineData("javascript:alert(1)", false)]
        [InlineData("ftp://example.com/x", false)]
        [InlineData("https://example.com/a b", false)]
        [InlineData("", false)]
        [InlineData(null, false)]
        public void A_forum_link_is_an_http_address_with_a_host(string? link, bool isLink) =>
            Assert.Equal(isLink, CollectionRules.IsForumLink(link));

        [Fact]
        public void A_link_that_is_not_one_is_refused_and_too_many_are_too()
        {
            CollectionContentRequest content = Content(Image(Sha)) with
            {
                ForumLinks = [.. Enumerable.Repeat("https://www.lemon64.com/forum", CollectionRules.MaximumForumLinks), "javascript:alert(1)"]
            };

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, Now);

            Assert.Contains($"A collection links to at most {CollectionRules.MaximumForumLinks} forum threads.", problems);
            Assert.Contains("[javascript:alert(1)] is no web address - a forum thread's link starts with https://.", problems);
        }

        [Fact]
        public void Words_longer_than_their_limit_are_refused_naming_them()
        {
            CollectionContentRequest content = Content(Image(Sha) with { Caption = new string('x', CollectionRules.MaximumCaptionLength + 1) }) with
            {
                Title = new string('t', CollectionRules.MaximumTitleLength + 1)
            };

            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(content, Now);

            Assert.Contains($"The title is at most {CollectionRules.MaximumTitleLength} characters.", problems);
            Assert.Contains($"U7, pin 12: the note is at most {CollectionRules.MaximumCaptionLength} characters.", problems);
        }

        [Fact]
        public void Nothing_sent_is_said() => Assert.Equal(["Nothing was sent."], CollectionRules.ProblemsOf(null, Now));

        [Fact]
        public void A_capture_is_named_by_its_component_and_pin()
        {
            Assert.Equal("U7, pin 12", CollectionRules.WhereOf(Image(Sha)));
            Assert.Equal("C12", CollectionRules.WhereOf(Image(Sha, "C12", pin: null)));
        }

        [Fact]
        public void A_date_is_written_as_CRT_writes_dates() =>
            Assert.Equal("2026-November-26", CollectionRules.DateText(Now));

        [Fact]
        public void The_page_is_under_the_site_by_its_id()
        {
            Assert.Equal("https://classic-repair-toolbox.dk/collections/abc", CollectionRules.PageUrl("https://classic-repair-toolbox.dk", "abc"));
            Assert.Equal("https://classic-repair-toolbox.dk/crt-test/collections/abc", CollectionRules.PageUrl("https://classic-repair-toolbox.dk/crt-test/", "abc"));
        }

        // A capture's anchor never takes one of the page's own - a comment's "comment-12", its list's "comments-...", its box's
        // "comment-on-..." - so a mail's link to comment #12 never jumps to a component labelled "COMMENT" (code review, 2026-11-27).
        [Fact]
        public void A_capture_s_anchor_never_takes_a_comment_s()
        {
            static CollectionImageEntry Capture(string label, string? pin) =>
                new(new string('a', 64), label, pin, null, null, null, null, null, null, null, false, null);

            IReadOnlyList<string> anchors = CollectionRules.AnchorsOf(
                [Capture("COMMENT", "12"), Capture("Comments", "3"), Capture("comment-on", null), Capture("locked", null), Capture("U7", "12")]);

            Assert.Equal(["capture-COMMENT-12", "capture-Comments-3", "capture-comment-on", "capture-locked", "U7-12"], anchors);
        }
    }
}
