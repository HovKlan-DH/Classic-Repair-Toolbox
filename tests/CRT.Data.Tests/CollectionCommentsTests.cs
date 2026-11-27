using Handlers.DataHandling;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // A comment's text (CollectionComments): needed, at most MaximumLength characters as it is kept, its line ends "\n".
    // ###########################################################################################
    public sealed class CollectionCommentsTests
    {
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   \r\n  ")]
        public void A_comment_with_no_words_is_refused(string? body) =>
            Assert.Equal(CollectionComments.Missing, CollectionComments.Problem(body));

        // A browser posts every line end as "\r\n" while its maxlength counts one character - so a comment the page let through
        // must not be refused for its line ends (code review, 2026-11-27).
        [Fact]
        public void Line_ends_count_as_one_character_as_the_page_s_box_counts_them()
        {
            string tenLines = string.Join("\r\n", Enumerable.Repeat(new string('a', 199), 10));

            Assert.Equal(1999, CollectionComments.Tidied(tenLines).Length);
            Assert.Null(CollectionComments.Problem(tenLines));
        }

        [Fact]
        public void A_comment_longer_than_the_limit_as_kept_is_refused()
        {
            string text = new('a', CollectionComments.MaximumLength + 1);

            Assert.Equal($"A comment is at most {CollectionComments.MaximumLength} characters - this one has {text.Length}.", CollectionComments.Problem(text));
            Assert.Null(CollectionComments.Problem(new string('a', CollectionComments.MaximumLength)));
        }

        [Fact]
        public void A_comment_is_kept_trimmed_with_its_line_ends_one_kind() =>
            Assert.Equal("one\ntwo\nthree", CollectionComments.Tidied("  one\r\ntwo\rthree \n"));
    }
}
