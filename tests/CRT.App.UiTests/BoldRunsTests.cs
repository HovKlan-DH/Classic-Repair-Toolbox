using Avalonia.Controls;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// BoldRuns.Marked - a sentence saying its own bold pieces as "<b>...</b>" (2026-11-03), so a wording
// class, which holds no control, can say what is bold. No word may be lost to a slip in the tags.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoldRunsTests
{
    [Fact]
    public void The_pieces_between_the_tags_are_bold()
    {
        Assert.Equal(
            [("Whether the board comes in a ", false), ("PAL", true), (" and ", false), ("NTSC", true), (" version.", false)],
            BoldRuns.Marked("Whether the board comes in a <b>PAL</b> and <b>NTSC</b> version."));

        Assert.Equal([("3", true), (" files", false)], BoldRuns.Marked("<b>3</b> files"));
        Assert.Equal([("Added ", false), ("3", true)], BoldRuns.Marked("Added <b>3</b>"));
    }

    [Fact]
    public void A_sentence_with_no_tags_is_one_plain_piece_and_nothing_is_no_piece()
    {
        Assert.Equal([("No regions.", false)], BoldRuns.Marked("No regions."));
        Assert.Empty(BoldRuns.Marked(string.Empty));
        Assert.Empty(BoldRuns.Marked(null));
    }

    [Theory]
    [InlineData("Half <b>open", "Half open")]
    [InlineData("Half closed</b> here", "Half closed here")]
    [InlineData("A <b>bold</b> and a stray</b>", "A bold and a stray")]
    [InlineData("<b>a<b>b</b>", "ab")]
    [InlineData("<b>a<b>b</b>c</b>", "abc")]
    public void A_tag_with_no_partner_loses_no_word(string marked, string read)
    {
        Assert.Equal(read, BoldRuns.Plain(marked));
    }

    // A tag inside another is read as if it were not there, as the header says (code review, 2026-11-04: it stood
    // in the bold piece, "<b>" on screen).
    [Fact]
    public void A_tag_inside_another_is_not_shown()
    {
        Assert.Equal([("ab", true)], BoldRuns.Marked("<b>a<b>b</b>"));
    }

    // Marked pieces on a block: Text null, Inlines holding them, read through TextOf - plain, plain Text.
    [Fact]
    public void A_block_shows_them_and_a_line_break_stays_in_its_piece()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();

            BoldRuns.ShowMarked(block, "The board now has <b>25</b> KiCad files:\n- 1 PCB file");

            Assert.Null(block.Text);
            Assert.Equal("The board now has 25 KiCad files:\n- 1 PCB file", BoldRuns.TextOf(block));
            Assert.Equal(["25"], BoldRuns.BoldPiecesOf(block));

            BoldRuns.ShowMarked(block, "Nothing bold");

            Assert.Equal("Nothing bold", block.Text);
            Assert.Empty(BoldRuns.BoldPiecesOf(block));
        });
    }
}
