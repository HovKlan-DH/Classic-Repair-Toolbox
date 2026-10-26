using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PictureToolWording - the picture tools' bar, the buttons that start them, and the menu a
// right-click on a component opens (CRT 4.0 step 2, slice 6 - contributor mode went).
// ###########################################################################################
public sealed class PictureToolWordingTests
{
    [Fact]
    public void The_bar_names_the_tool_and_its_picture()
    {
        Assert.Equal("Mark components on [Top]", PictureToolWording.Title(PictureTool.MarkComponents, " Top "));
        Assert.Equal("Align KiCad on [Top]", PictureToolWording.Title(PictureTool.AlignKiCad, "Top"));
        Assert.Equal("Opening [Top]...", PictureToolWording.Opening("Top"));
        Assert.Equal("KiCad cannot be aligned on [Top]", PictureToolWording.CannotStart(PictureTool.AlignKiCad, "Top"));
    }

    // Each tool's help says how it is used, and that Done saves into the draft.
    [Fact]
    public void Each_tools_help_says_how_it_is_used_and_that_Done_saves_into_the_draft()
    {
        Assert.Contains("right-click it to remove it", PictureToolWording.Help(PictureTool.MarkComponents), StringComparison.Ordinal);
        Assert.Contains("Ctrl+Z undoes", PictureToolWording.Help(PictureTool.MarkComponents), StringComparison.Ordinal);
        Assert.Contains("Esc cancels", PictureToolWording.Help(PictureTool.AlignKiCad), StringComparison.Ordinal);
        Assert.All([PictureTool.MarkComponents, PictureTool.AlignKiCad], tool =>
            Assert.Contains("Done saves it", PictureToolWording.Help(tool), StringComparison.Ordinal));
    }

    [Fact]
    public void The_component_menu_says_what_its_highlight_item_does()
    {
        Assert.Equal("Highlight [U8]", PictureToolWording.Highlight(" U8 "));
        Assert.Equal("Take the highlight away from [U8]", PictureToolWording.TakeHighlightAway("U8"));
        Assert.Equal("Edit this component", PictureToolWording.EditComponent);
    }

    // Plain ASCII, no Excel, no "contributor mode" (it went), and no button label with "...".
    [Fact]
    public void Every_sentence_is_plain_ASCII_and_names_no_contributor_mode()
    {
        Assert.All(PictureToolWording.All(), sentence =>
        {
            Assert.All(sentence, character => Assert.True(character < 128, $"[{sentence}] is not plain ASCII"));
            Assert.DoesNotContain("Excel", sentence, StringComparison.Ordinal);
            Assert.DoesNotContain("contributor mode", sentence, StringComparison.OrdinalIgnoreCase);
        });

        Assert.All(
            [PictureToolWording.MarkButton, PictureToolWording.AlignButton, PictureToolWording.Done, PictureToolWording.Cancel, PictureToolWording.Close, PictureToolWording.EditComponent],
            label => Assert.DoesNotContain("...", label, StringComparison.Ordinal));
    }
}
