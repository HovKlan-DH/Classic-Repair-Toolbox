using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PictureToolStart - when "Mark components" or "Align KiCad", asked for on a picture from the
// Contribute tab, can start on the Schematics tab (CRT 4.0 step 2, slice 6). The picture's image and
// the board's KiCad data are read in the background: the tool waits for what it works on, and what
// can never come is refused with the reason - a tool left waiting for ever is the failure to avoid.
// ###########################################################################################
public sealed class PictureToolStartTests
{
    private static PictureToolStart Decide(
        PictureTool tool,
        bool onBoard = true,
        bool selected = true,
        PictureImage image = PictureImage.Shown,
        BoardKiCad kiCad = BoardKiCad.Loaded,
        bool hasView = true) =>
        PictureToolStart.Decide(tool, " Top ", onBoard, selected, image, kiCad, hasView);

    [Fact]
    public void Marking_starts_once_the_picture_is_shown_whatever_the_KiCad_data()
    {
        Assert.Equal(PictureToolStep.Start, Decide(PictureTool.MarkComponents, kiCad: BoardKiCad.None, hasView: false).Step);
        Assert.Equal(PictureToolStep.Start, Decide(PictureTool.MarkComponents, kiCad: BoardKiCad.Loading).Step);
    }

    // The picture not yet selected, or its image still being read: wait.
    [Fact]
    public void A_picture_not_yet_shown_is_waited_for()
    {
        Assert.Equal(PictureToolStep.Wait, Decide(PictureTool.MarkComponents, selected: false).Step);
        Assert.Equal(PictureToolStep.Wait, Decide(PictureTool.MarkComponents, image: PictureImage.Loading).Step);
        Assert.Equal(PictureToolStep.Wait, Decide(PictureTool.AlignKiCad, image: PictureImage.Loading).Step);
    }

    // What can never come is refused, the reason naming the picture - never waited for.
    [Fact]
    public void A_picture_not_on_the_board_or_whose_image_cannot_be_read_is_refused()
    {
        PictureToolStart gone = Decide(PictureTool.MarkComponents, onBoard: false, selected: false);
        Assert.Equal(PictureToolStep.Refuse, gone.Step);
        Assert.Equal("[Top] is not a picture of the board shown.", gone.Reason);

        PictureToolStart missing = Decide(PictureTool.AlignKiCad, image: PictureImage.Missing);
        Assert.Equal(PictureToolStep.Refuse, missing.Step);
        Assert.Equal("The image of [Top] could not be read.", missing.Reason);
    }

    [Fact]
    public void Aligning_waits_for_the_KiCad_data_and_needs_a_view_of_it_for_the_picture()
    {
        Assert.Equal(PictureToolStep.Wait, Decide(PictureTool.AlignKiCad, kiCad: BoardKiCad.Loading, hasView: false).Step);
        Assert.Equal(PictureToolStep.Start, Decide(PictureTool.AlignKiCad).Step);

        PictureToolStart none = Decide(PictureTool.AlignKiCad, kiCad: BoardKiCad.None, hasView: false);
        Assert.Equal((PictureToolStep.Refuse, PictureToolWording.NoKiCadData), (none.Step, none.Reason));

        PictureToolStart noView = Decide(PictureTool.AlignKiCad, hasView: false);
        Assert.Equal(PictureToolStep.Refuse, noView.Step);
        Assert.StartsWith("[Top] names no view of the board's KiCad data", noView.Reason, StringComparison.Ordinal);
    }
}
