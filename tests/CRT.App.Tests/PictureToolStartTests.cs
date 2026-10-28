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

    // ---- "Done - next picture" (step 3, slice 3d) -----------------------------------------------------

    private static readonly string[] Pictures = ["Layout", "Top", "Bottom", "Sheet 1", "Sheet 2"];

    private static string? Next(string current, string[] withView, params string[] aligned) =>
        PictureToolStart.NextToAlign(
            Pictures,
            current,
            picture => withView.Contains(picture, StringComparer.OrdinalIgnoreCase),
            picture => aligned.Contains(picture, StringComparer.OrdinalIgnoreCase));

    // The next picture after this one with a KiCad view and no alignment yet - a picture with no
    // view (a photo of the layout) is passed over.
    [Fact]
    public void The_next_picture_is_the_next_with_a_KiCad_view_not_aligned_yet()
    {
        string[] withView = ["Top", "Bottom", "Sheet 1", "Sheet 2"];

        Assert.Equal("Bottom", Next("Top", withView));
        Assert.Equal("Sheet 2", Next("Top", withView, "Bottom", "Sheet 1"));
        Assert.Equal("Top", Next(" layout ", withView));
    }

    // Round to the start: the last picture's next is the first not aligned.
    [Fact]
    public void After_the_last_picture_the_next_is_found_from_the_start()
    {
        Assert.Equal("Top", Next("Sheet 2", ["Top", "Bottom", "Sheet 1", "Sheet 2"], "Bottom"));
    }

    // Nothing left but this one: no next - Done is the way on. This one itself is never its own next,
    // aligned or not.
    [Fact]
    public void With_every_other_picture_aligned_there_is_no_next()
    {
        Assert.Null(Next("Top", ["Top", "Bottom"], "Bottom"));
        Assert.Null(Next("Top", ["Top"]));
        Assert.Null(PictureToolStart.NextToAlign([], "Top", _ => true, _ => false));
    }

    // A picture not in the list (one hidden meanwhile) starts the search from the first.
    [Fact]
    public void A_picture_not_listed_looks_from_the_first()
    {
        Assert.Equal("Top", Next("Gone", ["Top", "Bottom"]));
    }
}
