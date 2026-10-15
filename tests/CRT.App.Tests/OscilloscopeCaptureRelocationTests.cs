using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// OscilloscopeCaptureRelocation - which files in the data folder are oscilloscope captures the old
// default saved there, moving them beside it, and the folder test the orphan cleanup keeps away
// from (code review, 2026-10-13: the cleanup deleted every capture). DataManagerOrphanCleanupTests
// covers the cleanup itself.
public sealed class OscilloscopeCaptureRelocationTests : IDisposable
{
    private readonly TempWorkspace thisData = new();
    private readonly TempWorkspace thisBeside = new();

    public void Dispose()
    {
        this.thisData.Dispose();
        this.thisBeside.Dispose();
    }

    private string Target => this.thisBeside.Path_("Oscilloscope images");

    [Theory]
    [InlineData("U1_5.png", true)]
    [InlineData("U14_Pin 3_PAL.png", true)]
    [InlineData("U1_5.PNG", true)]
    [InlineData("U1.png", false)]
    [InlineData("_5.png", false)]
    [InlineData("U1_5.jpg", false)]
    [InlineData("Classic-Repair-Toolbox.v3.0.0.xlsx", false)]
    [InlineData("", false)]
    public void A_capture_is_a_png_named_board_label_underscore_pin(string fileName, bool expected)
    {
        Assert.Equal(expected, OscilloscopeCaptureRelocation.LooksLikeCapture(fileName));
    }

    [Theory]
    [InlineData("photo.jpg", true)]
    [InlineData("scope.png", true)]
    [InlineData("U1_5.png", false)]
    [InlineData("dataChecksums.json", false)]
    public void An_image_that_is_not_shaped_like_a_capture_is_in_doubt(string fileName, bool expected)
    {
        Assert.Equal(expected, OscilloscopeCaptureRelocation.IsImageInDoubt(fileName));
    }

    [Fact]
    public void Same_or_inside_is_a_folder_boundary_not_a_name_prefix()
    {
        string root = this.thisData.Path_("Data");

        Assert.True(OscilloscopeCaptureRelocation.IsSameOrInside(root, root));
        Assert.True(OscilloscopeCaptureRelocation.IsSameOrInside(Path.Combine(root, "a", "b.png"), root));
        Assert.True(OscilloscopeCaptureRelocation.IsSameOrInside(root + Path.DirectorySeparatorChar, root));

        // A sibling whose name merely starts with the folder's name is NOT inside it.
        Assert.False(OscilloscopeCaptureRelocation.IsSameOrInside(this.thisData.Path_("DataBackup", "x.png"), root));
        Assert.False(OscilloscopeCaptureRelocation.IsSameOrInside(this.thisData.Root, root));
        Assert.False(OscilloscopeCaptureRelocation.IsSameOrInside(string.Empty, root));
    }

    [Fact]
    public void Only_top_level_captures_the_manifest_does_not_list_are_moved()
    {
        this.thisData.WriteFile("U1_5.png", "capture");
        this.thisData.WriteFile("Listed_image.png", "data");
        this.thisData.WriteFile("Commodore/C64/U2_1.png", "deeper - published data lives there");
        this.thisData.WriteFile("master.xlsx", "workbook");

        var outcome = OscilloscopeCaptureRelocation.MoveCapturesOutOfDataRoot(
            this.thisData.Root,
            this.Target,
            name => name == "Listed_image.png");

        var moved = Assert.Single(outcome.Moved);
        Assert.Equal(Path.Combine(this.Target, "U1_5.png"), moved.To);
        Assert.Equal("capture", File.ReadAllText(moved.To));
        Assert.Empty(outcome.LeftInPlace);

        Assert.True(File.Exists(this.thisData.Path_("Listed_image.png")));
        Assert.True(File.Exists(this.thisData.Path_("Commodore", "C64", "U2_1.png")));
        Assert.True(File.Exists(this.thisData.Path_("master.xlsx")));
    }

    [Fact]
    public void A_name_already_taken_is_never_overwritten_and_is_left_in_place()
    {
        this.thisData.WriteFile("U1_5.png", "older");
        Directory.CreateDirectory(this.Target);
        File.WriteAllText(Path.Combine(this.Target, "U1_5.png"), "newer");

        var outcome = OscilloscopeCaptureRelocation.MoveCapturesOutOfDataRoot(this.thisData.Root, this.Target, _ => false);

        Assert.Empty(outcome.Moved);
        Assert.Equal(this.thisData.Path_("U1_5.png"), Assert.Single(outcome.LeftInPlace));
        Assert.Equal("newer", File.ReadAllText(Path.Combine(this.Target, "U1_5.png")));
        Assert.Equal("older", this.thisData.ReadFile("U1_5.png"));
    }

    [Fact]
    public void Images_in_doubt_are_left_in_place_and_reported()
    {
        this.thisData.WriteFile("board photo.jpg", "photo");

        var outcome = OscilloscopeCaptureRelocation.MoveCapturesOutOfDataRoot(this.thisData.Root, this.Target, _ => false);

        Assert.Empty(outcome.Moved);
        Assert.Equal(this.thisData.Path_("board photo.jpg"), Assert.Single(outcome.LeftInPlace));
        Assert.False(Directory.Exists(this.Target));
    }

    // A target inside the data folder is no way out of it: nothing moves.
    [Fact]
    public void Nothing_moves_into_a_folder_inside_the_data_folder()
    {
        this.thisData.WriteFile("U1_5.png", "capture");

        var outcome = OscilloscopeCaptureRelocation.MoveCapturesOutOfDataRoot(
            this.thisData.Root,
            this.thisData.Path_("Scope"),
            _ => false);

        Assert.Empty(outcome.Moved);
        Assert.True(File.Exists(this.thisData.Path_("U1_5.png")));
    }

    [Fact]
    public void A_capture_folder_inside_the_data_folder_gets_no_move_target()
    {
        string dataRoot = this.thisData.Root;

        Assert.Null(OscilloscopeCaptureRelocation.Plan.ForFolder(dataRoot, dataRoot).MoveTarget);
        Assert.Null(OscilloscopeCaptureRelocation.Plan.ForFolder(this.thisData.Path_("Scope"), dataRoot).MoveTarget);
        Assert.Null(OscilloscopeCaptureRelocation.Plan.ForFolder(string.Empty, dataRoot).MoveTarget);
        Assert.Equal(this.Target, OscilloscopeCaptureRelocation.Plan.ForFolder(this.Target, dataRoot).MoveTarget);
    }

    [Fact]
    public void A_missing_data_folder_moves_nothing_and_does_not_throw()
    {
        var outcome = OscilloscopeCaptureRelocation.MoveCapturesOutOfDataRoot(
            this.thisData.Path_("not there"),
            this.Target,
            _ => false);

        Assert.Empty(outcome.Moved);
        Assert.Empty(outcome.LeftInPlace);
    }
}
