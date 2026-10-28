using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPaneFileNames - the name a file added in the component pane gets in its board (CRT 4.0
// step 2, slice 3b; the plan: "No 'File location' box. Every file belongs to its board"). In its
// kind's usual folder, never a name the board already has - given one, it would replace that file
// for every row citing it - and shaped so the server takes it (no leading dot, no folder in it).
// ###########################################################################################
public sealed class ComponentPaneFileNamesTests
{
    private static string Picked(string fileName) => Path.Combine(Path.GetTempPath(), "picked", fileName);

    [Fact]
    public void A_picture_keeps_its_own_name_in_the_component_images_folder()
    {
        Assert.Equal("Component images/HotCPU.png", ComponentPaneFileNames.ForPicture(Picked("HotCPU.png"), []));
        Assert.Equal("Component images/HotCPU (2).png", ComponentPaneFileNames.ForPicture(Picked("HotCPU.png"), ["component images/hotcpu.png"]));
    }

    [Fact]
    public void A_file_keeps_its_own_name_in_the_component_files_folder()
    {
        Assert.Equal("Component local files/6510 datasheet.pdf", ComponentPaneFileNames.ForFile(Picked("6510 datasheet.pdf"), []));
    }

    // A file of the board itself (slice 4a) - in the folder every converted board keeps its own in,
    // never a name the board has.
    [Fact]
    public void A_board_file_keeps_its_own_name_in_the_board_files_folder()
    {
        Assert.Equal("Board local files/Service manual.pdf", ComponentPaneFileNames.ForBoardFile(Picked("Service manual.pdf"), []));
        Assert.Equal("Board local files/Service manual (2).pdf", ComponentPaneFileNames.ForBoardFile(Picked("Service manual.pdf"), ["board local files/service manual.pdf"]));
    }

    // A picture of the board lies at the board's top under its own name, as every live board keeps
    // them - a second file of that name numbered, and a leading dot never kept (the server refuses one).
    [Fact]
    public void A_board_picture_keeps_its_own_name_at_the_boards_top()
    {
        Assert.Equal("Top side.png", ComponentPaneFileNames.ForBoardPicture(Picked("Top side.png"), []));
        Assert.Equal("Top side (2).png", ComponentPaneFileNames.ForBoardPicture(Picked("Top side.png"), ["top side.png"]));
        Assert.Equal("scan.jpg", ComponentPaneFileNames.ForBoardPicture(Picked(".scan.jpg"), []));
    }

    // ###########################################################################################
    // A device's photo gets a PLAIN name - no separator, no leading dot - whatever the picked file
    // is called (code review, 2026-10-28): on Linux and macOS "\" is part of a name, and FreeName
    // made "x\..\..\board.png" into "../../board.png", a copy above the device's folder.
    // ###########################################################################################
    [Fact]
    public void A_device_photo_gets_a_plain_name_of_its_own()
    {
        // windows-path-literal: a Linux file NAME holding backslashes, the case being guarded against.
        string climbing = Path.Combine(Path.GetTempPath(), "x\\..\\..\\board.png");

        string name = ComponentPaneFileNames.ForDevicePhoto(climbing, []);

        Assert.DoesNotContain('/', name);
        Assert.DoesNotContain('\\', name);
        Assert.False(name.StartsWith('.'), name);
        Assert.EndsWith("board.png", name, StringComparison.Ordinal);
        Assert.Equal("front (2).jpg", ComponentPaneFileNames.ForDevicePhoto(Picked("front.jpg"), ["Front.jpg"]));
    }

    // ###########################################################################################
    // A scope baseline is named as the published ones are - "<label>_<pin>_<name>" with the picked
    // file's type ("Scope baseline/IC107_1_PA3.jpg") - whatever the screenshot was called.
    // ###########################################################################################
    [Fact]
    public void A_scope_baseline_is_named_by_its_label_pin_and_signal_as_the_published_ones_are()
    {
        Assert.Equal("Scope baseline/U8_1_PA3.png", ComponentPaneFileNames.ForBaseline("U8", "1", "PA3", Picked("IMG_0001.png"), []));
        Assert.Equal("Scope baseline/U8_1.jpg", ComponentPaneFileNames.ForBaseline("U8", " 1 ", " ", Picked("capture.jpg"), []));
        Assert.Equal("Scope baseline/U8_1 (2).png", ComponentPaneFileNames.ForBaseline("U8", "1", null, Picked("a.png"), ["Scope baseline/U8_1.png"]));
    }

    // What a file system or the server would refuse: a folder sign, a reserved sign, a leading dot.
    [Fact]
    public void A_name_is_shaped_so_the_board_can_keep_it()
    {
        Assert.Equal("Scope baseline/U8_A_B_1_X_Y.png", ComponentPaneFileNames.ForBaseline("U8", "A/B:1", "X?Y", Picked("a.png"), []));
        Assert.Equal("Component local files/hidden notes.txt", ComponentPaneFileNames.ForFile(Picked(".hidden notes.txt"), []));
        Assert.Equal("Component local files/file", ComponentPaneFileNames.ForFile(Picked("..."), []));
    }
}
