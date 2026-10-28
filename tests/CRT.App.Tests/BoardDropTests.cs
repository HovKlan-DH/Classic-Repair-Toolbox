using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardDrop - what files dropped on the Contribute tab are (CRT 4.0 step 3, slice 3e): pictures of
// the board, its KiCad project, or neither. What was dropped decides which window opens with them.
// ###########################################################################################
public sealed class BoardDropTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    // A small file system: folders, and the files at each folder's top - a folder of neither said of
    // as "why <folder>".
    private static BoardDrop Of(IReadOnlyList<string> dropped, Dictionary<string, string[]> folders) =>
        BoardDrop.Of(
            dropped,
            folders.ContainsKey,
            folder => folders[folder].Any(ComponentListBuilder.IsSupportedKiCadRawFile),
            folder => folders[folder],
            folder => $"why {folder}");

    [Fact]
    public void A_folder_holding_a_KiCad_project_is_the_KiCad_project()
    {
        BoardDrop drop = Of(["Open128"], new() { ["Open128"] = ["Open128/Open128.kicad_pro", "Open128/top.png"] });

        Assert.Equal(BoardDropKind.KiCad, drop.Kind);
        Assert.Equal(["Open128"], drop.Paths);
    }

    // A folder of board scans - no KiCad file, pictures at its top - is those pictures.
    [Fact]
    public void A_folder_of_scans_is_its_pictures()
    {
        BoardDrop drop = Of(["Scans"], new() { ["Scans"] = ["Scans/top.JPG", "Scans/notes.txt", "Scans/bottom.png"] });

        Assert.Equal(BoardDropKind.Pictures, drop.Kind);
        Assert.Equal(["Scans/top.JPG", "Scans/bottom.png"], drop.Paths);
    }

    // ###########################################################################################
    // A folder with neither is NOTHING, said for what it is (code review, 2026-10-28 - it was tried
    // as the KiCad project, and a published board got a draft made for nothing): no window opens.
    // ###########################################################################################
    [Fact]
    public void A_folder_with_neither_is_nothing_said_for_what_it_is()
    {
        BoardDrop drop = Of(["Old"], new() { ["Old"] = ["Old/c64.sch", "Old/c64.pro"] });

        Assert.Equal(BoardDropKind.Nothing, drop.Kind);
        Assert.Empty(drop.Paths);
        Assert.Equal("why Old", drop.Problem);
    }

    // ###########################################################################################
    // A KiCad FILE on its own is no project (code review, 2026-10-28 - it stood for its folder, which
    // was walked in full): said, and nothing opens. Of several folders the first is taken.
    // ###########################################################################################
    [Fact]
    public void A_KiCad_file_on_its_own_is_said_and_the_first_folder_is_taken()
    {
        string pro = Path.Combine("Projects", "Open128", "Open128.kicad_pro");

        BoardDrop byFile = Of([pro], []);
        Assert.Equal(BoardDropKind.Nothing, byFile.Kind);
        Assert.Equal(BoardDrop.KiCadFileNotFolder, byFile.Problem);

        // With pictures dropped beside it, the pictures are added - the file said to be left out.
        BoardDrop beside = Of([pro, "top.png"], []);
        Assert.Equal(BoardDropKind.Pictures, beside.Kind);
        Assert.Equal(BoardDrop.LeftWithoutPictures(1), beside.Note);

        // One KiCad project at a time - the other one said to be left.
        BoardDrop twoFolders = Of(["A", "B"], new() { ["A"] = ["A/a.kicad_pcb"], ["B"] = ["B/b.kicad_pcb"] });
        Assert.Equal(["A"], twoFolders.Paths);
        Assert.Equal(BoardDrop.LeftBesideKiCad(1), twoFolders.Note);
    }

    // Picture files are pictures - anything else dropped with them is left out, and said.
    [Fact]
    public void Picture_files_are_pictures_and_the_rest_is_said_to_be_left_out()
    {
        BoardDrop drop = Of(["top.png", "notes.txt", "bottom.webp", " "], []);

        Assert.Equal(BoardDropKind.Pictures, drop.Kind);
        Assert.Equal(["top.png", "bottom.webp"], drop.Paths);
        Assert.Equal(BoardDrop.LeftWithoutPictures(1), drop.Note);
        Assert.Equal(string.Empty, Of(["top.png"], []).Note);
    }

    // ###########################################################################################
    // Loose pictures dropped with a folder of scans go in TOGETHER (code review, 2026-10-28: the
    // first folder decided alone, and the loose photos were skipped without a word). A folder of
    // neither among them is said to be left out.
    // ###########################################################################################
    [Fact]
    public void Loose_pictures_and_folders_of_scans_go_in_together()
    {
        BoardDrop drop = Of(
            ["Top.jpg", "Scans", "Bottom.jpg", "Notes"],
            new() { ["Scans"] = ["Scans/side.png", "Scans/readme.txt"], ["Notes"] = ["Notes/notes.txt"] });

        Assert.Equal(BoardDropKind.Pictures, drop.Kind);
        Assert.Equal(["Top.jpg", "Scans/side.png", "Bottom.jpg"], drop.Paths);
        Assert.Equal(BoardDrop.LeftWithoutPictures(1), drop.Note);
    }

    // A KiCad project dropped with pictures is taken on its own - the pictures said to be left.
    [Fact]
    public void A_KiCad_project_dropped_with_pictures_is_taken_alone_and_the_rest_said()
    {
        BoardDrop drop = Of(["Top.jpg", "Open128", "Bottom.jpg"], new() { ["Open128"] = ["Open128/Open128.kicad_pro"] });

        Assert.Equal(BoardDropKind.KiCad, drop.Kind);
        Assert.Equal(["Open128"], drop.Paths);
        Assert.Equal(BoardDrop.LeftBesideKiCad(2), drop.Note);
    }

    [Fact]
    public void Nothing_usable_is_nothing()
    {
        BoardDrop drop = Of(["notes.txt", "datasheet.pdf"], []);

        Assert.Equal(BoardDropKind.Nothing, drop.Kind);
        Assert.Equal(BoardDrop.NothingUsable, drop.Problem);
        Assert.Empty(drop.Paths);
        Assert.Empty(Of([], []).Paths);
    }

    // ###########################################################################################
    // A KiCad FILE dropped is no project (code review, 2026-10-28): taken as its folder, a file from
    // Downloads imported every KiCad project under Downloads, and a sheet from Pages/ only Pages/.
    // ###########################################################################################
    [Fact]
    public void A_KiCad_file_dropped_on_its_own_is_no_project()
    {
        string pro = this.thisWorkspace.WriteFile(Path.Combine("Downloads", "Open128.kicad_pro"), "{}");
        this.thisWorkspace.WriteFile(Path.Combine("Downloads", "Other", "Other.kicad_pcb"), "(kicad_pcb)");

        BoardDrop drop = BoardDrop.OnDisk([pro]);
        Assert.Equal(BoardDropKind.Nothing, drop.Kind);
        Assert.Equal(BoardDrop.KiCadFileNotFolder, drop.Problem);
    }

    // ###########################################################################################
    // A folder holding neither a KiCad file nor pictures is nothing to add (code review, 2026-10-28):
    // tried as the KiCad project, a published board got a draft made for nothing, which stayed.
    // ###########################################################################################
    [Fact]
    public void A_folder_with_neither_KiCad_files_nor_pictures_is_nothing()
    {
        string empty = this.thisWorkspace.Path_("Empty");
        Directory.CreateDirectory(empty);

        string eagle = this.thisWorkspace.Path_("Eagle");
        this.thisWorkspace.WriteFile(Path.Combine("Eagle", "board.brd"), "<?xml version=\"1.0\"?>");

        BoardDrop none = BoardDrop.OnDisk([empty]);
        Assert.Equal(BoardDropKind.Nothing, none.Kind);
        Assert.Equal("That folder holds neither KiCad files nor pictures CRT can show. " + BoardDrop.NothingUsable, none.Problem);

        // Another program's project is told how KiCad imports it.
        BoardDrop other = BoardDrop.OnDisk([eagle]);
        Assert.Equal(BoardDropKind.Nothing, other.Kind);
        Assert.Equal("No KiCad files CRT can read in that folder. " + KiCadImportSummary.OtherProgram, other.Problem);
    }

    // ###########################################################################################
    // A hidden file is never one of the pictures (code review, 2026-10-28): macOS leaves "._IMG_0001.JPG"
    // beside every picture on a USB stick or an SD card, each became a picture row with no preview, and
    // Submit was refused for names starting with a dot the contributor never saw.
    // ###########################################################################################
    [Fact]
    public void A_hidden_file_in_a_folder_of_scans_is_no_picture()
    {
        string scans = this.thisWorkspace.Path_("Card");
        this.thisWorkspace.WriteFile(Path.Combine("Card", "top.png"), "x");
        this.thisWorkspace.WriteFile(Path.Combine("Card", "._top.png"), "x");

        BoardDrop drop = BoardDrop.OnDisk([scans]);

        Assert.Equal(BoardDropKind.Pictures, drop.Kind);
        Assert.Equal([Path.Combine(scans, "top.png")], drop.Paths);

        // Dropped as files, the same.
        Assert.Equal(["top.png"], BoardDrop.Of(["top.png", "._top.png"], _ => false, _ => false, _ => [], _ => "").Paths);
    }

    // The rule on a real folder: a project folder, and a folder of scans.
    [Fact]
    public void On_disk_a_project_folder_and_a_folder_of_scans_are_told_apart()
    {
        string project = this.thisWorkspace.Path_("Open128");
        this.thisWorkspace.WriteFile(Path.Combine("Open128", "Pages", "vic.kicad_sch"), "(kicad_sch)");

        string scans = this.thisWorkspace.Path_("Scans");
        this.thisWorkspace.WriteFile(Path.Combine("Scans", "top.png"), "x");
        Directory.CreateDirectory(project);

        Assert.Equal(BoardDropKind.KiCad, BoardDrop.OnDisk([project]).Kind);

        BoardDrop pictures = BoardDrop.OnDisk([scans]);
        Assert.Equal(BoardDropKind.Pictures, pictures.Kind);
        Assert.Equal([Path.Combine(scans, "top.png")], pictures.Paths);
    }
}
