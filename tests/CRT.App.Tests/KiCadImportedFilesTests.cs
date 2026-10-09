using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadImportedFiles - removing one imported KiCad file from a draft's "KiCad data" folder, the
// per-file "Remove" in BoardFilesWindow (owner request, 2026-09-24).
//
// This DELETES FILES, so half of these are about what it must refuse: anything not strictly inside
// the KiCad folder - the draft's own workbook beside it, a sibling folder whose name merely starts
// the same way, the folder itself. The other half pin the tidy-up: folders the delete empties go,
// up to "KiCad data" and never beyond it.
// ###########################################################################################
public sealed class KiCadImportedFilesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string BoardFolder => Path.Combine(this.thisWorkspace.Root, "Drafts", "Test Manu4", "HW4", "Board4");

    private string KiCadFolder => Path.Combine(this.BoardFolder, "KiCad data");

    public void Dispose() => this.thisWorkspace.Dispose();

    private string Write(string folder, string relative)
    {
        string full = Path.Combine(folder, relative.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, "(kicad_sch)");
        return full;
    }

    [Fact]
    public void The_file_is_deleted_and_its_neighbours_are_kept()
    {
        string pcb = this.Write(this.KiCadFolder, "Open128.kicad_pcb");
        string pro = this.Write(this.KiCadFolder, "Open128.kicad_pro");

        Assert.True(KiCadImportedFiles.TryRemove(this.KiCadFolder, pcb));

        Assert.False(File.Exists(pcb));
        Assert.True(File.Exists(pro));
        Assert.True(Directory.Exists(this.KiCadFolder));
    }

    // Removing a multi-sheet project's last page must not leave an empty Pages/ behind - the
    // component editor's "File location" drop-down lists every end folder it finds.
    [Fact]
    public void A_sub_folder_the_delete_EMPTIES_is_removed()
    {
        this.Write(this.KiCadFolder, "Open128.kicad_pcb");
        string page = this.Write(this.KiCadFolder, "Pages/vic.kicad_sch");

        Assert.True(KiCadImportedFiles.TryRemove(this.KiCadFolder, page));

        Assert.False(Directory.Exists(Path.Combine(this.KiCadFolder, "Pages")));
        Assert.True(Directory.Exists(this.KiCadFolder));
    }

    [Fact]
    public void A_sub_folder_that_still_holds_files_is_kept()
    {
        string vic = this.Write(this.KiCadFolder, "Pages/vic.kicad_sch");
        string sid = this.Write(this.KiCadFolder, "Pages/sid.kicad_sch");

        Assert.True(KiCadImportedFiles.TryRemove(this.KiCadFolder, vic));

        Assert.True(File.Exists(sid));
    }

    // Removing everything returns the draft to how it was before the import - no "KiCad data"
    // folder at all - but the tidy-up stops THERE: the board folder holds the workbook.
    [Fact]
    public void Removing_the_last_file_removes_KiCad_data_but_never_the_board_folder()
    {
        this.Write(this.BoardFolder, "Data HW4 Board4.xlsx");
        string page = this.Write(this.KiCadFolder, "Pages/vic.kicad_sch");

        Assert.True(KiCadImportedFiles.TryRemove(this.KiCadFolder, page));

        Assert.False(Directory.Exists(this.KiCadFolder));
        Assert.True(File.Exists(Path.Combine(this.BoardFolder, "Data HW4 Board4.xlsx")));
    }

    // Even with an EMPTY board folder above it, the walk must stop at "KiCad data".
    [Fact]
    public void The_tidy_up_never_climbs_past_KiCad_data_even_into_an_empty_parent()
    {
        string pcb = this.Write(this.KiCadFolder, "board.kicad_pcb");

        Assert.True(KiCadImportedFiles.TryRemove(this.KiCadFolder, pcb));

        Assert.True(Directory.Exists(this.BoardFolder));
    }

    // ------------------------------------------------------------------ What it must refuse

    [Fact]
    public void A_file_BESIDE_the_KiCad_folder_is_refused_and_left_alone()
    {
        this.Write(this.KiCadFolder, "board.kicad_pcb");
        string workbook = this.Write(this.BoardFolder, "Data HW4 Board4.xlsx");

        Assert.False(KiCadImportedFiles.TryRemove(this.KiCadFolder, workbook));
        Assert.True(File.Exists(workbook));
    }

    // The case a prefix test gets wrong: "KiCad data2" starts with "KiCad data".
    [Fact]
    public void A_SIBLING_folder_sharing_the_name_prefix_is_refused()
    {
        string sibling = this.Write(Path.Combine(this.BoardFolder, "KiCad data2"), "board.kicad_pcb");

        Assert.False(KiCadImportedFiles.TryRemove(this.KiCadFolder, sibling));
        Assert.True(File.Exists(sibling));
    }

    [Fact]
    public void A_path_that_TRAVERSES_out_of_the_folder_is_refused()
    {
        string workbook = this.Write(this.BoardFolder, "Data HW4 Board4.xlsx");
        Directory.CreateDirectory(this.KiCadFolder);

        string traversing = Path.Combine(this.KiCadFolder, "..", "Data HW4 Board4.xlsx");

        Assert.False(KiCadImportedFiles.TryRemove(this.KiCadFolder, traversing));
        Assert.True(File.Exists(workbook));
    }

    [Fact]
    public void The_KiCad_folder_itself_is_not_a_file_to_remove()
    {
        this.Write(this.KiCadFolder, "board.kicad_pcb");

        Assert.False(KiCadImportedFiles.TryRemove(this.KiCadFolder, this.KiCadFolder));
        Assert.True(Directory.Exists(this.KiCadFolder));
    }

    [Fact]
    public void A_file_that_is_already_gone_reports_failure_rather_than_success()
    {
        Directory.CreateDirectory(this.KiCadFolder);

        Assert.False(KiCadImportedFiles.TryRemove(
            this.KiCadFolder,
            Path.Combine(this.KiCadFolder, "never-there.kicad_pcb")));
    }

    [Fact]
    public void Blank_inputs_are_refused()
    {
        Assert.False(KiCadImportedFiles.TryRemove(string.Empty, "x.kicad_pcb"));
        Assert.False(KiCadImportedFiles.TryRemove(this.KiCadFolder, "  "));
    }
}
