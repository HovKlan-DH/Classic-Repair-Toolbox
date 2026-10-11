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

    // ------------------------------------------------------------------ copying a project IN

    private string SourceFolder => Path.Combine(this.thisWorkspace.Root, "My KiCad project");

    [Fact]
    public void A_project_is_copied_in_with_its_sub_folders_and_overwrites_what_was_there()
    {
        string pcb = this.Write(this.SourceFolder, "Open128.kicad_pcb");
        string page = this.Write(this.SourceFolder, "Pages/vic.kicad_sch");
        File.WriteAllText(pcb, "(kicad_pcb new)");

        string old = this.Write(this.KiCadFolder, "Open128.kicad_pcb");

        (int copied, bool complete) = KiCadImportedFiles.CopyIn(this.SourceFolder, [pcb, page], this.KiCadFolder);

        Assert.Equal(2, copied);
        Assert.True(complete);
        Assert.Equal("(kicad_pcb new)", File.ReadAllText(old));
        Assert.True(File.Exists(Path.Combine(this.KiCadFolder, "Pages", "vic.kicad_sch")));
    }

    // ###########################################################################################
    // *** A COPY THAT STOPS PART-WAY SAYS HOW FAR IT GOT (code review, 2026-10-10). *** The files
    // before the one that failed are already overwritten, so the window must count the draft as
    // changed - a plain "failed" left the board on screen drawing the old project. A FOLDER where
    // the second file goes makes that copy fail on every system.
    // ###########################################################################################
    [Fact]
    public void A_copy_that_stops_part_way_says_how_many_files_it_copied()
    {
        string first = this.Write(this.SourceFolder, "a.kicad_pcb");
        string second = this.Write(this.SourceFolder, "b.kicad_sch");
        Directory.CreateDirectory(Path.Combine(this.KiCadFolder, "b.kicad_sch"));

        (int copied, bool complete) = KiCadImportedFiles.CopyIn(this.SourceFolder, [first, second], this.KiCadFolder);

        Assert.Equal(1, copied);
        Assert.False(complete);
        Assert.True(File.Exists(Path.Combine(this.KiCadFolder, "a.kicad_pcb")));
    }

    // The report reads the published KiCad data too, and says so - only then.
    [Fact]
    public void The_report_says_it_read_the_published_KiCad_data_only_when_it_did()
    {
        Assert.Equal(string.Empty, KiCadImportedFiles.ReportSourceLine(publishedFiles: 0, draftFiles: 3));
        Assert.Equal(string.Empty, KiCadImportedFiles.ReportSourceLine(publishedFiles: 0, draftFiles: 0));
        Assert.Contains("published KiCad data - your draft has none of its own", KiCadImportedFiles.ReportSourceLine(publishedFiles: 2, draftFiles: 0));
        Assert.Contains("published KiCad data together with your draft's", KiCadImportedFiles.ReportSourceLine(publishedFiles: 2, draftFiles: 1));
    }
}
