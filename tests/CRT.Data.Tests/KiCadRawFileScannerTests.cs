using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Finding a KiCad project's real files, sub-folders included (owner request, 2026-09-24).
//
// *** THE BUG THESE EXIST FOR. *** Both the importer and the app's own reader enumerated with
// SearchOption.TopDirectoryOnly, while DataManager SYNCED the same folder with AllDirectories.
// So a multi-sheet project's pages were downloaded to every client and then never read.
//
// The shipped Commodore C128 "310378 Open128" board is exactly that shape - a root .kicad_sch at
// the top and 22 more under Pages/ - so this was live on real data, not hypothetical. The fixture
// below reproduces that layout deliberately.
//
// Real files in a temp folder throughout: this walks a directory tree, so there is nothing to
// fake.
// ###########################################################################################
public sealed class KiCadRawFileScannerTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string Project => Path.Combine(this.thisWorkspace.Root, "Open128");

    private string Write(string relativePath, string content = "(kicad_sch)")
    {
        string full = Path.Combine(
            this.Project,
            relativePath.Replace('/', Path.DirectorySeparatorChar));

        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, content);

        return full;
    }

    private List<string> ScanNames() =>
        KiCadRawFileScanner.Scan(this.Project)
            .Select(path => KiCadRawFileScanner.RelativeDestinationFor(this.Project, path)
                .Replace(Path.DirectorySeparatorChar, '/'))
            .ToList();

    // ------------------------------------------------------------------ the reported bug

    // ###########################################################################################
    // *** THE REPORTED BUG, AS A TEST. *** This fails against the TopDirectoryOnly version, which
    // returned the three top-level files and none of the pages.
    // ###########################################################################################
    [Fact]
    public void A_multi_sheet_project_finds_the_pages_in_its_SUB_FOLDER()
    {
        this.Write("Open128.kicad_pcb");
        this.Write("Open128.kicad_pro");
        this.Write("Open128.kicad_sch");
        this.Write("Pages/vic.kicad_sch");
        this.Write("Pages/cpu-8500.kicad_sch");
        this.Write("Pages/ram.kicad_sch");

        Assert.Equal(
            ["Open128.kicad_pcb", "Open128.kicad_pro", "Open128.kicad_sch",
             "Pages/cpu-8500.kicad_sch", "Pages/ram.kicad_sch", "Pages/vic.kicad_sch"],
            this.ScanNames());
    }

    // Nesting is not limited to one level - a contributor may organise pages however they like.
    [Fact]
    public void A_deeply_nested_sheet_is_still_found()
    {
        this.Write("Sheets/analog/audio/sid.kicad_sch");

        Assert.Equal(["Sheets/analog/audio/sid.kicad_sch"], this.ScanNames());
    }

    // ------------------------------------------------------------------ what is left behind

    // ###########################################################################################
    // *** THE WHOLE REASON RECURSING IS SAFE. *** The project owner's requirement is that picking a
    // full KiCad project yields only the few relevant files. That is the EXTENSION filter's doing,
    // not the depth limit's - so it has to hold at depth, which is what this asserts.
    //
    // Every path here is a real thing found in a KiCad project folder.
    // ###########################################################################################
    [Theory]
    [InlineData("footprints.pretty/R_0805.kicad_mod")]
    [InlineData("3dmodels/DIP-40.step")]
    [InlineData("3dmodels/DIP-40.wrl")]
    [InlineData("gerbers/Open128-F_Cu.gbr")]
    [InlineData("gerbers/Open128.drl")]
    [InlineData("Open128.kicad_prl")]
    [InlineData("Open128.net")]
    [InlineData("bom/parts.csv")]
    [InlineData("docs/notes.txt")]
    [InlineData("Pages/vic.kicad_sch-bak")]
    public void Everything_that_is_not_a_KiCad_source_file_is_left_behind(string relativePath)
    {
        // A real file alongside it, so this proves exclusion rather than an empty scan.
        this.Write("Open128.kicad_pcb");
        this.Write(relativePath);

        Assert.Equal(["Open128.kicad_pcb"], this.ScanNames());
    }

    // ###########################################################################################
    // *** THE ONE CASE THE EXTENSION FILTER CANNOT CATCH. ***
    //
    // KiCad's own backup folder holds copies of the REAL sheets, which pass the extension filter
    // on their own merits. Importing them would duplicate every page, and a duplicated page parses
    // as real circuitry - so a net would be found twice over rather than being harmlessly ignored.
    //
    // This is the reason the walk is hand-written instead of using SearchOption.AllDirectories,
    // which cannot skip a sub-tree.
    // ###########################################################################################
    [Fact]
    public void KiCads_own_BACKUP_folder_is_not_descended_into()
    {
        this.Write("Open128.kicad_sch");
        this.Write("Open128-backups/Open128-2026-09-01.kicad_sch");
        this.Write("Open128-backups/Pages/vic.kicad_sch");

        Assert.Equal(["Open128.kicad_sch"], this.ScanNames());
    }

    [Theory]
    [InlineData("Open128-backups")]
    [InlineData("SomethingElse-BACKUPS")]
    [InlineData("fp-info-cache")]
    [InlineData(".git")]
    [InlineData(".vscode")]
    public void The_skipped_folder_names(string folderName) =>
        Assert.True(KiCadRawFileScanner.IsSkippedFolder(folderName));

    [Theory]
    [InlineData("Pages")]
    [InlineData("Sheets")]
    [InlineData("backups-of-something")]
    [InlineData("")]
    public void A_folder_that_is_NOT_skipped(string folderName) =>
        Assert.False(KiCadRawFileScanner.IsSkippedFolder(folderName));

    // ------------------------------------------------------------------ the destination path

    // ###########################################################################################
    // *** THE STRUCTURE IS PRESERVED, NOT FLATTENED. ***
    //
    // Two reasons, either of which is enough: a root sheet references its pages by RELATIVE path,
    // and two sub-folders may hold a same-named page. Flattening breaks the first and silently
    // drops one of the second.
    // ###########################################################################################
    [Fact]
    public void A_sub_folder_file_keeps_its_sub_folder_in_the_destination()
    {
        string source = this.Write("Pages/vic.kicad_sch");

        Assert.Equal(
            Path.Combine("Pages", "vic.kicad_sch"),
            KiCadRawFileScanner.RelativeDestinationFor(this.Project, source));
    }

    [Fact]
    public void A_top_level_file_keeps_its_bare_name()
    {
        string source = this.Write("Open128.kicad_pcb");

        Assert.Equal(
            "Open128.kicad_pcb",
            KiCadRawFileScanner.RelativeDestinationFor(this.Project, source));
    }

    // ###########################################################################################
    // *** A DESTINATION CAN NEVER ESCAPE THE BOARD FOLDER. *** The result is combined with the
    // board's "KiCad data" path, so a relative path climbing out of the scanned folder would write
    // outside the board entirely. It cannot arise from Scan (which only ever walks downward), but
    // this is a path being built for a File.Copy destination and the guard is cheap.
    // ###########################################################################################
    [Fact]
    public void A_file_OUTSIDE_the_scanned_folder_falls_back_to_its_bare_name()
    {
        string outside = Path.Combine(this.thisWorkspace.Root, "elsewhere", "stray.kicad_sch");

        string destination = KiCadRawFileScanner.RelativeDestinationFor(this.Project, outside);

        Assert.Equal("stray.kicad_sch", destination);
        Assert.DoesNotContain("..", destination, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ ordering and edges

    [Fact]
    public void Results_are_sorted_so_two_scans_agree()
    {
        this.Write("Pages/zebra.kicad_sch");
        this.Write("Pages/alpha.kicad_sch");
        this.Write("Open128.kicad_pcb");

        Assert.Equal(this.ScanNames(), this.ScanNames());
        Assert.Equal(
            ["Open128.kicad_pcb", "Pages/alpha.kicad_sch", "Pages/zebra.kicad_sch"],
            this.ScanNames());
    }

    [Fact]
    public void A_missing_folder_scans_to_nothing_rather_than_throwing()
    {
        // Both callers run on a user-chosen path, and most boards have no KiCad data at all.
        Assert.Empty(KiCadRawFileScanner.Scan(Path.Combine(this.thisWorkspace.Root, "nope")));
        Assert.Empty(KiCadRawFileScanner.Scan(string.Empty));
    }

    [Fact]
    public void A_folder_holding_no_KiCad_files_scans_to_nothing()
    {
        this.Write("readme.txt");

        Assert.Empty(KiCadRawFileScanner.Scan(this.Project));
    }
}
