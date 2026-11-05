using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadImportSummary - what a KiCad import found, left out and leaves to do (CRT 4.0 step 3, slice
// 3e - the plan: "KiCad: drop the project folder and get a summary of what was found, what was
// ignored, and what is left to do"). A project from an older KiCad used to be met with "No KiCad
// files in that folder" and nothing more; it is told what to do now.
// ###########################################################################################
public sealed class KiCadImportSummaryTests
{
    private static KiCadPcbFootprint Chip(string reference) => new() { Reference = reference, Pads = [new KiCadPcbPad { Number = "1" }] };

    private static KiCadPcb Pcb(int nets, params string[] references) => new()
    {
        Footprints = [.. references.Select(Chip)],
        Nets = new KiCadPcbNets
        {
            List =
            [
                new KiCadNetRef { Id = "0", Name = "", IsNoNet = true },
                .. Enumerable.Range(1, nets).Select(number => new KiCadNetRef { Id = $"{number}", Name = $"/N{number}" }),
            ],
        },
    };

    // ---- found ----------------------------------------------------------------------------------

    // The first PCB's components - the very ones "Add N from KiCad" counts, a logo with no pad none -
    // and its nets, KiCad's "no net" not one of them; and the sheets.
    [Fact]
    public void What_was_found_counts_the_first_PCBs_components_and_nets_and_the_sheets()
    {
        KiCadPcb pcb = Pcb(2, "U1", "U2");
        pcb.Footprints.Add(new KiCadPcbFootprint { Reference = "LOGO1" });

        var root = new KiCadProjectRoot
        {
            Pcb = [pcb],
            Schematics = [new KiCadSchematic(), new KiCadSchematic(), new KiCadSchematic()],
        };

        Assert.Equal("Found 1 PCB with 2 components and 2 nets, and 3 schematic sheets.", KiCadImportSummary.Found(root));
    }

    // KiCad files left out as they hold no bytes are named (owner decision, 2026-11-02) - none, none said.
    [Fact]
    public void Empty_KiCad_files_left_out_are_named()
    {
        Assert.Equal(string.Empty, KiCadImportSummary.EmptyFiles([]));
        Assert.Equal(
            "Left out, as it is empty: [a.kicad_sch]. Save it again in KiCad and import the folder again.",
            KiCadImportSummary.EmptyFiles(["a.kicad_sch"]));
        Assert.Equal(
            "Left out, as they are empty: [a.kicad_sch], [b.kicad_pcb]. Save them again in KiCad and import the folder again.",
            KiCadImportSummary.EmptyFiles(["a.kicad_sch", "b.kicad_pcb"]));
    }

    [Fact]
    public void No_PCB_one_sheet_and_several_PCBs_are_each_said()
    {
        Assert.Equal(
            "Found no PCB file, and 1 schematic sheet.",
            KiCadImportSummary.Found(new KiCadProjectRoot { Schematics = [new KiCadSchematic()] }));

        Assert.Equal(
            "Found 2 PCB files - the first, the one CRT draws, with 1 component and 1 net, and no schematic sheet.",
            KiCadImportSummary.Found(new KiCadProjectRoot { Pcb = [Pcb(1, "U1"), Pcb(5, "U1", "U2", "U3")] }));
    }

    // ---- left out -------------------------------------------------------------------------------

    // Counted by what each file is, in a fixed order, singular for one. A .sch or .brd is an older
    // KiCad's only when the scan said so by its first line - else another program's (EAGLE's, say).
    [Fact]
    public void What_was_left_out_is_counted_by_what_it_is()
    {
        string[] leftOut =
        [
            "Open128.pretty/DIP-40.kicad_mod", "Open128.pretty/DIP-28.kicad_mod",
            "3d/socket.step", "3d/crystal.WRL",
            "gerbers/Open128-F_Cu.gtl", "gerbers/Open128.drl",
            "Open128.kicad_pcb-bak", "_autosave-Open128.kicad_sch", "Open128.kicad_sch.lck",
            "fp-lib-table", "Open128.kicad_prl",
            "symbols.kicad_sym",
            "Open128.sch",
            "eagle/board.brd",
            "README.md", "Open128.pdf",
        ];

        Assert.Equal(
            "Left out, as CRT does not read them: 2 footprints, 1 symbol library file, 2 3D models, 2 Gerber and drill files, " +
            "1 file from an older KiCad, 1 schematic or board file of another program, 2 KiCad settings files, 3 backups and 2 other files.",
            KiCadImportSummary.LeftOut(leftOut, ["Open128.sch"]));

        Assert.Equal("Left out, as CRT does not read them: 1 footprint.", KiCadImportSummary.LeftOut(["x.kicad_mod"], []));
        Assert.Equal(string.Empty, KiCadImportSummary.LeftOut([], []));
    }

    // ###########################################################################################
    // A schematic or board an older KiCad wrote (.sch, .brd - told by the scan) is said, with what to
    // do - but not for a project saved again in KiCad 6 or later, which keeps the old files beside
    // the new.
    // ###########################################################################################
    [Fact]
    public void A_schematic_or_board_from_an_older_KiCad_is_said_with_what_to_do()
    {
        Assert.Equal(
            ["Its schematic files (.sch) are from an older KiCad, which CRT cannot read - open the project in KiCad 6 or later, save it, and import the folder again."],
            KiCadImportSummary.CannotRead(["board.sch", "board.pro"], ["board.sch"], ["board.kicad_pcb"]));

        Assert.Equal(
            ["Its board file (.brd) is from an older KiCad, which CRT cannot read - open the project in KiCad 6 or later, save it, and import the folder again."],
            KiCadImportSummary.CannotRead(["board.brd"], ["board.brd"], ["board.kicad_sch"]));

        Assert.StartsWith(
            "Its schematic (.sch) and board (.brd) files are",
            Assert.Single(KiCadImportSummary.CannotRead(["a.sch", "a.brd"], ["a.sch", "a.brd"], [])),
            StringComparison.Ordinal);

        // Saved again in a newer KiCad: the new files are read, nothing to say.
        Assert.Empty(KiCadImportSummary.CannotRead(["board.sch", "board.pro"], ["board.sch"], ["board.kicad_sch", "board.kicad_pcb"]));
        Assert.Empty(KiCadImportSummary.CannotRead(["notes.txt"], [], []));
    }

    // ###########################################################################################
    // .sch and .brd are EAGLE's too (code review, 2026-10-28): files the scan did not find an older
    // KiCad's are never called that - told how KiCad imports an EAGLE project instead; and not at all
    // when the KiCad files beside them are read.
    // ###########################################################################################
    [Fact]
    public void Another_programs_schematic_and_board_are_told_how_KiCad_imports_them()
    {
        Assert.Equal([KiCadImportSummary.OtherProgram], KiCadImportSummary.CannotRead(["board.sch", "board.brd"], [], []));
        Assert.Contains("EAGLE", KiCadImportSummary.OtherProgram, StringComparison.Ordinal);
        Assert.DoesNotContain("older KiCad", KiCadImportSummary.OtherProgram, StringComparison.Ordinal);

        Assert.Empty(KiCadImportSummary.CannotRead(["board.sch", "board.brd"], [], ["board.kicad_pcb"]));
    }

    // A folder with nothing CRT reads: said for what it is.
    [Fact]
    public void A_folder_with_nothing_to_import_says_why()
    {
        Assert.Equal(
            "No KiCad files CRT can read in that folder. Its schematic files (.sch) are from an older KiCad, which CRT cannot read - open the project in KiCad 6 or later, save it, and import the folder again.",
            KiCadImportSummary.NothingToImport(["C64.sch", "C64.pro", "C64.kicad_pcb-bak"], ["C64.sch"]));

        Assert.Equal(
            "No KiCad files CRT can read in that folder. " + KiCadImportSummary.OtherProgram,
            KiCadImportSummary.NothingToImport(["C64.sch", "C64.brd"], []));

        Assert.Equal(
            "No KiCad files in that folder. Only modern .kicad_pcb / .kicad_pro / .kicad_sch files are imported - sub-folders are searched too.",
            KiCadImportSummary.NothingToImport(["photo.jpg"], []));
    }

    // The summary's button for the KiCad components still missing (2026-11-01 - Board status's "Add N
    // from KiCad" until then), one or several.
    [Fact]
    public void The_button_for_the_missing_KiCad_components_counts_them()
    {
        Assert.Equal("Add 1 component from KiCad", KiCadImportSummary.AddComponentsButton(1));
        Assert.Equal("Add 214 components from KiCad", KiCadImportSummary.AddComponentsButton(214));
    }

    // ###########################################################################################
    // WHERE THE COUNT BESIDE "KiCad data" COMES FROM (owner request, 2026-10-30: "4" read as four KiCad
    // pictures): the Summary counts the board's KiCad files by kind - the very names the left column
    // counts (ContributeBoardParts.KiCadFileCount), each once.
    // ###########################################################################################
    [Fact]
    public void The_summary_says_the_boards_KiCad_files_by_kind_as_the_left_column_counts_them()
    {
        string[] names =
        [
            "KiCad data/c64.kicad_pcb",
            "KiCad data/c64.kicad_sch",
            "KiCad data/Pages/vic.kicad_sch",
            "KiCad data/c64.kicad_pro",
            "kicad data/c64.kicad_pcb",
            "Images/top.png",
        ];

        // The number in bold and a line per kind (owner request, 2026-11-03).
        Assert.Equal(
            "The board now has <b>4</b> KiCad files:\n- 1 PCB file\n- 2 schematic files\n- 1 project file",
            KiCadImportSummary.Files(names));
        Assert.Equal(4, ContributeBoardParts.KiCadFileCount(names));

        Assert.Equal(
            "The board now has <b>2</b> KiCad files:\n- 1 PCB file\n- 1 other file",
            KiCadImportSummary.Files(["KiCad data/a.kicad_pcb", "KiCad data/KiCad-traces.json"]));

        Assert.Equal("The board now has <b>1</b> KiCad file:\n- 1 PCB file", KiCadImportSummary.Files(["KiCad data/a.kicad_pcb"]));

        Assert.Equal(string.Empty, KiCadImportSummary.Files(["Images/top.png"]));
    }
}
