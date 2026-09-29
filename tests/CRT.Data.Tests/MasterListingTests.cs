using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// MasterListing - a new system's ONE row in the main Excel data file (owner request, 2026-09-27:
// "When a system is added to BETA, and it is a NEW system, can you then make sure it gets added
// also to the main Excel data file"). The fixture is shaped like the real master: a preamble, the
// header on row 9, and an Oscilloscope sheet that must come through untouched.
// ###########################################################################################
public sealed class MasterListingTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    static MasterListingTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private const string C64Long = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";
    private const string C128 = "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx";
    private const string C128Dcr = "Commodore/C128/250477/Data C128DCR 250477 v2.0.0.xlsx";
    private const string Spectrum = "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Spectrum Issue 4B v2.0.0.xlsx";
    private const string Open128 = "Commodore/C128/310378 Open128/Data C128 310378 Open128 v2.0.0.xlsx";

    private static readonly MasterListingRow Open128Row =
        new("Commodore 128", "310378 Open128", Open128, "Open-source C128 replica.");

    private string Master => Path.Combine(this.thisWorkspace.Root, "Classic-Repair-Toolbox.v2.0.0.xlsx");

    // (hardware, board, workbook, notes) per data row; a null hardware leaves the cell BLANK, the
    // carry-forward shape CRT also reads.
    private void WriteMaster(params (string? Hardware, string Board, string Workbook)[] rows)
    {
        using var package = new ExcelPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add(MasterWorkbookSchema.SheetName);

        sheet.Cells[1, 1].Value = "# Commodore Repair Toolbox";
        sheet.Cells[3, 1].Value = "# Revision date: 2026-August-7";
        sheet.Cells[8, 1].Value = "Hardware and boards";
        sheet.Cells[9, 1].Value = MasterWorkbookSchema.ColHardwareName;
        sheet.Cells[9, 2].Value = MasterWorkbookSchema.ColBoardName;
        sheet.Cells[9, 3].Value = MasterWorkbookSchema.ColExcelDataFile;
        sheet.Cells[9, 4].Value = MasterWorkbookSchema.ColHardwareNotes;

        for (int i = 0; i < rows.Length; i++)
        {
            if (rows[i].Hardware is not null)
            {
                sheet.Cells[10 + i, 1].Value = rows[i].Hardware;
            }

            sheet.Cells[10 + i, 2].Value = rows[i].Board;
            sheet.Cells[10 + i, 3].Value = rows[i].Workbook;
            sheet.Cells[10 + i, 4].Value = "notes " + rows[i].Board;
        }

        ExcelWorksheet scopes = package.Workbook.Worksheets.Add("Oscilloscope");
        scopes.Cells[1, 1].Value = "Brand";
        scopes.Cells[2, 1].Value = "Rigol";

        package.SaveAs(new FileInfo(this.Master));
    }

    private void WriteRealisticMaster() =>
        this.WriteMaster(
            ("Commodore 64", "250407 (long board)", C64Long),
            ("Commodore 128", "310378 (C128 & C128D)", C128),
            ("Commodore 128", "250477 (C128DCR)", C128Dcr),
            ("ZX Spectrum 16K/48K", "Issue 4B", Spectrum));

    private IReadOnlyList<MasterListingRow> Rows()
    {
        Assert.True(MasterListing.TryRead(this.Master, out IReadOnlyList<MasterListingRow> rows, out string why), why);
        return rows;
    }

    private IReadOnlyList<string> Boards() => this.Rows().Select(row => row.BoardName).ToList();

    // ------------------------------------------------------------------ reading

    [Fact]
    public void The_rows_are_read_in_order_as_CRT_reads_them()
    {
        this.WriteMaster(("Commodore 128", "310378", C128), (null, "250477", C128Dcr));

        IReadOnlyList<MasterListingRow> rows = this.Rows();

        Assert.Equal(["310378", "250477"], rows.Select(row => row.BoardName));

        // A blank hardware cell takes the name above it, like CRT's drop-downs.
        Assert.Equal("Commodore 128", rows[1].HardwareName);
        Assert.Equal("Commodore/C128/250477", rows[1].SystemId);
    }

    [Fact]
    public void An_unreadable_file_is_a_failure_not_an_empty_list()
    {
        File.WriteAllText(this.Master, "not a workbook");

        Assert.False(MasterListing.TryRead(this.Master, out IReadOnlyList<MasterListingRow> rows, out string why));
        Assert.Empty(rows);
        Assert.NotEmpty(why);
    }

    // ------------------------------------------------------------------ inserting

    [Fact]
    public void A_new_system_goes_straight_after_the_row_it_was_placed_after()
    {
        this.WriteRealisticMaster();

        MasterListingEdit edit = MasterListing.Insert(this.Master, Open128Row, C128Dcr);

        Assert.True(edit.IsDone, edit.Failure);
        Assert.True(edit.Changed);
        Assert.Equal(["250407 (long board)", "310378 (C128 & C128D)", "250477 (C128DCR)", "310378 Open128", "Issue 4B"], this.Boards());
        Assert.Equal(Open128Row, this.Rows()[3]);
    }

    [Fact]
    public void Placed_after_nothing_it_is_the_first_row()
    {
        this.WriteRealisticMaster();

        Assert.True(MasterListing.Insert(this.Master, Open128Row, null).IsDone);

        Assert.Equal("310378 Open128", this.Boards()[0]);
        Assert.Equal(5, this.Rows().Count);
    }

    // Re-running an interrupted publish must not list the system twice.
    [Fact]
    public void Inserting_a_system_already_listed_updates_its_row_instead_of_adding_one()
    {
        this.WriteRealisticMaster();
        Assert.True(MasterListing.Insert(this.Master, Open128Row, C128Dcr).IsDone);

        MasterListingEdit again = MasterListing.Insert(this.Master, Open128Row, C128Dcr);

        Assert.True(again.IsDone);
        Assert.False(again.Changed);
        Assert.Single(this.Rows(), row => row.SystemId == Open128Row.SystemId);

        MasterListingEdit renamed = MasterListing.Insert(this.Master, Open128Row with { BoardName = "310378 Open128 (replica)" }, C64Long);

        Assert.True(renamed.Changed);
        Assert.Equal(["250407 (long board)", "310378 (C128 & C128D)", "250477 (C128DCR)", "310378 Open128 (replica)", "Issue 4B"], this.Boards());
    }

    // The maintainer placed it relative to that row; anywhere else is a place nobody chose.
    [Fact]
    public void Placed_after_a_row_no_longer_listed_it_is_refused_and_the_file_is_untouched()
    {
        this.WriteRealisticMaster();
        byte[] before = File.ReadAllBytes(this.Master);

        MasterListingEdit edit = MasterListing.Insert(this.Master, Open128Row, "Commodore/C128/999999/Data C128 999999 v2.0.0.xlsx");

        Assert.False(edit.IsDone);
        Assert.Contains("no longer in the list", edit.Failure);
        Assert.Equal(before, File.ReadAllBytes(this.Master));
    }

    // ###########################################################################################
    // *** TWO SYSTEMS UNDER THE SAME NAMES WOULD BE ONE BOARD TWICE. *** CRT keys a board by
    // "hardware name|board name" without regard to case, so a second row with the same pair would
    // share every setting and workbook with the first. Refused, with the file left alone.
    // ###########################################################################################
    [Fact]
    public void A_system_under_names_another_system_is_listed_under_is_refused_and_the_file_is_untouched()
    {
        this.WriteRealisticMaster();
        byte[] before = File.ReadAllBytes(this.Master);

        MasterListingEdit edit = MasterListing.Insert(
            this.Master,
            Open128Row with { HardwareName = "commodore 128", BoardName = " 250477 (C128DCR) " },
            C128);

        Assert.False(edit.IsDone);
        Assert.Contains("[Commodore 128] / [250477 (C128DCR)] is already in the drop-down lists, for Commodore/C128/250477", edit.Failure);
        Assert.Equal(before, File.ReadAllBytes(this.Master));
    }

    [Fact]
    public void Only_ANOTHER_systems_row_under_the_same_names_is_a_clash()
    {
        this.WriteRealisticMaster();
        IReadOnlyList<MasterListingRow> rows = this.Rows();

        Assert.Equal(C128Dcr, MasterListing.NamesTakenBy(rows, "Commodore/C128/310378 Open128", "COMMODORE 128", "250477 (c128dcr)")?.ExcelDataFile);

        // Its own row - updating a listed system keeps its names.
        Assert.Null(MasterListing.NamesTakenBy(rows, "Commodore/C128/250477", "Commodore 128", "250477 (C128DCR)"));

        // The same hardware with a board name of its own, or the same board name under other hardware.
        Assert.Null(MasterListing.NamesTakenBy(rows, "Commodore/C128/310378 Open128", "Commodore 128", "310378 Open128"));
        Assert.Null(MasterListing.NamesTakenBy(rows, "Commodore/C128/310378 Open128", "Commodore 64", "250477 (C128DCR)"));
    }

    // Updating a listed system's own row under its own names is not a clash with itself.
    [Fact]
    public void A_listed_system_keeps_its_names_when_its_row_is_updated()
    {
        this.WriteRealisticMaster();

        MasterListingEdit edit = MasterListing.Insert(
            this.Master,
            new MasterListingRow("Commodore 128", "250477 (C128DCR)", C128Dcr, "New notes."),
            afterExcelDataFile: null);

        Assert.True(edit.IsDone, edit.Failure);
        Assert.Equal("New notes.", this.Rows().Single(row => row.ExcelDataFile == C128Dcr).Notes);
    }

    // ###########################################################################################
    // CRT carries a blank hardware cell forward from the row above. A new row inserted above a
    // blank-named board must not hand that board ITS name - the board would move to another
    // hardware in everybody's drop-down.
    // ###########################################################################################
    [Fact]
    public void A_board_below_whose_hardware_cell_is_blank_keeps_its_own_hardware()
    {
        this.WriteMaster(("Commodore 64", "250407", C64Long), (null, "250425", "Commodore/C64/250425/Data C64 250425 v2.0.0.xlsx"));

        Assert.True(MasterListing.Insert(this.Master, Open128Row, C64Long).IsDone);

        IReadOnlyList<MasterListingRow> rows = this.Rows();
        Assert.Equal("Commodore 128", rows[1].HardwareName);
        Assert.Equal("Commodore 64", rows[2].HardwareName);
    }

    // The preamble, the Oscilloscope sheet and every other row come through as they were.
    [Fact]
    public void Nothing_but_the_one_row_changes()
    {
        this.WriteRealisticMaster();

        Assert.True(MasterListing.Insert(this.Master, Open128Row, C128Dcr).IsDone);

        using var package = new ExcelPackage(new FileInfo(this.Master));
        ExcelWorksheet sheet = package.Workbook.Worksheets[MasterWorkbookSchema.SheetName];

        Assert.Equal("# Revision date: 2026-August-7", sheet.Cells[3, 1].Text);
        Assert.Equal(MasterWorkbookSchema.ColHardwareName, sheet.Cells[9, 1].Text);
        Assert.Equal("notes Issue 4B", sheet.Cells[14, 4].Text);
        Assert.Equal("Rigol", package.Workbook.Worksheets["Oscilloscope"].Cells[2, 1].Text);

        // Written in place, with no temporary file left beside it.
        Assert.Single(Directory.GetFiles(this.thisWorkspace.Root));
    }

    [Theory]
    [InlineData("", "310378 Open128", "hardware name")]
    [InlineData("Commodore 128", "  ", "board name")]
    [InlineData("Commodore\n128", "310378 Open128", "one line")]
    public void A_row_that_cannot_be_listed_is_refused(string hardware, string board, string expected)
    {
        this.WriteRealisticMaster();

        MasterListingEdit edit = MasterListing.Insert(this.Master, Open128Row with { HardwareName = hardware, BoardName = board }, null);

        Assert.False(edit.IsDone);
        Assert.Contains(expected, edit.Failure);
    }

    [Fact]
    public void A_master_without_all_four_columns_takes_no_row()
    {
        using (var package = new ExcelPackage())
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(MasterWorkbookSchema.SheetName);
            sheet.Cells[1, 3].Value = MasterWorkbookSchema.ColExcelDataFile;
            package.SaveAs(new FileInfo(this.Master));
        }

        Assert.False(MasterListing.Insert(this.Master, Open128Row, null).IsDone);
    }

    // ------------------------------------------------------------------ removing

    [Fact]
    public void Removing_takes_the_systems_row_out_and_nothing_else()
    {
        this.WriteRealisticMaster();
        Assert.True(MasterListing.Insert(this.Master, Open128Row, C128Dcr).IsDone);

        MasterListingEdit edit = MasterListing.Remove(this.Master, Open128Row.SystemId);

        Assert.True(edit.IsDone);
        Assert.True(edit.Changed);
        Assert.Equal(["250407 (long board)", "310378 (C128 & C128D)", "250477 (C128DCR)", "Issue 4B"], this.Boards());
    }

    [Fact]
    public void Removing_a_row_does_not_move_a_blank_named_board_below_it()
    {
        this.WriteMaster(("Commodore 128", "310378 Open128", Open128), (null, "250477", C128Dcr), ("Commodore 64", "250407", C64Long));
        // The blank cell below now leans on the removed row for "Commodore 128".

        Assert.True(MasterListing.Remove(this.Master, Open128Row.SystemId).IsDone);

        Assert.Equal("Commodore 128", this.Rows()[0].HardwareName);
        Assert.Equal("250477", this.Rows()[0].BoardName);
    }

    [Fact]
    public void Removing_a_system_not_listed_changes_nothing()
    {
        this.WriteRealisticMaster();

        MasterListingEdit edit = MasterListing.Remove(this.Master, Open128Row.SystemId);

        Assert.True(edit.IsDone);
        Assert.False(edit.Changed);
    }

    // ------------------------------------------------------------------ the same place in another tree

    private static MasterListingRow Row(string workbook) => new("H", "B", workbook, string.Empty);

    // ###########################################################################################
    // Production's list, when the system is promoted (owner decision: "insert at the same place").
    // After the nearest row above it that production also lists.
    // ###########################################################################################
    [Fact]
    public void In_production_it_goes_after_the_nearest_row_above_it_that_production_lists()
    {
        IReadOnlyList<MasterListingRow> beta = [Row(C64Long), Row(C128), Row(C128Dcr), Row(Open128), Row(Spectrum)];
        IReadOnlyList<MasterListingRow> production = [Row(C64Long), Row(C128), Row(Spectrum)];

        Assert.True(MasterListing.TryResolvePlacement(beta, Row(Open128).SystemId, production, out string? after));
        Assert.Equal(C128, after);
    }

    [Fact]
    public void With_nothing_above_it_in_production_it_goes_before_the_nearest_row_below()
    {
        IReadOnlyList<MasterListingRow> beta = [Row(Open128), Row(C128), Row(Spectrum)];
        IReadOnlyList<MasterListingRow> production = [Row(C64Long), Row(Spectrum)];

        Assert.True(MasterListing.TryResolvePlacement(beta, Row(Open128).SystemId, production, out string? after));
        Assert.Equal(C64Long, after);

        IReadOnlyList<MasterListingRow> spectrumFirst = [Row(Spectrum), Row(C64Long)];
        Assert.True(MasterListing.TryResolvePlacement(beta, Row(Open128).SystemId, spectrumFirst, out after));
        Assert.Null(after);
    }

    [Fact]
    public void A_system_the_source_does_not_list_has_no_place()
    {
        Assert.False(MasterListing.TryResolvePlacement([Row(C64Long)], Row(Open128).SystemId, [Row(C64Long)], out _));
    }

    // ------------------------------------------------------------------ which file

    // Only the newest VERSIONED master is ever written - older generations are frozen.
    [Fact]
    public void Only_the_newest_versioned_master_is_named_for_writing()
    {
        File.WriteAllText(Path.Combine(this.thisWorkspace.Root, "Classic-Repair-Toolbox.xlsx"), "old");

        Assert.Null(MasterListing.NewestMasterPath(this.thisWorkspace.Root));

        this.WriteRealisticMaster();

        Assert.Equal(this.Master, MasterListing.NewestMasterPath(this.thisWorkspace.Root));
    }
}
