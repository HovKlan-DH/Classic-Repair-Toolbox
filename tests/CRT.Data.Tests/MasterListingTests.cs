using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// MasterListing - a new board's ONE row in the main Excel data file (owner request, 2026-09-27:
// "When a board is added to BETA, and it is a NEW board, can you then make sure it gets added
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
        using var package = EpplusLicense.NewPackage();
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
        Assert.Equal("Commodore/C128/250477", rows[1].BoardId);
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
    public void A_new_board_goes_straight_after_the_row_it_was_placed_after()
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

    // Re-running an interrupted publish must not list the board twice.
    [Fact]
    public void Inserting_a_board_already_listed_updates_its_row_instead_of_adding_one()
    {
        this.WriteRealisticMaster();
        Assert.True(MasterListing.Insert(this.Master, Open128Row, C128Dcr).IsDone);

        MasterListingEdit again = MasterListing.Insert(this.Master, Open128Row, C128Dcr);

        Assert.True(again.IsDone);
        Assert.False(again.Changed);
        Assert.Single(this.Rows(), row => row.BoardId == Open128Row.BoardId);

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
    // *** TWO BOARDS UNDER THE SAME NAMES WOULD BE ONE BOARD TWICE. *** CRT keys a board by
    // "hardware name|board name" without regard to case, so a second row with the same pair would
    // share every setting and workbook with the first. Refused, with the file left alone.
    // ###########################################################################################
    [Fact]
    public void A_board_under_names_another_board_is_listed_under_is_refused_and_the_file_is_untouched()
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
    public void Only_ANOTHER_boards_row_under_the_same_names_is_a_clash()
    {
        this.WriteRealisticMaster();
        IReadOnlyList<MasterListingRow> rows = this.Rows();

        Assert.Equal(C128Dcr, MasterListing.NamesTakenBy(rows, "Commodore/C128/310378 Open128", "COMMODORE 128", "250477 (c128dcr)")?.ExcelDataFile);

        // Its own row - updating a listed board keeps its names.
        Assert.Null(MasterListing.NamesTakenBy(rows, "Commodore/C128/250477", "Commodore 128", "250477 (C128DCR)"));

        // The same hardware with a board name of its own, or the same board name under other hardware.
        Assert.Null(MasterListing.NamesTakenBy(rows, "Commodore/C128/310378 Open128", "Commodore 128", "310378 Open128"));
        Assert.Null(MasterListing.NamesTakenBy(rows, "Commodore/C128/310378 Open128", "Commodore 64", "250477 (C128DCR)"));
    }

    // Updating a listed board's own row under its own names is not a clash with itself.
    [Fact]
    public void A_listed_board_keeps_its_names_when_its_row_is_updated()
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

    // The rest of the preamble, the Oscilloscope sheet and every other row come through as they
    // were - the date line is the one thing a write changes besides its row (see below).
    [Fact]
    public void Nothing_but_the_one_row_and_the_date_changes()
    {
        this.WriteRealisticMaster();

        Assert.True(MasterListing.Insert(this.Master, Open128Row, C128Dcr, MasterListingTests.Now).IsDone);

        using var package = EpplusLicense.OpenPackage(new FileInfo(this.Master));
        ExcelWorksheet sheet = package.Workbook.Worksheets[MasterWorkbookSchema.SheetName];

        Assert.Equal("# Commodore Repair Toolbox", sheet.Cells[1, 1].Text);
        Assert.Equal("# Revision date: 2026-October-4", sheet.Cells[3, 1].Text);
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
        using (var package = EpplusLicense.NewPackage())
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(MasterWorkbookSchema.SheetName);
            sheet.Cells[1, 3].Value = MasterWorkbookSchema.ColExcelDataFile;
            package.SaveAs(new FileInfo(this.Master));
        }

        Assert.False(MasterListing.Insert(this.Master, Open128Row, null).IsDone);
    }

    // ------------------------------------------------------------------ removing

    [Fact]
    public void Removing_takes_the_boards_row_out_and_nothing_else()
    {
        this.WriteRealisticMaster();
        Assert.True(MasterListing.Insert(this.Master, Open128Row, C128Dcr).IsDone);

        MasterListingEdit edit = MasterListing.Remove(this.Master, Open128Row.BoardId);

        Assert.True(edit.IsDone);
        Assert.True(edit.Changed);
        Assert.Equal(["250407 (long board)", "310378 (C128 & C128D)", "250477 (C128DCR)", "Issue 4B"], this.Boards());
    }

    [Fact]
    public void Removing_a_row_does_not_move_a_blank_named_board_below_it()
    {
        this.WriteMaster(("Commodore 128", "310378 Open128", Open128), (null, "250477", C128Dcr), ("Commodore 64", "250407", C64Long));
        // The blank cell below now leans on the removed row for "Commodore 128".

        Assert.True(MasterListing.Remove(this.Master, Open128Row.BoardId).IsDone);

        Assert.Equal("Commodore 128", this.Rows()[0].HardwareName);
        Assert.Equal("250477", this.Rows()[0].BoardName);
    }

    // ###########################################################################################
    // Two rows whose workbooks sit in the same folder are one board: deleting it takes the whole
    // folder, so a row left behind would be a board CRT offers but cannot load (code review,
    // 2026-10-04). Every row goes, the rows between them stay, and a blank-named board after either
    // keeps its own hardware name.
    // ###########################################################################################
    [Fact]
    public void Removing_takes_out_every_row_of_the_board()
    {
        const string Open128Second = "Commodore/C128/310378 Open128/Data C128 310378 Open128 rev B v2.0.0.xlsx";

        this.WriteMaster(
            ("Commodore 128", "310378 Open128", Open128),
            (null, "250477", C128Dcr),
            ("Commodore 128", "310378 Open128 rev B", Open128Second),
            (null, "310378", C128),
            ("Commodore 64", "250407", C64Long));

        MasterListingEdit edit = MasterListing.Remove(this.Master, Open128Row.BoardId);

        Assert.True(edit.IsDone, edit.Failure);
        Assert.True(edit.Changed);
        Assert.Equal(["250477", "310378", "250407"], this.Boards());
        Assert.Equal(["Commodore 128", "Commodore 128", "Commodore 64"], this.Rows().Select(row => row.HardwareName));
    }

    [Fact]
    public void Removing_a_board_not_listed_changes_nothing()
    {
        this.WriteRealisticMaster();

        MasterListingEdit edit = MasterListing.Remove(this.Master, Open128Row.BoardId);

        Assert.True(edit.IsDone);
        Assert.False(edit.Changed);
    }

    // ------------------------------------------------------------------ the same place in another tree

    private static MasterListingRow Row(string workbook) => new("H", "B", workbook, string.Empty);

    // ###########################################################################################
    // Production's list, when the board is promoted (owner decision: "insert at the same place").
    // After the nearest row above it that production also lists.
    // ###########################################################################################
    [Fact]
    public void In_production_it_goes_after_the_nearest_row_above_it_that_production_lists()
    {
        IReadOnlyList<MasterListingRow> beta = [Row(C64Long), Row(C128), Row(C128Dcr), Row(Open128), Row(Spectrum)];
        IReadOnlyList<MasterListingRow> production = [Row(C64Long), Row(C128), Row(Spectrum)];

        Assert.True(MasterListing.TryResolvePlacement(beta, Row(Open128).BoardId, production, out string? after));
        Assert.Equal(C128, after);
    }

    [Fact]
    public void With_nothing_above_it_in_production_it_goes_before_the_nearest_row_below()
    {
        IReadOnlyList<MasterListingRow> beta = [Row(Open128), Row(C128), Row(Spectrum)];
        IReadOnlyList<MasterListingRow> production = [Row(C64Long), Row(Spectrum)];

        Assert.True(MasterListing.TryResolvePlacement(beta, Row(Open128).BoardId, production, out string? after));
        Assert.Equal(C64Long, after);

        IReadOnlyList<MasterListingRow> spectrumFirst = [Row(Spectrum), Row(C64Long)];
        Assert.True(MasterListing.TryResolvePlacement(beta, Row(Open128).BoardId, spectrumFirst, out after));
        Assert.Null(after);
    }

    [Fact]
    public void A_board_the_source_does_not_list_has_no_place()
    {
        Assert.False(MasterListing.TryResolvePlacement([Row(C64Long)], Row(Open128).BoardId, [Row(C64Long)], out _));
    }

    // ------------------------------------------------------------------ every write: dates and panes

    private static readonly DateTimeOffset Now = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

    // ###########################################################################################
    // The file as it is on the server (owner, 2026-10-04): both sheets open with three "#" lines, the
    // third the revision date; the boards' header on row 5, the oscilloscopes' on row 5 with many
    // columns to its right.
    // ###########################################################################################
    private void WriteServerShapedMaster(bool oscilloscopeDate = true)
    {
        using var package = EpplusLicense.NewPackage();
        ExcelWorksheet boards = package.Workbook.Worksheets.Add(MasterWorkbookSchema.SheetName);

        boards.Cells[1, 1].Value = "# Commodore Repair Toolbox";
        boards.Cells[2, 1].Value = "# Hardware and board definitions";
        boards.Cells[3, 1].Value = "# Revision date: 2026-August-7";
        boards.Cells[5, 1].Value = MasterWorkbookSchema.ColHardwareName;
        boards.Cells[5, 2].Value = MasterWorkbookSchema.ColBoardName;
        boards.Cells[5, 3].Value = MasterWorkbookSchema.ColExcelDataFile;
        boards.Cells[5, 4].Value = MasterWorkbookSchema.ColHardwareNotes;
        boards.Cells[6, 1].Value = "Commodore 64";
        boards.Cells[6, 2].Value = "250407 (long board)";
        boards.Cells[6, 3].Value = C64Long;

        ExcelWorksheet scopes = package.Workbook.Worksheets.Add(MasterWorkbookSchema.OscilloscopeSheetName);

        scopes.Cells[1, 1].Value = "# Commodore Repair Toolbox";
        scopes.Cells[2, 1].Value = "# Oscilloscope configurations";

        if (oscilloscopeDate)
        {
            scopes.Cells[3, 1].Value = "# Revision date: 2026-August-7";
        }

        scopes.Cells[5, 1].Value = MasterWorkbookSchema.ColBrand;
        scopes.Cells[5, 2].Value = MasterWorkbookSchema.ColSeriesOrModel;
        scopes.Cells[5, 3].Value = MasterWorkbookSchema.ColPort;
        scopes.Cells[6, 1].Value = "Rigol";
        scopes.Cells[6, 2].Value = "DS1000Z";

        package.SaveAs(new FileInfo(this.Master));
    }

    // The <pane> element Excel freezes a sheet with, as written in the file - sheet1 is the first sheet.
    private string Pane(int sheet)
    {
        using ZipArchive zip = ZipFile.OpenRead(this.Master);
        using var reader = new StreamReader(zip.GetEntry($"xl/worksheets/sheet{sheet}.xml")!.Open());

        return Regex.Match(reader.ReadToEnd(), "<pane [^>]*/>").Value;
    }

    private string Cell(string sheet, int row, int column)
    {
        using var package = EpplusLicense.OpenPackage(new FileInfo(this.Master));
        return package.Workbook.Worksheets[sheet].Cells[row, column].Text;
    }

    // ###########################################################################################
    // *** EVERY WRITE DATES BOTH SHEETS (owner request, 2026-10-04: "make sure to update the date in
    // the 'Revision date:'" - and for the Oscilloscope sheet, "Again, update the date"). ***
    // ###########################################################################################
    [Fact]
    public void A_write_dates_the_revision_line_of_both_sheets()
    {
        this.WriteServerShapedMaster();

        Assert.True(MasterListing.Insert(this.Master, Open128Row, C64Long, MasterListingTests.Now).IsDone);

        Assert.Equal("# Revision date: 2026-October-4", this.Cell(MasterWorkbookSchema.SheetName, 3, 1));
        Assert.Equal("# Revision date: 2026-October-4", this.Cell(MasterWorkbookSchema.OscilloscopeSheetName, 3, 1));

        // The lines around it are left as they were.
        Assert.Equal("# Hardware and board definitions", this.Cell(MasterWorkbookSchema.SheetName, 2, 1));
        Assert.Equal("# Oscilloscope configurations", this.Cell(MasterWorkbookSchema.OscilloscopeSheetName, 2, 1));
        Assert.Equal("DS1000Z", this.Cell(MasterWorkbookSchema.OscilloscopeSheetName, 6, 2));
    }

    // ###########################################################################################
    // *** AND FREEZES THEIR PANES (owner request, 2026-10-04). *** "Hardware & Board" under its
    // header row, as every board workbook is; "Oscilloscope" under its header row AND after its
    // second column, so the brand and model stay on screen while its many columns scroll by. Read
    // from the file's own XML: ySplit is the rows frozen, xSplit the columns.
    // ###########################################################################################
    [Fact]
    public void A_write_freezes_the_boards_under_the_header_and_the_oscilloscopes_after_two_columns_too()
    {
        this.WriteServerShapedMaster();

        Assert.True(MasterListing.Insert(this.Master, Open128Row, C64Long, MasterListingTests.Now).IsDone);

        string boards = this.Pane(1);
        Assert.Contains("ySplit=\"5\"", boards);
        Assert.Contains("state=\"frozen\"", boards);
        Assert.DoesNotContain("xSplit", boards);

        string scopes = this.Pane(2);
        Assert.Contains("ySplit=\"5\"", scopes);
        Assert.Contains("xSplit=\"2\"", scopes);
        Assert.Contains("state=\"frozen\"", scopes);
    }

    // ###########################################################################################
    // *** THE DATE LINE KEEPS ITS LOOK. *** The shipped file's line is rich text, as a board
    // workbook's is: a plain "# Revision date: " and the date in bold, size 16. Writing the cell's
    // value would flatten it to one plain run - so only the runs' text changes. (Checked against the
    // shipped file's own XML, 2026-10-04.)
    // ###########################################################################################
    [Fact]
    public void A_rich_text_date_line_keeps_its_bold_date()
    {
        this.WriteServerShapedMaster();

        using (var package = EpplusLicense.OpenPackage(new FileInfo(this.Master)))
        {
            foreach (string name in new[] { MasterWorkbookSchema.SheetName, MasterWorkbookSchema.OscilloscopeSheetName })
            {
                ExcelRange cell = package.Workbook.Worksheets[name].Cells[3, 1];
                cell.Value = null;
                cell.RichText.Add("# Revision date: ");
                OfficeOpenXml.Style.ExcelRichText date = cell.RichText.Add("2026-August-7");
                date.Bold = true;
                date.Size = 16;
            }

            package.Save();
        }

        Assert.True(MasterListing.Insert(this.Master, Open128Row, C64Long, MasterListingTests.Now).IsDone);

        using var after = EpplusLicense.OpenPackage(new FileInfo(this.Master));

        foreach (string name in new[] { MasterWorkbookSchema.SheetName, MasterWorkbookSchema.OscilloscopeSheetName })
        {
            ExcelRange cell = after.Workbook.Worksheets[name].Cells[3, 1];

            Assert.Equal("# Revision date: 2026-October-4", cell.Text);
            Assert.True(cell.IsRichText);
            Assert.False(cell.RichText[0].Bold);
            Assert.Equal("2026-October-4", cell.RichText[1].Text);
            Assert.True(cell.RichText[1].Bold);
            Assert.Equal(16f, cell.RichText[1].Size);
        }
    }

    // A preamble with no date line is not given one - nothing is invented in a preamble the project
    // owner wrote; the panes are still set.
    [Fact]
    public void A_sheet_without_a_date_line_is_not_given_one()
    {
        this.WriteServerShapedMaster(oscilloscopeDate: false);

        Assert.True(MasterListing.Remove(this.Master, C64Long[..C64Long.LastIndexOf('/')], MasterListingTests.Now).IsDone);

        Assert.Equal(string.Empty, this.Cell(MasterWorkbookSchema.OscilloscopeSheetName, 3, 1));
        Assert.Equal(MasterWorkbookSchema.ColBrand, this.Cell(MasterWorkbookSchema.OscilloscopeSheetName, 5, 1));
        Assert.Contains("xSplit=\"2\"", this.Pane(2));
    }

    // A write that changes nothing writes nothing - not even the date.
    [Fact]
    public void A_write_that_changes_nothing_leaves_the_file_and_its_date_alone()
    {
        this.WriteServerShapedMaster();
        byte[] before = File.ReadAllBytes(this.Master);

        MasterListingEdit edit = MasterListing.Insert(
            this.Master, new MasterListingRow("Commodore 64", "250407 (long board)", C64Long, string.Empty), null, MasterListingTests.Now);

        Assert.True(edit.IsDone);
        Assert.False(edit.Changed);
        Assert.Equal(before, File.ReadAllBytes(this.Master));
    }

    // ------------------------------------------------------------------ the order of the lists

    private static IReadOnlyList<string> Ids(params string[] workbooks) =>
        workbooks.Select(workbook => Row(workbook).BoardId).ToList();

    // ###########################################################################################
    // *** THE ORDER OF THE DROP-DOWN LISTS (owner request, 2026-10-04: "sort the list of systems,
    // which then gets saved to both sources (BETA + stable)"). *** Each row the order names goes
    // where the order says.
    // ###########################################################################################
    [Fact]
    public void Arranged_as_an_order_every_named_row_goes_where_the_order_puts_it()
    {
        IReadOnlyList<MasterListingRow> rows = [Row(C64Long), Row(C128), Row(C128Dcr), Row(Spectrum)];

        Assert.Equal([3, 2, 0, 1], MasterListing.ArrangeAs(rows, Ids(Spectrum, C128Dcr, C64Long, C128)));
    }

    // ###########################################################################################
    // The STABLE source's list follows BETA's order, but need not hold the same boards: one it
    // lists that BETA's order does not name stays straight after the row it followed - and first
    // when it was first. A board BETA names that stable lacks is simply not there.
    // ###########################################################################################
    [Fact]
    public void A_row_the_order_does_not_name_stays_after_the_row_it_followed()
    {
        // Stable: C64, Spectrum (not in BETA's order), C128 - BETA's order: C128, Open128, C64.
        IReadOnlyList<MasterListingRow> stable = [Row(C64Long), Row(Spectrum), Row(C128)];

        Assert.Equal([2, 0, 1], MasterListing.ArrangeAs(stable, Ids(C128, Open128, C64Long)));

        // First and unnamed: it stays first.
        IReadOnlyList<MasterListingRow> spectrumFirst = [Row(Spectrum), Row(C64Long), Row(C128)];

        Assert.Equal([0, 2, 1], MasterListing.ArrangeAs(spectrumFirst, Ids(C128, C64Long)));
    }

    [Fact]
    public void Reordering_moves_whole_rows_and_keeps_each_ones_notes()
    {
        this.WriteRealisticMaster();

        MasterListingEdit edit = MasterListing.Reorder(this.Master, Ids(Spectrum, C128Dcr, C64Long, C128), MasterListingTests.Now);

        Assert.True(edit.IsDone, edit.Failure);
        Assert.True(edit.Changed);

        IReadOnlyList<MasterListingRow> rows = this.Rows();

        Assert.Equal(["Issue 4B", "250477 (C128DCR)", "250407 (long board)", "310378 (C128 & C128D)"], rows.Select(row => row.BoardName));
        Assert.Equal(["notes Issue 4B", "notes 250477 (C128DCR)", "notes 250407 (long board)", "notes 310378 (C128 & C128D)"], rows.Select(row => row.Notes));
        Assert.Equal(["ZX Spectrum 16K/48K", "Commodore 128", "Commodore 64", "Commodore 128"], rows.Select(row => row.HardwareName));

        // In place: the header where it was, nothing left below the list, no temporary file.
        Assert.Equal(MasterWorkbookSchema.ColHardwareName, this.Cell(MasterWorkbookSchema.SheetName, 9, 1));
        Assert.Equal(string.Empty, this.Cell(MasterWorkbookSchema.SheetName, 14, 2));
        Assert.Single(Directory.GetFiles(this.thisWorkspace.Root));
    }

    // ###########################################################################################
    // *** A BLANK HARDWARE CELL KEEPS ITS OWN HARDWARE. *** CRT carries a blank hardware cell forward
    // from the row above, so a blank-named C128 board moved under a C64 row would turn into a C64
    // board in everybody's drop-down. Its own name is written into its cell instead.
    // ###########################################################################################
    [Fact]
    public void A_moved_board_whose_hardware_cell_is_blank_keeps_its_hardware()
    {
        this.WriteMaster(("Commodore 128", "310378", C128), (null, "250477", C128Dcr), ("Commodore 64", "250407", C64Long));

        Assert.True(MasterListing.Reorder(this.Master, Ids(C128, C64Long, C128Dcr), MasterListingTests.Now).IsDone);

        IReadOnlyList<MasterListingRow> rows = this.Rows();

        Assert.Equal(["310378", "250407", "250477"], rows.Select(row => row.BoardName));
        Assert.Equal(["Commodore 128", "Commodore 64", "Commodore 128"], rows.Select(row => row.HardwareName));
    }

    // An order the file already has writes nothing at all.
    [Fact]
    public void Reordering_into_the_order_the_file_has_writes_nothing()
    {
        this.WriteRealisticMaster();
        byte[] before = File.ReadAllBytes(this.Master);

        MasterListingEdit edit = MasterListing.Reorder(this.Master, Ids(C64Long, C128, C128Dcr, Spectrum), MasterListingTests.Now);

        Assert.True(edit.IsDone);
        Assert.False(edit.Changed);
        Assert.Equal(before, File.ReadAllBytes(this.Master));
    }

    [Fact]
    public void Reordering_dates_the_file()
    {
        this.WriteServerShapedMaster();
        Assert.True(MasterListing.Insert(this.Master, Open128Row, C64Long, MasterListingTests.Now.AddDays(-30)).IsDone);

        Assert.True(MasterListing.Reorder(this.Master, Ids(Open128, C64Long), MasterListingTests.Now).IsDone);

        Assert.Equal(["310378 Open128", "250407 (long board)"], this.Boards());
        Assert.Equal("# Revision date: 2026-October-4", this.Cell(MasterWorkbookSchema.SheetName, 3, 1));
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
