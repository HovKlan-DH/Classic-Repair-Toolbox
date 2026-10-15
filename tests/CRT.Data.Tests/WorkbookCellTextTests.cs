using System.Globalization;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// WorkbookCellText is the one way a workbook cell becomes text, so that a comma-decimal machine
// reads what the server reads (code review, 2026-10-13). BoardDataReaderTests proves it end to end
// on a board; these pin the rule cell by cell.
public sealed class WorkbookCellTextTests
{
    static WorkbookCellTextTests() => EpplusLicense.Ensure();

    private static string ReadUnder(string cultureName, Action<ExcelRange> fill)
    {
        CultureInfo original = CultureInfo.CurrentCulture;

        using ExcelPackage package = EpplusLicense.NewPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet");
        fill(sheet.Cells[1, 1]);

        try
        {
            CultureInfo.CurrentCulture = new CultureInfo(cultureName);
            string text = WorkbookCellText.Of(sheet.Cells[1, 1]);

            // The caller's culture is put back.
            Assert.Equal(cultureName, CultureInfo.CurrentCulture.Name);
            return text;
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData("da-DK")]
    [InlineData("de-DE")]
    [InlineData("fr-FR")]
    [InlineData("en-GB")]
    public void A_decimal_number_reads_with_a_point_in_every_culture(string culture)
    {
        Assert.Equal("0.35", ReadUnder(culture, cell => cell.Value = 0.35));
        Assert.Equal("-1.2", ReadUnder(culture, cell => cell.Value = -1.2));
    }

    [Fact]
    public void A_whole_number_reads_without_separators()
    {
        Assert.Equal("6526", ReadUnder("da-DK", cell => cell.Value = 6526d));
        Assert.Equal("1234567", ReadUnder("de-DE", cell => cell.Value = 1234567d));
    }

    [Fact]
    public void A_cell_number_format_still_applies()
    {
        // A percentage typed in Excel keeps its format - the opacity parser accepts "40%".
        Assert.Equal("40%", ReadUnder("da-DK", cell =>
        {
            cell.Value = 0.4;
            cell.Style.Numberformat.Format = "0%";
        }));

        Assert.Equal("0.30", ReadUnder("da-DK", cell =>
        {
            cell.Value = 0.3;
            cell.Style.Numberformat.Format = "0.00";
        }));
    }

    [Fact]
    public void A_text_cell_is_returned_as_typed()
    {
        // A comma typed AS TEXT is the contributor's own text and is not reinterpreted here -
        // SubmissionValidator reports such a value rather than guessing.
        Assert.Equal("0,3", ReadUnder("en-US", cell => cell.Value = "0,3"));
        Assert.Equal(" padded ", ReadUnder("da-DK", cell => cell.Value = " padded "));
    }

    [Fact]
    public void An_empty_cell_reads_as_empty()
    {
        Assert.Equal(string.Empty, ReadUnder("da-DK", _ => { }));
    }

    [Fact]
    public void Trimmed_trims()
    {
        using ExcelPackage package = EpplusLicense.NewPackage();
        ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet");
        sheet.Cells[2, 3].Value = "  U8  ";

        Assert.Equal("U8", WorkbookCellText.Trimmed(sheet, 2, 3));
    }
}
