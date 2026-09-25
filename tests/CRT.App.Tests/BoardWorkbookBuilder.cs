using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// Builds a board .xlsx in the shape BoardDataReader expects, so reader tests do not depend on
// any contributed board file in Assets/Data staying unchanged.
//
// Sheet and column names here MUST match the constants in BoardDataReader. If a rename there
// is not mirrored here, the reader tests fail - which is the point: the workbook schema is a
// contract with every contributor's Excel file.
internal sealed class BoardWorkbookBuilder
{
    private readonly ExcelPackage thisPackage;

    static BoardWorkbookBuilder()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public BoardWorkbookBuilder()
    {
        this.thisPackage = new ExcelPackage();
    }

    /// <summary>Adds a sheet whose first row is the header and whose later rows are the data.</summary>
    public BoardWorkbookBuilder Sheet(string name, string[] headers, params string?[][] rows)
    {
        return this.SheetWithLeadingRows(name, Array.Empty<string>(), headers, rows);
    }

    /// <summary>
    /// Adds a sheet with some free-text rows above the header row - real board files carry
    /// "# Revision date:" style comments there, and the reader scans for the header.
    /// </summary>
    public BoardWorkbookBuilder SheetWithLeadingRows(
        string name, string[] leadingRows, string[] headers, params string?[][] rows)
    {
        var sheet = this.thisPackage.Workbook.Worksheets.Add(name);

        int row = 1;
        foreach (string leading in leadingRows)
        {
            sheet.Cells[row, 1].Value = leading;
            row++;
        }

        for (int c = 0; c < headers.Length; c++)
        {
            sheet.Cells[row, c + 1].Value = headers[c];
        }

        row++;

        foreach (string?[] dataRow in rows)
        {
            for (int c = 0; c < dataRow.Length; c++)
            {
                sheet.Cells[row, c + 1].Value = dataRow[c];
            }

            row++;
        }

        return this;
    }

    public string SaveTo(string path)
    {
        this.thisPackage.SaveAs(new FileInfo(path));
        this.thisPackage.Dispose();
        return path;
    }

    // ------------------------------------------------------------------ sheet schemas

    public static readonly string[] SchematicsHeaders =
    {
        "UUID v4", "Schematic name", "CAD name", "Schematic image file",
        "Schematic highlight color", "Schematic highlight opacity",
        "Opposite trace highlight color", "Thumbnail highlight color", "Thumbnail highlight opacity"
    };

    public static readonly string[] ComponentsHeaders =
    {
        "UUID v4", "Board label", "Friendly name", "Technical name or value",
        "Part-number", "Category", "Region",
        "Short one-liner description (one short line only!)"
    };

    public static readonly string[] ComponentImagesHeaders =
    {
        "UUID v4", "Board label", "Region", "Pin", "Name",
        "Expected oscilloscope reading", "File", "Note", "T/DIV", "V/DIV", "T.LVL"
    };

    public static readonly string[] ComponentLocalFilesHeaders = { "UUID v4", "Board label", "Name", "File" };
    public static readonly string[] ComponentLinksHeaders = { "UUID v4", "Board label", "Name", "URL" };
    public static readonly string[] BoardLocalFilesHeaders = { "Category", "Name", "File" };
    public static readonly string[] BoardLinksHeaders = { "Category", "Name", "URL" };
    public static readonly string[] CreditsHeaders = { "Category", "Sub-category", "Name or handle", "Contact (email or web page)" };
    public static readonly string[] ImportantSignalsHeaders = { "Display name", "KiCad net name" };

    /// <summary>A workbook with every sheet the reader looks for, populated with one row each.</summary>
    public static string WriteCompleteBoard(string path, string revisionDate = "2026-01-15")
    {
        return new BoardWorkbookBuilder()
            .SheetWithLeadingRows(
                "Board schematics",
                new[] { $"# Revision date: {revisionDate}" },
                SchematicsHeaders,
                new[] { "uuid-1", "Sheet 1", "board.kicad_pcb", "sheet1.png", "#FF0000", "0.5", "#00FF00", "#0000FF", "0.25" })
            .Sheet("Components", ComponentsHeaders,
                new[] { "uuid-2", "U1", "PLA", "906114-01", "906114", "IC", "PAL", "Programmable logic array" },
                new[] { "uuid-3", "C1", "Filter cap", "100nF", null, "Capacitor", "", "Decoupling" })
            .Sheet("Component images", ComponentImagesHeaders,
                new[] { "uuid-4", "U1", "PAL", "1", "Clock in", "1MHz square", "u1-pin1.png", "note", "5ms", "500mV", "1.65V" })
            .Sheet("Component local files", ComponentLocalFilesHeaders,
                new[] { "uuid-5", "U1", "Datasheet", "906114.pdf" })
            .Sheet("Component links", ComponentLinksHeaders,
                new[] { "uuid-6", "U1", "Reference", "https://example.org/pla" })
            .Sheet("Board local files", BoardLocalFilesHeaders,
                new[] { "Schematics", "Service manual", "manual.pdf" })
            .Sheet("Board links", BoardLinksHeaders,
                new[] { "Forums", "Thread", "https://example.org/thread" })
            .Sheet("Credits", CreditsHeaders,
                new[] { "Data", "Schematics", "Someone", "someone@example.org" })
            .Sheet("Important signals", ImportantSignalsHeaders,
                new[] { "Clock", "/Sheet1/CLK" })
            .SaveTo(path);
    }
}
