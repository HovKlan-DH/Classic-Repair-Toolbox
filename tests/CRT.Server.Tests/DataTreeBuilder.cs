using Handlers.DataHandling;
using OfficeOpenXml;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Writes a REAL data tree for the tests of what a publish removes: master workbooks listing
    // board workbooks, boards citing files, and the files themselves. DataTreeUsage reads every
    // one of these, so a placeholder master (the plain-text "master" other fixtures write, which is
    // enough for the generation rule) would make the whole tree unreadable - and an unreadable
    // tree removes nothing, which is the one thing these tests need to see NOT happen.
    // ###########################################################################################
    internal static class DataTreeBuilder
    {
        public const string BoardFolder = "Commodore/C64/250407";
        public const string Workbook = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";

        static DataTreeBuilder()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
        }

        // The v2.0.0 master, listing the given board workbooks in its "Excel data file" column.
        public static void Master(string root, params string[] listed) =>
            DataTreeBuilder.ListingMaster(
                root,
                listed.Select(workbook => new MasterListingRow("Hardware", Path.GetFileNameWithoutExtension(workbook), workbook, string.Empty)).ToArray());

        // ###########################################################################################
        // The v2.0.0 master as CRT reads it - all four columns under the real names, a preamble above
        // the header - listing these rows in order. What a NEW system's publish adds its row to
        // (2026-09-27); a placeholder text file cannot take a row, and a new system is then refused.
        // ###########################################################################################
        public static void ListingMaster(string root, params MasterListingRow[] rows)
        {
            using var package = EpplusLicense.NewPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(MasterWorkbookSchema.SheetName);

            sheet.Cells[1, 1].Value = "# Commodore Repair Toolbox";
            sheet.Cells[3, 1].Value = MasterWorkbookSchema.ColHardwareName;
            sheet.Cells[3, 2].Value = MasterWorkbookSchema.ColBoardName;
            sheet.Cells[3, 3].Value = MasterWorkbookSchema.ColExcelDataFile;
            sheet.Cells[3, 4].Value = MasterWorkbookSchema.ColHardwareNotes;

            for (int i = 0; i < rows.Length; i++)
            {
                sheet.Cells[4 + i, 1].Value = rows[i].HardwareName;
                sheet.Cells[4 + i, 2].Value = rows[i].BoardName;
                sheet.Cells[4 + i, 3].Value = rows[i].ExcelDataFile;
                sheet.Cells[4 + i, 4].Value = rows[i].Notes;
            }

            package.Workbook.Worksheets.Add("Oscilloscope").Cells[1, 1].Value = "Brand";

            package.SaveAs(new FileInfo(DataTreeBuilder.Full(root, "Classic-Repair-Toolbox.v2.0.0.xlsx")));
        }

        // The rows of a tree's v2.0.0 master, top to bottom.
        public static IReadOnlyList<MasterListingRow> ListedIn(string root)
        {
            Assert.True(
                MasterListing.TryRead(DataTreeBuilder.Full(root, "Classic-Repair-Toolbox.v2.0.0.xlsx"), out IReadOnlyList<MasterListingRow> rows, out string why),
                why);

            return rows;
        }

        // A board workbook whose "Board local files" rows cite the given files, with one component.
        public static void Board(string root, string workbook, params string[] cited)
        {
            var data = new BoardData
            {
                RevisionDate = "2026-August-21",
                Components =
                [
                    new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-01" }
                ]
            };

            foreach (string file in cited)
                data.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = Path.GetFileName(file), File = file });

            BoardWorkbookWriter.Write(DataTreeBuilder.Full(root, workbook), data);
        }

        public static void Files(string root, params string[] paths)
        {
            foreach (string path in paths)
                File.WriteAllText(DataTreeBuilder.Full(root, path), "bytes of " + path);
        }

        public static string Full(string root, string relative)
        {
            string full = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            return full;
        }
    }
}
