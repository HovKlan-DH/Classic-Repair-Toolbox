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
        public static void Master(string root, params string[] listed)
        {
            using var package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(DataTreeUsage.MasterSheetName);
            sheet.Cells[1, 1].Value = "Hardware name";
            sheet.Cells[1, 2].Value = "Board name";
            sheet.Cells[1, 3].Value = DataTreeUsage.ExcelDataFileColumn;
            sheet.Cells[1, 4].Value = "Hardware notes";

            for (int i = 0; i < listed.Length; i++)
                sheet.Cells[2 + i, 3].Value = listed[i];

            package.SaveAs(new FileInfo(DataTreeBuilder.Full(root, "Classic-Repair-Toolbox.v2.0.0.xlsx")));
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
