using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClassicRepairToolbox.Tests;
using Handlers.DataHandling;
using OfficeOpenXml;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // DataTreeUsage - which files in a data tree are used, and so which may be removed (owner
    // decision, 2026-09-25: "there must be no orphan files").
    //
    // Every test builds a real tree with real workbooks: the rule is only as good as its reading of
    // the master's "Excel data file" column and the board sheets, and a fake would test neither.
    // The two failure directions are not equal - calling a used file unused DELETES it from every
    // user's download - so most tests here are about what must be KEPT.
    // ###########################################################################################
    [Collection("BoardData")]
    public sealed class DataTreeUsageTests
    {
        internal const string Board = "Commodore/C64/250407";
        internal const string Workbook = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";

        static DataTreeUsageTests()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
        }

        // ---- building a tree -------------------------------------------------------------------

        // A master listing the given board workbooks under a free-text row, as the real ones do.
        internal static void WriteMaster(TempWorkspace tree, string name, params string[] listed)
        {
            using var package = EpplusLicense.NewPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(DataTreeUsage.MasterSheetName);
            sheet.Cells[1, 1].Value = "# Revision date: 2026-09-25";
            sheet.Cells[2, 1].Value = "Hardware name";
            sheet.Cells[2, 2].Value = "Board name";
            sheet.Cells[2, 3].Value = DataTreeUsage.ExcelDataFileColumn;
            sheet.Cells[2, 4].Value = "Hardware notes";

            for (int i = 0; i < listed.Length; i++)
            {
                sheet.Cells[3 + i, 1].Value = "Hardware";
                sheet.Cells[3 + i, 2].Value = "Board " + i;
                sheet.Cells[3 + i, 3].Value = listed[i];
            }

            package.SaveAs(new FileInfo(tree.Path_(name)));
        }

        // A board workbook citing one file from each of the four sheets that can cite one.
        internal static void WriteBoard(
            TempWorkspace tree,
            string workbook,
            string schematic,
            string? componentImage = null,
            string? componentFile = null,
            string? boardFile = null)
        {
            var builder = new BoardWorkbookBuilder()
                .SheetWithLeadingRows(
                    BoardWorkbookSchema.SheetBoardSchematics,
                    ["# Revision date: 2026-09-25"],
                    BoardWorkbookBuilder.SchematicsHeaders,
                    ["uuid-1", "Sheet 1", "", schematic, "", "", "", "", ""])
                .Sheet(BoardWorkbookSchema.SheetComponents, BoardWorkbookBuilder.ComponentsHeaders,
                    ["uuid-2", "U1", "PLA", "906114-01", "", "IC", "", "PLA"]);

            if (componentImage is not null)
            {
                builder.Sheet(BoardWorkbookSchema.SheetComponentImages, BoardWorkbookBuilder.ComponentImagesHeaders,
                    ["uuid-3", "U1", "", "1", "Pin 1", "", componentImage, "", "", "", ""]);
            }

            if (componentFile is not null)
            {
                builder.Sheet(BoardWorkbookSchema.SheetComponentLocalFiles, BoardWorkbookBuilder.ComponentLocalFilesHeaders,
                    ["uuid-4", "U1", "Datasheet", componentFile]);
            }

            if (boardFile is not null)
            {
                builder.Sheet(BoardWorkbookSchema.SheetBoardLocalFiles, BoardWorkbookBuilder.BoardLocalFilesHeaders,
                    ["Service", "Manual", boardFile]);
            }

            builder.SaveTo(tree.Path_(workbook.Split('/')));
        }

        internal static void Touch(TempWorkspace tree, params string[] paths)
        {
            foreach (string path in paths)
                tree.WriteFile(path, "x");
        }

        // The common shape: one master, one board citing a schematic image, and a spare file.
        internal static TempWorkspace OneBoard()
        {
            var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook);
            DataTreeUsageTests.WriteBoard(tree, Workbook, $"{Board}/sheet1.png");
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png", $"{Board}/old.png");
            return tree;
        }

        // ---- the rule ------------------------------------------------------------------------

        [Fact]
        public void A_file_a_board_cites_is_used_and_one_nothing_cites_is_not()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsComplete, string.Join(" ", usage.Problems));
            Assert.True(usage.IsUsed($"{Board}/sheet1.png"));
            Assert.Equal([$"{Board}/old.png"], usage.UnusedFiles);
            Assert.Equal(1, usage.MasterCount);
            Assert.Equal(1, usage.BoardWorkbookCount);
        }

        [Fact]
        public void All_four_citing_sheets_count()
        {
            using var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook);
            DataTreeUsageTests.WriteBoard(tree, Workbook,
                $"{Board}/sheet1.png",
                componentImage: $"{Board}/Scope baseline/U1_1.png",
                componentFile: "Commodore/Shared files/Component local files/906114.pdf",
                boardFile: $"{Board}/Board local files/service.pdf");
            DataTreeUsageTests.Touch(tree,
                $"{Board}/sheet1.png",
                $"{Board}/Scope baseline/U1_1.png",
                "Commodore/Shared files/Component local files/906114.pdf",
                $"{Board}/Board local files/service.pdf");

            Assert.Empty(DataTreeUsage.Compute(tree.Root).UnusedFiles);
        }

        // Older generations serve older CRT builds and are never written, so what they cite stays
        // for as long as they exist - even when the newest workbook has moved on.
        [Fact]
        public void A_file_only_an_OLDER_generation_cites_is_still_used()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            const string legacy = "Commodore/C64/250407/Data C64 250407.xlsx";
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.xlsx", legacy);
            DataTreeUsageTests.WriteBoard(tree, legacy, $"{Board}/old.png");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsUsed($"{Board}/old.png"));
            Assert.Empty(usage.UnusedFiles);
            Assert.Equal(2, usage.MasterCount);
        }

        // A new system the server publishes is not added to any master (done by hand later). A
        // rule that trusted the masters alone would delete the whole new board.
        [Fact]
        public void A_board_in_the_tree_that_NO_master_lists_still_counts()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            const string unlisted = "Amstrad/CPC/464/Data CPC 464 v2.0.0.xlsx";
            DataTreeUsageTests.WriteBoard(tree, unlisted, "Amstrad/CPC/464/main.png");
            DataTreeUsageTests.Touch(tree, "Amstrad/CPC/464/main.png");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsComplete);
            Assert.True(usage.IsUsed(unlisted));
            Assert.True(usage.IsUsed("Amstrad/CPC/464/main.png"));
        }

        [Fact]
        public void The_sidecar_the_KiCad_folder_the_MiniPro_folder_and_bang_files_are_used()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            DataTreeUsageTests.Touch(tree,
                $"{Board}/Data C64 250407 v2.0.0.json",
                $"{Board}/KiCad data/board.kicad_pcb",
                $"{Board}/KiCad data/sub/part.kicad_sch",
                $"{DataTreeUsage.MiniProTestsFolder}/Catalogue/catalogue.json",
                $"{DataTreeUsage.MiniProTestsFolder}/Vectors/906114.xml.gz",
                "Commodore/Shared files/!README.txt");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.Equal([$"{Board}/old.png"], usage.UnusedFiles);
        }

        // Every Windows and macOS client finds "U8.PNG" when a row says "u8.png", so that file is
        // used - the safe direction for a rule that deletes.
        [Fact]
        public void A_citation_matches_its_file_whatever_the_case()
        {
            using var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook);
            DataTreeUsageTests.WriteBoard(tree, Workbook, $"{Board}/SHEET1.png");
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.PNG");

            Assert.Empty(DataTreeUsage.Compute(tree.Root).UnusedFiles);
        }

        // A citation written with backslashes (typed by hand in Excel on Windows) is the same path.
        [Fact]
        public void A_citation_with_backslashes_is_the_same_path()
        {
            using var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook.Replace('/', '\\'));
            DataTreeUsageTests.WriteBoard(tree, Workbook, @"Commodore\C64\250407\sheet1.png");
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsComplete, string.Join(" ", usage.Problems));
            Assert.Empty(usage.UnusedFiles);
        }

        [Fact]
        public void Hidden_and_half_written_files_are_never_considered()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            DataTreeUsageTests.Touch(tree, $"{Board}/.sheet1.png.tmp_1", $"{Board}/.hidden", $"{Board}/Thumbs.db");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.DoesNotContain(usage.Files, path => path.Contains(".tmp_", StringComparison.Ordinal) || path.EndsWith(".hidden", StringComparison.Ordinal));
            Assert.Equal([$"{Board}/old.png"], usage.UnusedFiles);
        }

        // ---- failing closed ------------------------------------------------------------------

        [Fact]
        public void A_master_listing_a_workbook_the_tree_lacks_makes_the_result_incomplete_and_removes_nothing()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook, "Amstrad/CPC/464/Data CPC 464 v2.0.0.xlsx");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.False(usage.IsComplete);
            Assert.Contains(usage.Problems, problem => problem.Contains("Data CPC 464 v2.0.0.xlsx", StringComparison.Ordinal));
            Assert.Empty(usage.UnusedFiles);
            Assert.Empty(usage.RemovableFrom([$"{Board}/old.png"]));
        }

        // An unreadable workbook cites an UNKNOWN set of files, not none. Treating it as none
        // would delete everything that board uses.
        [Fact]
        public void An_unreadable_board_workbook_fails_closed()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            tree.WriteFile("Amstrad/CPC/464/Data CPC 464 v2.0.0.xlsx", "this is not a workbook");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.False(usage.IsComplete);
            Assert.Empty(usage.UnusedFiles);
        }

        [Fact]
        public void A_master_without_its_sheet_fails_closed()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            using (var package = EpplusLicense.NewPackage())
            {
                package.Workbook.Worksheets.Add("Something else").Cells[1, 1].Value = "x";
                package.SaveAs(new FileInfo(tree.Path_("Classic-Repair-Toolbox.xlsx")));
            }

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.False(usage.IsComplete);
            Assert.Contains(usage.Problems, problem => problem.Contains("Classic-Repair-Toolbox.xlsx", StringComparison.Ordinal));
            Assert.Empty(usage.UnusedFiles);
        }

        [Fact]
        public void A_missing_tree_is_incomplete_rather_than_empty()
        {
            DataTreeUsageResult usage = DataTreeUsage.Compute(Path.Combine(Path.GetTempPath(), "crt-no-such-tree-" + Guid.NewGuid().ToString("N")));

            Assert.False(usage.IsComplete);
            Assert.Empty(usage.UnusedFiles);
        }

        // ---- previewing a publish ------------------------------------------------------------

        // The maintainer is shown what a publish would remove before anything is written, from the
        // SAME rule - with the workbook's new citations standing in for the file on disk.
        [Fact]
        public void A_preview_uses_what_the_workbook_WILL_cite_and_leaves_the_real_tree_alone()
        {
            using var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook);
            DataTreeUsageTests.WriteBoard(tree, Workbook, $"{Board}/sheet1.png", componentImage: $"{Board}/old.png");
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png", $"{Board}/old.png");

            var after = new Dictionary<string, IReadOnlyCollection<string>> { [Workbook] = [$"{Board}/sheet1.png"] };

            Assert.Equal([$"{Board}/old.png"], DataTreeUsage.Compute(tree.Root, after).RemovableFrom([$"{Board}/old.png"]));
            Assert.Empty(DataTreeUsage.Compute(tree.Root).UnusedFiles);
        }

        [Fact]
        public void A_preview_can_add_a_workbook_the_publish_has_not_written_yet()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();

            var after = new Dictionary<string, IReadOnlyCollection<string>>
            {
                ["Amstrad/CPC/464/Data CPC 464 v2.0.0.xlsx"] = [$"{Board}/old.png"]
            };

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root, after);

            Assert.True(usage.IsComplete, string.Join(" ", usage.Problems));
            Assert.True(usage.IsUsed($"{Board}/old.png"));
        }

        // A shared file one board stops citing stays while any other board - of any generation,
        // in any folder - still cites it.
        [Fact]
        public void A_shared_file_another_board_still_cites_is_not_removable()
        {
            using var tree = new TempWorkspace();
            const string shared = "Generic shared files/Component images/7408.jpg";
            const string other = "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx";
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook, other);
            DataTreeUsageTests.WriteBoard(tree, Workbook, $"{Board}/sheet1.png", componentImage: shared);
            DataTreeUsageTests.WriteBoard(tree, other, "Commodore/C128/310378/main.png", componentImage: shared);
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png", "Commodore/C128/310378/main.png", shared);

            var after = new Dictionary<string, IReadOnlyCollection<string>> { [Workbook] = [$"{Board}/sheet1.png"] };

            Assert.Empty(DataTreeUsage.Compute(tree.Root, after).RemovableFrom([shared]));
        }

        [Fact]
        public void RemovableFrom_answers_in_the_trees_own_spelling_and_skips_files_it_does_not_have()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.Equal([$"{Board}/old.png"], usage.RemovableFrom([$"{Board}/OLD.png", $"{Board}/never-existed.png", $"{Board}/sheet1.png"]));
        }

        // ---- the helpers ---------------------------------------------------------------------

        [Fact]
        public void NoLongerCited_is_what_the_board_dropped_ignoring_case()
        {
            IReadOnlyList<string> dropped = DataTreeUsage.NoLongerCited(
                ["a/one.png", "a/two.png", "a/three.png", "a/two.png"],
                ["A/ONE.png", @"a\three.png"]);

            Assert.Equal(["a/two.png"], dropped);
            Assert.Empty(DataTreeUsage.NoLongerCited(null, ["a.png"]));
        }

        [Theory]
        [InlineData("Classic-Repair-Toolbox.xlsx", true)]
        [InlineData("Classic-Repair-Toolbox.v2.0.0.xlsx", true)]
        [InlineData("classic-repair-toolbox.V2.0.0.XLSX", true)]
        [InlineData("Classic-Repair-Toolbox backup.xlsx", false)]
        [InlineData("Classic-Repair-Toolbox.v2.0.0.xls", false)]
        [InlineData("Data C64 250407.xlsx", false)]
        [InlineData("", false)]
        public void Master_workbooks_are_recognised_by_name(string name, bool expected)
        {
            Assert.Equal(expected, DataTreeUsage.IsMasterFileName(name));
        }

        [Theory]
        [InlineData("Commodore/C64/250407/Data C64 250407.xlsx", true)]
        [InlineData("Commodore/C64/250407/~$Data C64 250407.xlsx", false)]
        [InlineData("Commodore/Shared files/Board local files/list.xlsx", false)]
        [InlineData("Generic shared files/MiniPro/IC tests/x.xlsx", false)]
        [InlineData("Commodore/C64/250407/Sub/Data.xlsx", false)]
        [InlineData("Commodore/C64/250407/sheet1.png", false)]
        public void Board_workbooks_are_the_workbooks_at_the_top_of_a_board_folder(string path, bool expected)
        {
            Assert.Equal(expected, DataTreeUsage.IsBoardFolderWorkbook(path));
        }

        // ###########################################################################################
        // *** A MASTER NAMING A BOARD WORKBOOK IN ANOTHER CASE (owner request, 2026-10-04: "prove that
        // the files listed in here really can be deleted"). *** Every Windows and macOS CRT opens such
        // a workbook and shows what it cites. The rule opened it by the MASTER's spelling - not there
        // on the Linux server, and a missing workbook cites nothing - so everything only it cites was
        // listed as unused. It is read at its spelling on disk now.
        //
        // Windows finds the file by either spelling, so on this machine the whole-tree test passes
        // with or without the fix; it is the Linux run (GitHub's) that fails without it. The spelling
        // lookup below is what holds the fix on every OS.
        // ###########################################################################################
        [Fact]
        public void A_board_workbook_a_master_names_in_another_case_is_read_and_what_it_cites_is_used()
        {
            using var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", "commodore/c64/250407/DATA c64 250407 V2.0.0.xlsx");
            DataTreeUsageTests.WriteBoard(tree, Workbook, $"{Board}/sheet1.png");
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png", $"{Board}/old.png");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsComplete, string.Join(" ", usage.Problems));
            Assert.Equal(1, usage.BoardWorkbookCount);
            Assert.True(usage.IsUsed($"{Board}/sheet1.png"));
            Assert.Equal([$"{Board}/old.png"], usage.UnusedFiles);
        }

        [Fact]
        public void A_board_workbook_is_read_at_every_spelling_the_tree_holds_and_never_at_another()
        {
            ILookup<string, string> onDisk = new[]
            {
                Workbook,
                $"{Board}/sheet1.png",
                // Two spellings of one name - only a case-sensitive file system holds both.
                "Commodore/C128/310378/Data.xlsx",
                "Commodore/C128/310378/data.xlsx"
            }.ToLookup(path => path, StringComparer.OrdinalIgnoreCase);

            Assert.Equal([Workbook], DataTreeUsage.SpellingsOnDisk("commodore/C64/250407/data C64 250407 v2.0.0.XLSX", onDisk));
            Assert.Equal(
                ["Commodore/C128/310378/Data.xlsx", "Commodore/C128/310378/data.xlsx"],
                DataTreeUsage.SpellingsOnDisk("Commodore/C128/310378/DATA.xlsx", onDisk));
            Assert.Empty(DataTreeUsage.SpellingsOnDisk("Commodore/C64/250407/Other.xlsx", onDisk));
        }

        // ###########################################################################################
        // The legacy master's "KiCad data file" column, which CRT 1.x reads for a board's traces: a file
        // named there is used wherever it is (today every one is inside a "KiCad data" folder anyway).
        // ###########################################################################################
        [Fact]
        public void A_file_a_legacy_masters_KiCad_data_file_column_names_is_used()
        {
            using var tree = new TempWorkspace();

            using (var package = EpplusLicense.NewPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add(DataTreeUsage.MasterSheetName);
                sheet.Cells[1, 1].Value = "Hardware name";
                sheet.Cells[1, 2].Value = DataTreeUsage.ExcelDataFileColumn;
                sheet.Cells[1, 3].Value = DataTreeUsage.LegacyKiCadDataFileColumn;
                sheet.Cells[2, 1].Value = "Commodore 64";
                sheet.Cells[2, 2].Value = Workbook;
                sheet.Cells[2, 3].Value = $"{Board}/traces/KiCad-traces.json";
                package.SaveAs(new FileInfo(tree.Path_("Classic-Repair-Toolbox.xlsx")));
            }

            DataTreeUsageTests.WriteBoard(tree, Workbook, $"{Board}/sheet1.png");
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png", $"{Board}/traces/KiCad-traces.json", $"{Board}/traces/other.json");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsComplete, string.Join(" ", usage.Problems));
            Assert.True(usage.IsUsed($"{Board}/traces/KiCad-traces.json"));
            Assert.Equal([$"{Board}/traces/other.json"], usage.UnusedFiles);

            // ###########################################################################################
            // And through the cache, read and then remembered (code review, 2026-10-04): the named
            // files used to ride inside the master's listing and be recovered by a downcast, so a cache
            // that ever copied the listing would have dropped them without a word - the fail-open
            // direction. The second Compute is served from the cache, and still keeps the file.
            // ###########################################################################################
            var cache = new WorkbookReadCache();

            DataTreeUsageResult read = DataTreeUsage.Compute(tree.Root, cache: cache);
            int reads = cache.Reads;
            DataTreeUsageResult remembered = DataTreeUsage.Compute(tree.Root, cache: cache);

            Assert.Equal(reads, cache.Reads);
            Assert.True(read.IsUsed($"{Board}/traces/KiCad-traces.json"));
            Assert.True(remembered.IsUsed($"{Board}/traces/KiCad-traces.json"));
            Assert.Equal([$"{Board}/traces/other.json"], remembered.UnusedFiles);
        }

        // ###########################################################################################
        // A hand-edited cell spelling a path oddly - "./", "//", "x/..", a folder name with a trailing
        // dot - is found by every client's operating system, so the file it reaches is used.
        // ###########################################################################################
        [Theory]
        [InlineData("Commodore/C64/250407/./sheet1.png")]
        [InlineData("Commodore/C64/250407//sheet1.png")]
        [InlineData("Commodore/C64/250407/Sub/../sheet1.png")]
        [InlineData("Commodore/C64/250407./sheet1.png")]
        public void A_path_written_oddly_counts_the_file_it_reaches_as_used(string cited)
        {
            using var tree = new TempWorkspace();
            DataTreeUsageTests.WriteMaster(tree, "Classic-Repair-Toolbox.v2.0.0.xlsx", Workbook);
            DataTreeUsageTests.WriteBoard(tree, Workbook, cited);
            DataTreeUsageTests.Touch(tree, $"{Board}/sheet1.png");

            DataTreeUsageResult usage = DataTreeUsage.Compute(tree.Root);

            Assert.True(usage.IsComplete, string.Join(" ", usage.Problems));
            Assert.True(usage.IsUsed($"{Board}/sheet1.png"), $"[{cited}] does not keep the file it reaches.");
            Assert.Empty(usage.UnusedFiles);
        }

        [Theory]
        [InlineData("a/b.png", "a/b.png")]
        [InlineData("a/./b.png", "a/b.png")]
        [InlineData("a//b.png", "a/b.png")]
        [InlineData("a/x/../b.png", "a/b.png")]
        [InlineData("a./b.png", "a/b.png")]
        [InlineData("a /b.png", "a/b.png")]
        [InlineData("../b.png", null)]
        [InlineData("a/../../b.png", null)]
        public void A_path_resolves_as_the_operating_system_finds_it(string path, string? expected)
        {
            Assert.Equal(expected, DataTreeUsage.Resolved(path));
        }
    }
}
