using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // THE SHIPPED DATA HAS NO ORPHAN FILES (owner decision, 2026-09-25: "there must be no
    // orphan files").
    //
    // Runs DataTreeUsage over the real Assets/Data. The first run found 50 files nothing used
    // (image-editor .fsc files, annotation projects, component images and PDFs no board cites,
    // uncited scope captures and readme texts), and the project owner approved removing all of them.
    // A file added to the tree without a row citing it now fails here, naming the file - before it
    // is uploaded to every user.
    //
    // It also proves the rule reads every shipped workbook, masters and both generations
    // included: an incomplete result would make the live trees remove nothing at all.
    // ###########################################################################################
    [Collection("BoardData")]
    public sealed class DataTreeUsageShippedDataTests
    {
        private static string DataRoot()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "Assets", "Data");
        }

        [Fact]
        public void Every_shipped_file_is_used()
        {
            var clock = Stopwatch.StartNew();
            DataTreeUsageResult usage = DataTreeUsage.Compute(DataTreeUsageShippedDataTests.DataRoot());
            clock.Stop();

            Assert.True(usage.IsComplete, "The shipped data could not be read completely: " + string.Join(" | ", usage.Problems));

            IReadOnlyList<string> unused = usage.UnusedFiles;

            Assert.True(
                unused.Count == 0,
                $"{unused.Count} shipped file(s) are used by nothing (checked in {clock.Elapsed.TotalSeconds:F1} s) - cite each from a " +
                "board, rename it to start with '!' if it is documentation, or remove it:" + Environment.NewLine +
                string.Join(Environment.NewLine, unused));
        }

        // Both generations are read - the older one serves older CRT builds, and what only it cites
        // must stay.
        [Fact]
        public void Both_master_generations_and_every_board_workbook_are_read()
        {
            DataTreeUsageResult usage = DataTreeUsage.Compute(DataTreeUsageShippedDataTests.DataRoot());

            Assert.True(usage.MasterCount >= 2, $"Expected the legacy and the v2.0.0 master, found {usage.MasterCount}.");
            Assert.True(usage.BoardWorkbookCount >= 20, $"Only {usage.BoardWorkbookCount} board workbooks were read.");
        }

        // No workbook cites the MiniPro tests; CRT reads them by folder. A sweep that forgot them
        // would delete real data.
        [Fact]
        public void The_MiniPro_tests_are_used_although_no_workbook_cites_them()
        {
            DataTreeUsageResult usage = DataTreeUsage.Compute(DataTreeUsageShippedDataTests.DataRoot());

            List<string> miniPro = usage.Files
                .Where(path => path.StartsWith(DataTreeUsage.MiniProTestsFolder + "/", StringComparison.Ordinal))
                .ToList();

            Assert.NotEmpty(miniPro);
            Assert.All(miniPro, path => Assert.True(usage.IsUsed(path), path));
        }
    }
}
