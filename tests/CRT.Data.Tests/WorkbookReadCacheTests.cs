using System;
using System.IO;
using ClassicRepairToolbox.Tests;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // WorkbookReadCache - what a workbook cites, read once per version of the file (code review,
    // 2026-09-25). The maintainer application's submission detail computes a removal preview on every
    // click, and that used to parse every workbook in the BETA tree each time.
    //
    // What matters: an unchanged workbook is not read again, a CHANGED one always is (a stale
    // answer would show the maintainer the wrong list), and a failed read is never remembered. Real
    // workbooks in a real temp tree, built by DataTreeUsageTests' own builders.
    // ###########################################################################################
    [Collection("BoardData")]
    public sealed class WorkbookReadCacheTests
    {
        private static readonly DateTime First = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        private static readonly DateTime Second = new(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc);

        private static string WorkbookPath(TempWorkspace tree) =>
            tree.Path_(DataTreeUsageTests.Workbook.Split('/'));

        [Fact]
        public void An_unchanged_tree_is_read_once_and_answers_as_a_fresh_read_does()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            var cache = new WorkbookReadCache();

            DataTreeUsageResult first = DataTreeUsage.Compute(tree.Root, cache: cache);
            int readsAfterFirst = cache.Reads;

            DataTreeUsageResult second = DataTreeUsage.Compute(tree.Root, cache: cache);
            DataTreeUsageResult fresh = DataTreeUsage.Compute(tree.Root);

            // One master and one board workbook, each opened once.
            Assert.Equal(2, readsAfterFirst);
            Assert.Equal(readsAfterFirst, cache.Reads);

            Assert.True(second.IsComplete, string.Join(" ", second.Problems));
            Assert.Equal(fresh.UnusedFiles, second.UnusedFiles);
            Assert.Equal(first.UnusedFiles, second.UnusedFiles);
        }

        // A publish rewrites a workbook; the next preview must see what it cites NOW.
        [Fact]
        public void A_rewritten_workbook_is_read_again()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            string board = DataTreeUsageTests.Board;
            File.SetLastWriteTimeUtc(WorkbookReadCacheTests.WorkbookPath(tree), WorkbookReadCacheTests.First);

            var cache = new WorkbookReadCache();

            Assert.Equal([$"{board}/old.png"], DataTreeUsage.Compute(tree.Root, cache: cache).UnusedFiles);

            // The board now cites old.png too.
            DataTreeUsageTests.WriteBoard(tree, DataTreeUsageTests.Workbook, $"{board}/sheet1.png", boardFile: $"{board}/old.png");
            File.SetLastWriteTimeUtc(WorkbookReadCacheTests.WorkbookPath(tree), WorkbookReadCacheTests.Second);

            DataTreeUsageResult after = DataTreeUsage.Compute(tree.Root, cache: cache);

            Assert.Empty(after.UnusedFiles);
            Assert.Equal(3, cache.Reads);
        }

        // An unreadable workbook is tried again every time: remembering the failure would keep the
        // preview blocked after the file was fixed.
        [Fact]
        public void A_workbook_that_could_not_be_read_is_not_remembered()
        {
            using TempWorkspace tree = DataTreeUsageTests.OneBoard();
            File.WriteAllText(WorkbookReadCacheTests.WorkbookPath(tree), "this is not a workbook");

            var cache = new WorkbookReadCache();

            Assert.False(DataTreeUsage.Compute(tree.Root, cache: cache).IsComplete);
            int readsAfterFirst = cache.Reads;

            Assert.False(DataTreeUsage.Compute(tree.Root, cache: cache).IsComplete);

            // The master is remembered; the broken board workbook is opened again.
            Assert.Equal(readsAfterFirst + 1, cache.Reads);
        }
    }
}
