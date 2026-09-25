using System;
using Handlers.DataHandling;
using Xunit;

// ###############################################################################################
// Covers DataGenerationRules - which workbook GENERATION the tree is on and what a file of that
// generation is called.
//
// The rule these tests exist to protect is the project owner's: publishing writes ONLY the newest
// generation and never touches an older one, because older generations still serve older
// application builds. Writing into a frozen generation fails silently - the write succeeds and a
// compatibility target is quietly edited - so the ordering and naming rules below are the only
// thing standing between that decision and a very quiet bug.
//
// Both real naming conventions are pinned here on purpose: the master workbook separates its
// version with a DOT and a board workbook with a SPACE. Both are shipped, so a "tidy-up" that
// unified them would break whichever tree it did not match.
// ###############################################################################################
namespace CRT.Data.Tests
{
    public class DataGenerationRulesTests
    {
        // -----------------------------------------------------------------------------------
        // Reading a generation out of a file name
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("Classic-Repair-Toolbox.v2.0.0.xlsx", "2.0.0")]
        [InlineData("Data C64 250407 v2.0.0.xlsx", "2.0.0")]
        [InlineData("Data CPC 664 MC0005A v2.0.0.xlsx", "2.0.0")]
        [InlineData("Data C64 250407 v3.1.4.xlsx", "3.1.4")]
        public void Both_shipped_naming_conventions_yield_their_generation(string fileName, string expected)
        {
            Assert.Equal(Version.Parse(expected), DataGenerationRules.TryReadGeneration(fileName));
        }

        [Theory]
        [InlineData("Classic-Repair-Toolbox.xlsx")]
        [InlineData("Data C64 250407.xlsx")]
        public void The_unversioned_original_reports_no_generation(string fileName)
        {
            // Null rather than a fabricated 0.0.0: the caller must not be able to order the
            // original against real versions by accident, or print a version that never existed.
            Assert.Null(DataGenerationRules.TryReadGeneration(fileName));
        }

        [Fact]
        public void A_board_name_containing_a_v_is_not_mistaken_for_a_version()
        {
            // "Data VIC20 250403" - the V of VIC20 is a letter in the board's name, not a version
            // marker. Without the separator check this would try to parse "IC20 250403".
            Assert.Null(DataGenerationRules.TryReadGeneration("Data VIC20 250403.xlsx"));
        }

        [Fact]
        public void A_real_version_still_reads_on_a_board_whose_name_contains_a_v()
        {
            // The same board WITH a generation must still resolve - the guard above must not have
            // cost us the ordinary case.
            Assert.Equal(
                Version.Parse("2.0.0"),
                DataGenerationRules.TryReadGeneration("Data VIC20 250403 v2.0.0.xlsx"));
        }

        [Theory]
        [InlineData("Data C64 250407 vNotAVersion.xlsx")]
        [InlineData("Data C64 250407 v.xlsx")]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void An_unparseable_name_is_ignored_rather_than_throwing(string? fileName)
        {
            // The tree is synced, so a stray or half-written file can appear in it. None of them
            // may stop a publish.
            Assert.Null(DataGenerationRules.TryReadGeneration(fileName));
        }

        [Fact]
        public void A_name_that_is_only_a_version_marker_is_refused()
        {
            // marker < 1: there is no name before the "v", so this is not a board file at all.
            Assert.Null(DataGenerationRules.TryReadGeneration("v2.0.0.xlsx"));
        }

        // -----------------------------------------------------------------------------------
        // Resolving the newest generation present
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_newest_generation_wins_regardless_of_the_order_files_arrive_in()
        {
            // A directory listing returns whatever order the filesystem feels like, so the answer
            // must not depend on it.
            string[] files =
            [
                "Data C64 250407 v2.0.0.xlsx",
                "Data C64 250407 v3.0.0.xlsx",
                "Data C64 250407 v2.6.0.xlsx"
            ];

            Assert.Equal(Version.Parse("3.0.0"), DataGenerationRules.ResolveNewestGeneration(files));
            Assert.Equal(Version.Parse("3.0.0"), DataGenerationRules.ResolveNewestGeneration([.. files.Reverse()]));
        }

        [Fact]
        public void Versions_are_compared_numerically_not_as_text()
        {
            // "10.0.0" sorts BEFORE "9.0.0" as a string. A publish picking the text-newest would
            // write into a frozen generation.
            string[] files = ["Data C64 250407 v9.0.0.xlsx", "Data C64 250407 v10.0.0.xlsx"];

            Assert.Equal(Version.Parse("10.0.0"), DataGenerationRules.ResolveNewestGeneration(files));
        }

        [Fact]
        public void A_tree_holding_only_the_unversioned_original_reports_no_generation()
        {
            Assert.Null(DataGenerationRules.ResolveNewestGeneration(["Classic-Repair-Toolbox.xlsx"]));
        }

        [Fact]
        public void A_stray_file_does_not_displace_a_real_generation()
        {
            string[] files =
            [
                "Data C64 250407 v2.0.0.xlsx",
                "notes.txt",
                "Data C64 250407 vDRAFT.xlsx",
                "~$Data C64 250407 v2.0.0.xlsx"
            ];

            Assert.Equal(Version.Parse("2.0.0"), DataGenerationRules.ResolveNewestGeneration(files));
        }

        [Fact]
        public void An_empty_or_null_listing_reports_no_generation()
        {
            Assert.Null(DataGenerationRules.ResolveNewestGeneration([]));
            Assert.Null(DataGenerationRules.ResolveNewestGeneration(null));
        }

        // -----------------------------------------------------------------------------------
        // Building names, and round-tripping them
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_master_workbook_is_built_with_a_DOT_separator()
        {
            Assert.Equal(
                "Classic-Repair-Toolbox.v2.0.0.xlsx",
                DataGenerationRules.BuildMasterFileName(Version.Parse("2.0.0")));
        }

        [Fact]
        public void A_board_workbook_is_built_with_a_SPACE_separator()
        {
            Assert.Equal(
                "Data C64 250407 v2.0.0.xlsx",
                DataGenerationRules.BuildBoardFileName("Data C64 250407", Version.Parse("2.0.0")));
        }

        [Fact]
        public void A_null_generation_builds_the_unversioned_original()
        {
            Assert.Equal("Classic-Repair-Toolbox.xlsx", DataGenerationRules.BuildMasterFileName(null));
            Assert.Equal("Data C64 250407.xlsx", DataGenerationRules.BuildBoardFileName("Data C64 250407", null));
        }

        [Theory]
        [InlineData("Data C64 250407 v2.0.0.xlsx", "Data C64 250407")]
        [InlineData("Data C64 250407.xlsx", "Data C64 250407")]
        [InlineData("Data CPC 664 MC0005A v2.0.0.xlsx", "Data CPC 664 MC0005A")]
        [InlineData("Data VIC20 250403.xlsx", "Data VIC20 250403")]
        public void A_board_stem_is_read_back_without_its_generation(string fileName, string expected)
        {
            Assert.Equal(expected, DataGenerationRules.ReadBoardStem(fileName));
        }

        [Fact]
        public void A_board_file_survives_a_stem_and_rebuild_round_trip()
        {
            // This is the actual publish operation: take the file of the current generation, and
            // name the file of the target one. It must be exactly the shipped spelling.
            const string shipped = "Data C64 250407 v2.0.0.xlsx";

            string stem = DataGenerationRules.ReadBoardStem(shipped);
            string rebuilt = DataGenerationRules.BuildBoardFileName(stem, Version.Parse("2.0.0"));

            Assert.Equal(shipped, rebuilt);
        }

        [Fact]
        public void A_blank_stem_builds_nothing_rather_than_a_bare_extension()
        {
            // ".xlsx" is a hidden file on Linux and a nameless one everywhere. Returning empty
            // lets the caller refuse, rather than creating it.
            Assert.Equal(string.Empty, DataGenerationRules.BuildBoardFileName("", Version.Parse("2.0.0")));
            Assert.Equal(string.Empty, DataGenerationRules.BuildBoardFileName("   ", null));
        }

        // -----------------------------------------------------------------------------------
        // The guard that protects a frozen generation
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_older_generation_is_recognised_as_older()
        {
            Assert.True(DataGenerationRules.IsOlderGeneration(Version.Parse("2.0.0"), Version.Parse("3.0.0")));
        }

        [Fact]
        public void The_target_generation_itself_is_not_older_than_itself()
        {
            // Equal must be writable - it is the ordinary publish.
            Assert.False(DataGenerationRules.IsOlderGeneration(Version.Parse("3.0.0"), Version.Parse("3.0.0")));
        }

        [Fact]
        public void A_newer_generation_than_the_target_is_not_older()
        {
            // Not "writable" - just not the case this guard answers. Discovery makes it
            // unreachable in practice, since the target IS the newest found.
            Assert.False(DataGenerationRules.IsOlderGeneration(Version.Parse("4.0.0"), Version.Parse("3.0.0")));
        }

        [Fact]
        public void The_unversioned_original_is_older_than_every_real_generation()
        {
            // This is the case the project owner named outright: "the no-version one, which is the
            // first file, which still works for an older application version". Publishing must
            // never write it.
            Assert.True(DataGenerationRules.IsOlderGeneration(null, Version.Parse("2.0.0")));
        }

        [Fact]
        public void The_unversioned_original_is_not_older_than_itself()
        {
            // A tree that has never been versioned at all: the original is the target, and
            // publishing into it is correct.
            Assert.False(DataGenerationRules.IsOlderGeneration(null, null));
        }

        [Fact]
        public void A_real_generation_is_never_older_than_the_unversioned_original()
        {
            Assert.False(DataGenerationRules.IsOlderGeneration(Version.Parse("2.0.0"), null));
        }
    }
}
