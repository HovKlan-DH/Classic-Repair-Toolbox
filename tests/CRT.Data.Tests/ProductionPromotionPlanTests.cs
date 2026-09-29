using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers ProductionPromotionPlan - what publishing one system from BETA to Production copies,
    // and what it refuses (owner request, 2026-09-25).
    //
    // Production is what every user downloads, so the refusals matter as much as the copies: a
    // board promoted while pointing at a file Production does not have is a broken board on every
    // machine, and a shared file changed by a maintainer reaches boards nobody asked them about.
    // ###########################################################################################
    public sealed class ProductionPromotionPlanTests
    {
        private const string Folder = "Commodore/C64/250407";

        private static PublishedTreeView Tree(params (string Path, string Sha256)[] files)
        {
            var hashes = files.ToDictionary(file => file.Path, file => file.Sha256, StringComparer.Ordinal);

            return new PublishedTreeView(
                path => hashes.TryGetValue(path, out string? sha) ? sha : null,
                folder =>
                {
                    string prefix = folder.Length == 0 ? string.Empty : folder + "/";

                    var names = hashes.Keys
                        .Where(path => path.StartsWith(prefix, StringComparison.Ordinal))
                        .Select(path => path[prefix.Length..].Split('/')[0])
                        .Distinct(StringComparer.Ordinal)
                        .ToList();

                    return names.Count == 0 ? null : names;
                });
        }

        private static ProductionPromotionResult Plan(
            IEnumerable<string> own,
            IEnumerable<string> cited,
            PublishedTreeView beta,
            PublishedTreeView production) =>
            ProductionPromotionPlan.Build("Commodore", "C64", "250407", own, cited, beta, production);

        private static string P(string rest) => $"{ProductionPromotionPlanTests.Folder}/{rest}";

        // -----------------------------------------------------------------------------------
        // What is copied
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Only_files_that_DIFFER_from_production_are_copied()
        {
            // A typo fix changes the workbook and system.json; the 1,200 images are identical and
            // must not be rewritten on every user's machine for nothing.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("Data C64 250407 v2.0.0.xlsx"), "new-workbook"),
                (P("Data C64 250407 v2.0.0.json"), "new-sidecar"),
                (P("main.png"), "same"));

            PublishedTreeView production = ProductionPromotionPlanTests.Tree(
                (P("Data C64 250407 v2.0.0.xlsx"), "old-workbook"),
                (P("Data C64 250407 v2.0.0.json"), "old-sidecar"),
                (P("main.png"), "same"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("Data C64 250407 v2.0.0.xlsx"), P("Data C64 250407 v2.0.0.json"), P("main.png")], [P("main.png")], beta, production);

            Assert.True(result.CanPromote);
            Assert.Equal([P("Data C64 250407 v2.0.0.json"), P("Data C64 250407 v2.0.0.xlsx")], result.Files.Select(file => file.Path));
            Assert.All(result.Files, file => Assert.Equal(PromotionChange.Replaced, file.Change));
            Assert.Equal(1, result.UnchangedCount);
        }

        // ###########################################################################################
        // *** WHAT IS ALREADY THE SAME IS NAMED, NOT ONLY COUNTED (owner request, 2026-09-28). *** The
        // maintainer's file tree draws production after the publish from the copies, the removals
        // and these - own files and shared ones alike - so the count and the list must agree.
        // ###########################################################################################
        [Fact]
        public void The_files_already_the_same_are_listed_as_well_as_counted()
        {
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("Data C64 250407 v2.0.0.xlsx"), "new"),
                (P("Images/b.png"), "same-b"),
                (P("Images/a.png"), "same-a"),
                ("Commodore/Shared files/manual.pdf", "same-manual"));

            PublishedTreeView production = ProductionPromotionPlanTests.Tree(
                (P("Data C64 250407 v2.0.0.xlsx"), "old"),
                (P("Images/b.png"), "same-b"),
                (P("Images/a.png"), "same-a"),
                ("Commodore/Shared files/manual.pdf", "same-manual"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("Data C64 250407 v2.0.0.xlsx"), P("Images/b.png"), P("Images/a.png")],
                [P("Images/a.png"), "Commodore/Shared files/manual.pdf"],
                beta,
                production);

            Assert.Equal(
                ["Commodore/C64/250407/Images/a.png", "Commodore/C64/250407/Images/b.png", "Commodore/Shared files/manual.pdf"],
                result.Unchanged);
            Assert.Equal(result.Unchanged!.Count, result.UnchangedCount);
        }

        [Fact]
        public void A_file_production_does_not_have_is_ADDED()
        {
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree((P("new.png"), "n"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("new.png")], [], beta, ProductionPromotionPlanTests.Tree());

            PromotionFile file = Assert.Single(result.Files);
            Assert.Equal(PromotionChange.Added, file.Change);
            Assert.Equal("n", file.Sha256);
        }

        [Fact]
        public void The_copies_are_in_WRITE_ORDER_content_then_the_board()
        {
            // The publish's order and for its reason: an interrupted promotion must leave a board
            // that loads - a workbook never points at a file that has not arrived yet.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("Data C64 250407 v2.0.0.xlsx"), "w"),
                (P("Data C64 250407 v2.0.0.json"), "s"),
                (P("Images/a.png"), "a"),
                (P("z.pdf"), "z"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("Data C64 250407 v2.0.0.xlsx"), P("Data C64 250407 v2.0.0.json"), P("Images/a.png"), P("z.pdf")],
                [],
                beta,
                ProductionPromotionPlanTests.Tree());

            Assert.Equal(
                [P("Images/a.png"), P("z.pdf"), P("Data C64 250407 v2.0.0.json"), P("Data C64 250407 v2.0.0.xlsx")],
                result.Files.Select(file => file.Path));

            Assert.Equal(PromotionStage.Board, result.Files[^1].Stage);
        }

        [Fact]
        public void A_system_json_left_in_BETA_by_an_earlier_build_is_NEVER_promoted()
        {
            // Retired by the project owner (2026-09-25): it is not relevant to users and must not be
            // downloaded by them. One still in BETA is left behind, not carried to production.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("system.json"), "d"),
                (P("main.png"), "m"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("system.json"), P("main.png")], [], beta, ProductionPromotionPlanTests.Tree());

            Assert.Equal([P("main.png")], result.Files.Select(file => file.Path));
        }

        [Fact]
        public void An_xlsx_BELOW_the_folders_top_level_is_content_not_the_board()
        {
            Assert.Equal(PromotionStage.Content, ProductionPromotionPlan.StageOf(ProductionPromotionPlanTests.Folder, P("KiCad data/x.json")));
            Assert.Equal(PromotionStage.Board, ProductionPromotionPlan.StageOf(ProductionPromotionPlanTests.Folder, P("Data C64 250407.xlsx")));
        }

        [Fact]
        public void Hidden_files_are_never_promoted()
        {
            // A publish writes ".<name>.crt-publish-<guid>.tmp" beside its target; one left by a
            // crash must not travel to every user.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P(".main.png.crt-publish-abc.tmp"), "t"),
                (P("main.png"), "m"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P(".main.png.crt-publish-abc.tmp"), P("main.png")], [], beta, ProductionPromotionPlanTests.Tree());

            Assert.Equal([P("main.png")], result.Files.Select(file => file.Path));
        }

        [Fact]
        public void Nothing_to_copy_is_still_a_promotion_that_records_the_revision()
        {
            // The case where BETA and Production were brought level by hand before this existed:
            // promoting copies nothing and marks the system as in production.
            PublishedTreeView tree = ProductionPromotionPlanTests.Tree((P("main.png"), "m"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan([P("main.png")], [], tree, tree);

            Assert.True(result.CanPromote);
            Assert.Empty(result.Files);
            Assert.Equal(1, result.UnchangedCount);
        }

        // -----------------------------------------------------------------------------------
        // Shared files
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_CHANGED_shared_file_is_copied_and_marks_the_promotion_as_touching_shared_files()
        {
            // It reaches every board that cites it, which is what makes it the administrator's.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/Shared files/74LS08.png", "new"));

            PublishedTreeView production = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/Shared files/74LS08.png", "old"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], ["Commodore/Shared files/74LS08.png"], beta, production);

            Assert.True(result.TouchesSharedFiles);
            PromotionFile shared = Assert.Single(result.Files);
            Assert.True(shared.IsShared);
            Assert.Equal("Commodore/Shared files/74LS08.png", shared.Path);
        }

        // A NEW shared file is copied like any other, but reaches no board that did not ask for it -
        // so it does not make the promotion the administrator's too (owner decision, 2026-09-27).
        [Fact]
        public void A_NEW_shared_file_is_copied_but_does_not_need_the_administrator()
        {
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/Shared files/74LS08.png", "new"));

            PublishedTreeView production = ProductionPromotionPlanTests.Tree((P("main.png"), "m"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], ["Commodore/Shared files/74LS08.png"], beta, production);

            Assert.False(result.TouchesSharedFiles);
            PromotionFile shared = Assert.Single(result.Files);
            Assert.True(shared.IsShared);
            Assert.Equal(PromotionChange.Added, shared.Change);
        }

        [Fact]
        public void A_shared_file_used_UNCHANGED_does_not_touch_shared_files()
        {
            PublishedTreeView tree = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Generic shared files/7400.png", "same"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], ["Generic shared files/7400.png"], tree, tree);

            Assert.False(result.TouchesSharedFiles);
            Assert.Empty(result.Files);
        }

        [Fact]
        public void Shared_files_the_board_does_NOT_cite_are_left_alone()
        {
            // Another board's shared image changed in BETA is that board's promotion, not this one.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/Shared files/other.png", "new"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], [], beta, ProductionPromotionPlanTests.Tree((P("main.png"), "m")));

            Assert.Empty(result.Files);
            Assert.False(result.TouchesSharedFiles);
        }

        // -----------------------------------------------------------------------------------
        // Refusals
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Another_boards_file_NOT_yet_in_production_refuses_the_whole_promotion()
        {
            // C128DCR 250477 cites texts in the C128 310378 folder. If those changed in BETA and
            // were never promoted, this board would go out pointing at the old ones.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/C128/310378/scope.txt", "new"));

            PublishedTreeView production = ProductionPromotionPlanTests.Tree(
                ("Commodore/C128/310378/scope.txt", "old"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], ["Commodore/C128/310378/scope.txt"], beta, production);

            Assert.False(result.CanPromote);
            Assert.Empty(result.Files);
            Assert.Equal("promote.foreign_not_in_production", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void Another_boards_file_ALREADY_identical_in_production_is_fine_and_never_copied()
        {
            PublishedTreeView tree = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/C128/310378/scope.txt", "same"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], ["Commodore/C128/310378/scope.txt"], tree, tree);

            Assert.True(result.CanPromote);
            Assert.Empty(result.Files);
        }

        [Fact]
        public void A_cited_file_BETA_itself_does_not_have_refuses_the_promotion()
        {
            // The BETA board is broken; a broken board does not go to everyone.
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree((P("main.png"), "m"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], [P("gone.png"), "Commodore/Shared files/gone.png"], beta, ProductionPromotionPlanTests.Tree());

            Assert.False(result.CanPromote);
            Assert.Equal(2, result.Problems.Count(problem => problem.Code == "promote.cited_missing_in_beta"));
        }

        [Fact]
        public void A_path_that_differs_from_production_only_by_CASE_is_refused()
        {
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree((P("Main.png"), "m"));
            PublishedTreeView production = ProductionPromotionPlanTests.Tree((P("main.png"), "old"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan([P("Main.png")], [], beta, production);

            Assert.False(result.CanPromote);
            Assert.Equal("promote.case_collision", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_system_with_NOTHING_in_BETA_is_refused()
        {
            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [], [], ProductionPromotionPlanTests.Tree(), ProductionPromotionPlanTests.Tree());

            Assert.False(result.CanPromote);
            Assert.Equal("promote.not_in_beta", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_walked_file_OUTSIDE_the_system_folder_is_refused()
        {
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree(
                (P("main.png"), "m"),
                ("Commodore/C64/250425/main.png", "x"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png"), "Commodore/C64/250425/main.png"], [], beta, ProductionPromotionPlanTests.Tree());

            Assert.False(result.CanPromote);
            Assert.Equal("promote.outside_system", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_refused_plan_carries_NO_files_so_nothing_can_be_half_copied_from_it()
        {
            PublishedTreeView beta = ProductionPromotionPlanTests.Tree((P("main.png"), "m"));

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [P("main.png")], [P("gone.png")], beta, ProductionPromotionPlanTests.Tree());

            Assert.False(result.CanPromote);
            Assert.Empty(result.Files);
        }
    }
}
