using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers ProductionPromotionPlan - what publishing one board from BETA to Production copies,
    // and what it refuses (owner request, 2026-09-25).
    //
    // Production is what every user downloads, so the refusals matter as much as the copies: a
    // board promoted while pointing at a file Production does not have is a broken board on every
    // machine. In the CRT 4.0 tree (2026-10-25) a board is its board file and the stored files it
    // lists ("Files/<ab>/<sha256>.<ext>") - so those are what is compared and copied.
    // ###########################################################################################
    public sealed class ProductionPromotionPlanTests
    {
        private const string BoardFile = "Commodore/C64/250407/board.json";

        private static readonly string ShaA = new('a', 64);
        private static readonly string ShaB = new('b', 64);
        private static readonly string BoardSha = new('1', 64);
        private static readonly string OldBoardSha = new('2', 64);

        private static BoardDocument.StoredFileRow Row(string name, string sha) => new() { Name = name, Sha256 = sha, Size = 10 };

        private static string Stored(string sha, string name) => DataTreeLayout.StoredFilePath(sha, name);

        private static ProductionPromotionResult Plan(
            IReadOnlyList<BoardDocument.StoredFileRow>? beta,
            IEnumerable<string>? inBeta = null,
            IEnumerable<string>? inProduction = null,
            string? betaBoardFile = null,
            string? productionBoardFile = null,
            Func<string, string?>? productionVariant = null)
        {
            var betaHolds = new HashSet<string>(inBeta ?? [], StringComparer.Ordinal);
            var productionHolds = new HashSet<string>(inProduction ?? [], StringComparer.Ordinal);

            return ProductionPromotionPlan.Build(
                "Commodore", "C64", "250407",
                beta,
                beta is null ? null : betaBoardFile ?? ProductionPromotionPlanTests.BoardSha,
                productionBoardFile,
                betaHolds.Contains,
                productionHolds.Contains,
                productionVariant);
        }

        // -----------------------------------------------------------------------------------
        // What is copied
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Only_stored_files_production_lacks_and_a_changed_board_file_are_copied()
        {
            // A typo fix changes the board file; the 1,200 images are stored already and must not be
            // copied to every user's machine for nothing.
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");
            string b = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaB, "b.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("a.png", ProductionPromotionPlanTests.ShaA), ProductionPromotionPlanTests.Row("b.png", ProductionPromotionPlanTests.ShaB)],
                inBeta: [a, b],
                inProduction: [a],
                productionBoardFile: ProductionPromotionPlanTests.OldBoardSha);

            Assert.True(result.CanPromote);
            Assert.Equal([b, ProductionPromotionPlanTests.BoardFile], result.Files.Select(file => file.Path));
            Assert.Equal(PromotionChange.Added, result.Files[0].Change);
            Assert.Equal(PromotionChange.Replaced, result.Files[1].Change);
            Assert.Equal([a], result.Unchanged);
            Assert.Equal(1, result.UnchangedCount);
        }

        [Fact]
        public void A_stored_files_checksum_is_the_one_its_name_carries()
        {
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("Images/a.png", ProductionPromotionPlanTests.ShaA)], inBeta: [a]);

            PromotionFile file = result.Files.First(copy => copy.Stage == PromotionStage.Content);

            Assert.Equal(ProductionPromotionPlanTests.ShaA, file.Sha256);
            Assert.Equal("Images/a.png", file.Name);
        }

        [Fact]
        public void The_copies_are_in_WRITE_ORDER_content_then_the_board_file()
        {
            // The publish's order and for its reason: an interrupted promotion must leave a board
            // that loads - a board file never points at bytes that have not arrived yet.
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");
            string b = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaB, "z.pdf");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("z.pdf", ProductionPromotionPlanTests.ShaB), ProductionPromotionPlanTests.Row("a.png", ProductionPromotionPlanTests.ShaA)],
                inBeta: [a, b]);

            Assert.Equal(PromotionStage.Board, result.Files[^1].Stage);
            Assert.Equal(ProductionPromotionPlanTests.BoardFile, result.Files[^1].Path);
            Assert.All(result.Files.Take(result.Files.Count - 1), file => Assert.Equal(PromotionStage.Content, file.Stage));
        }

        [Fact]
        public void A_board_production_does_not_have_adds_its_board_file()
        {
            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan([]);

            PromotionFile file = Assert.Single(result.Files);
            Assert.Equal(PromotionChange.Added, file.Change);
            Assert.Equal(ProductionPromotionPlanTests.BoardSha, file.Sha256);
            Assert.Equal(DataTreeLayout.BoardFileName, file.Name);
        }

        [Fact]
        public void Two_names_of_the_same_bytes_are_copied_once()
        {
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("a.png", ProductionPromotionPlanTests.ShaA), ProductionPromotionPlanTests.Row("copy of a.png", ProductionPromotionPlanTests.ShaA)],
                inBeta: [a]);

            Assert.Single(result.Files, file => file.Path == a);
        }

        [Fact]
        public void Nothing_to_copy_is_still_a_promotion_that_records_the_revision()
        {
            // BETA and production brought level by hand: promoting copies nothing and marks the
            // board as in production.
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("a.png", ProductionPromotionPlanTests.ShaA)],
                inBeta: [a],
                inProduction: [a],
                productionBoardFile: ProductionPromotionPlanTests.BoardSha);

            Assert.True(result.CanPromote);
            Assert.Empty(result.Files);
            Assert.Equal([ProductionPromotionPlanTests.BoardFile, a], result.Unchanged);
        }

        [Fact]
        public void Nothing_waits_for_the_administrator_any_more()
        {
            // No shared files in the 4.0 tree: a promotion never changes another board's file.
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("a.png", ProductionPromotionPlanTests.ShaA)], inBeta: [a]);

            Assert.False(result.TouchesSharedFiles);
            Assert.All(result.Files, file => Assert.False(file.IsShared));
        }

        [Fact]
        public void The_plan_names_every_path_by_the_boards_own_name()
        {
            // What the Maintainer tab draws the tree with - a stored file's path is its checksum.
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("Images/a.png", ProductionPromotionPlanTests.ShaA)], inBeta: [a]);

            Assert.Equal("Images/a.png", result.Names![a]);
            Assert.Equal(DataTreeLayout.BoardFileName, result.Names[ProductionPromotionPlanTests.BoardFile]);
        }

        // -----------------------------------------------------------------------------------
        // Refusals
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_listed_file_BETA_itself_does_not_have_refuses_the_promotion()
        {
            // The BETA board is broken; a broken board does not go to everyone.
            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("gone.png", ProductionPromotionPlanTests.ShaA)]);

            Assert.False(result.CanPromote);
            Assert.Equal("promote.cited_missing_in_beta", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_listed_file_without_a_usable_checksum_refuses_the_promotion()
        {
            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("a.png", "not-a-hash")]);

            Assert.False(result.CanPromote);
            Assert.Equal("promote.unreadable_in_beta", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_board_folder_that_differs_from_productions_only_by_CASE_is_refused()
        {
            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [],
                productionVariant: folder => folder == "Commodore/C64/250407" ? "commodore/C64/250407" : null);

            Assert.False(result.CanPromote);
            Assert.Equal("promote.case_collision", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_board_with_NOTHING_in_BETA_is_refused()
        {
            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(null);

            Assert.False(result.CanPromote);
            Assert.Equal("promote.not_in_beta", Assert.Single(result.Problems).Code);
        }

        [Fact]
        public void A_refused_plan_carries_NO_files_so_nothing_can_be_half_copied_from_it()
        {
            string a = ProductionPromotionPlanTests.Stored(ProductionPromotionPlanTests.ShaA, "a.png");

            ProductionPromotionResult result = ProductionPromotionPlanTests.Plan(
                [ProductionPromotionPlanTests.Row("a.png", ProductionPromotionPlanTests.ShaA), ProductionPromotionPlanTests.Row("gone.png", ProductionPromotionPlanTests.ShaB)],
                inBeta: [a]);

            Assert.False(result.CanPromote);
            Assert.Empty(result.Files);
        }
    }
}
