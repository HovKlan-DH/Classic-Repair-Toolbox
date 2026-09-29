using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ProductionPromoter - the copy half of a production promotion - against real temp
    // trees. The plan decides what; these pin that the copy is verified, in order, and that a
    // refusal found before writing leaves production untouched.
    // ###########################################################################################
    public sealed class ProductionPromoterTests : IDisposable
    {
        private readonly string thisRoot;
        private readonly string thisBeta;
        private readonly string thisProduction;

        public ProductionPromoterTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-promoter", Guid.NewGuid().ToString("N"));
            this.thisBeta = Path.Combine(this.thisRoot, "beta");
            this.thisProduction = Path.Combine(this.thisRoot, "production");

            Directory.CreateDirectory(this.thisBeta);
            Directory.CreateDirectory(this.thisProduction);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private PromotionFile PutInBeta(string relative, string content, PromotionStage stage = PromotionStage.Content)
        {
            string path = Path.Combine(this.thisBeta, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

            byte[] bytes = Encoding.UTF8.GetBytes(content);
            File.WriteAllBytes(path, bytes);

            return new PromotionFile(relative, Convert.ToHexStringLower(SHA256.HashData(bytes)), PromotionChange.Added, stage, IsShared: false);
        }

        private string InProduction(string relative) =>
            Path.Combine(this.thisProduction, relative.Replace('/', Path.DirectorySeparatorChar));

        [Fact]
        public async Task Every_planned_file_lands_with_BETAs_exact_bytes_and_no_temporary_file_is_left()
        {
            PromotionFile image = this.PutInBeta("Commodore/C64/250407/Images/a.png", "image");
            PromotionFile sidecar = this.PutInBeta("Commodore/C64/250407/Data C64 250407 v2.0.0.json", "{}", PromotionStage.Board);

            PromotionCopyOutcome outcome = await ProductionPromoter.CopyAsync([image, sidecar], this.thisBeta, this.thisProduction);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(2, outcome.FilesCopied);
            Assert.Equal("image", await File.ReadAllTextAsync(this.InProduction("Commodore/C64/250407/Images/a.png")));
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(this.InProduction("Commodore/C64/250407/Data C64 250407 v2.0.0.json"))!));
        }

        // ###########################################################################################
        // *** A PRODUCTION FOLDER THE SERVICE MAY NOT WRITE STOPS THE PROMOTION BEFORE THE FIRST
        // COPY (owner report, 2026-09-28). *** A folder copied into production by hand as root
        // refuses every write, and meeting it at file 600 leaves production half-promoted. The
        // folders are asked first and the refusal names them - and hands them to the flow, which
        // logs the command that fixes them. (The probe is told the answer: a folder that may not
        // be written cannot be made on every OS.)
        // ###########################################################################################
        [Fact]
        public async Task A_production_folder_the_service_may_not_write_is_refused_before_anything_is_copied()
        {
            PromotionFile image = this.PutInBeta("Commodore/C64/250407/Images/a.png", "image");
            PromotionFile workbook = this.PutInBeta("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", "board", PromotionStage.Board);

            string board = this.InProduction("Commodore/C64/250407");
            Directory.CreateDirectory(board);

            PromotionCopyOutcome outcome = await ProductionPromoter.CopyAsync(
                [image, workbook], this.thisBeta, this.thisProduction, canWriteFolder: folder => folder != board);

            Assert.False(outcome.IsDone);
            Assert.Equal(0, outcome.FilesCopied);
            Assert.Equal([board], outcome.FoldersRefusing);
            Assert.Contains("[Commodore/C64/250407] in the production data, so nothing was changed", outcome.Error, StringComparison.Ordinal);
            Assert.Empty(Directory.GetFileSystemEntries(board));
        }

        [Fact]
        public async Task A_file_that_CHANGED_in_BETA_since_the_plan_is_not_copied_and_nothing_after_it_is()
        {
            // The maintainer checked one set of bytes; different ones must not reach every user under
            // the same name.
            PromotionFile first = this.PutInBeta("Commodore/C64/250407/a.png", "checked");
            PromotionFile board = this.PutInBeta("Commodore/C64/250407/Data.xlsx", "workbook", PromotionStage.Board);

            await File.WriteAllTextAsync(Path.Combine(this.thisBeta, "Commodore", "C64", "250407", "a.png"), "edited since");

            PromotionCopyOutcome outcome = await ProductionPromoter.CopyAsync([first, board], this.thisBeta, this.thisProduction);

            Assert.False(outcome.IsDone);
            Assert.Contains("changed in BETA", outcome.Error, StringComparison.Ordinal);
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/a.png")));
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/Data.xlsx")));
        }

        [Fact]
        public async Task A_file_gone_from_BETA_refuses_BEFORE_anything_is_written()
        {
            PromotionFile present = this.PutInBeta("Commodore/C64/250407/a.png", "a");
            var missing = new PromotionFile("Commodore/C64/250407/gone.png", new string('0', 64), PromotionChange.Added, PromotionStage.Content, false);

            PromotionCopyOutcome outcome = await ProductionPromoter.CopyAsync([present, missing], this.thisBeta, this.thisProduction);

            Assert.False(outcome.IsDone);
            Assert.Equal(0, outcome.FilesCopied);
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/a.png")));
        }

        [Fact]
        public async Task A_path_that_escapes_the_tree_refuses_BEFORE_anything_is_written()
        {
            PromotionFile present = this.PutInBeta("Commodore/C64/250407/a.png", "a");
            var escaping = new PromotionFile("../outside.png", new string('0', 64), PromotionChange.Added, PromotionStage.Content, false);

            PromotionCopyOutcome outcome = await ProductionPromoter.CopyAsync([present, escaping], this.thisBeta, this.thisProduction);

            Assert.False(outcome.IsDone);
            Assert.False(File.Exists(this.InProduction("Commodore/C64/250407/a.png")));
            Assert.False(File.Exists(Path.Combine(this.thisRoot, "outside.png")));
        }

        [Fact]
        public async Task An_existing_production_file_is_REPLACED_not_appended_to()
        {
            PromotionFile file = this.PutInBeta("Commodore/C64/250407/a.png", "new");

            Directory.CreateDirectory(Path.GetDirectoryName(this.InProduction("Commodore/C64/250407/a.png"))!);
            await File.WriteAllTextAsync(this.InProduction("Commodore/C64/250407/a.png"), "the old, longer content");

            PromotionCopyOutcome outcome = await ProductionPromoter.CopyAsync([file with { Change = PromotionChange.Replaced }], this.thisBeta, this.thisProduction);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal("new", await File.ReadAllTextAsync(this.InProduction("Commodore/C64/250407/a.png")));
        }
    }
}
