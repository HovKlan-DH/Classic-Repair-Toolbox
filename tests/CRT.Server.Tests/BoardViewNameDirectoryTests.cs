using CRT.Server.Configuration;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers BoardViewNameDirectory - which boards a view may name, and their names, from the
    // published main Excel data files: production's first, then BETA's.
    //
    // A real temp tree holds the master FILES (the directory finds the newest generation and watches
    // its time and size); what a file holds is handed in by a fake reader, since reading a workbook
    // is MasterListing's and tested there.
    // ###########################################################################################
    public sealed class BoardViewNameDirectoryTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-board-view-names", Guid.NewGuid().ToString("N"));

        private DateTimeOffset thisNow = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private readonly Dictionary<string, List<MasterListingRow>> thisContents = new(StringComparer.Ordinal);

        private int thisReads;

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

        // A data root holding one master of the given generation; returns its path.
        private string Master(string tree, string generation = "2.0.0")
        {
            string root = Path.Combine(this.thisRoot, tree);
            Directory.CreateDirectory(root);

            string path = Path.Combine(root, DataGenerationRules.BuildMasterFileName(Version.Parse(generation)));
            File.WriteAllText(path, "x");
            return path;
        }

        private static MasterListingRow Row(string hardware, string board, string file) => new(hardware, board, file, string.Empty);

        private BoardViewNameDirectory Names(params string[] roots) =>
            new(roots.Select(root => Path.Combine(this.thisRoot, root)).ToList(), this.Read, () => this.thisNow);

        private bool Read(string masterPath, out IReadOnlyList<MasterListingRow> rows)
        {
            this.thisReads++;

            if (this.thisContents.TryGetValue(masterPath, out List<MasterListingRow>? listed))
            {
                rows = listed;
                return true;
            }

            rows = [];
            return false;
        }

        [Fact]
        public void A_listed_system_is_found_with_the_drop_down_names()
        {
            this.thisContents[this.Master("prod")] = [BoardViewNameDirectoryTests.Row("Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407.xlsx")];

            BoardViewNames? names = this.Names("prod").Find("Commodore/C64/250407");

            Assert.Equal(new BoardViewNames("Commodore 64", "250407"), names);
        }

        // An id no listing has names nothing - the guard that keeps made-up ids out of the table.
        [Theory]
        [InlineData("Commodore/C64/999999")]
        [InlineData("commodore/c64/250407")]
        [InlineData("")]
        [InlineData(null)]
        public void A_system_no_listing_has_is_not_found(string? systemId)
        {
            this.thisContents[this.Master("prod")] = [BoardViewNameDirectoryTests.Row("Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407.xlsx")];

            Assert.Null(this.Names("prod").Find(systemId));
        }

        // Production names win; a system only BETA lists (placed, not yet promoted) is still found.
        [Fact]
        public void Production_is_asked_first_and_beta_after_it()
        {
            this.thisContents[this.Master("prod")] = [BoardViewNameDirectoryTests.Row("Commodore 64", "250407", "Commodore/C64/250407/a.xlsx")];
            this.thisContents[this.Master("beta")] =
            [
                BoardViewNameDirectoryTests.Row("C64 (renamed in BETA)", "250407", "Commodore/C64/250407/a.xlsx"),
                BoardViewNameDirectoryTests.Row("Commodore 16", "310156", "Commodore/C16/310156/b.xlsx")
            ];

            BoardViewNameDirectory directory = this.Names("prod", "beta");

            Assert.Equal("Commodore 64", directory.Find("Commodore/C64/250407")!.HardwareName);
            Assert.Equal("Commodore 16", directory.Find("Commodore/C16/310156")!.HardwareName);
        }

        // Only the NEWEST generation is read - older masters serve older builds.
        [Fact]
        public void Only_the_newest_generations_master_is_read()
        {
            this.thisContents[this.Master("prod", "1.0.0")] = [BoardViewNameDirectoryTests.Row("Old name", "250407", "Commodore/C64/250407/a.xlsx")];
            this.thisContents[this.Master("prod", "2.0.0")] = [BoardViewNameDirectoryTests.Row("New name", "250407", "Commodore/C64/250407/a.xlsx")];

            Assert.Equal("New name", this.Names("prod").Find("Commodore/C64/250407")!.HardwareName);
        }

        // ###########################################################################################
        // A workbook is read once per file version: many views in a minute read it once, and an
        // unchanged file is not read again after RecheckAfter either - only a changed one is.
        // ###########################################################################################
        [Fact]
        public void A_master_is_read_again_only_when_the_file_changes()
        {
            string master = this.Master("prod");
            this.thisContents[master] = [BoardViewNameDirectoryTests.Row("Commodore 64", "250407", "Commodore/C64/250407/a.xlsx")];
            BoardViewNameDirectory directory = this.Names("prod");

            directory.Find("Commodore/C64/250407");
            directory.Find("Commodore/C64/250407");
            this.thisNow += BoardViewNameDirectory.RecheckAfter + TimeSpan.FromSeconds(1);
            directory.Find("Commodore/C64/250407");

            Assert.Equal(1, this.thisReads);

            // Changed on disk: read again once the recheck is due.
            File.WriteAllText(master, "a longer file");
            this.thisContents[master] = [BoardViewNameDirectoryTests.Row("Commodore 64C", "250407", "Commodore/C64/250407/a.xlsx")];
            this.thisNow += BoardViewNameDirectory.RecheckAfter + TimeSpan.FromSeconds(1);

            Assert.Equal("Commodore 64C", directory.Find("Commodore/C64/250407")!.HardwareName);
            Assert.Equal(2, this.thisReads);
        }

        // A master that cannot be read just now keeps the names last read, and is tried again - it
        // never turns into an empty listing refusing every view until the file changes.
        [Fact]
        public void A_master_that_cannot_be_read_keeps_the_names_last_read_and_is_tried_again()
        {
            string master = this.Master("prod");
            this.thisContents[master] = [BoardViewNameDirectoryTests.Row("Commodore 64", "250407", "Commodore/C64/250407/a.xlsx")];
            BoardViewNameDirectory directory = this.Names("prod");

            directory.Find("Commodore/C64/250407");

            File.WriteAllText(master, "being replaced");
            this.thisContents.Remove(master);
            this.thisNow += BoardViewNameDirectory.RecheckAfter + TimeSpan.FromSeconds(1);

            Assert.Equal("Commodore 64", directory.Find("Commodore/C64/250407")!.HardwareName);

            this.thisContents[master] = [BoardViewNameDirectoryTests.Row("Commodore 64C", "250407", "Commodore/C64/250407/a.xlsx")];
            this.thisNow += BoardViewNameDirectory.RecheckAfter + TimeSpan.FromSeconds(1);

            Assert.Equal("Commodore 64C", directory.Find("Commodore/C64/250407")!.HardwareName);
        }

        // Production's promotion root first, the older setting when only it is set, BETA last.
        [Fact]
        public void The_roots_are_production_then_beta()
        {
            string production = Path.Combine(Path.GetTempPath(), "prod", "Data");
            string older = Path.Combine(Path.GetTempPath(), "prod-old", "Data");
            string beta = Path.Combine(Path.GetTempPath(), "beta", "Data");

            Assert.Equal(
                [production, beta],
                BoardViewNameDirectory.RootsOf(new ServerOptions { ProductionDataTreeRoot = production, ProductionTreeRoot = older, DataTreeRoot = beta }));

            Assert.Equal([older, beta], BoardViewNameDirectory.RootsOf(new ServerOptions { ProductionTreeRoot = older, DataTreeRoot = beta }));
            Assert.Equal([beta], BoardViewNameDirectory.RootsOf(new ServerOptions { DataTreeRoot = beta }));
        }
    }
}
