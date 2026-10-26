using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionFileRules - WHICH files a submission may carry (security review, 2026-09-25).
    //
    // In the CRT 4.0 tree (2026-10-25) a submitted file is the board's NAME for it, its bytes stored
    // under their checksum - so no submission names a place in the tree, and the rules about another
    // board's files and the shared folders went with them. What stays: a file no row uses (no
    // maintainer ever saw it), a type board data never uses (a dot-file the web server reads, an
    // executable that syncs to every disk), and a NEW board whose folder differs from a published
    // one only by case.
    // ###########################################################################################
    public sealed class SubmissionFileRulesTests
    {
        // A manifest for Commodore/C64/250407 carrying these files, each cited by a board-level
        // row unless the caller says otherwise.
        private static SubmissionManifest Manifest(params (string Name, string Hash)[] files)
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407"
            };

            foreach ((string name, string hash) in files)
            {
                manifest.Files.Add(new SubmissionFile { Path = name, Sha256 = hash, SizeBytes = 10 });
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = name });
            }

            return manifest;
        }

        private static string Hash(char c) => new(c, 64);

        private static IReadOnlyList<string> Codes(SubmissionManifest manifest, PublishedTreeView? board = null) =>
            SubmissionFileRules.ValidateManifestFiles(manifest, board).Select(finding => finding.Code).ToList();

        // The published board listing these names (name -> hash).
        private static PublishedTreeView Listing(params (string Name, string Hash)[] files)
        {
            Dictionary<string, string> hashes = files.ToDictionary(file => file.Name, file => file.Hash);
            return new PublishedTreeView(name => hashes.TryGetValue(name, out string? hash) ? hash : null, _ => null);
        }

        [Fact]
        public void Files_of_allowed_types_that_rows_use_pass()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(
                ("Sheet1.png", SubmissionFileRulesTests.Hash('1')),
                ("Board local files/CBM Approved Cross Reference.html", SubmissionFileRulesTests.Hash('2')),
                ("Datasheets/7805.PDF", SubmissionFileRulesTests.Hash('3')));

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));
        }

        // ---------------------------------------------------------------------- type and name

        [Theory]
        [InlineData(".htaccess", "file.hidden_name")]
        [InlineData(".well-known/x.png", "file.hidden_name")]
        [InlineData(".crt-base.json", "file.hidden_name")]
        [InlineData("Makefile", "file.no_type")]
        [InlineData("run.exe", "file.type_not_allowed")]
        [InlineData("page.php", "file.type_not_allowed")]
        [InlineData("drawing.svg", "file.type_not_allowed")]
        [InlineData("board.json", "file.type_not_allowed")]
        [InlineData("Data C64 250407 v3.0.0.xlsx", "file.type_not_allowed")]
        public void A_file_of_a_type_board_data_never_uses_is_refused(string name, string code)
        {
            // A dot-file is read by Apache as configuration for its folder, and ".crt-" names are
            // CRT's own bookkeeping in a draft; a .json would be a board file of the contributor's
            // making; a .php or .exe is simply not board data.
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((name, SubmissionFileRulesTests.Hash('1')));

            Assert.Contains(code, SubmissionFileRulesTests.Codes(manifest));
        }

        // ---------------------------------------------------------------------- KiCad data (2026-09-26)

        // ###########################################################################################
        // *** THE BOARD'S KiCad DATA IS THE ONE EXEMPTION FROM THE CITATION RULE. *** No row cites
        // it - CRT reads the files by name - so before this, a submission could not carry it at all
        // and a new board was published without its traces.
        // ###########################################################################################
        [Fact]
        public void KiCad_data_is_accepted_although_no_row_cites_it()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(("Sheet1.png", SubmissionFileRulesTests.Hash('1')));

            manifest.Files.Add(new SubmissionFile { Path = "KiCad data/board.kicad_pcb", Sha256 = SubmissionFileRulesTests.Hash('2'), SizeBytes = 10 });
            manifest.Files.Add(new SubmissionFile { Path = "KiCad data/Pages/vic.kicad_sch", Sha256 = SubmissionFileRulesTests.Hash('3'), SizeBytes = 10 });

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));
        }

        // The exemption is exactly as wide as SubmissionKiCadFiles.IsSubmittable: a KiCad type
        // anywhere else, or a type CRT never reads inside the folder, is refused as before.
        [Theory]
        [InlineData("board.kicad_pcb", "file.type_not_allowed")]
        [InlineData("Images/KiCad data/board.kicad_pcb", "file.type_not_allowed")]
        [InlineData("KiCad data/KiCad-traces.json", "file.type_not_allowed")]
        public void A_KiCad_type_outside_the_KiCad_folder_is_refused(string name, string code)
        {
            var manifest = new SubmissionManifest { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" };
            manifest.Files.Add(new SubmissionFile { Path = name, Sha256 = SubmissionFileRulesTests.Hash('1'), SizeBytes = 10 });

            Assert.Contains(code, SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));
        }

        // A converted board carries a few files of its KiCad folder that nothing reads; they go along
        // as the published board has them, and cannot be changed or brought in new.
        [Fact]
        public void A_KiCad_folder_straggler_passes_only_as_the_published_board_lists_it()
        {
            var manifest = new SubmissionManifest { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" };
            manifest.Files.Add(new SubmissionFile { Path = "KiCad data/KiCad-traces.json", Sha256 = SubmissionFileRulesTests.Hash('1'), SizeBytes = 10 });

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, SubmissionFileRulesTests.Listing(("KiCad data/KiCad-traces.json", SubmissionFileRulesTests.Hash('1')))));
            Assert.NotEmpty(SubmissionFileRulesTests.Codes(manifest, SubmissionFileRulesTests.Listing(("KiCad data/KiCad-traces.json", SubmissionFileRulesTests.Hash('2')))));
            Assert.NotEmpty(SubmissionFileRulesTests.Codes(manifest, board: null));
        }

        // An allowed ORDINARY type inside the KiCad folder gets no exemption: uncited, it is a file
        // no maintainer would ever see.
        [Fact]
        public void An_uncited_image_inside_the_KiCad_folder_is_still_refused()
        {
            var manifest = new SubmissionManifest { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" };
            manifest.Files.Add(new SubmissionFile { Path = "KiCad data/photo.png", Sha256 = SubmissionFileRulesTests.Hash('1'), SizeBytes = 10 });

            Assert.Contains("file.not_used", SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));
        }

        // ###########################################################################################
        // *** ONE BAD NAME, ONE FINDING (code review, 2026-09-25). *** A name the path rules refuse
        // is not reported again here under another code - two contradictory explanations of one
        // fault. The path rules' finding stands.
        // ###########################################################################################
        [Theory]
        [InlineData("Images/../../etc/x.png")]
        [InlineData("/etc/x.png")]
        [InlineData("Images//x.png")]
        public void A_name_the_path_rules_refuse_is_reported_once_by_them_alone(string name)
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((name, SubmissionFileRulesTests.Hash('1')));

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));

            Assert.Equal(
                "path.rejected",
                Assert.Single(SubmissionPathRules.ValidateManifestPaths(manifest, System.IO.Path.GetTempPath())).Code);
        }

        [Fact]
        public void The_refusal_names_the_types_that_ARE_allowed()
        {
            SubmissionFileRules.TryCheckName("a.exe", out _, out string reason);

            Assert.Contains(".png", reason, System.StringComparison.Ordinal);
            Assert.Contains(".pdf", reason, System.StringComparison.Ordinal);
        }

        // --------------------------------------------------------------------- used by a row

        // Nothing would show a file no row names - so no maintainer could have seen it.
        [Fact]
        public void A_file_NO_ROW_USES_is_refused()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(("a.png", SubmissionFileRulesTests.Hash('1')));
            manifest.Rows.BoardLocalFiles.Clear();

            Assert.Contains("file.not_used", SubmissionFileRulesTests.Codes(manifest));
        }

        // Every row kind that names a file counts, through the same collector the client uses to
        // build its file list - so a client sending exactly what its rows cite is never refused.
        [Fact]
        public void A_file_named_by_ANY_kind_of_row_counts_as_used()
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407",
                Rows = new SubmissionRows
                {
                    Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "s.png" } },
                    ComponentImages = { new ComponentImageEntry { BoardLabel = "U8", File = "i.jpg" } },
                    ComponentLocalFiles = { new ComponentLocalFileEntry { BoardLabel = "U8", File = "d.pdf" } },
                    BoardLocalFiles = { new BoardLocalFileEntry { File = "n.txt" } }
                }
            };

            foreach (string name in new[] { "s.png", "i.jpg", "d.pdf", "n.txt" })
                manifest.Files.Add(new SubmissionFile { Path = name, Sha256 = SubmissionFileRulesTests.Hash('1') });

            Assert.DoesNotContain("file.not_used", SubmissionFileRulesTests.Codes(manifest));
        }

        // A name that once meant another board's folder is just a name of THIS board now - its
        // bytes are stored under their checksum, so nothing of the other board can change.
        [Fact]
        public void A_name_shaped_like_another_boards_path_is_this_boards_own_file()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(("Commodore/C64/250425/Sheet1.png", SubmissionFileRulesTests.Hash('1')));

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, board: null));
        }

        // --------------------------------------------------------------------- case variants

        // The squat this closes: an anonymous "new board" whose folder is a case-variant of a real
        // board, whose board file would then replace the real board's on every Windows and macOS
        // client.
        [Fact]
        public void A_board_whose_folder_is_a_case_variant_of_a_published_one_is_refused()
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "commodore/c64/250407",
                Manufacturer = "commodore",
                Hardware = "c64",
                Board = "250407"
            };

            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("Commodore/C64/250407/board.json");

            Assert.Contains("identity.case_collision", SubmissionFileRulesTests.Codes(manifest, tree));
        }

        // A duplicate name is SubmissionPathRules' finding; this class must not report it twice.
        [Fact]
        public void A_duplicate_name_is_judged_once()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(("a.exe", SubmissionFileRulesTests.Hash('1')));
            manifest.Files.Add(new SubmissionFile { Path = "a.exe", Sha256 = SubmissionFileRulesTests.Hash('1') });

            Assert.Single(SubmissionFileRulesTests.Codes(manifest));
        }
    }
}
