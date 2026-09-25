using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionFileRules - WHICH files a submission may carry and WHERE (security review,
    // 2026-09-25).
    //
    // Each rule closes a way a perfectly SHAPED path still did harm: overwriting another board's
    // file, carrying a file no row uses (so no maintainer ever saw it), carrying a dot-file the web
    // server reads as configuration, or landing beside a published path that differs only by case.
    // SubmissionRulesShippedDataTests proves none of them refuses a board that is already published.
    // ###########################################################################################
    public sealed class SubmissionFileRulesTests
    {
        private const string Own = "Commodore/C64/250407/";

        // A manifest for Commodore/C64/250407 carrying these files, each cited by a board-level
        // row unless the caller says otherwise.
        private static SubmissionManifest Manifest(params (string Path, string Hash)[] files)
        {
            var manifest = new SubmissionManifest
            {
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407"
            };

            foreach ((string path, string hash) in files)
            {
                manifest.Files.Add(new SubmissionFile { Path = path, Sha256 = hash, SizeBytes = 10 });
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = path });
            }

            return manifest;
        }

        private static string Hash(char c) => new(c, 64);

        private static IReadOnlyList<string> Codes(SubmissionManifest manifest, PublishedTreeView? tree = null) =>
            SubmissionFileRules.ValidateManifestFiles(manifest, tree).Select(finding => finding.Code).ToList();

        [Fact]
        public void Own_and_shared_files_of_allowed_types_that_rows_use_pass()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(
                (SubmissionFileRulesTests.Own + "Sheet1.png", SubmissionFileRulesTests.Hash('1')),
                ("Commodore/Shared files/Board local files/CBM Approved Cross Reference.html", SubmissionFileRulesTests.Hash('2')),
                ("Generic shared files/Datasheets/7805.PDF", SubmissionFileRulesTests.Hash('3')));

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));
        }

        // ---------------------------------------------------------------------- type and name

        [Theory]
        [InlineData("Commodore/C64/250407/.htaccess", "file.hidden_name")]
        [InlineData("Commodore/C64/250407/.well-known/x.png", "file.hidden_name")]
        [InlineData("Commodore/C64/250407/Makefile", "file.no_type")]
        [InlineData("Commodore/C64/250407/run.exe", "file.type_not_allowed")]
        [InlineData("Commodore/C64/250407/page.php", "file.type_not_allowed")]
        [InlineData("Commodore/C64/250407/drawing.svg", "file.type_not_allowed")]
        [InlineData("Commodore/C64/250407/system.json", "file.type_not_allowed")]
        [InlineData("Commodore/C64/250407/Data C64 250407 v2.0.0.json", "file.type_not_allowed")]
        public void A_file_of_a_type_board_data_never_uses_is_refused(string path, string code)
        {
            // A dot-file is read by Apache as configuration for its folder; an uploaded .json would
            // rewrite a board's generated highlight sidecar or its system.json; a .php or .exe is
            // simply not board data.
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((path, SubmissionFileRulesTests.Hash('1')));

            Assert.Contains(code, SubmissionFileRulesTests.Codes(manifest));
        }

        // ###########################################################################################
        // *** ONE BAD PATH, ONE FINDING (code review, 2026-09-25). *** A path the path rules refuse
        // used to be reported again here under another code - "Commodore/../etc/x.png" as a hidden
        // name, an absolute path as unused - two contradictory explanations of one fault. The file
        // rules now skip it, as their header always said they did; the path rules' finding stands.
        // ###########################################################################################
        [Theory]
        [InlineData("Commodore/../etc/x.png")]
        [InlineData("/etc/x.png")]
        [InlineData("Commodore//x.png")]
        public void A_path_the_path_rules_refuse_is_reported_once_by_them_alone(string path)
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((path, SubmissionFileRulesTests.Hash('1')));

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));

            Assert.Equal(
                "path.rejected",
                Assert.Single(SubmissionPathRules.ValidateManifestPaths(manifest, System.IO.Path.GetTempPath())).Code);
        }

        [Fact]
        public void The_refusal_names_the_types_that_ARE_allowed()
        {
            SubmissionFileRules.TryCheckName("Commodore/C64/250407/a.exe", out _, out string reason);

            Assert.Contains(".png", reason, System.StringComparison.Ordinal);
            Assert.Contains(".pdf", reason, System.StringComparison.Ordinal);
        }

        // --------------------------------------------------------------------- used by a row

        // Nothing would show a file no row names - so no maintainer could have seen it.
        [Fact]
        public void A_file_NO_ROW_USES_is_refused()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(
                (SubmissionFileRulesTests.Own + "a.png", SubmissionFileRulesTests.Hash('1')));

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
                SystemId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407",
                Rows = new SubmissionRows
                {
                    Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = SubmissionFileRulesTests.Own + "s.png" } },
                    ComponentImages = { new ComponentImageEntry { BoardLabel = "U8", File = SubmissionFileRulesTests.Own + "i.jpg" } },
                    ComponentLocalFiles = { new ComponentLocalFileEntry { BoardLabel = "U8", File = SubmissionFileRulesTests.Own + "d.pdf" } },
                    BoardLocalFiles = { new BoardLocalFileEntry { File = SubmissionFileRulesTests.Own + "n.txt" } }
                }
            };

            foreach (string path in new[] { "s.png", "i.jpg", "d.pdf", "n.txt" })
                manifest.Files.Add(new SubmissionFile { Path = SubmissionFileRulesTests.Own + path, Sha256 = SubmissionFileRulesTests.Hash('1') });

            Assert.DoesNotContain("file.not_used", SubmissionFileRulesTests.Codes(manifest));
        }

        // -------------------------------------------------------------------- another board

        // THE finding that motivated this class: a submission to one board carrying another
        // board's file, which the publish then copied over the real one.
        [Fact]
        public void ANOTHER_boards_file_is_refused_when_the_tree_cannot_be_consulted()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(
                ("Commodore/C64/250425/Sheet1.png", SubmissionFileRulesTests.Hash('1')));

            Assert.Contains("file.other_board", SubmissionFileRulesTests.Codes(manifest, tree: null));
        }

        [Fact]
        public void ANOTHER_boards_file_that_is_not_published_is_refused()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(
                ("Commodore/C64/250425/new.png", SubmissionFileRulesTests.Hash('1')));

            Assert.Contains("file.other_board", SubmissionFileRulesTests.Codes(manifest, PublishedTreeView.Empty));
        }

        [Fact]
        public void ANOTHER_boards_file_that_would_CHANGE_is_refused()
        {
            string path = "Commodore/C64/250425/Sheet1.png";
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((path, SubmissionFileRulesTests.Hash('1')));

            var tree = new PublishedTreeView(p => p == path ? SubmissionFileRulesTests.Hash('2') : null, _ => null);

            Assert.Contains("file.other_board_changed", SubmissionFileRulesTests.Codes(manifest, tree));
        }

        // Real published data does this - C128DCR 250477 cites texts in the C128 310378 folder -
        // so citing another board's file UNCHANGED must pass, or that board could never be
        // submitted again.
        [Fact]
        public void ANOTHER_boards_file_cited_UNCHANGED_passes()
        {
            string path = "Commodore/C128/310378/Scope baseline/notes.txt";
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((path, SubmissionFileRulesTests.Hash('4')));

            var tree = new PublishedTreeView(p => p == path ? SubmissionFileRulesTests.Hash('4') : null, _ => null);

            Assert.Empty(SubmissionFileRulesTests.Codes(manifest, tree));
        }

        // --------------------------------------------------------------------- case variants

        [Fact]
        public void A_path_differing_from_a_published_one_only_by_case_is_refused()
        {
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest(
                (SubmissionFileRulesTests.Own + "sheet1.png", SubmissionFileRulesTests.Hash('1')));

            PublishedTreeView tree = PublishedTreeViewTests.TreeOf(SubmissionFileRulesTests.Own + "Sheet1.png");

            Assert.Contains("path.case_collision_published", SubmissionFileRulesTests.Codes(manifest, tree));
        }

        // The squat this closes: an anonymous "new system" whose folder is a case-variant of a real
        // board, whose files would then replace the real board's on every Windows and macOS client.
        [Fact]
        public void A_board_whose_folder_is_a_case_variant_of_a_published_one_is_refused()
        {
            var manifest = new SubmissionManifest
            {
                SystemId = "commodore/c64/250407",
                Manufacturer = "commodore",
                Hardware = "c64",
                Board = "250407"
            };

            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("Commodore/C64/250407/Sheet1.png");

            Assert.Contains("identity.case_collision", SubmissionFileRulesTests.Codes(manifest, tree));
        }

        // A duplicate path is SubmissionPathRules' finding; this class must not report it twice.
        [Fact]
        public void A_duplicate_path_is_judged_once()
        {
            string path = SubmissionFileRulesTests.Own + "a.exe";
            SubmissionManifest manifest = SubmissionFileRulesTests.Manifest((path, SubmissionFileRulesTests.Hash('1')));
            manifest.Files.Add(new SubmissionFile { Path = path, Sha256 = SubmissionFileRulesTests.Hash('1') });

            Assert.Single(SubmissionFileRulesTests.Codes(manifest));
        }
    }
}
