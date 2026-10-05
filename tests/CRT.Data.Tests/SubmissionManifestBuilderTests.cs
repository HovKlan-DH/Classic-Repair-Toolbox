using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers turning a drafted system into the manifest the server receives.
    //
    // TWO PROPERTIES CARRY THE MOST WEIGHT HERE:
    //
    // 1. WHICH FILES GET SENT is derived from the ROWS, never from a directory walk. A walk would
    //    sweep up editor backups and cache files and upload a contributor's unrelated files to a
    //    public server. The tests below assert that an unreferenced file is not included even when
    //    its hash is offered.
    //
    // 2. THE MANIFEST IS THE COMPLETE INTENDED STATE. A file the draft did not touch still has to
    //    appear, because omitting it reads as a deletion on the server.
    // ###########################################################################################
    public class SubmissionManifestBuilderTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private static SubmissionIdentity Identity()
        {
            return new SubmissionIdentity
            {
                // "Manufacturer/Hardware/Board" - the `systems` table's own primary key.
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                BaseRevision = "2026-09-01",
                Summary = "Corrected R12.",
                ApplicationVersion = "2.6.0",
                CreatedUtc = SubmissionManifestBuilderTests.Now
            };
        }

        private static BoardData Board()
        {
            return new BoardData
            {
                Schematics =
                {
                    new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "main.png" }
                },
                Components =
                {
                    new ComponentEntry { BoardLabel = "U8" },
                    new ComponentEntry { BoardLabel = "R12" }
                },
                ComponentImages =
                {
                    new ComponentImageEntry { BoardLabel = "U8", File = "Images/U8-pin3.png" }
                }
            };
        }

        private static Dictionary<string, SubmissionFileInfo> Hashes(params string[] paths)
        {
            var map = new Dictionary<string, SubmissionFileInfo>(StringComparer.Ordinal);

            foreach (string path in paths)
                map[path] = new SubmissionFileInfo(new string('a', 64), 1024);

            return map;
        }

        // -----------------------------------------------------------------------------------
        // What gets included.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Every_file_the_rows_reference_is_included()
        {
            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png"));

            Assert.True(result.IsComplete);
            Assert.Equal(2, result.Manifest.Files.Count);
            Assert.Contains(result.Manifest.Files, file => file.Path == "main.png");
            Assert.Contains(result.Manifest.Files, file => file.Path == "Images/U8-pin3.png");
        }

        // ###########################################################################################
        // *** THE BOARD'S KiCad DATA TRAVELS TOO (owner decision, 2026-09-26). *** No row cites the
        // "KiCad data" folder, so the rows-only list left it behind: a new system was published
        // without its traces. The caller collects the paths (SubmissionKiCadFiles.Collect) and each
        // is hashed and listed like any referenced file - and one that is missing on disk is a named
        // problem, not a silent omission.
        // ###########################################################################################
        [Fact]
        public void The_boards_KiCad_data_is_included_beside_the_referenced_files()
        {
            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png", "KiCad data/board.kicad_pcb"),
                kiCadFiles: ["KiCad data/board.kicad_pcb"]);

            Assert.True(result.IsComplete);
            Assert.Contains(result.Manifest.Files, file => file.Path == "KiCad data/board.kicad_pcb");
        }

        [Fact]
        public void A_KiCad_file_missing_on_disk_is_reported_by_name()
        {
            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png"),
                kiCadFiles: ["KiCad data/board.kicad_pcb"]);

            Assert.False(result.IsComplete);
            Assert.Contains(result.Problems, problem => problem.Contains("KiCad data/board.kicad_pcb"));
        }

        [Fact]
        public void A_file_the_rows_do_NOT_reference_is_never_sent()
        {
            // The contributor's folder may hold editor backups, thumbnail caches and personal
            // files. Only what the data names goes to a public server.
            Dictionary<string, SubmissionFileInfo> hashes =
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png", "my-notes.txt", "main.png.bak");

            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(), SubmissionManifestBuilderTests.Identity(), hashes);

            Assert.DoesNotContain(result.Manifest.Files, file => file.Path == "my-notes.txt");
            Assert.DoesNotContain(result.Manifest.Files, file => file.Path == "main.png.bak");
        }

        [Fact]
        public void Every_board_data_section_reaches_the_manifest()
        {
            // A section silently dropped here would delete that whole section on the server, since
            // the manifest is the complete intended state. Asserted section by section rather than
            // by a count, so adding a section to BoardData without adding it here is caught.
            var board = new BoardData
            {
                Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "main.png" } },
                Components = { new ComponentEntry { BoardLabel = "U8" } },
                ComponentImages = { new ComponentImageEntry { BoardLabel = "U8", File = "main.png" } },
                ComponentHighlights = { new ComponentHighlightEntry { SchematicName = "Main", BoardLabel = "U8" } },
                ComponentLocalFiles = { new ComponentLocalFileEntry { BoardLabel = "U8", File = "main.png" } },
                ComponentLinks = { new ComponentLinkEntry { BoardLabel = "U8", Url = "https://example.com" } },
                BoardLocalFiles = { new BoardLocalFileEntry { Name = "Manual", File = "main.png" } },
                BoardLinks = { new BoardLinkEntry { Name = "Forum", Url = "https://example.com" } },
                Credits = { new CreditEntry { NameOrHandle = "Dennis" } },
                KiCadImportantSignals = { new KiCadImportantSignalEntry { DisplayName = "VCC" } }
            };

            SubmissionRows rows = SubmissionManifestBuilder.Build(
                board, SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png")).Manifest.Rows;

            Assert.Single(rows.Schematics);
            Assert.Single(rows.Components);
            Assert.Single(rows.ComponentImages);
            Assert.Single(rows.ComponentHighlights);
            Assert.Single(rows.ComponentLocalFiles);
            Assert.Single(rows.ComponentLinks);
            Assert.Single(rows.BoardLocalFiles);
            Assert.Single(rows.BoardLinks);
            Assert.Single(rows.Credits);
            Assert.Single(rows.KiCadImportantSignals);
        }

        [Fact]
        public void A_file_referenced_by_several_rows_appears_once()
        {
            // A shared image referenced from three components is one file, and listing it three
            // times would make the upload estimate wrong and the server store it repeatedly.
            var board = new BoardData
            {
                Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "shared.png" } },
                ComponentImages =
                {
                    new ComponentImageEntry { BoardLabel = "U8", File = "shared.png" },
                    new ComponentImageEntry { BoardLabel = "U9", File = "shared.png" }
                }
            };

            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                board, SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("shared.png"));

            Assert.Single(result.Manifest.Files);
        }

        [Fact]
        public void Files_differing_only_by_CASE_are_kept_as_two()
        {
            // They are two files to the server. Collapsing them would silently drop one and leave
            // the row that named it pointing at nothing.
            var board = new BoardData
            {
                Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "U8.png" } },
                ComponentImages = { new ComponentImageEntry { BoardLabel = "U8", File = "u8.PNG" } }
            };

            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                board, SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("U8.png", "u8.PNG"));

            Assert.Equal(2, result.Manifest.Files.Count);
        }

        [Fact]
        public void The_file_list_is_in_a_stable_order()
        {
            // The same system must produce the same manifest twice - it makes two submissions
            // comparable and the negotiation step reproducible when something has to be diagnosed.
            IReadOnlyList<string> first =
                SubmissionManifestBuilder.CollectReferencedFiles(SubmissionManifestBuilderTests.Board());

            IReadOnlyList<string> second =
                SubmissionManifestBuilder.CollectReferencedFiles(SubmissionManifestBuilderTests.Board());

            Assert.Equal(first, second);
        }

        // -----------------------------------------------------------------------------------
        // Missing files.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_referenced_file_that_is_missing_is_REPORTED_and_names_itself()
        {
            // A manifest naming a file the server will never receive fails at finalise with a far
            // less helpful message, so it is caught here where the path is still in hand.
            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png"));

            Assert.False(result.IsComplete);

            string problem = Assert.Single(result.Problems);
            Assert.Contains("Images/U8-pin3.png", problem, StringComparison.Ordinal);
        }

        [Fact]
        public void A_missing_file_does_not_stop_the_others_being_listed()
        {
            // The contributor should see every problem at once, and the manifest should still show
            // what WOULD be sent.
            SubmissionManifestBuildResult result = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png"));

            Assert.Single(result.Manifest.Files);
            Assert.Equal("main.png", result.Manifest.Files[0].Path);
        }

        // -----------------------------------------------------------------------------------
        // Identity.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_identity_reaches_the_manifest_unchanged()
        {
            SubmissionManifest manifest = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png")).Manifest;

            Assert.Equal("Commodore/C64/250407", manifest.SystemId);
            Assert.Equal("Commodore", manifest.Manufacturer);
            Assert.Equal("C64", manifest.Hardware);
            Assert.Equal("250407", manifest.Board);
            Assert.Equal("2026-09-01", manifest.BaseRevision);
            Assert.Equal("Corrected R12.", manifest.Summary);

            // A draft of a published board carries no notes.
            Assert.Equal(string.Empty, manifest.HardwareNotes);
        }

        // A new system's notes from "Create system" travel with it, trimmed (owner request, 2026-10-05).
        [Fact]
        public void A_new_systems_notes_reach_the_manifest_trimmed()
        {
            SubmissionIdentity identity = SubmissionManifestBuilderTests.Identity();

            SubmissionManifest manifest = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                new SubmissionIdentity
                {
                    SystemId = identity.SystemId,
                    Manufacturer = identity.Manufacturer,
                    Hardware = identity.Hardware,
                    Board = identity.Board,
                    Summary = identity.Summary,
                    HardwareNotes = "  Open-source replica.\nRev. B only.  ",
                    CreatedUtc = identity.CreatedUtc
                },
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png")).Manifest;

            Assert.Equal("Open-source replica.\nRev. B only.", manifest.HardwareNotes);
        }

        [Fact]
        public void The_manifest_carries_the_CURRENT_format_version()
        {
            // A client that sent an old version number would be rejected by a server that has
            // moved on - so this must track the constant rather than being written by hand.
            SubmissionManifest manifest = SubmissionManifestBuilder.Build(
                SubmissionManifestBuilderTests.Board(),
                SubmissionManifestBuilderTests.Identity(),
                SubmissionManifestBuilderTests.Hashes("main.png", "Images/U8-pin3.png")).Manifest;

            Assert.Equal(SubmissionFormat.CurrentVersion, manifest.FormatVersion);
        }

        // -----------------------------------------------------------------------------------
        // The preview shown before anything uploads.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_preview_separates_new_files_from_ones_the_server_already_has()
        {
            // "2 new, 238 already on the server" is what makes a typo fix visibly cheap.
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    new SubmissionFile { Path = "a.png", Sha256 = new string('a', 64), SizeBytes = 100 },
                    new SubmissionFile { Path = "b.png", Sha256 = new string('b', 64), SizeBytes = 200 }
                }
            };

            SubmissionPreview preview = SubmissionManifestBuilder.Preview(manifest, [new string('a', 64)]);

            Assert.Equal(1, preview.NewFileCount);
            Assert.Equal(1, preview.AlreadyOnServer);
            Assert.Equal(200, preview.BytesToUpload);
        }

        [Fact]
        public void The_preview_counts_a_shared_blob_ONCE_towards_the_upload()
        {
            // One blob referenced from three paths uploads once. Counting it three times would
            // overstate the upload and make the estimate useless.
            string hash = new('c', 64);

            var manifest = new SubmissionManifest
            {
                Files =
                {
                    new SubmissionFile { Path = "a.png", Sha256 = hash, SizeBytes = 500 },
                    new SubmissionFile { Path = "b.png", Sha256 = hash, SizeBytes = 500 },
                    new SubmissionFile { Path = "c.png", Sha256 = hash, SizeBytes = 500 }
                }
            };

            SubmissionPreview preview = SubmissionManifestBuilder.Preview(manifest, []);

            Assert.Equal(1, preview.NewFileCount);
            Assert.Equal(500, preview.BytesToUpload);

            // ...but the manifest still lists all three paths, because all three exist.
            Assert.Equal(3, preview.FileCount);
        }

        [Fact]
        public void A_submission_needing_no_upload_reports_zero_bytes()
        {
            // The typo fix: a manifest and nothing else.
            string hash = new('d', 64);

            var manifest = new SubmissionManifest
            {
                Files = { new SubmissionFile { Path = "a.png", Sha256 = hash, SizeBytes = 999 } }
            };

            SubmissionPreview preview = SubmissionManifestBuilder.Preview(manifest, [hash]);

            Assert.Equal(0, preview.NewFileCount);
            Assert.Equal(0, preview.BytesToUpload);
        }

        [Theory]
        [InlineData(0, "0 bytes")]
        [InlineData(512, "512 bytes")]
        [InlineData(1536, "1.5 KB")]
        [InlineData(1048576, "1.0 MB")]
        [InlineData(5242880, "5.0 MB")]
        public void A_size_is_formatted_for_a_person_to_read(long bytes, string expected)
        {
            // Invariant culture throughout: the app has no localisation, so a machine with a comma
            // decimal separator would be the only inconsistency on screen.
            Assert.Equal(expected, SubmissionManifestBuilder.FormatSize(bytes));
        }
    }
}
