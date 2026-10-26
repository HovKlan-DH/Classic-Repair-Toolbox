using System;
using System.Collections.Generic;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionKiCadFiles - a board's KiCad data travelling in its submission (owner
    // decision, 2026-09-26). Before it, the "KiCad data" folder - the one part of a board no row
    // cites - silently stayed on the contributor's machine, and a new board was published to BETA
    // without its traces. In the CRT 4.0 tree (2026-10-25) a board names its KiCad files itself,
    // "KiCad data/<...>", in its list of files.
    // ###########################################################################################
    public sealed class SubmissionKiCadFilesTests
    {
        private static readonly string Sha = new('a', 64);

        private static SubmissionManifest Manifest() =>
            new() { BoardId = "Manu1/Hardware1/Board1", Manufacturer = "Manu1", Hardware = "Hardware1", Board = "Board1" };

        // ---------------------------------------------------------------------- The rule

        // Under the board's "KiCad data" (sub-folders included, as a multi-sheet project has), and
        // of a type CRT reads.
        [Theory]
        [InlineData("KiCad data/board.kicad_pcb", true)]
        [InlineData("KiCad data/board.kicad_sch", true)]
        [InlineData("KiCad data/board.kicad_pro", true)]
        [InlineData("KiCad data/Pages/vic.kicad_sch", true)]
        [InlineData("kicad DATA/board.kicad_pcb", true)]
        // The types the converted trees still hold but CRT never reads - not new in a submission.
        [InlineData("KiCad data/KiCad-traces.json", false)]
        [InlineData("KiCad data/Issue 4.sch", false)]
        // The right type in the wrong place, or a folder name alone.
        [InlineData("board.kicad_pcb", false)]
        [InlineData("Images/KiCad data/board.kicad_pcb", false)]
        [InlineData("KiCad data/", false)]
        [InlineData("KiCad data/board.png", false)]
        public void A_submissions_KiCad_file_is_in_its_KiCad_folder_and_a_type_CRT_reads(string name, bool submittable)
        {
            Assert.Equal(submittable, SubmissionKiCadFiles.IsSubmittable(SubmissionKiCadFilesTests.Manifest(), name));
        }

        // A file nothing reads that the published board already has goes along as it is - and a
        // changed one, or a new one, does not.
        [Theory]
        [InlineData("KiCad data/KiCad-traces.json", true, true)]
        [InlineData("KiCad data/KiCad-traces.json", false, false)]
        [InlineData("Images/KiCad-traces.json", true, false)]
        public void A_straggler_of_the_KiCad_folder_travels_only_as_published(string name, bool samePublished, bool submittable)
        {
            string published = samePublished ? SubmissionKiCadFilesTests.Sha : new string('b', 64);

            Assert.Equal(submittable, SubmissionKiCadFiles.IsSubmittable(SubmissionKiCadFilesTests.Manifest(), name, published, SubmissionKiCadFilesTests.Sha));
        }

        // ###########################################################################################
        // A straggler is the KiCad folder's file of a type CRT does not read - which a submission
        // carries only as published, so the content check passes it (MatchesWholeFile) while the type
        // itself stays refused (Matches). A KiCad file CRT reads is checked as always.
        // ###########################################################################################
        [Theory]
        [InlineData("KiCad data/KiCad-traces.json", true)]
        [InlineData("KiCad data/Issue 4.sch", true)]
        [InlineData("kicad DATA/Pages/notes.txt", true)]
        [InlineData("KiCad data/board.kicad_pcb", false)]
        [InlineData("Images/KiCad-traces.json", false)]
        [InlineData("KiCad-traces.json", false)]
        [InlineData(null, false)]
        public void A_straggler_is_a_file_of_the_KiCad_folder_CRT_does_not_read(string? name, bool straggler)
        {
            Assert.Equal(straggler, SubmissionKiCadFiles.IsStraggler(name));
        }

        [Fact]
        public void A_straggler_passes_the_content_check_while_its_type_stays_refused_everywhere_else()
        {
            byte[] notJson = [0x4D, 0x5A, 0x90, 0x00];

            Assert.True(SubmissionContentRules.MatchesWholeFile("KiCad data/KiCad-traces.json", notJson, () => null, out _));
            Assert.False(SubmissionContentRules.Matches("KiCad data/KiCad-traces.json", notJson, out _));
            Assert.False(SubmissionContentRules.MatchesWholeFile("Images/KiCad-traces.json", notJson, () => null, out _));
            Assert.False(SubmissionContentRules.MatchesWholeFile("KiCad data/board.kicad_pcb", notJson, () => null, out _));
        }

        // The Maintainer tab's side of the same question, asked of a name alone.
        [Theory]
        [InlineData("KiCad data/board.kicad_pcb", true)]
        [InlineData("Sheet1.png", false)]
        // A cited ordinary file inside the folder is a normal file with a row, not KiCad data.
        [InlineData("KiCad data/readme.txt", false)]
        [InlineData("", false)]
        public void A_KiCad_data_path_is_recognised_by_its_folder(string name, bool inside)
        {
            Assert.Equal(inside, SubmissionKiCadFiles.IsKiCadDataPath(name));
        }

        // ---------------------------------------------------------------------- The collection

        // ###########################################################################################
        // Every name the draft's board file lists under "KiCad data/" - the published board's KiCad
        // files as well as the draft's own, since a draft lists the whole board - sorted, once each.
        // ###########################################################################################
        [Fact]
        public void Every_KiCad_name_the_board_file_lists_is_collected_in_order()
        {
            var board = new BoardDocument
            {
                Files =
                [
                    new BoardDocument.StoredFileRow { Name = "KiCad data/Pages/vic.kicad_sch", Sha256 = SubmissionKiCadFilesTests.Sha },
                    new BoardDocument.StoredFileRow { Name = "Images/board.png", Sha256 = SubmissionKiCadFilesTests.Sha },
                    new BoardDocument.StoredFileRow { Name = "KiCad data/board.kicad_pcb", Sha256 = SubmissionKiCadFilesTests.Sha },
                    new BoardDocument.StoredFileRow { Name = "KiCad data/KiCad-traces.json", Sha256 = SubmissionKiCadFilesTests.Sha },
                ]
            };

            IReadOnlyList<string> collected = SubmissionKiCadFiles.Collect(board);

            // The straggler travels too: the server keeps it only if it is the published one's bytes.
            Assert.Equal(
                ["KiCad data/KiCad-traces.json", "KiCad data/Pages/vic.kicad_sch", "KiCad data/board.kicad_pcb"],
                collected);
        }

        // Most boards have no KiCad data at all, and a board with no board file has nothing.
        [Fact]
        public void A_board_with_no_KiCad_files_or_no_board_file_contributes_nothing()
        {
            Assert.Empty(SubmissionKiCadFiles.Collect(new BoardDocument()));
            Assert.Empty(SubmissionKiCadFiles.Collect(null));
        }

        // ###########################################################################################
        // A project imported into the draft and not yet listed by a write goes too (code review,
        // 2026-10-25): the Schematics tab draws it and the draft's fingerprint covers it, so leaving
        // it out sent a submission without it and then greyed Submit out as already sent. Once each,
        // and only names inside "KiCad data".
        // ###########################################################################################
        [Fact]
        public void KiCad_files_lying_in_the_draft_not_yet_listed_are_collected_too()
        {
            var board = new BoardDocument
            {
                Files = [new BoardDocument.StoredFileRow { Name = "KiCad data/board.kicad_pcb", Sha256 = SubmissionKiCadFilesTests.Sha }]
            };

            IReadOnlyList<string> collected = SubmissionKiCadFiles.Collect(
                board,
                ["KiCad data/Pages/cpu.kicad_sch", "KiCad data/board.kicad_pcb", "Images/board.png"]);

            Assert.Equal(["KiCad data/Pages/cpu.kicad_sch", "KiCad data/board.kicad_pcb"], collected);
            Assert.Equal(["KiCad data/Pages/cpu.kicad_sch"], SubmissionKiCadFiles.Collect(null, ["KiCad data/Pages/cpu.kicad_sch"]));
        }
    }
}
