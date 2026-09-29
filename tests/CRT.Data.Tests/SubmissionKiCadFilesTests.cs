using System;
using System.Collections.Generic;
using System.IO;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionKiCadFiles - a board's KiCad data travelling in its submission (owner
    // decision, 2026-09-26). Before it, the "KiCad data" folder - the one part of a board no row
    // cites - silently stayed on the contributor's machine, and a new system was published to BETA
    // without its traces.
    // ###########################################################################################
    public sealed class SubmissionKiCadFilesTests : IDisposable
    {
        private const string SystemId = "Manu1/Hardware1/Board1";

        private readonly string thisRoot =
            Path.Combine(Path.GetTempPath(), "crt-kicad-files-" + Guid.NewGuid().ToString("N"));

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(this.thisRoot))
                    Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless; failing a test over cleanup is not.
            }
        }

        private string Folder(string name) => Path.Combine(this.thisRoot, name);

        private void Put(string root, string relativePath)
        {
            string path = Path.Combine(this.thisRoot, root, "KiCad data", relativePath.Replace('/', Path.DirectorySeparatorChar));

            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "(kicad_pcb)");
        }

        private static SubmissionManifest Manifest() =>
            new() { SystemId = SystemId, Manufacturer = "Manu1", Hardware = "Hardware1", Board = "Board1" };

        // ---------------------------------------------------------------------- The rule

        // Inside the board's own "KiCad data" folder (sub-folders included, as a multi-sheet
        // project has), and of a type CRT reads.
        [Theory]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/board.kicad_pcb", true)]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/board.kicad_sch", true)]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/board.kicad_pro", true)]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/Pages/vic.kicad_sch", true)]
        [InlineData("Manu1/Hardware1/Board1/kicad DATA/board.kicad_pcb", true)]
        // The types the shipped trees still hold but CRT never reads - not what a submission carries.
        [InlineData("Manu1/Hardware1/Board1/KiCad data/KiCad-traces.json", false)]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/Issue 4.sch", false)]
        // The right type in the wrong place: outside the folder, in another board, in a shared folder.
        [InlineData("Manu1/Hardware1/Board1/board.kicad_pcb", false)]
        [InlineData("Commodore/C64/250407/KiCad data/board.kicad_pcb", false)]
        [InlineData("Manu1/Shared files/KiCad data/board.kicad_pcb", false)]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/board.png", false)]
        public void A_submissions_KiCad_file_is_its_own_boards_and_a_type_CRT_reads(string path, bool submittable)
        {
            Assert.Equal(submittable, SubmissionKiCadFiles.IsSubmittable(SubmissionKiCadFilesTests.Manifest(), path));
        }

        // The Maintainer tab's side of the same question, asked of a path alone.
        [Theory]
        [InlineData("Manu1/Hardware1/Board1/KiCad data/board.kicad_pcb", true)]
        [InlineData("Manu1/Hardware1/Board1/Sheet1.png", false)]
        // A cited ordinary file inside the folder is a normal file with a row, not KiCad data.
        [InlineData("Manu1/Hardware1/Board1/KiCad data/readme.txt", false)]
        [InlineData("", false)]
        public void A_KiCad_data_path_is_recognised_by_its_folder(string path, bool inside)
        {
            Assert.Equal(inside, SubmissionKiCadFiles.IsKiCadDataPath(path));
        }

        // ---------------------------------------------------------------------- The collection

        // ###########################################################################################
        // The union of the draft's folder and the synced official one - a draft over a published
        // board carries the whole folder, its own copy of a same-named file winning, the same
        // draft-first rule the upload's byte resolution applies.
        // ###########################################################################################
        [Fact]
        public void The_draft_and_the_official_folders_are_joined_with_the_draft_winning()
        {
            this.Put("draft", "board.kicad_pcb");
            this.Put("draft", "Pages/vic.kicad_sch");
            this.Put("official", "BOARD.kicad_pcb");
            this.Put("official", "old.kicad_sch");

            IReadOnlyList<string> collected = SubmissionKiCadFiles.Collect(
                SystemId, this.Folder("draft"), this.Folder("official"));

            // ###########################################################################################
            // "BOARD.kicad_pcb" is the draft's "board.kicad_pcb" under another case - one file to
            // every Windows and macOS machine. It travels ONCE, under the PUBLISHED spelling, so
            // the publish REPLACES the official file; sending the draft's spelling would write a
            // second file beside it on the case-sensitive server (code review, 2026-09-26). The
            // BYTES are still the draft's - SubmissionFileLocator resolves each path draft-first.
            // ###########################################################################################
            Assert.Equal(
                [
                    "Manu1/Hardware1/Board1/KiCad data/BOARD.kicad_pcb",
                    "Manu1/Hardware1/Board1/KiCad data/Pages/vic.kicad_sch",
                    "Manu1/Hardware1/Board1/KiCad data/old.kicad_sch"
                ],
                collected);
        }

        // Only what CRT reads travels: the generated traces report and a stray image stay behind.
        [Fact]
        public void Files_of_types_CRT_does_not_read_are_left_behind()
        {
            this.Put("draft", "board.kicad_pcb");

            string folder = Path.Combine(this.Folder("draft"), "KiCad data");
            File.WriteAllText(Path.Combine(folder, "KiCad-traces.json"), "{}");
            File.WriteAllText(Path.Combine(folder, "photo.png"), "not read");

            Assert.Equal(
                ["Manu1/Hardware1/Board1/KiCad data/board.kicad_pcb"],
                SubmissionKiCadFiles.Collect(SystemId, this.Folder("draft"), officialSystemFolder: null));
        }

        // Most boards have no KiCad data at all, and a draft-only system has no official folder.
        [Fact]
        public void Missing_folders_contribute_nothing()
        {
            Assert.Empty(SubmissionKiCadFiles.Collect(SystemId, this.Folder("nowhere"), this.Folder("also-nowhere")));
            Assert.Empty(SubmissionKiCadFiles.Collect(SystemId, null, null));
            Assert.Empty(SubmissionKiCadFiles.Collect(string.Empty, this.Folder("draft"), null));
        }
    }
}
