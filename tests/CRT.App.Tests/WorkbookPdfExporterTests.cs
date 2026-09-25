using Handlers.DataHandling;
using QuestPDF.Drawing;

namespace ClassicRepairToolbox.Tests;

// The QuestPDF CONFIGURATION the workbook export stands on - not its layout, which stays untested
// for the reason WorkbookPdfExporter's header gives (asserting on PDF bytes tests QuestPDF rather
// than this app). The archive half has its own file, WorkbookZipExportTests.
//
// WHY THIS EXISTS: QuestPDF 2026.9.0 changed its font defaults in one release. It stopped using
// the system's fonts, and it started REFUSING to generate a document when a glyph or a font family
// is missing, where before it drew a blank and carried on. A workbook export prints repair notes
// exactly as the user typed them - an emoji, a CJK part marking, an arrow - so under the new
// defaults one such character in one note fails the whole export. Nothing in the layout tests
// would notice: they all use plain ASCII titles.
//
// It also removed registering a font under a name of the app's own choosing, so the icon font is
// now reached by the family name stored inside the .otf - a name the code has to get right, with
// nothing but a missing icon to show when it does not.
public class WorkbookPdfExporterTests : IDisposable
{
    private const string SolidFontPath = "Assets/Fonts/Font Awesome 7 Free-Solid-900.otf";

    private readonly TempWorkspace thisWorkspace = new();

    // The app's own configuration, not a copy of it - a test setting only the licence ran the
    // exporter under QuestPDF's new defaults, which is how the first test below was seen to fail.
    public WorkbookPdfExporterTests() => WorkbookPdfExporter.ConfigureQuestPdf();

    public void Dispose() => this.thisWorkspace.Dispose();

    // ###########################################################################################
    // THE REGRESSION TEST for the QuestPDF 2026.9.0 upgrade: text no available font can draw must
    // cost a glyph, never the export.
    //
    // The characters are chosen to be missing from QuestPDF's bundled Lato on every platform CI
    // runs on: an emoji, CJK, and a Private Use Area code point that no text font assigns.
    // ###########################################################################################
    [Fact]
    public void A_worklog_holding_characters_no_bundled_font_can_draw_still_exports()
    {
        var workbook = new WorkbookRecord
        {
            Id = 1,
            BoardKey = "C64|250469",
            Title = "Repair 中文",
            Note = "Smells burnt \U0001F525",
            Status = "Open",
            StartDate = new DateTime(2026, 9, 3)
        };

        var entry = new WorklogEntryRecord
        {
            Id = 1,
            SchematicName = "Sheet 1",
            Title = "U8 marked 中文 ",
            Description = "Replaced it \U0001F525 and the fault moved → U9",
            Category = "Issue",
            State = "Open"
        };

        var document = WorkbookExportModel.Build(
            workbook,
            new[] { entry },
            null,
            entryId => Path.Combine(this.thisWorkspace.Root, $"worklog_{entryId}"),
            new DateTime(2026, 9, 4),
            "DKK");

        string path = Path.Combine(this.thisWorkspace.Root, "export.pdf");

        WorkbookPdfExporter.WritePdf(document, path);

        Assert.True(new FileInfo(path).Length > 0);
    }

    // ###########################################################################################
    // The icon font's family name is the one QuestPDF reads out of the .otf, and its face is the
    // weight-900 Solid one every glyph site asks for with .Black().
    //
    // Registers the same file the exporter reads through Avalonia's AssetLoader (the csproj links
    // it in as that resource), since AssetLoader itself is not available to a plain test.
    // ###########################################################################################
    [Fact]
    public void The_icon_font_family_the_export_draws_with_is_the_one_its_font_file_declares()
    {
        FontManager.RegisterFontFromBinaryData(File.ReadAllBytes(ResolveAssetPath(SolidFontPath)));

        var registered = FontManager.GetRegisteredFonts().ToList();
        var iconFaces = registered
            .Where(face => string.Equals(face.FamilyName, WorkbookPdfExporter.IconFontFamily, StringComparison.OrdinalIgnoreCase))
            .ToList();

        Assert.True(
            iconFaces.Count > 0,
            $"No registered face is in family [{WorkbookPdfExporter.IconFontFamily}]. Registered: " +
            string.Join(", ", registered.Select(face => $"[{face.FamilyName}] {face.Weight}")));

        Assert.Contains(iconFaces, face => Convert.ToInt32(face.Weight) == 900);
    }

    // ###########################################################################################
    // An export with the icon font loaded really draws its icons IN that font - the whole path, from
    // the cached bytes through registration to the family and weight each glyph site asks for.
    //
    // The only evidence is which font the PDF embeds, so that is what this reads: with
    // ThrowOnMissingFontFamilies off (see ConfigureQuestPdf), a glyph site naming a family nobody
    // registered does not throw - it falls back to Lato and prints a blank box where the padlock
    // should be. This exporter has shipped without its icons once already, unnoticed until the PDF
    // was opened. Fails if IconFontFamily is set back to the old private "CRT Export Icons".
    // ###########################################################################################
    [Fact]
    public void An_export_with_the_icon_font_loaded_embeds_that_font_for_its_icons()
    {
        WorkbookPdfExporter.SetIconFontBytesForTests(File.ReadAllBytes(ResolveAssetPath(SolidFontPath)));

        try
        {
            var workbook = new WorkbookRecord
            {
                Id = 1,
                BoardKey = "C64|250469",
                Title = "Repair",
                Status = "Open",
                StartDate = new DateTime(2026, 9, 3)
            };

            // Open and Issue both carry an icon on their pills.
            var entry = new WorklogEntryRecord
            {
                Id = 1,
                SchematicName = "Sheet 1",
                Title = "Bad cap",
                Category = "Issue",
                State = "Open"
            };

            var document = WorkbookExportModel.Build(
                workbook,
                new[] { entry },
                null,
                entryId => Path.Combine(this.thisWorkspace.Root, $"worklog_{entryId}"),
                new DateTime(2026, 9, 4),
                "DKK");

            string path = Path.Combine(this.thisWorkspace.Root, "export.pdf");
            WorkbookPdfExporter.WritePdf(document, path);

            // Font names sit in the PDF's font dictionaries as plain text, subset-prefixed
            // ("ABCDEF+FontAwesome7Free-Solid"); Latin1 maps every byte to one char, so the search
            // cannot be thrown off by the binary streams around them.
            string pdf = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(path));

            Assert.Contains("FontAwesome7Free-Solid", pdf);
        }
        finally
        {
            WorkbookPdfExporter.SetIconFontBytesForTests(null);
        }
    }

    // Walks up from the test binary to the folder holding the asset - the same approach as
    // FontAwesomeAssetTests, so it works in every configuration and on CI.
    private static string ResolveAssetPath(string relativePath)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory != null)
        {
            string candidate = Path.Combine(directory.FullName, relativePath);
            if (File.Exists(candidate))
            {
                return candidate;
            }
            directory = directory.Parent;
        }

        throw new FileNotFoundException($"Could not find {relativePath} above {AppContext.BaseDirectory}");
    }
}
