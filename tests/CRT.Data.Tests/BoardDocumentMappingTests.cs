using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Between the Excel-era board (BoardData and the sidecar) and the CRT 4.0 board file. Every value
// must come back as it went in - checked by reflection, so a column added to BoardData later
// fails here until the mapping carries it - apart from the two deliberate changes: the scope
// baselines split from the pictures, and the files named by the board.
// ###########################################################################################
public sealed class BoardDocumentMappingTests
{
    private static readonly string[] FileProperties =
    [
        nameof(BoardSchematicEntry.SchematicImageFile),
        nameof(ComponentImageEntry.File),
    ];

    // Every string property of a row filled with a value naming the row and the property, so
    // nothing can pass by being blank on both sides.
    private static T Filled<T>(string tag) where T : new()
    {
        var row = new T();

        foreach (PropertyInfo property in typeof(T).GetProperties().Where(p => p.PropertyType == typeof(string) && p.CanWrite))
            property.SetValue(row, $"{tag}.{property.Name}");

        return row;
    }

    private static ComponentImageEntry Image(string tag, bool baseline) => new()
    {
        BoardLabel = $"{tag}.label",
        Region = $"{tag}.region",
        Pin = baseline ? "1" : string.Empty,
        Name = $"{tag}.name",
        ExpectedOscilloscopeReading = $"{tag}.expected",
        File = $"{tag}.file",
        Note = $"{tag}.note",
        TimeDiv = baseline ? "1uS" : string.Empty,
        VoltsDiv = baseline ? "2V" : string.Empty,
        TriggerLevelVolts = baseline ? "1.4V" : string.Empty,

        // CRT 4.0's conditions (2026-10-25) - on pictures too: one typed in the table on a picture
        // row must not be lost on the way through the file.
        MachineState = $"{tag}.machine state",
        Source = $"{tag}.source",
    };

    private static BoardContent Content()
    {
        var board = new BoardData
        {
            RevisionDate = "2026-October-4",
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Schematics = [Filled<BoardSchematicEntry>("schematic")],
            Components = [Filled<ComponentEntry>("component")],
            ComponentImages = [Image("picture 1", false), Image("baseline 1", true), Image("picture 2", false), Image("baseline 2", true)],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "S", BoardLabel = "U1", X = "10", Y = "-20", Width = "30", Height = "40" }],
            ComponentLocalFiles = [Filled<ComponentLocalFileEntry>("component file")],
            ComponentLinks = [Filled<ComponentLinkEntry>("component link")],
            BoardLocalFiles = [Filled<BoardLocalFileEntry>("board file")],
            BoardLinks = [Filled<BoardLinkEntry>("board link")],
            Credits = [Filled<CreditEntry>("credit")],
            KiCadImportantSignals = [Filled<KiCadImportantSignalEntry>("signal")],
            ComponentPins = [Filled<ComponentPinEntry>("pin")],
            ComponentReplacements = [Filled<ComponentReplacementEntry>("replacement")],
        };

        return new BoardContent(
            board,
            [new KiCadCalibrationEntry { SchematicName = "S", CadName = "C", OffsetX = 1.25, OffsetY = 2.5, ScaleX = 0.75, ScaleY = 0.5, MirrorX = true, MirrorY = false }],
            [new SchematicRenameRecord("Old", "New", "rename.file", "2026-October-10")]);
    }

    // The board's own name for a cited path, in these tests: "name of <path>".
    private static string NameOf(string path) => "name of " + path;

    private static BoardContent RoundTrip(BoardContent content) =>
        BoardDocumentMapping.ToContent(
            DataTreeJson.ReadBoard(DataTreeJson.Write(BoardDocumentMapping.ToDocument(content, NameOf, []))));

    [Fact]
    public void The_scope_baselines_are_split_from_the_pictures_each_in_its_order()
    {
        BoardDocument document = BoardDocumentMapping.ToDocument(Content(), NameOf, []);

        Assert.Equal(["picture 1.label", "picture 2.label"], document.ComponentImages.Select(row => row.BoardLabel));
        Assert.Equal(["baseline 1.label", "baseline 2.label"], document.ScopeBaselines.Select(row => row.BoardLabel));
        Assert.Equal("1.4V", document.ScopeBaselines[0].TriggerLevel);
    }

    // Back again the pictures come first and the baselines after - the popup shows them apart.
    [Fact]
    public void Back_again_the_pictures_come_first_and_the_baselines_after()
    {
        BoardContent after = RoundTrip(Content());

        Assert.Equal(
            ["picture 1.label", "picture 2.label", "baseline 1.label", "baseline 2.label"],
            after.Board.ComponentImages.Select(row => row.BoardLabel));
    }

    [Fact]
    public void Every_value_of_every_row_comes_back()
    {
        BoardContent before = Content();
        BoardContent after = RoundTrip(before);

        Assert.Equal(before.Board.RevisionDate, after.Board.RevisionDate);
        Assert.Equal(before.Board.HardwareName, after.Board.HardwareName);
        Assert.Equal(before.Board.BoardName, after.Board.BoardName);

        SameRows(before.Board.Schematics, after.Board.Schematics);
        SameRows(before.Board.Components, after.Board.Components);
        SameRows(before.Board.ComponentHighlights, after.Board.ComponentHighlights);
        SameRows(before.Board.ComponentLocalFiles, after.Board.ComponentLocalFiles);
        SameRows(before.Board.ComponentLinks, after.Board.ComponentLinks);
        SameRows(before.Board.BoardLocalFiles, after.Board.BoardLocalFiles);
        SameRows(before.Board.BoardLinks, after.Board.BoardLinks);
        SameRows(before.Board.Credits, after.Board.Credits);
        SameRows(before.Board.KiCadImportantSignals, after.Board.KiCadImportantSignals);
        SameRows(before.Board.ComponentPins, after.Board.ComponentPins);
        SameRows(before.Board.ComponentReplacements, after.Board.ComponentReplacements);
        SameRows(
            before.Board.ComponentImages.Where(row => !ComponentImageKinds.IsScopeBaseline(row)).Concat(before.Board.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline)).ToList(),
            after.Board.ComponentImages);
        SameRows(before.Calibrations.ToList(), after.Calibrations.ToList());
        SameRows(before.Renames.ToList(), after.Renames.ToList(), fileProperty: nameof(SchematicRenameRecord.File));
    }

    // A new property on a BoardData row type must be carried by the mapping, or every board
    // loses that column in the conversion.
    [Fact]
    public void Every_row_property_of_BoardData_is_a_column_the_mapping_knows()
    {
        var known = new Dictionary<Type, int>
        {
            [typeof(BoardSchematicEntry)] = 8,
            [typeof(ComponentEntry)] = 9,
            [typeof(ComponentImageEntry)] = 12,
            [typeof(ComponentPinEntry)] = 5,
            [typeof(ComponentReplacementEntry)] = 6,
            [typeof(ComponentHighlightEntry)] = 6,
            [typeof(ComponentLocalFileEntry)] = 3,
            [typeof(ComponentLinkEntry)] = 3,
            [typeof(BoardLocalFileEntry)] = 3,
            [typeof(BoardLinkEntry)] = 3,
            [typeof(CreditEntry)] = 4,
            [typeof(KiCadImportantSignalEntry)] = 2,
            [typeof(KiCadCalibrationEntry)] = 8,
            [typeof(BoardData)] = 15,
        };

        foreach ((Type type, int count) in known)
            Assert.True(type.GetProperties().Length == count, $"{type.Name} has {type.GetProperties().Length} properties; the mapping carries {count} - map the new one in BoardDocumentMapping and BoardDocument.");
    }

    [Fact]
    public void A_row_cites_its_file_by_the_name_the_board_gives_it()
    {
        BoardDocument document = BoardDocumentMapping.ToDocument(Content(), NameOf, []);

        Assert.Equal("name of schematic.SchematicImageFile", document.Schematics[0].ImageFile);
        Assert.Equal("name of picture 1.file", document.ComponentImages[0].File);
        Assert.Equal("name of baseline 1.file", document.ScopeBaselines[0].File);
        Assert.Equal("name of component file.File", document.ComponentFiles[0].File);
        Assert.Equal("name of board file.File", document.BoardFiles[0].File);
        Assert.Equal("name of rename.file", document.SchematicRenames[0].File);
    }

    // A note-only image row cites no file - and the naming is never asked about nothing.
    [Fact]
    public void A_row_citing_no_file_names_none()
    {
        BoardContent content = Content();
        content.Board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U9", Name = "Pinout", Note = "Same as 6526" });

        var asked = new List<string>();
        BoardDocument document = BoardDocumentMapping.ToDocument(content, path => { asked.Add(path); return NameOf(path); }, []);

        Assert.Equal(string.Empty, document.ComponentImages.Single(row => row.BoardLabel == "U9").File);
        Assert.DoesNotContain(asked, string.IsNullOrWhiteSpace);
    }

    [Fact]
    public void The_board_files_go_into_the_document_as_given()
    {
        BoardDocument.StoredFileRow file = new() { Name = "Board.png", Sha256 = new string('a', 64), Size = 7 };

        BoardDocument document = BoardDocumentMapping.ToDocument(Content(), NameOf, [file]);

        Assert.Same(file, Assert.Single(document.Files));
    }

    // The label editor works in real geometry and hands over a rectangle drawn at any zoom as
    // "123.4567" (LabelEditorBoardWriter). The board file holds whole pixels, rounded as the
    // Excel-era sidecar writer and the app's own highlight store always rounded them - away from
    // zero. Refused, as it was until 2026-10-26, every rectangle drawn or moved in a 4.0 draft
    // failed to save.
    [Theory]
    [InlineData("123.4567", 123)]
    [InlineData("12.5", 13)]
    [InlineData("-12.5", -13)]
    [InlineData("7", 7)]
    [InlineData("1E2", 100)]
    public void A_fractional_highlight_coordinate_is_rounded_to_whole_pixels(string coordinate, int expected)
    {
        BoardContent content = Content();
        content.Board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U7", X = coordinate, Y = coordinate, Width = coordinate, Height = coordinate });

        BoardDocument.HighlightRow row = BoardDocumentMapping.ToDocument(content, NameOf, []).Highlights.Single(highlight => highlight.BoardLabel == "U7");

        Assert.Equal([expected, expected, expected, expected], new[] { row.X, row.Y, row.Width, row.Height });
    }

    // What is no number at all is never quietly a zero - the board's top-left corner is a real
    // place - so the write stops and says which highlight it was.
    [Theory]
    [InlineData("twelve")]
    [InlineData("")]
    [InlineData("12,5")]
    [InlineData("NaN")]
    [InlineData("1E12")]
    public void A_highlight_coordinate_that_is_no_number_is_refused_naming_it(string coordinate)
    {
        BoardContent content = Content();
        content.Board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U7", X = coordinate, Y = "1", Width = "1", Height = "1" });

        FormatException refused = Assert.Throws<FormatException>(() => BoardDocumentMapping.ToDocument(content, NameOf, []));

        Assert.Contains("[U7]", refused.Message);
        Assert.Contains("[Top]", refused.Message);
    }

    // ---- helpers ---------------------------------------------------------------------------

    // Equal property by property; `fileProperty` (and the row types' own file columns) compared as
    // "the name the board gave that path".
    private static void SameRows<T>(List<T> before, List<T> after, string? fileProperty = null)
    {
        Assert.Equal(before.Count, after.Count);

        foreach (PropertyInfo property in typeof(T).GetProperties())
        {
            for (int i = 0; i < before.Count; i++)
            {
                object? a = property.GetValue(before[i]);
                object? b = property.GetValue(after[i]);

                bool isFile = property.Name == fileProperty || FileProperties.Contains(property.Name);

                if (isFile && a is string path && path.Length > 0)
                    Assert.Equal(NameOf(path), b);
                else
                    Assert.True(Equals(a, b), $"{typeof(T).Name}.{property.Name}: [{a}] became [{b}]");
            }
        }
    }
}
