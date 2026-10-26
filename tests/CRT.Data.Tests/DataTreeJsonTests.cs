using System.IO;
using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The CRT 4.0 board and main files on disk. What matters, in order: the same content is always
// the same bytes (the manifest, the sync and git compare bytes); what a later CRT adds survives
// this one; and a file this CRT cannot read correctly is refused, never read as empty.
// ###########################################################################################
public sealed class DataTreeJsonTests
{
    private const string Sha = "3f7a0c19b2d44e8f9a6b5c4d3e2f1a0b9c8d7e6f5a4b3c2d1e0f9a8b7c6d5e4f";

    // A board touching every list, with text that needs care: quotes, a backslash, a micro sign,
    // Danish letters, and a calibration with all its decimals.
    private static BoardDocument Board() => new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        RevisionDate = "2026-October-4",
        Schematics = [new() { Name = "Board layout", ImageFile = "Board layout.png", HighlightColor = "#FF0000", HighlightOpacity = "0.3", CadName = "250407_ - PCB Top" }],
        Components = [new() { BoardLabel = "C15", FriendlyName = "Capacitor", TechnicalNameOrValue = "10µF \"tantal\" \\ 16V", Category = "Capacitor", Description = "Glat og flot - æøå" }],
        ComponentImages = [new() { BoardLabel = "U1", Name = "Pinout", File = "Component images/6510.png", Note = "Top view" }],
        ScopeBaselines = [new() { BoardLabel = "U1", Pin = "1", File = "Scope baseline/U1_1_PAL.png", TimeDiv = "1uS", VoltsDiv = "2V", TriggerLevel = "1.4V" }],
        ComponentFiles = [new() { BoardLabel = "U1", Name = "Datasheet", File = "Component local files/6510.pdf" }],
        ComponentLinks = [new() { BoardLabel = "U1", Name = "Wiki", Url = "https://example.org/6510" }],
        BoardFiles = [new() { Category = "Service manuals", Name = "Service manual", File = "Board local files/manual.pdf" }],
        BoardLinks = [new() { Category = "Videos", Name = "Repair", Url = "https://example.org/v" }],
        Credits = [new() { Category = "Schematics", SubCategory = "Redrawn", NameOrHandle = "Anna", Contact = "anna@example.org" }],
        ImportantSignals = [new() { DisplayName = "+5VDC", KiCadNetName = "+5V" }],
        Highlights = [new() { Schematic = "Board layout", BoardLabel = "C15", X = 1448, Y = -3, Width = 102, Height = 36 }],
        KiCadCalibrations = [new() { Schematic = "Board layout", CadName = "250407_ - PCB Top", OffsetX = 50.87193460490473, OffsetY = 451.9972752043597, ScaleX = 0.9769080633767282, ScaleY = 0.812766517468314, MirrorX = true }],
        SchematicRenames = [new() { From = "Top (replica)", To = "PCB top side", File = "Board layout.png", Date = "2026-October-10" }],
        Files = [new() { Name = "Board layout.png", Sha256 = Sha, Size = 123456 }],
    };

    [Fact]
    public void The_same_board_is_always_the_same_bytes()
    {
        Assert.Equal(DataTreeJson.Write(Board()), DataTreeJson.Write(Board()));
    }

    [Fact]
    public void A_board_read_back_writes_the_same_bytes_again()
    {
        string written = DataTreeJson.Write(Board());

        Assert.Equal(written, DataTreeJson.Write(DataTreeJson.ReadBoard(written)));
    }

    // Every value survives - the calibration's decimals included, which is where a sloppy number
    // format would move a KiCad overlay.
    [Fact]
    public void Every_value_survives_a_write_and_a_read()
    {
        BoardDocument read = DataTreeJson.ReadBoard(DataTreeJson.Write(Board()));

        Assert.Equal("10µF \"tantal\" \\ 16V", read.Components[0].TechnicalNameOrValue);
        Assert.Equal("Glat og flot - æøå", read.Components[0].Description);
        Assert.Equal(50.87193460490473, read.KiCadCalibrations[0].OffsetX);
        Assert.Equal(0.812766517468314, read.KiCadCalibrations[0].ScaleY);
        Assert.True(read.KiCadCalibrations[0].MirrorX);
        Assert.Equal(-3, read.Highlights[0].Y);
        Assert.Equal(123456, read.Files[0].Size);
        Assert.Equal("1.4V", read.ScopeBaselines[0].TriggerLevel);
        Assert.Equal("2026-October-10", read.SchematicRenames[0].Date);
    }

    // One row per line: a changed row is one changed line in a diff.
    [Fact]
    public void Each_row_is_one_line_and_lines_end_in_a_bare_newline()
    {
        string written = DataTreeJson.Write(Board());
        string[] lines = written.Split('\n');

        Assert.DoesNotContain('\r', written);
        Assert.EndsWith("}\n", written);
        Assert.Contains(lines, line => line.StartsWith("    {\"boardLabel\":\"C15\",", StringComparison.Ordinal) && line.EndsWith("}", StringComparison.Ordinal));
        Assert.Contains("  \"components\": [", lines);
        Assert.Contains("  \"credits\": [", lines);
    }

    // Text is written as it is - no \u escapes for a reader of the file to puzzle over.
    [Fact]
    public void Text_is_written_as_it_is()
    {
        string written = DataTreeJson.Write(Board());

        Assert.Contains("10µF", written);
        Assert.Contains("æøå", written);
        Assert.DoesNotContain("\\u00", written);
    }

    [Fact]
    public void An_empty_list_is_written_as_an_empty_list()
    {
        string written = DataTreeJson.Write(new BoardDocument());

        Assert.Contains("  \"highlights\": [],", written.Split('\n'));
        Assert.Contains("  \"files\": []", written.Split('\n'));
        Assert.Empty(DataTreeJson.ReadBoard(written).Files);
    }

    // A picture writes no scope setting it does not have - and keeps one it has.
    [Fact]
    public void A_picture_writes_only_the_scope_settings_it_has()
    {
        BoardDocument board = Board();
        board.ComponentImages.Add(new BoardDocument.ComponentImageRow { BoardLabel = "U2", Name = "Odd", TimeDiv = "5mS" });

        string written = DataTreeJson.Write(board);
        BoardDocument read = DataTreeJson.ReadBoard(written);

        Assert.DoesNotContain(written.Split('\n'), line => line.Contains("\"6510.png\"", StringComparison.Ordinal) && line.Contains("timeDiv", StringComparison.Ordinal));
        Assert.Null(read.ComponentImages[0].TimeDiv);
        Assert.Equal("5mS", read.ComponentImages[1].TimeDiv);
    }

    // ###########################################################################################
    // *** CRT 4.0's NEW FIELDS ARE WRITTEN ONLY WHEN FILLED (step 2, 2026-10-25). *** A board that
    // has none - every board converted from Excel - keeps its component and baseline rows line for
    // line, so publishing it again does not rewrite every row of the file; one that has them keeps
    // them.
    // ###########################################################################################
    [Fact]
    public void A_component_or_baseline_writes_a_package_notes_state_or_source_only_when_it_has_one()
    {
        string plain = DataTreeJson.Write(Board());

        foreach (string field in new[] { "package", "repairNotes", "machineState", "source" })
            Assert.DoesNotContain($"\"{field}\":", plain, StringComparison.Ordinal);

        BoardDocument board = Board();
        board.Components[0].Package = "DIP-40";
        board.Components[0].RepairNotes = "Socket it.\nSee https://example.org";
        board.ScopeBaselines[0].MachineState = "READY prompt";
        board.ScopeBaselines[0].Source = "Known-good board";

        BoardDocument read = DataTreeJson.ReadBoard(DataTreeJson.Write(board));

        Assert.Equal("DIP-40", read.Components[0].Package);
        Assert.Equal("Socket it.\nSee https://example.org", read.Components[0].RepairNotes);
        Assert.Equal("READY prompt", read.ScopeBaselines[0].MachineState);
        Assert.Equal("Known-good board", read.ScopeBaselines[0].Source);
    }

    // The two new lists are written and read like every other; a file written before them reads
    // them as empty.
    [Fact]
    public void Pins_and_replacements_survive_and_a_file_without_them_reads_none()
    {
        BoardDocument board = Board();
        board.ComponentPins.Add(new BoardDocument.ComponentPinRow { BoardLabel = "U1", Pin = "40", Name = "VCC", Direction = "Power" });
        board.ComponentReplacements.Add(new BoardDocument.ComponentReplacementRow { BoardLabel = "U1", PartNumber = "8500", Kind = "Drop-in", Source = "Datasheet" });

        BoardDocument read = DataTreeJson.ReadBoard(DataTreeJson.Write(board));

        Assert.Equal("VCC", Assert.Single(read.ComponentPins).Name);
        Assert.Equal("8500", Assert.Single(read.ComponentReplacements).PartNumber);

        BoardDocument old = DataTreeJson.ReadBoard("{ \"formatVersion\": 1, \"components\": [] }");

        Assert.Empty(old.ComponentPins);
        Assert.Empty(old.ComponentReplacements);
    }

    // ---- what a later CRT adds ---------------------------------------------------------------

    // A field added by a later CRT - at the top, in a row, nested - is written back, unchanged.
    [Fact]
    public void Fields_this_CRT_does_not_know_are_kept()
    {
        string later = DataTreeJson.Write(Board())
            .Replace("  \"formatVersion\": 1,", "  \"formatVersion\": 1,\n  \"replacements\": [{\"partNumber\":\"ZTX650\",\"dropIn\":true}],")
            .Replace("{\"boardLabel\":\"C15\",", "{\"boardLabel\":\"C15\",\"pins\":{\"count\":2,\"names\":[\"+\",\"-\"]},");

        BoardDocument read = DataTreeJson.ReadBoard(later);
        string again = DataTreeJson.Write(read);

        Assert.Contains("\"replacements\": [{\"partNumber\":\"ZTX650\",\"dropIn\":true}]", again);
        Assert.Contains("\"pins\":{\"count\":2,\"names\":[\"+\",\"-\"]}", again);
        Assert.Equal(again, DataTreeJson.Write(DataTreeJson.ReadBoard(again)));
    }

    [Fact]
    public void A_file_from_a_newer_format_is_refused_saying_so()
    {
        string newer = DataTreeJson.Write(Board()).Replace("\"formatVersion\": 1", "\"formatVersion\": 2");

        InvalidDataException refused = Assert.Throws<InvalidDataException>(() => DataTreeJson.ReadBoard(newer));
        Assert.Contains("newer CRT", refused.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData("null")]
    [InlineData("[]")]
    [InlineData("{\"formatVersion\": 0}")]
    [InlineData("{\"components\": [null]}")]
    [InlineData("{\"components\": [{\"boardLabel\": 7}]}")]
    [InlineData("{\"highlights\": [{\"x\": \"left\"}]}")]
    [InlineData("{\"components\": [")]
    public void A_file_that_cannot_be_read_correctly_is_refused(string json)
    {
        Assert.Throws<InvalidDataException>(() => DataTreeJson.ReadBoard(json));
    }

    // A null text is a blank cell, and a missing list is an empty one.
    [Fact]
    public void A_null_text_reads_as_empty_and_a_missing_list_as_empty()
    {
        BoardDocument read = DataTreeJson.ReadBoard("{\"formatVersion\": 1, \"revisionDate\": null, \"components\": [{\"boardLabel\": \"U1\", \"region\": null}]}");

        Assert.Equal(string.Empty, read.RevisionDate);
        Assert.Equal(string.Empty, read.Components[0].Region);
        Assert.Empty(read.Highlights);
        Assert.Empty(read.Files);
    }

    // ---- the main file -----------------------------------------------------------------------

    [Fact]
    public void The_main_file_reads_back_to_the_same_bytes()
    {
        var main = new MainDocument
        {
            BoardsRevisionDate = "2026-October-11",
            Boards = [new() { BoardId = "Commodore/C64/250407", HardwareName = "Commodore 64", BoardName = "250407 (long board)", HardwareNotes = "8-pin audio/video port." }],
            OscilloscopesRevisionDate = "2026-October-4",
            Oscilloscopes = [new() { Brand = "Rigol", SeriesOrModel = "DHO800", Port = "5555", SetTimeDiv = ":TIMEBASE:SCALE {0}", TimeDivList = "2nS,5nS" }],
        };

        string written = DataTreeJson.Write(main);
        MainDocument read = DataTreeJson.ReadMain(written);

        Assert.Equal(written, DataTreeJson.Write(read));
        Assert.Equal(":TIMEBASE:SCALE {0}", read.Oscilloscopes[0].SetTimeDiv);
        Assert.Equal("250407 (long board)", read.Boards[0].BoardName);
    }

    [Fact]
    public void A_main_file_from_a_newer_format_is_refused()
    {
        Assert.Throws<InvalidDataException>(() => DataTreeJson.ReadMain("{\"formatVersion\": 9}"));
    }

    // The writer's own output is JSON any reader can parse - not just this one.
    [Fact]
    public void What_is_written_is_plain_JSON()
    {
        using JsonDocument board = JsonDocument.Parse(DataTreeJson.Write(Board()));
        using JsonDocument main = JsonDocument.Parse(DataTreeJson.Write(new MainDocument()));

        Assert.Equal(1, board.RootElement.GetProperty("formatVersion").GetInt32());
        Assert.Equal(JsonValueKind.Array, main.RootElement.GetProperty("oscilloscopes").ValueKind);
    }
}
