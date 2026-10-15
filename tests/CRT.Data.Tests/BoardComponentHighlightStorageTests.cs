using System.Text.Json;
using System.Text.Json.Nodes;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for BoardComponentHighlightStorage - the per-board ".json" file that
// sits beside each board ".xlsx" and holds component highlight rectangles plus KiCad box
// calibration. Contributors edit these through the app, so the two behaviours that matter most
// are (a) saving one feature must never destroy another feature's data in the same file, and
// (b) board labels sort naturally (C2 before C10), because the file is read by humans too.
public sealed class BoardComponentHighlightStorageTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string ExcelPath => Path.Combine(this.thisWorkspace.Root, "Data C64 250407 v1.0.0.xlsx");
    private string JsonPath => Path.ChangeExtension(this.ExcelPath, ".json");

    private static LabelEditorSaveRow Row(
        string schematic, string label, double x, double y, double w = 10, double h = 10) =>
        new()
        {
            SchematicName = schematic,
            BoardLabel = label,
            X = x,
            Y = y,
            Width = w,
            Height = h
        };

    // ------------------------------------------------------------------------ GetJsonPath

    [Fact]
    public void GetJsonPath_swaps_the_excel_extension_for_json()
    {
        Assert.Equal(
            Path.Combine("C64", "Data C64 250407 v1.0.0.json"),
            BoardComponentHighlightStorage.GetJsonPath(Path.Combine("C64", "Data C64 250407 v1.0.0.xlsx")));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void GetJsonPath_returns_empty_for_an_empty_excel_path(string excelPath)
    {
        Assert.Equal(string.Empty, BoardComponentHighlightStorage.GetJsonPath(excelPath));
    }

    // ---------------------------------------------------------- LoadComponentHighlights

    [Fact]
    public void LoadComponentHighlights_returns_empty_when_the_json_file_is_missing()
    {
        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    [Fact]
    public void LoadComponentHighlights_returns_empty_for_a_blank_file()
    {
        File.WriteAllText(this.JsonPath, "   ");
        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    [Fact]
    public void LoadComponentHighlights_returns_empty_for_malformed_json_instead_of_throwing()
    {
        // A corrupt contribution must not take the app down; it degrades to "no highlights".
        File.WriteAllText(this.JsonPath, "{ this is not json");
        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    [Fact]
    public void LoadComponentHighlights_flattens_the_nested_json_into_entries()
    {
        File.WriteAllText(this.JsonPath, """
        {
          "Component highlights": {
            "Sheet 1": {
              "U1": [ { "X": 10, "Y": 20, "Width": 30, "Height": 40 } ]
            }
          }
        }
        """);

        var entries = BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath);

        ComponentHighlightEntry entry = Assert.Single(entries);
        Assert.Equal("Sheet 1", entry.SchematicName);
        Assert.Equal("U1", entry.BoardLabel);
        Assert.Equal("10", entry.X);
        Assert.Equal("20", entry.Y);
        Assert.Equal("30", entry.Width);
        Assert.Equal("40", entry.Height);
    }

    [Fact]
    public void LoadComponentHighlights_keeps_every_rectangle_for_a_label()
    {
        // One component can be highlighted in several places on the same sheet.
        File.WriteAllText(this.JsonPath, """
        {
          "Component highlights": {
            "Sheet 1": {
              "U1": [ { "X": 1, "Y": 1, "Width": 5, "Height": 5 },
                      { "X": 9, "Y": 9, "Width": 5, "Height": 5 } ]
            }
          }
        }
        """);

        Assert.Equal(2, BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath).Count);
    }

    [Fact]
    public void LoadComponentHighlights_orders_board_labels_naturally_not_lexicographically()
    {
        // C10 must come after C2, not before it. This ordering is what the JSON file and the
        // label editor list both rely on.
        File.WriteAllText(this.JsonPath, """
        {
          "Component highlights": {
            "Sheet 1": {
              "C105": [ { "X": 1, "Y": 1, "Width": 1, "Height": 1 } ],
              "C10":  [ { "X": 1, "Y": 1, "Width": 1, "Height": 1 } ],
              "C2":   [ { "X": 1, "Y": 1, "Width": 1, "Height": 1 } ]
            }
          }
        }
        """);

        var labels = BoardComponentHighlightStorage
            .LoadComponentHighlights(this.ExcelPath)
            .Select(e => e.BoardLabel)
            .ToList();

        Assert.Equal(new[] { "C2", "C10", "C105" }, labels);
    }

    [Fact]
    public void LoadComponentHighlights_skips_entries_with_a_blank_schematic_or_label()
    {
        File.WriteAllText(this.JsonPath, """
        {
          "Component highlights": {
            "": { "U1": [ { "X": 1, "Y": 1, "Width": 1, "Height": 1 } ] },
            "Sheet 1": { "  ": [ { "X": 1, "Y": 1, "Width": 1, "Height": 1 } ] }
          }
        }
        """);

        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    [Fact]
    public void LoadComponentHighlights_ignores_other_feature_roots_in_the_same_file()
    {
        File.WriteAllText(this.JsonPath, """
        {
          "Component highlights": { "Sheet 1": { "U1": [ { "X": 1, "Y": 2, "Width": 3, "Height": 4 } ] } },
          "KiCad calibration points": { "Sheet 1": { "CadName": "board.kicad_pcb" } }
        }
        """);

        Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    // ###########################################################################################
    // *** A SIDECAR THAT CANNOT BE READ IN FULL IS NOT WRITTEN OVER (code review, 2026-10-13). ***
    // Both savers rebuild the whole file from readers that answer "nothing" for what they cannot
    // parse: one hand-edited rectangle ("X": 10.5) and a calibration save wrote every highlight on
    // the board away. Each such file now refuses the save and is left byte for byte.
    // ###########################################################################################
    [Theory]
    [InlineData("""{ "Component highlights": { "Sheet 1": { "U8": [ { "X": 10.5, "Y": 1, "Width": 5, "Height": 5 } ] } } }""")]
    [InlineData("""{ "Component highlights": { "Sheet 1": { "U8": [ { "X": 1, "Y": 1, "Width": 5, "Height": 5 } ] } }, "KiCad calibration points": { "Sheet 1": { "OffsetX": "left" } } }""")]
    [InlineData("""{ "Component highlights": { "Sheet 1": { "U8": [ { "X": 1, "Y": 1, "Width": 5, "Height": 5 } ] } }, "KiCad calibration points": { "Sheet 1": 5 } }""")]
    [InlineData("""{ "Schematic renames": { "From": "a" } }""")]
    public void A_sidecar_that_does_not_read_in_full_refuses_both_saves_and_is_left_as_it_was(string sidecar)
    {
        string excelPath = this.ExcelPath;
        string jsonPath = BoardComponentHighlightStorage.GetJsonPath(excelPath);
        File.WriteAllText(jsonPath, sidecar);

        Assert.Throws<InvalidOperationException>(() =>
            BoardComponentHighlightStorage.SaveKiCadCalibration(excelPath, "Sheet 2", "pcb", 1, 2, 1, 1, false, false));

        Assert.Throws<InvalidOperationException>(() =>
            BoardComponentHighlightStorage.SaveComponentHighlights(excelPath, "Sheet 2", []));

        Assert.Equal(sidecar, File.ReadAllText(jsonPath));
    }

    // ---------------------------------------------------------- SaveComponentHighlights

    [Fact]
    public void SaveComponentHighlights_round_trips_through_LoadComponentHighlights()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath,
            "Sheet 1",
            new[] { Row("Sheet 1", "U1", 10, 20, 30, 40) });

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));

        Assert.Equal("Sheet 1", entry.SchematicName);
        Assert.Equal("U1", entry.BoardLabel);
        Assert.Equal("10", entry.X);
        Assert.Equal("40", entry.Height);
    }

    [Fact]
    public void SaveComponentHighlights_rounds_coordinates_away_from_zero()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath,
            "Sheet 1",
            new[] { Row("Sheet 1", "U1", 10.5, 20.4, 30.5, 40.6) });

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));

        Assert.Equal("11", entry.X);   // .5 rounds up, not to even
        Assert.Equal("20", entry.Y);
        Assert.Equal("31", entry.Width);
        Assert.Equal("41", entry.Height);
    }

    [Fact]
    public void SaveComponentHighlights_only_writes_rows_for_the_named_schematic()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath,
            "Sheet 1",
            new[]
            {
                Row("Sheet 1", "U1", 1, 1),
                Row("Sheet 2", "U2", 2, 2)   // belongs to a different sheet - must be ignored
            });

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
        Assert.Equal("U1", entry.BoardLabel);
    }

    [Fact]
    public void SaveComponentHighlights_replaces_only_the_named_schematic_and_keeps_the_others()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U1", 1, 1) });
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 2", new[] { Row("Sheet 2", "U2", 2, 2) });

        // Re-saving Sheet 1 must not wipe Sheet 2.
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U9", 9, 9) });

        var entries = BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath);

        Assert.Equal(2, entries.Count);
        Assert.Contains(entries, e => e.SchematicName == "Sheet 1" && e.BoardLabel == "U9");
        Assert.Contains(entries, e => e.SchematicName == "Sheet 2" && e.BoardLabel == "U2");
        Assert.DoesNotContain(entries, e => e.BoardLabel == "U1");   // replaced, not merged
    }

    [Fact]
    public void SaveComponentHighlights_preserves_the_kicad_calibration_root()
    {
        // The two features share one file. This is the regression that would silently destroy
        // a contributor's calibration work.
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 1", "board.kicad_pcb", 1, 2, 3, 4, mirrorX: true, mirrorY: false);

        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U1", 1, 1) });

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 1",
            out string cadName, out double offsetX, out double offsetY,
            out double scaleX, out double scaleY, out bool mirrorX, out bool mirrorY));

        Assert.Equal("board.kicad_pcb", cadName);
        Assert.Equal(1, offsetX);
        Assert.Equal(4, scaleY);
        Assert.True(mirrorX);
        Assert.False(mirrorY);
    }

    [Fact]
    public void SaveComponentHighlights_groups_repeated_labels_and_orders_rectangles_by_position()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath,
            "Sheet 1",
            new[]
            {
                Row("Sheet 1", "U1", x: 50, y: 90),
                Row("Sheet 1", "U1", x: 10, y: 10),
                Row("Sheet 1", "U1", x: 30, y: 10)
            });

        var rects = JsonNode.Parse(File.ReadAllText(this.JsonPath))!
            ["Component highlights"]!["Sheet 1"]!["U1"]!.AsArray();

        Assert.Equal(3, rects.Count);
        // Sorted by Y then X.
        Assert.Equal(10, rects[0]!["Y"]!.GetValue<int>());
        Assert.Equal(10, rects[0]!["X"]!.GetValue<int>());
        Assert.Equal(30, rects[1]!["X"]!.GetValue<int>());
        Assert.Equal(90, rects[2]!["Y"]!.GetValue<int>());
    }

    // ###########################################################################################
    // *** THE ORDER IS THE SHARED WRITER'S NOW (code review, 2026-10-13). *** This method wrote its
    // own file shape - labels in natural order - while every draft write and the server's publish
    // write through BoardSidecarWriter, which orders labels by plain text. One content had several
    // byte forms, and the draft's fingerprint (which hashes the bytes) brought Submit back for a
    // draft that held what was sent. It now writes through BoardSidecarWriter; the shared order is
    // kept rather than changed, because changing it would re-hash every published sidecar at once.
    // ###########################################################################################
    [Fact]
    public void SaveComponentHighlights_writes_board_labels_in_the_shared_writers_order()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath,
            "Sheet 1",
            new[]
            {
                Row("Sheet 1", "C105", 1, 1),
                Row("Sheet 1", "C2", 1, 1),
                Row("Sheet 1", "C10", 1, 1)
            });

        var labels = JsonNode.Parse(File.ReadAllText(this.JsonPath))!
            ["Component highlights"]!["Sheet 1"]!.AsObject()
            .Select(p => p.Key)
            .ToList();

        Assert.Equal(new[] { "C10", "C105", "C2" }, labels);
    }

    [Fact]
    public void SaveComponentHighlights_skips_rows_with_a_blank_board_label()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath,
            "Sheet 1",
            new[] { Row("Sheet 1", "  ", 1, 1), Row("Sheet 1", "U1", 1, 1) });

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
        Assert.Equal("U1", entry.BoardLabel);
    }

    [Fact]
    public void SaveComponentHighlights_can_clear_a_schematic_by_saving_no_rows()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U1", 1, 1) });

        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", Array.Empty<LabelEditorSaveRow>());

        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SaveComponentHighlights_refuses_a_blank_schematic_name(string schematicName)
    {
        Assert.Throws<InvalidOperationException>(() =>
            BoardComponentHighlightStorage.SaveComponentHighlights(
                this.ExcelPath, schematicName, Array.Empty<LabelEditorSaveRow>()));
    }

    [Fact]
    public void SaveComponentHighlights_writes_valid_indented_json()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U1", 1, 1) });

        string json = File.ReadAllText(this.JsonPath);

        Assert.Contains("\n", json);                                     // indented, human-editable
        Exception? thrown = Record.Exception(() => JsonDocument.Parse(json));
        Assert.True(thrown is null, "saved board JSON must be parseable");
    }

    // ------------------------------------------------------------- KiCad calibration

    [Fact]
    public void TryLoadKiCadCalibration_returns_false_when_the_file_is_missing()
    {
        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 1", out _, out _, out _, out _, out _, out _, out _));
    }

    [Fact]
    public void TryLoadKiCadCalibration_reports_neutral_defaults_when_it_fails()
    {
        // Scale defaults to 1, not 0 - a zero scale would collapse the whole overlay.
        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 1",
            out string cadName, out double offsetX, out double offsetY,
            out double scaleX, out double scaleY, out bool mirrorX, out bool mirrorY));

        Assert.Equal(string.Empty, cadName);
        Assert.Equal(0, offsetX);
        Assert.Equal(0, offsetY);
        Assert.Equal(1, scaleX);
        Assert.Equal(1, scaleY);
        Assert.False(mirrorX);
        Assert.False(mirrorY);
    }

    [Fact]
    public void SaveKiCadCalibration_round_trips_every_field()
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 1", "  board.kicad_pcb  ",
            offsetX: 1.5, offsetY: -2.5, scaleX: 0.25, scaleY: 4.0,
            mirrorX: true, mirrorY: true);

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 1",
            out string cadName, out double offsetX, out double offsetY,
            out double scaleX, out double scaleY, out bool mirrorX, out bool mirrorY));

        Assert.Equal("board.kicad_pcb", cadName);   // trimmed on save
        Assert.Equal(1.5, offsetX);
        Assert.Equal(-2.5, offsetY);
        Assert.Equal(0.25, scaleX);
        Assert.Equal(4.0, scaleY);
        Assert.True(mirrorX);
        Assert.True(mirrorY);
    }

    [Fact]
    public void SaveKiCadCalibration_keeps_calibration_for_other_schematics()
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 1", "a.kicad_pcb", 1, 1, 1, 1, false, false);
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 2", "b.kicad_pcb", 2, 2, 2, 2, false, false);

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 1", out string first, out _, out _, out _, out _, out _, out _));
        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 2", out string second, out _, out _, out _, out _, out _, out _));

        Assert.Equal("a.kicad_pcb", first);
        Assert.Equal("b.kicad_pcb", second);
    }

    [Fact]
    public void SaveKiCadCalibration_preserves_the_component_highlights_root()
    {
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U1", 1, 1) });

        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 1", "board.kicad_pcb", 0, 0, 1, 1, false, false);

        Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(this.ExcelPath));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SaveKiCadCalibration_refuses_a_blank_schematic_name(string schematicName)
    {
        Assert.Throws<InvalidOperationException>(() =>
            BoardComponentHighlightStorage.SaveKiCadCalibration(
                this.ExcelPath, schematicName, "x", 0, 0, 1, 1, false, false));
    }

    [Fact]
    public void TryLoadKiCadCalibration_returns_false_for_an_unknown_schematic()
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 1", "board.kicad_pcb", 0, 0, 1, 1, false, false);

        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Sheet 99", out _, out _, out _, out _, out _, out _, out _));
    }

    [Fact]
    public void Root_keys_are_written_in_a_stable_order()
    {
        // Stable ordering keeps the JSON diff-friendly for contributions.
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Sheet 1", "board.kicad_pcb", 0, 0, 1, 1, false, false);
        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.ExcelPath, "Sheet 1", new[] { Row("Sheet 1", "U1", 1, 1) });

        var rootKeys = JsonNode.Parse(File.ReadAllText(this.JsonPath))!.AsObject()
            .Select(p => p.Key)
            .ToList();

        Assert.Equal(new[] { "Component highlights", "KiCad calibration points" }, rootKeys);
    }

    // ###########################################################################################
    // *** A CALIBRATION IS FOUND WHATEVER THE CAPITALS OF ITS SCHEMATIC'S NAME (owner request,
    // 2026-10-10 - "And fix all"). *** Every other schematic name compares without regard to
    // capitals; this exact lookup did not, so a schematic renamed from "top" to "Top" kept its
    // highlights and silently lost its KiCad calibration.
    // ###########################################################################################
    [Fact]
    public void TryLoadKiCadCalibration_finds_a_calibration_filed_under_other_capitals()
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(this.ExcelPath, "top", "B.Cu", 1.0, 2.0, 3.5, 4.0, true, false);

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Top", out string cadName, out _, out _, out double scaleX, out _, out bool mirrorX, out _));

        Assert.Equal("B.Cu", cadName);
        Assert.Equal(3.5, scaleX);
        Assert.True(mirrorX);
    }

    // The exact name wins over another capitalisation of it.
    [Fact]
    public void TryLoadKiCadCalibration_prefers_the_exact_name()
    {
        File.WriteAllText(this.JsonPath, """
        {
          "KiCad calibration points": {
            "TOP": { "CadName": "wrong", "ScaleX": 9.0 },
            "Top": { "CadName": "right", "ScaleX": 2.0 }
          }
        }
        """);

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.ExcelPath, "Top", out string cadName, out _, out _, out double scaleX, out _, out _, out _));

        Assert.Equal("right", cadName);
        Assert.Equal(2.0, scaleX);
    }

    // Saving under new capitals replaces the old key rather than keeping both.
    [Fact]
    public void SaveKiCadCalibration_replaces_the_same_schematic_under_other_capitals()
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(this.ExcelPath, "top", "B.Cu", 1.0, 2.0, 3.5, 4.0, false, false);
        BoardComponentHighlightStorage.SaveKiCadCalibration(this.ExcelPath, "Top", "F.Cu", 5.0, 6.0, 7.0, 8.0, false, false);

        JsonObject calibrations = (JsonObject)JsonNode.Parse(File.ReadAllText(this.JsonPath))!["KiCad calibration points"]!;

        Assert.Equal(["Top"], calibrations.Select(pair => pair.Key));
        Assert.Equal("F.Cu", calibrations["Top"]!["CadName"]!.GetValue<string>());
    }

    // ------------------------------------------------------------------ LoadKiCadCalibrations

    // ###########################################################################################
    // A board's calibrations from ONE read of the file (code review, 2026-10-10) - every draft
    // write collected them through TryLoadKiCadCalibration, parsing the file once per schematic,
    // twice over. The same answers as TryLoadKiCadCalibration: named as asked, found under other
    // capitals, a blank or unknown name skipped.
    // ###########################################################################################
    [Fact]
    public void LoadKiCadCalibrations_reads_every_name_asked_for_in_order()
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Top", "F.Cu", 1.5, -2.5, 0.25, 4.0, true, false);
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.ExcelPath, "Bottom", "B.Cu", 0, 0, 1, 1, false, true);

        List<KiCadCalibrationEntry> calibrations = BoardComponentHighlightStorage.LoadKiCadCalibrations(
            this.ExcelPath,
            [" bottom ", "Gone", "", null, "Top"]);

        Assert.Equal(["bottom", "Top"], calibrations.Select(entry => entry.SchematicName));
        Assert.Equal("B.Cu", calibrations[0].CadName);
        Assert.True(calibrations[0].MirrorY);
        Assert.Equal(1.5, calibrations[1].OffsetX);
        Assert.Equal(-2.5, calibrations[1].OffsetY);
        Assert.Equal(0.25, calibrations[1].ScaleX);
        Assert.Equal(4.0, calibrations[1].ScaleY);
        Assert.True(calibrations[1].MirrorX);
    }

    // One entry that cannot be read loses only itself.
    [Fact]
    public void LoadKiCadCalibrations_skips_an_unreadable_entry_and_keeps_the_others()
    {
        File.WriteAllText(this.JsonPath, """
            {
              "KiCad calibration points": {
                "Top": { "CadName": "F.Cu", "ScaleX": "not a number" },
                "Bottom": { "CadName": "B.Cu" }
              }
            }
            """);

        Assert.Equal(
            ["Bottom"],
            BoardComponentHighlightStorage.LoadKiCadCalibrations(this.ExcelPath, ["Top", "Bottom"]).Select(entry => entry.SchematicName));
    }

    [Fact]
    public void LoadKiCadCalibrations_is_empty_for_a_missing_or_malformed_file()
    {
        Assert.Empty(BoardComponentHighlightStorage.LoadKiCadCalibrations(this.ExcelPath, ["Top"]));

        File.WriteAllText(this.JsonPath, "{ not json");
        Assert.Empty(BoardComponentHighlightStorage.LoadKiCadCalibrations(this.ExcelPath, ["Top"]));
    }
}
