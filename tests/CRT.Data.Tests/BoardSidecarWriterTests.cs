using System.Globalization;
using System.Text.Json.Nodes;
using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// Covers BoardSidecarWriter - the `.json` file that sits beside every board workbook
// (NewContributeStrategy.md Phase 5, task 6).
//
// *** THIS EXISTS BECAUSE PUBLISHING WOULD OTHERWISE DESTROY DATA. *** A board is TWO files. The
// workbook holds the rows; the sidecar holds every COMPONENT HIGHLIGHT - the rectangles the whole
// app is built around - and every KiCad calibration. `BoardWorkbookWriter` writes only the
// workbook, so a publish that stopped there would leave a board whose highlights had silently
// vanished, with no retained revision to restore from (task 7 was struck).
//
// *** IT WRITES THE COMPLETE STATE, unlike BoardComponentHighlightStorage.SaveComponentHighlights.
// *** That method exists for the app's incremental editing: it replaces ONE schematic's highlights
// and deliberately preserves every other schematic already in the file. That is right for a label
// editor and exactly wrong for publishing, where the manifest is by contract the complete intended
// state - a schematic whose highlights a submission DELETED would keep its old ones forever.
//
// The round-trip tests go through the REAL reader (BoardComponentHighlightStorage.Load...), never
// a test-only parser, so a writer that agreed with a broken reader could not pass.
// ###########################################################################################
public sealed class BoardSidecarWriterTests : IDisposable
{
    private readonly string thisRoot;
    private readonly string thisWorkbookPath;

    public BoardSidecarWriterTests()
    {
        this.thisRoot = Path.Combine(
            Path.GetTempPath(), "crt-sidecar", Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(this.thisRoot);

        this.thisWorkbookPath = Path.Combine(this.thisRoot, "Data C64 250407 v2.0.0.xlsx");
    }

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

    private string SidecarPath => Path.ChangeExtension(this.thisWorkbookPath, ".json");

    private static ComponentHighlightEntry Highlight(
        string schematic, string label, string x, string y, string width = "40", string height = "20") =>
        new()
        {
            SchematicName = schematic,
            BoardLabel = label,
            X = x,
            Y = y,
            Width = width,
            Height = height
        };

    // -----------------------------------------------------------------------------------------
    // Highlights - the data whose loss this class exists to prevent.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_highlight_ROUND_TRIPS_through_the_real_reader()
    {
        // The strongest shape of test available here: written by the writer, read by the SHIPPED
        // reader. A writer checked against a test-only parser could agree with it while disagreeing
        // with the app.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "100", "200", "40", "20")],
            []);

        List<ComponentHighlightEntry> read =
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath);

        ComponentHighlightEntry entry = Assert.Single(read);

        Assert.Equal("Sheet 1", entry.SchematicName);
        Assert.Equal("U8", entry.BoardLabel);
        Assert.Equal("100", entry.X);
        Assert.Equal("200", entry.Y);
        Assert.Equal("40", entry.Width);
        Assert.Equal("20", entry.Height);
    }

    // ###########################################################################################
    // *** A SIDECAR COPIED IN BY HAND IS REPLACED (owner report, 2026-09-28). *** Production's
    // files were copied into BETA as root, and the next approval was refused at exactly this file:
    // it was OPENED for writing, which a root-owned file refuses, while REPLACING it needs only the
    // folder. Made here as a file its owner may not write. Not on Windows (no such mode) or as root
    // (which ignores it). Fails against the File.WriteAllText writer.
    // ###########################################################################################
    [Fact]
    public void A_sidecar_that_may_not_be_opened_for_writing_is_still_replaced()
    {
        // The return is for the platform analyzer, which cannot see that Skip throws.
        if (OperatingSystem.IsWindows())
        {
            Assert.Skip("Unix permissions only.");
            return;
        }
        Assert.SkipWhen(Environment.UserName == "root", "root may write anything.");

        File.WriteAllText(this.SidecarPath, "{}");
        File.SetUnixFileMode(this.SidecarPath, UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);

        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "100", "200")],
            []);

        Assert.Equal("U8", Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath)).BoardLabel);
    }

    // -----------------------------------------------------------------------------------------
    // Written only when its content changes (owner request, 2026-09-28)
    // -----------------------------------------------------------------------------------------

    private static string ShippedC128Sidecar()
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, "Assets", "Data", "Commodore", "C128", "310378", "Data C128 310378 v2.0.0.json");
    }

    // ###########################################################################################
    // *** THE CASE THAT WAS REPORTED. *** C128's shipped highlight file - written by CRT on
    // Windows: CRLF, labels in natural order (C2, C3, C10), an empty schematic, no calibration
    // section - published again with the SAME highlights. It used to be rewritten in the server's
    // format, and "Publish to production" then listed it as a file to copy though nothing in it had
    // changed. Now it keeps every byte.
    // ###########################################################################################
    [Fact]
    public void The_shipped_C128_highlights_published_again_unchanged_keep_every_byte()
    {
        File.Copy(BoardSidecarWriterTests.ShippedC128Sidecar(), this.SidecarPath);
        byte[] before = File.ReadAllBytes(this.SidecarPath);

        List<ComponentHighlightEntry> same = BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath);
        Assert.NotEmpty(same);

        Assert.False(BoardSidecarWriter.WriteIfChanged(this.thisWorkbookPath, same, []));
        Assert.Equal(before, File.ReadAllBytes(this.SidecarPath));
    }

    // Order, line endings and formatting are not content: the same highlights given in another
    // order, against a hand-written CRLF file, still count as the same.
    [Fact]
    public void Order_line_endings_and_layout_are_not_a_change()
    {
        File.WriteAllText(
            this.SidecarPath,
            "{\r\n  \"Component highlights\": { \"Sheet 1\": { \"C2\": [ { \"X\": 1, \"Y\": 2, \"Width\": 3, \"Height\": 4 } ],\r\n" +
            "  \"C10\": [ { \"X\": 5, \"Y\": 6, \"Width\": 7, \"Height\": 8 } ] } }\r\n}");

        Assert.True(BoardSidecarWriter.HoldsTheSameContent(
            this.thisWorkbookPath,
            [
                BoardSidecarWriterTests.Highlight("Sheet 1", "C10", "5", "6", "7", "8"),
                BoardSidecarWriterTests.Highlight("Sheet 1", "C2", "1", "2", "3", "4")
            ],
            []));
    }

    [Fact]
    public void A_moved_highlight_is_a_change_and_is_written()
    {
        BoardSidecarWriter.Write(this.thisWorkbookPath, [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "100", "200")], []);

        Assert.True(BoardSidecarWriter.WriteIfChanged(
            this.thisWorkbookPath, [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "101", "200")], []));

        Assert.Equal("101", Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath)).X);
    }

    // Every highlight removed is a change too - the empty board must be written, or the old
    // rectangles stay published.
    [Fact]
    public void Removing_every_highlight_is_a_change()
    {
        BoardSidecarWriter.Write(this.thisWorkbookPath, [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "100", "200")], []);

        Assert.True(BoardSidecarWriter.WriteIfChanged(this.thisWorkbookPath, [], []));
        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));
    }

    [Fact]
    public void A_changed_calibration_is_a_change()
    {
        var calibration = new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "pcb", OffsetX = 1.5, ScaleX = 2 };
        BoardSidecarWriter.Write(this.thisWorkbookPath, [], [calibration]);

        Assert.True(BoardSidecarWriter.HoldsTheSameContent(this.thisWorkbookPath, [], [calibration]));
        Assert.False(BoardSidecarWriter.HoldsTheSameContent(
            this.thisWorkbookPath, [], [new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "pcb", OffsetX = 1.5, ScaleX = 2, MirrorX = true }]));
    }

    // ###########################################################################################
    // *** ONLY A FILE IT CAN VOUCH FOR IS KEPT. *** A root nothing here reads could hold data no
    // maintainer saw, so a file carrying one is rewritten (and loses it) - as is one that does not
    // parse, or none at all.
    // ###########################################################################################
    [Fact]
    public void A_file_with_an_unknown_section_is_rewritten_without_it()
    {
        File.WriteAllText(this.SidecarPath, "{ \"Component highlights\": {}, \"Something else\": { \"a\": 1 } }");

        Assert.True(BoardSidecarWriter.WriteIfChanged(this.thisWorkbookPath, [], []));
        Assert.DoesNotContain("Something else", File.ReadAllText(this.SidecarPath), StringComparison.Ordinal);
    }

    [Fact]
    public void A_file_that_does_not_parse_or_is_missing_is_written()
    {
        Assert.True(BoardSidecarWriter.WriteIfChanged(this.thisWorkbookPath, [], []));
        Assert.True(File.Exists(this.SidecarPath));

        File.WriteAllText(this.SidecarPath, "{ not json");
        Assert.True(BoardSidecarWriter.WriteIfChanged(this.thisWorkbookPath, [], []));
        Assert.Contains(BoardSidecarWriter.HighlightsRoot, File.ReadAllText(this.SidecarPath), StringComparison.Ordinal);
    }

    [Fact]
    public void SEVERAL_schematics_and_labels_all_survive()
    {
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [
                BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20"),
                BoardSidecarWriterTests.Highlight("Sheet 1", "U9", "30", "40"),
                BoardSidecarWriterTests.Highlight("Sheet 2", "R1", "50", "60")
            ],
            []);

        List<ComponentHighlightEntry> read =
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath);

        Assert.Equal(3, read.Count);
        Assert.Contains(read, entry => entry.SchematicName == "Sheet 2" && entry.BoardLabel == "R1");
    }

    [Fact]
    public void ONE_label_with_TWO_rectangles_keeps_both()
    {
        // A component genuinely appears twice on a schematic - a connector drawn at both ends, a
        // chip shown split. The JSON shape is label to an ARRAY of rects precisely for this, and a
        // writer that assumed one rect per label would silently drop the second.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [
                BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20"),
                BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "300", "400")
            ],
            []);

        List<ComponentHighlightEntry> read =
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath);

        Assert.Equal(2, read.Count);
        Assert.All(read, entry => Assert.Equal("U8", entry.BoardLabel));
    }

    [Fact]
    public void A_REMOVED_schematics_highlights_are_GONE_after_a_rewrite()
    {
        // *** THE WHOLE REASON THIS IS NOT BoardComponentHighlightStorage.SaveComponentHighlights.
        // *** That method preserves every schematic it was not asked about, which is right for the
        // label editor and wrong here: the manifest is the COMPLETE intended state, so a submission
        // that deleted Sheet 2's highlights must actually delete them. A preserving writer would
        // leave them on the published board forever, and nothing would ever report it.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [
                BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20"),
                BoardSidecarWriterTests.Highlight("Sheet 2", "R1", "50", "60")
            ],
            []);

        Assert.Equal(2, BoardComponentHighlightStorage
            .LoadComponentHighlights(this.thisWorkbookPath).Count);

        // The second publish no longer mentions Sheet 2 at all.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20")],
            []);

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));

        Assert.Equal("Sheet 1", entry.SchematicName);
    }

    [Fact]
    public void A_board_with_NO_highlights_writes_a_file_rather_than_leaving_the_old_one()
    {
        // A submission can legitimately remove every highlight. Skipping the write "because there
        // is nothing to write" would leave the previous board's highlights in place - the exact
        // silent-stale-data failure this class exists to prevent, just at the whole-file level.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20")],
            []);

        BoardSidecarWriter.Write(this.thisWorkbookPath, [], []);

        Assert.True(File.Exists(this.SidecarPath));
        Assert.Empty(BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));
    }

    [Theory]
    [InlineData("", "U8")]
    [InlineData("Sheet 1", "")]
    [InlineData("   ", "U8")]
    public void A_highlight_missing_its_SCHEMATIC_or_LABEL_is_skipped(string schematic, string label)
    {
        // Contributed data. Either half missing makes the entry unreachable - the app looks
        // highlights up by schematic then by label - so writing it would put an unusable row in a
        // published board.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [
                BoardSidecarWriterTests.Highlight(schematic, label, "10", "20"),
                BoardSidecarWriterTests.Highlight("Sheet 1", "U9", "30", "40")
            ],
            []);

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));

        Assert.Equal("U9", entry.BoardLabel);
    }

    [Fact]
    public void Coordinates_are_parsed_INVARIANT_so_a_Danish_server_writes_the_same_file()
    {
        // *** THE PROJECT'S RECURRING BUG, and this is the worst place for it. *** These are
        // STRINGS in BoardData. Read under da-DK with culture-sensitive parsing, "100.5" becomes
        // 1005 - a highlight ten times across the board, written into a PUBLISHED file that every
        // user then downloads. Verified by actually switching the culture.
        CultureInfo original = CultureInfo.CurrentCulture;

        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("da-DK");

            BoardSidecarWriter.Write(
                this.thisWorkbookPath,
                [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "100.5", "200.4", "40", "20")],
                []);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));

        // Rounded to integers, as the shipped format stores them - 100.5 away from zero is 101.
        Assert.Equal("101", entry.X);
        Assert.Equal("200", entry.Y);
    }

    [Fact]
    public void A_MALFORMED_coordinate_skips_that_highlight_rather_than_writing_a_zero()
    {
        // Writing 0 would place the rectangle in the board's top-left corner, where it looks like
        // a real highlight somebody mis-drew rather than like data that never parsed.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [
                BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "not-a-number", "20"),
                BoardSidecarWriterTests.Highlight("Sheet 1", "U9", "30", "40")
            ],
            []);

        ComponentHighlightEntry entry = Assert.Single(
            BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));

        Assert.Equal("U9", entry.BoardLabel);
    }

    // -----------------------------------------------------------------------------------------
    // KiCad calibrations - the sidecar's other root.
    // -----------------------------------------------------------------------------------------

    private static KiCadCalibrationEntry Calibration(string schematic) => new()
    {
        SchematicName = schematic,
        CadName = "board.kicad_pcb",
        OffsetX = 1.5,
        OffsetY = 2.5,
        ScaleX = 0.75,
        ScaleY = 0.8,
        MirrorX = true,
        MirrorY = false
    };

    [Fact]
    public void A_calibration_ROUND_TRIPS_through_the_real_reader()
    {
        BoardSidecarWriter.Write(
            this.thisWorkbookPath, [], [BoardSidecarWriterTests.Calibration("Sheet 1")]);

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.thisWorkbookPath,
            "Sheet 1",
            out string cadName,
            out double offsetX,
            out double offsetY,
            out double scaleX,
            out double scaleY,
            out bool mirrorX,
            out bool mirrorY));

        Assert.Equal("board.kicad_pcb", cadName);
        Assert.Equal(1.5, offsetX);
        Assert.Equal(2.5, offsetY);
        Assert.Equal(0.75, scaleX);
        Assert.Equal(0.8, scaleY);
        Assert.True(mirrorX);
        Assert.False(mirrorY);
    }

    [Fact]
    public void BOTH_ROOTS_live_in_ONE_file_without_destroying_each_other()
    {
        // *** THE FAILURE MODE THIS PINS DOWN. *** Two roots, one file. A writer that wrote each
        // independently - open, set its own root, save - would have the second write load a file
        // the first had just created and then overwrite it, or would simply replace the whole
        // document. Either way one of the two is lost, and highlights are the half nobody notices
        // until they open the board.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20")],
            [BoardSidecarWriterTests.Calibration("Sheet 1")]);

        Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(this.thisWorkbookPath));

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.thisWorkbookPath, "Sheet 1",
            out _, out _, out _, out _, out _, out _, out _));
    }

    [Fact]
    public void A_REMOVED_calibration_is_gone_after_a_rewrite()
    {
        BoardSidecarWriter.Write(
            this.thisWorkbookPath, [], [BoardSidecarWriterTests.Calibration("Sheet 1")]);

        BoardSidecarWriter.Write(this.thisWorkbookPath, [], []);

        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.thisWorkbookPath, "Sheet 1",
            out _, out _, out _, out _, out _, out _, out _));
    }

    [Fact]
    public void A_calibration_with_NO_schematic_name_is_skipped()
    {
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [],
            [BoardSidecarWriterTests.Calibration(""), BoardSidecarWriterTests.Calibration("Sheet 1")]);

        JsonObject root = JsonNode.Parse(File.ReadAllText(this.SidecarPath))!.AsObject();

        Assert.Single(root["KiCad calibration points"]!.AsObject());
    }

    // -----------------------------------------------------------------------------------------
    // The file itself.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_sidecar_sits_beside_the_workbook_with_the_same_stem()
    {
        // The reader derives the path by changing the extension, so the writer must agree exactly.
        // A mismatch means a file nobody ever loads and a board that publishes with no highlights.
        BoardSidecarWriter.Write(
            this.thisWorkbookPath,
            [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20")],
            []);

        Assert.True(File.Exists(Path.Combine(this.thisRoot, "Data C64 250407 v2.0.0.json")));
    }

    [Fact]
    public void BOTH_ROOTS_are_always_present_even_when_empty()
    {
        // A board with no highlights and no calibrations still gets both roots, so the file has one
        // shape rather than four. The app's own loaders tolerate a missing root, but a consistent
        // file is far easier to diff by hand - which is how the project owner inspects the tree.
        BoardSidecarWriter.Write(this.thisWorkbookPath, [], []);

        JsonObject root = JsonNode.Parse(File.ReadAllText(this.SidecarPath))!.AsObject();

        Assert.True(root.ContainsKey("Component highlights"));
        Assert.True(root.ContainsKey("KiCad calibration points"));
    }

    [Fact]
    public void An_EXISTING_sidecar_is_REPLACED_rather_than_merged_into()
    {
        // Anything already in the file that this writer does not know about is dropped, on purpose.
        // The published sidecar must be exactly what the submission describes; carrying an unknown
        // root forward would publish data no maintainer ever saw.
        File.WriteAllText(this.SidecarPath, """{"Something else":{"kept":true}}""");

        BoardSidecarWriter.Write(this.thisWorkbookPath, [], []);

        JsonObject root = JsonNode.Parse(File.ReadAllText(this.SidecarPath))!.AsObject();

        Assert.False(root.ContainsKey("Something else"));
    }

    [Fact]
    public void The_folder_is_CREATED_when_it_does_not_exist()
    {
        // A brand-new board's folder does not exist before its first publish.
        string nested = Path.Combine(this.thisRoot, "new", "board", "Data X Y v2.0.0.xlsx");

        BoardSidecarWriter.Write(
            nested, [BoardSidecarWriterTests.Highlight("Sheet 1", "U8", "10", "20")], []);

        Assert.True(File.Exists(Path.ChangeExtension(nested, ".json")));
    }

    [Fact]
    public void A_BLANK_workbook_path_is_refused()
    {
        Assert.Throws<ArgumentException>(() => BoardSidecarWriter.Write("", [], []));
    }

    [Fact]
    public void NULL_lists_are_refused_rather_than_silently_writing_an_empty_board()
    {
        // *** REFUSED, NOT TREATED AS EMPTY. *** A caller passing null almost certainly failed to
        // load something rather than genuinely meaning "this board has no highlights", and
        // treating the two alike is how a publish quietly wipes a board.
        Assert.Throws<ArgumentNullException>(
            () => BoardSidecarWriter.Write(this.thisWorkbookPath, null!, []));

        Assert.Throws<ArgumentNullException>(
            () => BoardSidecarWriter.Write(this.thisWorkbookPath, [], null!));
    }
}
