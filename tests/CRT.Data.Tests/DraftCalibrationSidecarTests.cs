using System;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCad calibrations, now that a draft carries an ORDINARY JSON sidecar
// (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
//
// *** THIS IS WHERE BoardDraft's "ELEVENTH SECTION" RETIRED. *** Calibrations used to live in
// BoardDraft.KiCadCalibrations, a section added purely because a draft was a delta list and
// BoardData has no calibration field. They always had a published home - the sidecar's
// "KiCad calibration points" root, which BoardComponentHighlightStorage reads and writes and
// which the SERVER already publishes into.
//
// A draft folder now carries that same sidecar, so a drafted calibration is written and read by
// exactly the same calls a published one is. What these tests pin is the two things that
// followed from the change: the path resolution that chooses WHICH sidecar, and the submission
// collector that replaced KiCadCalibrationDraftWriter.CollectForSubmission.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftCalibrationSidecarTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData BoardWithSchematics(params string[] names)
    {
        var board = new BoardData();
        board.Schematics.AddRange(names.Select(name => new BoardSchematicEntry { SchematicName = name }));

        return board;
    }

    private void WritePublishedBoard(BoardData board)
    {
        string path = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftCalibrationSidecarTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);
    }

    private void WriteDraftBoard(BoardData board)
    {
        string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey),
            new DraftMarker { BoardKey = DraftCalibrationSidecarTests.BoardKey });
    }

    private static void SaveCalibration(string workbookPath, string schematicName, double scaleX)
    {
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            workbookPath,
            schematicName,
            cadName: "B.Cu",
            offsetX: 1.0,
            offsetY: 2.0,
            scaleX: scaleX,
            scaleY: 4.0,
            mirrorX: false,
            mirrorY: false);
    }

    // ------------------------------------------------------------------ Which sidecar

    [Fact]
    public void With_NO_draft_the_PUBLISHED_workbook_is_the_write_target()
    {
        this.WritePublishedBoard(DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1"));

        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftCalibrationSidecarTests.BoardKey),
            DraftBoardSource.ResolveWritablePath(this.DataRoot, this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey));
    }

    [Fact]
    public void With_a_draft_the_DRAFTS_workbook_is_the_write_target()
    {
        this.WritePublishedBoard(DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1"));
        this.WriteDraftBoard(DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1"));

        Assert.Equal(
            DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey),
            DraftBoardSource.ResolveWritablePath(this.DataRoot, this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey));
    }

    // ###########################################################################################
    // The marker is what makes a folder a draft here too, exactly as it is for the board itself.
    // A stray workbook under Drafts/ must not become the write target, or a calibration would be
    // saved into a file the application never reads back.
    // ###########################################################################################
    [Fact]
    public void A_stray_workbook_with_NO_marker_is_not_the_write_target()
    {
        this.WritePublishedBoard(DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1"));

        string stray = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(stray)!);
        File.WriteAllText(stray, "not a draft");

        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftCalibrationSidecarTests.BoardKey),
            DraftBoardSource.ResolveWritablePath(this.DataRoot, this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey));
    }

    // ------------------------------------------------------------------ Round trip

    // ###########################################################################################
    // *** A DRAFTED CALIBRATION DOES NOT LEAK INTO THE PUBLISHED SIDECAR. ***
    //
    // The whole point of writing into the draft folder. Before Phase 6 the calibration was written
    // into a draft ROW precisely to stop it touching the published tree; now the separation comes
    // from the path instead, and this proves the published sidecar is untouched.
    // ###########################################################################################
    [Fact]
    public void Saving_a_drafted_calibration_leaves_the_PUBLISHED_sidecar_alone()
    {
        BoardData board = DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1");
        this.WritePublishedBoard(board);
        this.WriteDraftBoard(board);

        string publishedPath = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftCalibrationSidecarTests.BoardKey);
        string draftPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);

        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet 1", scaleX: 3.5);

        // Read back from the draft: it is there.
        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            draftPath, "Sheet 1", out _, out _, out _, out double draftScaleX, out _, out _, out _));
        Assert.Equal(3.5, draftScaleX);

        // And the published side never heard about it.
        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            publishedPath, "Sheet 1", out _, out _, out _, out _, out _, out _, out _));
    }

    // ------------------------------------------------------------------ Collecting for submission

    [Fact]
    public void Every_saved_calibration_is_collected_for_submission()
    {
        BoardData board = DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1", "Sheet 2");
        this.WriteDraftBoard(board);

        string draftPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);

        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet 1", scaleX: 3.5);
        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet 2", scaleX: 7.0);

        var collected = DraftBoardSource.CollectCalibrations(draftPath, board);

        Assert.Equal(2, collected.Count);
        Assert.Equal(3.5, collected.Single(c => c.SchematicName == "Sheet 1").ScaleX);
    }

    // ###########################################################################################
    // *** A REMOVED CALIBRATION IS SIMPLY ABSENT - no tombstone needed. ***
    //
    // The delta writer had to skip Deleted rows explicitly: a removal was recorded AS a row, so
    // reading every row would have sent back the very calibration the contributor removed, and
    // because the manifest is the complete intended state, publishing would have put it straight
    // back. A sidecar has no such trap.
    // ###########################################################################################
    [Fact]
    public void A_schematic_with_NO_calibration_contributes_nothing()
    {
        BoardData board = DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1", "Sheet 2");
        this.WriteDraftBoard(board);

        string draftPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);

        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet 1", scaleX: 3.5);

        var collected = DraftBoardSource.CollectCalibrations(draftPath, board);

        Assert.Equal("Sheet 1", Assert.Single(collected).SchematicName);
    }

    // ###########################################################################################
    // A sidecar can carry an entry for a schematic the board no longer has - the contributor
    // deleted the page but the calibration stayed behind. Sending it would publish a calibration
    // for a page that does not exist, so the BOARD's schematic list decides what is collected.
    // ###########################################################################################
    [Fact]
    public void A_calibration_for_a_schematic_the_board_no_longer_has_is_NOT_collected()
    {
        BoardData withBoth = DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1", "Sheet 2");
        this.WriteDraftBoard(withBoth);

        string draftPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);

        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet 1", scaleX: 3.5);
        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet 2", scaleX: 7.0);

        // The contributor has since removed Sheet 2 from the board.
        BoardData withoutSheetTwo = DraftCalibrationSidecarTests.BoardWithSchematics("Sheet 1");

        var collected = DraftBoardSource.CollectCalibrations(draftPath, withoutSheetTwo);

        Assert.Equal("Sheet 1", Assert.Single(collected).SchematicName);
    }

    [Fact]
    public void Collected_calibrations_are_ORDERED_by_schematic_name()
    {
        // The manifest is hashed and diffed, so an unstable order reports a change nobody made.
        BoardData board = DraftCalibrationSidecarTests.BoardWithSchematics("Sheet Z", "Sheet A");
        this.WriteDraftBoard(board);

        string draftPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftCalibrationSidecarTests.BoardKey);

        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet Z", scaleX: 1.0);
        DraftCalibrationSidecarTests.SaveCalibration(draftPath, "Sheet A", scaleX: 2.0);

        var collected = DraftBoardSource.CollectCalibrations(draftPath, board);

        Assert.Equal(["Sheet A", "Sheet Z"], collected.Select(c => c.SchematicName));
    }

    [Fact]
    public void Collecting_with_NO_board_or_no_path_is_harmless()
    {
        Assert.Empty(DraftBoardSource.CollectCalibrations(string.Empty, new BoardData()));
        Assert.Empty(DraftBoardSource.CollectCalibrations("some/path.xlsx", board: null));
    }
}
