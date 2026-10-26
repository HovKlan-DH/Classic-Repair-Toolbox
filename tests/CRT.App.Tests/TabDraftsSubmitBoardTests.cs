using System.Text.Json.Nodes;
using CRT;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// WHAT A SUBMIT SENDS - TabDrafts.BoardToSendAsync, the board SubmitAsync puts in the manifest.
//
// *** OWNER REPORT, 2026-10-11. *** An image renamed to "PCB; Toppi Poppi" in "Schematic images"
// arrived in the queue as "PCB; Toppi", the name before it, with its highlights under the old name
// and its KiCad calibration missing. Submit read the board through DataManager.LoadBoardDataAsync,
// whose cache still held the board from before the rename (the window had been opened with the
// board not selected in the drop-down lists, and only the selected board's cache was cleared) -
// while the receipt's fingerprint was taken from the FILE, so the draft then read as "already sent".
//
// The second test is the same line's other way to send the wrong board: with "view boards as
// officially published" ticked, LoadBoardDataAsync answers the PUBLISHED board. The third is the
// same trap in "KiCad data" (TabDrafts.BoardForFilesWindowAsync; code review, 2026-10-11): its
// report judged the draft's KiCad data against the published "Important signals" rows.
//
// In the "DataManager" collection: the board is first read the way the rest of CRT reads it,
// through DataManager's real static state. Two tests also set UserSettings.ViewOfficialPublishedOnly,
// and every test points DraftManager at its own folder - a class can join one collection only, and
// collections never run in parallel (xunit.runner.json), so DataManagerDraftOverlayTests' choice is
// kept. All three statics are pointed away from the temp folder again in Dispose, so nothing a
// test here set can reach the next (code review, 2026-10-11).
// ###########################################################################################
[Collection("DataManager")]
public sealed class TabDraftsSubmitBoardTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    // Unique per test, so BoardDataReader's static cache never carries one test's board into the
    // next (see DataManagerDraftOverlayTests).
    private readonly string thisExcelDataFile =
        $"Commodore/C128/{Guid.NewGuid():N}/board.json";

    static TabDraftsSubmitBoardTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public TabDraftsSubmitBoardTests()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        DraftManager.LoadFrom(this.DraftsRoot);

        this.LoadFreshSettings();
    }

    public void Dispose()
    {
        DataManager.ClearBoardCache(this.thisExcelDataFile);
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");

        // A clean slate for whatever runs next - "view boards as officially published" off, and no
        // drafts root left pointing at a folder about to be deleted.
        this.LoadFreshSettings();
        DraftManager.LoadFrom(string.Empty);

        this.thisWorkspace.Dispose();
    }

    // "{}" first: LoadFrom on a missing file keeps the previous test's settings.
    private void LoadFreshSettings()
    {
        string settingsPath = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(settingsPath, "{}");
        UserSettings.LoadFrom(settingsPath);
    }

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DraftWorkbook => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, this.thisExcelDataFile);

    private string ImageFile => Path.GetDirectoryName(this.thisExcelDataFile)!.Replace('\\', '/') + "/C128 Neo.3.F.png";

    private DraftStatus Status => DraftStatusReader.Resolve(this.thisWorkspace.Root, this.DraftsRoot, this.thisExcelDataFile)!;

    private HardwareBoardEntry NewBoardEntry => new() { ExcelDataFile = this.thisExcelDataFile, IsDraftOnly = true };

    // A NEW board's draft, as the owner's was: one image "PCB; Top" with a highlight and a KiCad
    // calibration on it, both kept in the draft's JSON sidecar.
    private void CreateNewBoardDraft()
    {
        CachedWorkbooks.Write(this.DraftWorkbook, new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "PCB; Top", CadName = "Open128 - PCB Top", SchematicImageFile = this.ImageFile }],
            Components = [new ComponentEntry { BoardLabel = "C1", FriendlyName = "Capacitor", Category = "Capacitor" }],
        });

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, this.thisExcelDataFile),
            new DraftMarker
            {
                BoardKey = this.thisExcelDataFile,
                NewBoard = new NewBoardRegistration { HardwareName = "C128", BoardName = "Test", ExcelDataFile = this.thisExcelDataFile },
            });

        Assert.True(DraftWorkbookStore.Edit(
            this.DraftsRoot,
            this.thisExcelDataFile,
            board =>
            {
                var updated = new BoardData
                {
                    RevisionDate = board.RevisionDate,
                    Schematics = [.. board.Schematics],
                    Components = [.. board.Components],
                };

                updated.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "PCB; Top", BoardLabel = "C1", X = "10", Y = "20", Width = "30", Height = "40" });
                return updated;
            }));

        // Written into the draft's board file as CRT's calibration save files it - the storage
        // class is CRT.Data's own, not visible to this project. In CRT 4.0 the calibrations are a
        // list of the board file itself (BoardDocument.KiCadCalibrations).
        var boardFile = (JsonObject)JsonNode.Parse(File.ReadAllText(this.DraftWorkbook))!;
        boardFile["kiCadCalibrations"] = new JsonArray
        {
            new JsonObject
            {
                ["schematic"] = "PCB; Top",
                ["cadName"] = "Open128 - PCB Top",
                ["offsetX"] = 1.0,
                ["offsetY"] = 2.0,
                ["scaleX"] = 3.5,
                ["scaleY"] = 4.0,
                ["mirrorX"] = true,
                ["mirrorY"] = false,
            },
        };
        File.WriteAllText(this.DraftWorkbook, boardFile.ToJsonString());
    }

    [Fact]
    public async Task A_rename_made_after_the_board_was_read_is_what_the_submit_sends()
    {
        this.CreateNewBoardDraft();

        // Read as CRT reads a board - here as "Schematic images" does when it opens.
        BoardData? before = await DataManager.LoadBoardDataAsync(this.NewBoardEntry);
        Assert.Equal("PCB; Top", before!.Schematics.Single().SchematicName);

        Assert.Equal(
            SchematicRenameOutcome.Renamed,
            DraftWorkbookStore.RenameSchematic(this.DraftsRoot, this.thisExcelDataFile, new SchematicRowRef("PCB; Top", this.ImageFile), "PCB; Toppi Poppi"));

        BoardData? sent = await TabDrafts.BoardToSendAsync(this.Status);

        Assert.NotNull(sent);
        Assert.Equal("PCB; Toppi Poppi", sent!.Schematics.Single().SchematicName);
        Assert.Equal("PCB; Toppi Poppi", sent.ComponentHighlights.Single().SchematicName);

        // The calibration travels beside the board, collected under the board's names - under the
        // old name it found nothing, and the review said the calibration had been removed.
        Assert.Equal(
            ["PCB; Toppi Poppi"],
            DraftBoardSource.CollectCalibrations(this.DraftWorkbook, sent).Select(entry => entry.SchematicName));
    }

    [Fact]
    public async Task The_submit_sends_the_draft_while_boards_are_viewed_as_officially_published()
    {
        string published = Path.Combine(this.thisWorkspace.Root, this.thisExcelDataFile.Replace('/', Path.DirectorySeparatorChar));
        CachedWorkbooks.Write(published, new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "PLA", Category = "IC" }],
        });

        CachedWorkbooks.Write(this.DraftWorkbook, new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "PLA (corrected)", Category = "IC" }],
        });

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, this.thisExcelDataFile),
            new DraftMarker { BoardKey = this.thisExcelDataFile, BaseRevision = "2026-09-01" });

        UserSettings.ViewOfficialPublishedOnly = true;

        BoardData? sent = await TabDrafts.BoardToSendAsync(this.Status);

        Assert.NotNull(sent);
        Assert.Equal("PLA (corrected)", sent!.Components.Single().FriendlyName);
    }

    // ###########################################################################################
    // *** "KICAD DATA" IS HANDED THE DRAFT, NOT THE PUBLISHED BOARD (code review, 2026-10-11). *** Its
    // report says which "Important signals" rows will not light up and which nets they already use -
    // of the draft's KiCad data, so against the draft's rows. With "view boards as officially
    // published" ticked it was handed the published board's rows instead.
    // ###########################################################################################
    [Fact]
    public async Task KiCad_data_is_handed_the_drafts_signals_while_boards_are_viewed_as_officially_published()
    {
        string published = Path.Combine(this.thisWorkspace.Root, this.thisExcelDataFile.Replace('/', Path.DirectorySeparatorChar));
        CachedWorkbooks.Write(published, new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "PLA", Category = "IC" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "/PHI0" }],
        });

        CachedWorkbooks.Write(this.DraftWorkbook, new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", FriendlyName = "PLA", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", FriendlyName = "CPU", Category = "IC" },
            ],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "/PHI2" }],
        });

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, this.thisExcelDataFile),
            new DraftMarker { BoardKey = this.thisExcelDataFile, BaseRevision = "2026-09-01" });

        UserSettings.ViewOfficialPublishedOnly = true;

        BoardData? handed = await TabDrafts.BoardForFilesWindowAsync(new HardwareBoardEntry { ExcelDataFile = this.thisExcelDataFile });

        Assert.NotNull(handed);
        Assert.Equal(["/PHI2"], handed!.KiCadImportantSignals.Select(signal => signal.KiCadNetName));
        Assert.Equal(["U1", "U2"], handed.Components.Select(component => component.BoardLabel));
    }
}
