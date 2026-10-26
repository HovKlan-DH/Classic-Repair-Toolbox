using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Persisting an edit to a draft, now that a draft is a board workbook rather than a delta list
// (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
//
// *** THE TEST THAT MATTERS MOST IS THE STALENESS ONE. *** The project owner's requirement is that
// editing in the app and editing in Excel be interchangeable, which means the application can
// never assume the board it last read is still what is on disk. Every write here is
// read-modify-write against the FILE, and the test proving it is the one that edits the workbook
// behind the store's back before saving.
//
// Shares the "BoardData" collection: these write and re-read workbooks, and BoardDataReader
// keeps a process-wide cache.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftWorkbookStoreTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    private string WorkbookPath =>
        DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

    private string DraftFolder =>
        DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

    // The board with a datasheet row for U8, citing the file by its 4.0 name.
    private static BoardData WithDatasheet(BoardData board)
    {
        BoardData updated = DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU");
        updated.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Datasheets/U8.pdf" });
        return updated;
    }

    // Creates a draft holding the given components, the way DraftSeeder would.
    private void CreateDraft(params ComponentEntry[] components)
    {
        var board = new BoardData { RevisionDate = "2026-09-01" };
        board.Components.AddRange(components);

        Directory.CreateDirectory(Path.GetDirectoryName(this.WorkbookPath)!);
        BoardWorkbookWriter.Write(this.WorkbookPath, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey),
            new DraftMarker
            {
                BoardKey = DraftWorkbookStoreTests.BoardKey,
                BaseRevision = "2026-09-01",
            });
    }

    private static ComponentEntry Component(string label, string friendlyName) =>
        new() { BoardLabel = label, FriendlyName = friendlyName, Category = "IC" };

    // Replaces one component's friendly name - the shape every real draft writer has.
    private static BoardData WithFriendlyName(BoardData board, string label, string friendlyName)
    {
        var updated = new BoardData
        {
            RevisionDate = board.RevisionDate,
            Schematics = board.Schematics,
            ComponentImages = board.ComponentImages,
            ComponentHighlights = board.ComponentHighlights,
            ComponentLocalFiles = board.ComponentLocalFiles,
            ComponentLinks = board.ComponentLinks,
            BoardLocalFiles = board.BoardLocalFiles,
            BoardLinks = board.BoardLinks,
            Credits = board.Credits,
            KiCadImportantSignals = board.KiCadImportantSignals,
        };

        updated.Components.AddRange(board.Components.Select(component =>
            string.Equals(component.BoardLabel, label, StringComparison.OrdinalIgnoreCase)
                ? new ComponentEntry
                {
                    BoardLabel = component.BoardLabel,
                    FriendlyName = friendlyName,
                    Category = component.Category,
                }
                : component));

        return updated;
    }

    // A field a later CRT wrote on a row survives an edit made with this one (code review,
    // 2026-10-25) - the edited row's as well as the others'. BoardDocumentUnknowns has the rules.
    [Fact]
    public void Fields_a_later_CRT_wrote_on_rows_survive_an_edit()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"), DraftWorkbookStoreTests.Component("U9", "SID"));
        Assert.True(DataTreeLayout.IsBoardFile(this.WorkbookPath));

        BoardFiles.Change(this.WorkbookPath, document =>
        {
            foreach (BoardDocument.ComponentRow row in document.Components)
                row.Unknown = new() { ["pinCount"] = System.Text.Json.JsonSerializer.SerializeToElement(40) };
        });

        Assert.True(DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "6510 CPU")));

        BoardDocument written = BoardFiles.ReadDocument(this.WorkbookPath)!;

        Assert.Equal("6510 CPU", written.Components.Single(row => row.BoardLabel == "U8").FriendlyName);
        Assert.All(written.Components, row => Assert.Equal(40, row.Unknown!["pinCount"].GetInt32()));
    }

    // ------------------------------------------------------------------ Reading

    [Fact]
    public void The_draft_board_is_read_back_from_its_workbook()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        BoardData? board = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.NotNull(board);
        Assert.Equal("CPU", board!.Components.Single().FriendlyName);
    }

    [Fact]
    public void With_NO_draft_the_board_is_null_rather_than_empty()
    {
        // Null and "an empty board" mean very different things: the caller must be able to tell
        // "there is no draft" from "the draft has no rows", because the second is a real state a
        // contributor can create by emptying one.
        Assert.Null(DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    // ------------------------------------------------------------------ Writing

    [Fact]
    public void An_edit_is_written_back_to_the_workbook()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        bool saved = DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU (corrected)"));

        Assert.True(saved);

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        Assert.Equal("CPU (corrected)", reread!.Components.Single().FriendlyName);
    }

    [Fact]
    public void An_edit_leaves_the_boards_OTHER_sections_intact()
    {
        // The transform shape every writer uses returns a whole board, so a section left out of
        // that copy is a section erased. Cheap to get wrong and silent when it happens.
        var board = new BoardData { RevisionDate = "2026-09-01" };
        board.Components.Add(DraftWorkbookStoreTests.Component("U8", "CPU"));
        board.Credits.Add(new CreditEntry { Category = "Data", NameOrHandle = "Dennis" });

        Directory.CreateDirectory(Path.GetDirectoryName(this.WorkbookPath)!);
        BoardWorkbookWriter.Write(this.WorkbookPath, board);
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey),
            new DraftMarker { BoardKey = DraftWorkbookStoreTests.BoardKey });

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            current => DraftWorkbookStoreTests.WithFriendlyName(current, "U8", "CPU (corrected)"));

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.Single(reread!.Credits);
        Assert.Equal("Dennis", reread.Credits.Single().NameOrHandle);
    }

    // ###########################################################################################
    // *** THE ONE THAT MAKES EXCEL-EDITING SAFE. ***
    //
    // The contributor may change the workbook while the application is running - that is the whole
    // point of the new layout. So a save must transform what is ON DISK at that moment, never a
    // copy the application read earlier.
    //
    // This edits the workbook behind the store's back, then saves an unrelated change through it.
    // Both must survive. Against a version that cached the board, the out-of-band edit would be
    // silently overwritten - which is data loss the contributor would have no way to anticipate.
    // ###########################################################################################
    [Fact]
    public async Task An_edit_made_OUTSIDE_the_application_is_not_overwritten_by_the_next_save()
    {
        this.CreateDraft(
            DraftWorkbookStoreTests.Component("U8", "CPU"),
            DraftWorkbookStoreTests.Component("U9", "VIC"));

        // ###########################################################################################
        // Read it once THROUGH THE ORDINARY BOARD LOAD, which is what populates BoardDataReader's
        // process-wide cache - exactly what happens when the contributor opens the board in CRT.
        //
        // *** THIS LINE IS WHAT MAKES THE TEST NON-VACUOUS, and the first version of it was wrong.
        // *** It primed the cache through DraftWorkbookStore.LoadDraftBoard, which is UNCACHED by
        // design - so nothing was ever cached, and a sabotaged Edit() that read through the cache
        // still passed. Verified by making Edit() use the cached reader and watching this test go
        // red.
        // ###########################################################################################
        BoardData? asTheAppSawIt = await BoardDataReader.LoadAsync(this.WorkbookPath, this.WorkbookPath);

        Assert.NotNull(asTheAppSawIt);

        // Now "Excel" changes U9, with the application none the wiser.
        BoardWorkbookWriter.Write(
            this.WorkbookPath,
            DraftWorkbookStoreTests.WithFriendlyName(asTheAppSawIt!, "U9", "VIC (edited in Excel)"));

        // The application saves its own, unrelated change to U8.
        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            current => DraftWorkbookStoreTests.WithFriendlyName(current, "U8", "CPU (edited in app)"));

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.Equal("CPU (edited in app)", reread!.Components.Single(c => c.BoardLabel == "U8").FriendlyName);
        Assert.Equal("VIC (edited in Excel)", reread.Components.Single(c => c.BoardLabel == "U9").FriendlyName);
    }

    // ###########################################################################################
    // Editing a board with no draft REFUSES rather than creating one.
    //
    // Creating a draft means copying the published board (DraftSeeder). Quietly starting an empty
    // one here would produce a draft that compares as "every published row deleted" - and a
    // submission built from it would ask the server to delete the board.
    // ###########################################################################################
    [Fact]
    public void Editing_a_board_with_NO_draft_refuses_rather_than_creating_one()
    {
        bool saved = DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => board);

        Assert.False(saved);
        Assert.False(File.Exists(this.WorkbookPath));
    }

    [Fact]
    public void A_failed_write_leaves_the_PREVIOUS_draft_intact()
    {
        // A transform that throws stands in for any mid-write failure. The draft workbook is now
        // the only copy of the contributor's work, so a half-written file would be the worst
        // possible outcome.
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        Assert.Throws<InvalidOperationException>(() => DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            _ => throw new InvalidOperationException("the editor changed its mind")));

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.NotNull(reread);
        Assert.Equal("CPU", reread!.Components.Single().FriendlyName);
    }

    // ###########################################################################################
    // *** A NEW FILE IS TAKEN IN ONLY WHEN THE WRITE IS (code review, 2026-10-11; CRT 4.0,
    // 2026-10-24). *** A picked file lies in the draft folder under its name until the next write
    // stores its bytes and lists it. The Excel-era write failed half-way on its sidecar and left
    // rows naming files that were gone; a 4.0 write is one file - but a new file moved into the
    // store BEFORE a write that then failed would be unlisted, and the next write's tidying would
    // delete it. So it leaves its place only once the board file lists it.
    // ###########################################################################################
    [Fact]
    public void A_new_file_is_stored_and_listed_and_leaves_its_place_after_the_write()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string arrived = this.thisWorkspace.WriteFile(Path.Combine(this.DraftFolder, "Datasheets", "U8.pdf"), "the datasheet");

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.WithDatasheet));

        Assert.False(File.Exists(arrived));
        Assert.False(Directory.Exists(Path.Combine(this.DraftFolder, "Datasheets")));
        BoardDocument.StoredFileRow listed = Assert.Single(BoardFiles.ReadDocument(this.WorkbookPath)!.Files);
        Assert.Equal("Datasheets/U8.pdf", listed.Name);
        Assert.Equal("the datasheet", File.ReadAllText(DraftFileResolver.Resolve(string.Empty, this.DraftFolder, "Datasheets/U8.pdf")!));
    }

    [Fact]
    public void A_write_that_cannot_replace_the_board_file_leaves_a_new_file_where_it_arrived()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string arrived = this.thisWorkspace.WriteFile(Path.Combine(this.DraftFolder, "Datasheets", "U8.pdf"), "the datasheet");
        byte[] before = File.ReadAllBytes(this.WorkbookPath);

        bool saved;

        // Readable, but not replaceable - the write fails at its very end.
        using (new FileStream(this.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            saved = DraftWorkbookStore.Edit(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.WithDatasheet);
        }

        Assert.False(saved);
        Assert.Equal(before, File.ReadAllBytes(this.WorkbookPath));
        Assert.Equal("the datasheet", File.ReadAllText(arrived));

        // And the next write takes it in.
        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.WithDatasheet));
        Assert.False(File.Exists(arrived));
        Assert.Equal("the datasheet", File.ReadAllText(DraftFileResolver.Resolve(string.Empty, this.DraftFolder, "Datasheets/U8.pdf")!));
    }

    // ###########################################################################################
    // *** HIGHLIGHTS ARE SAVED, AND THIS TEST EXISTS BECAUSE THEY WERE NOT (found 2026-09-23). ***
    //
    // BoardWorkbookSchema has NO SHEET for ComponentHighlights - BoardDataReader loads them from
    // the JSON sidecar beside the workbook, and BoardWorkbookWriter silently drops them. So an
    // Edit that wrote only the workbook lost every rectangle the label editor had just saved, and
    // nothing anywhere said so: the save reported success, and the highlights were simply gone on
    // the next load.
    //
    // Caught while retiring the delta writers, by noticing that the only remaining caller of
    // BoardComponentHighlightStorage.SaveComponentHighlights was its own test file - which meant
    // nothing in production was writing highlights at all any more.
    // ###########################################################################################
    [Fact]
    public void An_edit_SAVES_component_highlights_which_live_in_the_sidecar()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        bool saved = DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board =>
            {
                var updated = DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU");
                updated.ComponentHighlights.Add(new ComponentHighlightEntry
                {
                    SchematicName = "Sheet 1",
                    BoardLabel = "U8",
                    X = "10",
                    Y = "20",
                    Width = "30",
                    Height = "40",
                });

                return updated;
            });

        Assert.True(saved);

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        ComponentHighlightEntry highlight = Assert.Single(reread!.ComponentHighlights);
        Assert.Equal("U8", highlight.BoardLabel);
        Assert.Equal("10", highlight.X);
    }

    // ###########################################################################################
    // *** AND A CALIBRATION IS NOT ERASED BY AN UNRELATED EDIT. ***
    //
    // Calibrations share that sidecar but are NOT part of BoardData, so a save that wrote the
    // sidecar from the board alone would pass none and wipe them - the contributor would lose
    // their KiCad alignment the next time they touched a label.
    // ###########################################################################################
    [Fact]
    public void An_edit_does_NOT_erase_a_calibration_stored_in_the_same_sidecar()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        // A board with the schematic the calibration belongs to, so it is collectable.
        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board =>
            {
                var updated = DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU");
                updated.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet 1" });

                return updated;
            });

        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.WorkbookPath, "Sheet 1", "B.Cu", 1.0, 2.0, 3.5, 4.0, false, false);

        // An unrelated edit.
        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU (corrected)"));

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "Sheet 1", out _, out _, out _, out double scaleX, out _, out _, out _));

        Assert.Equal(3.5, scaleX);
    }

    // A label fixed in the table keeps the component's rectangles with it (code review, 2026-10-13).
    [Fact]
    public void Relabelling_a_component_in_a_draft_moves_its_highlights()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, board =>
            board.WithComponentHighlights([new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" }])));

        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, board => new BoardData
        {
            RevisionDate = board.RevisionDate,
            Schematics = board.Schematics,
            Components = [DraftWorkbookStoreTests.Component("U08", "CPU")],
            ComponentHighlights = board.ComponentHighlights,
        }));

        BoardData? reloaded = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.Equal("U08", Assert.Single(reloaded!.ComponentHighlights).BoardLabel);
    }

    // ------------------------------------------------------------------ The marker

    // *** "I HAVE LOOKED AT THIS" NO LONGER MOVES THE BASE (code review, 2026-10-13). *** It used to
    // overwrite BaseRevision - which the submission tells the server - so a draft copied in
    // September claimed to be made from October's data. It now quiets the warning and nothing else.
    [Fact]
    public void Marking_drift_seen_keeps_the_base_revision_the_server_is_told()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string markerPath = DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        string before = DraftMarkerStore.Load(markerPath)!.BaseRevision;

        Assert.True(DraftWorkbookStore.MarkDriftSeen(
            this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, "2026-09-20"));

        DraftMarker? marker = DraftMarkerStore.Load(markerPath);

        Assert.Equal(before, marker!.BaseRevision);
        Assert.Equal("2026-09-20", marker.DriftSeenRevision);
        Assert.Equal(DraftWorkbookStoreTests.BoardKey, marker.BoardKey);
    }

    [Fact]
    public void Bringing_up_to_date_records_the_new_base_and_clears_the_seen_revision()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        Assert.True(DraftWorkbookStore.MarkDriftSeen(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, "2026-09-20"));

        Assert.True(DraftWorkbookStore.RecordBroughtUpToDate(
            this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, "2026-10-01", "ABC123"));

        DraftMarker? marker = DraftMarkerStore.Load(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        Assert.Equal("2026-10-01", marker!.BaseRevision);
        Assert.Equal("ABC123", marker.BaseContentHash);
        Assert.Equal(string.Empty, marker.DriftSeenRevision);
        Assert.Equal(DraftWorkbookStoreTests.BoardKey, marker.BoardKey);
    }

    [Fact]
    public void Updating_the_marker_of_a_board_with_NO_draft_answers_false()
    {
        Assert.False(DraftWorkbookStore.MarkDriftSeen(
            this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, "2026-09-20"));
        Assert.False(DraftWorkbookStore.RecordBroughtUpToDate(
            this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, "2026-09-20", "ABC"));
    }

    // ------------------------------------------------------------------ Removing files (RemoveFiles)

    // ###########################################################################################
    // *** REMOVE ON A PUBLISHED KiCad FILE IMPORTED AGAIN PUTS THE PUBLISHED ONE BACK (code review,
    // 2026-10-25). *** The KiCad data window lists such a file as the contributor's own, and Remove
    // means "not my copy" - but it took the name off the draft's list, so the board lost the file at
    // the next submission. A file the published board does not have is still dropped, and the bytes
    // the draft stored for either go with it.
    // ###########################################################################################
    [Fact]
    public void Removing_a_published_KiCad_file_imported_again_puts_the_published_one_back_and_drops_the_rest()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);

        BoardDocument.StoredFileRow mine = StoredFiles.Put(this.DraftFolder, "KiCad data/board.kicad_pcb", "(kicad_pcb mine)");
        BoardDocument.StoredFileRow extra = StoredFiles.Put(this.DraftFolder, "KiCad data/extra.kicad_sch", "(kicad_sch extra)");
        StoredFiles.WriteBoard(boardFile, mine, extra);

        string dataRoot = Path.Combine(this.thisWorkspace.Root, "Data");
        string publishedBoardFile = Path.Combine(dataRoot, "Commodore", "C64", "250407", DataTreeLayout.BoardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(publishedBoardFile)!);
        BoardDocument.StoredFileRow theirs = StoredFiles.Put(dataRoot, "KiCad data/board.kicad_pcb", "(kicad_pcb published)");
        StoredFiles.WriteBoard(publishedBoardFile, theirs);

        Assert.True(DraftWorkbookStore.RemoveFiles(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            ["KiCad data/board.kicad_pcb", "KiCad data/extra.kicad_sch"],
            publishedBoardFile));

        BoardDocument.StoredFileRow listed = Assert.Single(BoardFiles.ReadDocument(boardFile)!.Files);
        Assert.Equal("KiCad data/board.kicad_pcb", listed.Name);
        Assert.Equal(theirs.Sha256, listed.Sha256);

        // The draft's own bytes for both are gone; the published bytes are the tree's.
        Assert.False(File.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(mine.Sha256, mine.Name).Replace('/', Path.DirectorySeparatorChar))));
        Assert.False(File.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(extra.Sha256, extra.Name).Replace('/', Path.DirectorySeparatorChar))));
    }

    // With no published board to go back to, Remove takes the name off the list - as before.
    [Fact]
    public void Removing_a_KiCad_file_with_no_published_board_takes_it_off_the_list()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        StoredFiles.WriteBoard(boardFile, StoredFiles.Put(this.DraftFolder, "KiCad data/board.kicad_pcb", "(kicad_pcb mine)"));

        Assert.True(DraftWorkbookStore.RemoveFiles(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, ["KiCad data/board.kicad_pcb"]));

        Assert.Empty(BoardFiles.ReadDocument(boardFile)!.Files);
    }

    // ###########################################################################################
    // *** A NAME IN THE LIST CAN SAY ANYTHING (code review, 2026-10-26). *** A draft folder put into
    // Drafts by hand, or shared, can list "KiCad data/../../../x" - the KiCad data window lists it
    // as imported, and Remove deleted the file that name climbs to, outside the draft. It comes off
    // the list (written as it was listed, not only as Normalise spells it) and nothing outside the
    // draft folder is touched.
    // ###########################################################################################
    [Fact]
    public void Removing_a_listed_name_that_climbs_out_of_the_draft_deletes_nothing_outside_it()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        string outside = this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "notes.txt"), "the contributor's own notes");
        const string climbing = "KiCad data/../../../notes.txt";
        StoredFiles.WriteBoard(boardFile, StoredFiles.EntryFor(climbing, "anything"));

        Assert.True(DraftWorkbookStore.RemoveFiles(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, [climbing]));

        Assert.Equal("the contributor's own notes", File.ReadAllText(outside));
        Assert.Empty(BoardFiles.ReadDocument(boardFile)!.Files);
    }

    // Remove takes away the bytes of what it removes - the stored ones and a copy lying in the
    // folder - and nothing else: the draft held in memory keeps bytes no row cites while its undo
    // could bring the row back (DraftFileMemory), and a Remove in between deleted them.
    [Fact]
    public void Removing_a_KiCad_file_takes_its_own_bytes_and_leaves_the_rest_of_the_store()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        BoardDocument.StoredFileRow mine = StoredFiles.Put(this.DraftFolder, "KiCad data/board.kicad_pcb", "(kicad_pcb mine)");
        BoardDocument.StoredFileRow remembered = StoredFiles.Put(this.DraftFolder, "Component images/u8.png", "a picture whose row was removed");
        string lying = this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "KiCad data", "sheet.kicad_sch"), "(kicad_sch)");
        StoredFiles.WriteBoard(boardFile, mine);

        Assert.True(DraftWorkbookStore.RemoveFiles(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, ["KiCad data/board.kicad_pcb", "KiCad data/sheet.kicad_sch"]));

        Assert.False(File.Exists(this.StoredIn(mine)));
        Assert.False(File.Exists(lying));
        Assert.True(File.Exists(this.StoredIn(remembered)));
    }

    private string StoredIn(BoardDocument.StoredFileRow file) =>
        Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(file.Sha256, file.Name).Replace('/', Path.DirectorySeparatorChar));

    // ------------------------------------------------------------------ Discarding

    // ###########################################################################################
    // "One folder is the whole draft" - the same model DraftManager.DiscardDraft used.
    //
    // It matters MORE now than it did: the folder holds copied image bytes as well as rows, so
    // leaving them behind would strand megabytes per discarded draft.
    // ###########################################################################################
    [Fact]
    public void Discarding_removes_the_WHOLE_folder_including_copied_files()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string image = Path.Combine(
            DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey),
            "Sheet1.png");

        File.WriteAllText(image, "image bytes");

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        Assert.False(File.Exists(image));
        Assert.False(File.Exists(this.WorkbookPath));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    // ###########################################################################################
    // *** A DISCARD THAT STOPS PART-WAY LEAVES A DRAFT, NOT A HAND-PLACED FOLDER (code review,
    // 2026-09-27). *** The recursive delete removed the marker first and then stopped at the
    // workbook Excel held open - and a folder without a marker is taken in again at the next start
    // (DraftFolderImport), so the discarded draft came back as a new one. The marker goes last now:
    // the stopped discard says so, the folder is still this draft, and nothing re-imports it.
    // ###########################################################################################
    [Fact]
    public void A_discard_stopped_by_an_open_workbook_keeps_the_marker_so_the_folder_is_not_imported_again()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        using (new FileStream(this.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            Assert.False(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
        }

        Assert.True(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
        Assert.Empty(DraftFolderImport.ImportUnmarkedFolders(
            this.DraftsRoot, [new KnownDraftBoard(DraftWorkbookStoreTests.BoardKey, IsPublished: true)], DateTimeOffset.UtcNow));

        // Excel closed: discarding again finishes it.
        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    // ###########################################################################################
    // *** A DISCARD STOPPED BY ANY OTHER FILE LEAVES THE DRAFT WHOLE (owner report, 2026-10-02). ***
    // The folder was emptied in NAME order, and "Data ... .xlsx" sorts before "Issue 4.B.png" - so
    // a published draft being tidied away lost its workbook and sidecar and then stopped at an image
    // something had open. What was left - a marker and three images - could never be retired (there
    // was no workbook to compare) and read as a draft "0 rows changed" with the official data moved
    // under it, for good. The workbook now goes after every other file, so a stop leaves a draft.
    // ###########################################################################################
    [Fact]
    public void A_discard_stopped_by_an_open_image_keeps_the_workbook_and_its_sidecar()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string sidecar = BoardComponentHighlightStorage.GetJsonPath(this.WorkbookPath);
        File.WriteAllText(sidecar, "{}");

        // Sorts AFTER the workbook, as the real board's images did.
        string image = Path.Combine(
            DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey),
            "Issue 4.B.png");

        File.WriteAllText(image, "image bytes");

        // Held the way an ordinary reader holds a file - reading, sharing reads only.
        using (new FileStream(image, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            Assert.False(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
        }

        Assert.True(File.Exists(this.WorkbookPath));
        Assert.True(File.Exists(sidecar));
        Assert.True(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        // Nothing holding it: discarding again finishes it.
        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    // The workbook goes BEFORE its sidecar: held open in Excel - the usual stop - it must not be
    // left without the highlights and calibrations beside it, which would then read as changes.
    [Fact]
    public void A_discard_stopped_by_an_open_workbook_keeps_its_sidecar()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string sidecar = BoardComponentHighlightStorage.GetJsonPath(this.WorkbookPath);
        File.WriteAllText(sidecar, "{}");

        using (new FileStream(this.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            Assert.False(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
        }

        Assert.True(File.Exists(sidecar));
    }

    [Fact]
    public void Discarding_a_board_with_NO_draft_answers_false_rather_than_throwing()
    {
        Assert.False(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    // ###########################################################################################
    // *** NO EMPTY FOLDERS LEFT BEHIND (owner report, 2026-09-27). *** A published draft is tidied
    // away through this discard, and it used to leave "Drafts/Commodore/C64" (and "Drafts/
    // Commodore") standing empty. The folders above a discarded draft go too - while they are empty.
    // ###########################################################################################
    [Fact]
    public void Discarding_the_only_draft_removes_the_folders_it_leaves_empty_but_not_the_root()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        Assert.False(Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore", "C64")));
        Assert.False(Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore")));
        Assert.True(Directory.Exists(this.DraftsRoot));
    }

    [Fact]
    public void A_folder_still_holding_another_draft_or_anything_else_is_kept()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        // Another board of the same hardware, and something else beside the manufacturer's boards.
        string otherBoard = Path.Combine(this.DraftsRoot, "Commodore", "C64", "250425");
        Directory.CreateDirectory(otherBoard);
        File.WriteAllText(Path.Combine(otherBoard, "Data C64 250425.xlsx"), "x");

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        Assert.False(Directory.Exists(DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)));
        Assert.True(Directory.Exists(otherBoard));
        Assert.True(Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore")));
    }

    [Fact]
    public void A_manufacturer_folder_holding_a_shared_files_folder_is_kept()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string shared = Path.Combine(this.DraftsRoot, "Commodore", "Shared files");
        Directory.CreateDirectory(shared);

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        Assert.False(Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore", "C64")));
        Assert.True(Directory.Exists(shared));
    }

    // ------------------------------------------------------------------ Fingerprint / EditIfUnchanged
    //
    // The table editor's guard (2026-09-24): it holds every sheet and a save replaces all of them,
    // so it may only save onto the exact workbook it read. See EditIfUnchanged's header.

    [Fact]
    public void The_fingerprint_is_empty_without_a_draft_and_stable_while_nothing_changes()
    {
        Assert.Equal(string.Empty, DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));

        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string first = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);
        string second = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.Equal(64, first.Length);
        Assert.Equal(first, second);
    }

    [Fact]
    public void The_fingerprint_changes_when_the_workbooks_content_does()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string before = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU 6510"));

        Assert.NotEqual(before, DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    [Fact]
    public void EditIfUnchanged_writes_when_the_fingerprint_still_matches()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        DraftWorkbookEditOutcome outcome = DraftWorkbookStore.EditIfUnchanged(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            fingerprint,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU 6510"));

        Assert.Equal(DraftWorkbookEditOutcome.Saved, outcome);
        Assert.Equal(
            "CPU 6510",
            DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)!.Components.Single().FriendlyName);
    }

    [Fact]
    public void EditIfUnchanged_REFUSES_a_stale_fingerprint_and_writes_nothing()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string stale = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        // Edited behind the caller's back.
        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "Edited in Excel"));

        bool transformRan = false;
        DraftWorkbookEditOutcome outcome = DraftWorkbookStore.EditIfUnchanged(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            stale,
            board =>
            {
                transformRan = true;
                return DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "Stale overwrite");
            });

        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, outcome);
        Assert.False(transformRan);
        Assert.Equal(
            "Edited in Excel",
            DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)!.Components.Single().FriendlyName);
    }

    [Fact]
    public void EditIfUnchanged_says_NoDraft_when_there_is_nothing_to_edit()
    {
        Assert.Equal(
            DraftWorkbookEditOutcome.NoDraft,
            DraftWorkbookStore.EditIfUnchanged(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, "anything", board => board));
    }

    // ------------------------------------------------------------------ Renaming a schematic image
    //
    // Owner request, 2026-10-10: rename an image in the "Schematic images" window. The name is a key
    // in the sidecar too, so the rename has to carry the image's highlights and calibration.

    // A draft with two images, a highlight on each, and a calibration for the first.
    private void CreateDraftWithTwoImages()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board =>
            {
                var updated = DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU");
                updated.Schematics.Add(new BoardSchematicEntry { SchematicName = "Top", CadName = "Open128 - PCB Top", SchematicImageFile = "Commodore/C64/250407/top.png" });
                updated.Schematics.Add(new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Commodore/C64/250407/bottom.png" });
                updated.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "10", Y = "20", Width = "30", Height = "40" });
                updated.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Bottom", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" });

                return updated;
            });

        BoardComponentHighlightStorage.SaveKiCadCalibration(
            this.WorkbookPath, "Top", "Open128 - PCB Top", 1.0, 2.0, 3.5, 4.0, true, false);
    }

    // The "Top" row as the window points at it: its name and its image file.
    private static readonly SchematicRowRef TopRow = new("Top", "Commodore/C64/250407/top.png");

    // ###########################################################################################
    // *** THE CALIBRATION SURVIVES THE RENAME. *** An ordinary Edit collects calibrations under
    // the board's names AFTER the change, and the sidecar still filed this one under the old name -
    // it was dropped. The rename collects them before and moves the one.
    // ###########################################################################################
    [Fact]
    public void Renaming_an_image_carries_its_highlights_and_its_calibration_to_the_new_name()
    {
        this.CreateDraftWithTwoImages();

        Assert.Equal(
            SchematicRenameOutcome.Renamed,
            DraftWorkbookStore.RenameSchematic(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.TopRow, " PCB; Top "));

        BoardData board = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)!;

        Assert.Equal(["PCB; Top", "Bottom"], board.Schematics.Select(entry => entry.SchematicName));
        Assert.Equal("Open128 - PCB Top", board.Schematics[0].CadName);
        Assert.Equal(
            ["Bottom/U8", "PCB; Top/U8"],
            board.ComponentHighlights.Select(highlight => $"{highlight.SchematicName}/{highlight.BoardLabel}").Order());

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "PCB; Top", out string cadName, out _, out _, out double scaleX, out _, out bool mirrorX, out _));
        Assert.Equal("Open128 - PCB Top", cadName);
        Assert.Equal(3.5, scaleX);
        Assert.True(mirrorX);

        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "Top", out _, out _, out _, out _, out _, out _, out _));

        // The board is still a board: its other sheets and its caption untouched.
        Assert.Equal("CPU", board.Components.Single().FriendlyName);
    }

    // The sidecar files by the exact name, so "top" to "Top" has to move the calibration too.
    [Fact]
    public void A_change_of_capitals_alone_keeps_the_calibration()
    {
        this.CreateDraftWithTwoImages();

        Assert.Equal(
            SchematicRenameOutcome.Renamed,
            DraftWorkbookStore.RenameSchematic(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.TopRow, "TOP"));

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "TOP", out _, out _, out _, out double scaleX, out _, out _, out _));
        Assert.Equal(3.5, scaleX);
    }

    [Theory]
    [InlineData("bottom", SchematicRenameOutcome.NameTaken)]
    [InlineData("   ", SchematicRenameOutcome.Blank)]
    [InlineData("Top", SchematicRenameOutcome.Unchanged)]
    public void A_refused_or_needless_rename_writes_nothing(string newName, SchematicRenameOutcome expected)
    {
        this.CreateDraftWithTwoImages();
        string before = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey);

        Assert.Equal(expected, DraftWorkbookStore.RenameSchematic(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.TopRow, newName));

        Assert.Equal(before, DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey));
    }

    [Fact]
    public void Renaming_an_image_the_draft_does_not_have_says_so()
    {
        this.CreateDraftWithTwoImages();

        Assert.Equal(
            SchematicRenameOutcome.NotFound,
            DraftWorkbookStore.RenameSchematic(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, new SchematicRowRef("Gone", "Commodore/C64/250407/gone.png"), "New"));
    }

    // ###########################################################################################
    // *** OF TWO IMAGES SHARING A NAME, THE ONE MEANT IS RENAMED (code review, 2026-10-10). *** The
    // rename went to the FIRST row of the name, whichever the Name box belonged to. The row is found
    // by its image file too now, and what hangs on the name stays on the image that keeps it.
    // ###########################################################################################
    [Fact]
    public void Renaming_the_second_of_two_images_sharing_a_name_renames_that_one()
    {
        this.CreateDraftWithTwoImages();

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => board.WithSchematics(board.Schematics.Select(entry => entry.WithSchematicName("Top")).ToList()));

        Assert.Equal(
            SchematicRenameOutcome.Renamed,
            DraftWorkbookStore.RenameSchematic(
                this.DraftsRoot,
                DraftWorkbookStoreTests.BoardKey,
                new SchematicRowRef("Top", "Commodore/C64/250407/bottom.png"),
                "Bottom"));

        BoardData board = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)!;

        Assert.Equal(["Top", "Bottom"], board.Schematics.Select(entry => entry.SchematicName));
        Assert.Equal("Commodore/C64/250407/bottom.png", board.Schematics[1].SchematicImageFile);
        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "Top", out _, out _, out _, out _, out _, out _, out _));
    }

    // The name is right but the image file is not the one the window showed: changed outside it.
    [Fact]
    public void A_row_whose_image_file_changed_meanwhile_is_not_found()
    {
        this.CreateDraftWithTwoImages();

        Assert.Equal(
            SchematicRenameOutcome.NotFound,
            DraftWorkbookStore.RenameSchematic(
                this.DraftsRoot,
                DraftWorkbookStoreTests.BoardKey,
                new SchematicRowRef("Top", "Commodore/C64/250407/other.png"),
                "New"));
    }

    [Fact]
    public void Renaming_on_a_board_with_NO_draft_says_so()
    {
        Assert.Equal(
            SchematicRenameOutcome.NoDraft,
            DraftWorkbookStore.RenameSchematic(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey, DraftWorkbookStoreTests.TopRow, "New"));
    }

    // ###########################################################################################
    // *** EVERY EDIT FOLLOWS A RENAMED OR DELETED SCHEMATIC (owner request, 2026-10-10 - "And fix
    // all"). *** A "Board schematics" row renamed in the Drafts tab's table, or deleted there or with
    // Remove, used to leave its highlights on a name the board no longer has (an error at submit
    // that nothing in the table can fix) and lose its calibration. Now an edit follows them with the
    // server's rule for a table edit: same image file, new name - a rename; no such twin - deleted.
    // ###########################################################################################

    [Fact]
    public void An_edit_that_renames_a_schematic_row_carries_its_highlights_and_calibration()
    {
        this.CreateDraftWithTwoImages();

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => board.WithSchematics(board.Schematics
                .Select(entry => entry.SchematicName == "Top" ? entry.WithSchematicName("PCB; Top") : entry)
                .ToList()));

        BoardData board = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)!;

        Assert.Equal(
            ["Bottom/U8", "PCB; Top/U8"],
            board.ComponentHighlights.Select(highlight => $"{highlight.SchematicName}/{highlight.BoardLabel}").Order());
        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "PCB; Top", out _, out _, out _, out double scaleX, out _, out _, out _));
        Assert.Equal(3.5, scaleX);
    }

    [Fact]
    public void An_edit_that_deletes_a_schematic_row_takes_its_highlights_and_calibration_with_it()
    {
        this.CreateDraftWithTwoImages();

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => board.WithSchematics(board.Schematics.Where(entry => entry.SchematicName != "Top").ToList()));

        BoardData board = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)!;

        Assert.Equal(["Bottom/U8"], board.ComponentHighlights.Select(highlight => $"{highlight.SchematicName}/{highlight.BoardLabel}"));
        Assert.False(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "Top", out _, out _, out _, out _, out _, out _, out _));

        // And no highlight names a schematic the board does not have - the error the checks refuse.
        Assert.All(board.ComponentHighlights, highlight =>
            Assert.Contains(board.Schematics, schematic => SchematicRename.SameName(schematic.SchematicName, highlight.SchematicName)));
    }

    // An unrelated edit moves nothing.
    [Fact]
    public void An_edit_that_renames_no_schematic_keeps_the_calibration()
    {
        this.CreateDraftWithTwoImages();

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU (corrected)"));

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "Top", out _, out _, out _, out _, out _, out _, out _));
    }

    // The board's published renames, in the draft's copy of its .json, survive every draft write -
    // the writer builds the whole file, and a list it was not handed would be gone.
    [Fact]
    public void A_draft_write_keeps_the_boards_rename_list()
    {
        this.CreateDraftWithTwoImages();

        SchematicRenameRecord[] renames = [new("Old top", "Top", "Commodore/C64/250407/top.png", "2026-October-10")];
        BoardSidecarWriter.Write(
            this.WorkbookPath,
            BoardComponentHighlightStorage.LoadComponentHighlights(this.WorkbookPath),
            DraftBoardSource.CollectCalibrations(this.WorkbookPath, DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.BoardKey)),
            renames);

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.BoardKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU (corrected)"));

        Assert.Equal(renames, BoardComponentHighlightStorage.LoadSchematicRenames(this.WorkbookPath));
    }
}
