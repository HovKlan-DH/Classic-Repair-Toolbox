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

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    private string WorkbookPath =>
        DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

    // Creates a draft holding the given components, the way DraftSeeder would.
    private void CreateDraft(params ComponentEntry[] components)
    {
        var board = new BoardData { RevisionDate = "2026-09-01" };
        board.Components.AddRange(components);

        Directory.CreateDirectory(Path.GetDirectoryName(this.WorkbookPath)!);
        BoardWorkbookWriter.Write(this.WorkbookPath, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey),
            new DraftMarker
            {
                SystemKey = DraftWorkbookStoreTests.SystemKey,
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

    // ------------------------------------------------------------------ Reading

    [Fact]
    public void The_draft_board_is_read_back_from_its_workbook()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        BoardData? board = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        Assert.NotNull(board);
        Assert.Equal("CPU", board!.Components.Single().FriendlyName);
    }

    [Fact]
    public void With_NO_draft_the_board_is_null_rather_than_empty()
    {
        // Null and "an empty board" mean very different things: the caller must be able to tell
        // "there is no draft" from "the draft has no rows", because the second is a real state a
        // contributor can create by emptying one.
        Assert.Null(DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
    }

    // ------------------------------------------------------------------ Writing

    [Fact]
    public void An_edit_is_written_back_to_the_workbook()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        bool saved = DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU (corrected)"));

        Assert.True(saved);

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);
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
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey),
            new DraftMarker { SystemKey = DraftWorkbookStoreTests.SystemKey });

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
            current => DraftWorkbookStoreTests.WithFriendlyName(current, "U8", "CPU (corrected)"));

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

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
            DraftWorkbookStoreTests.SystemKey,
            current => DraftWorkbookStoreTests.WithFriendlyName(current, "U8", "CPU (edited in app)"));

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        Assert.Equal("CPU (edited in app)", reread!.Components.Single(c => c.BoardLabel == "U8").FriendlyName);
        Assert.Equal("VIC (edited in Excel)", reread.Components.Single(c => c.BoardLabel == "U9").FriendlyName);
    }

    // ###########################################################################################
    // Editing a system with no draft REFUSES rather than creating one.
    //
    // Creating a draft means copying the published board (DraftSeeder). Quietly starting an empty
    // one here would produce a draft that compares as "every published row deleted" - and a
    // submission built from it would ask the server to delete the board.
    // ###########################################################################################
    [Fact]
    public void Editing_a_system_with_NO_draft_refuses_rather_than_creating_one()
    {
        bool saved = DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
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
            DraftWorkbookStoreTests.SystemKey,
            _ => throw new InvalidOperationException("the editor changed its mind")));

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        Assert.NotNull(reread);
        Assert.Equal("CPU", reread!.Components.Single().FriendlyName);
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
            DraftWorkbookStoreTests.SystemKey,
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

        BoardData? reread = DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

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
            DraftWorkbookStoreTests.SystemKey,
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
            DraftWorkbookStoreTests.SystemKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU (corrected)"));

        Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
            this.WorkbookPath, "Sheet 1", out _, out _, out _, out double scaleX, out _, out _, out _));

        Assert.Equal(3.5, scaleX);
    }

    // ------------------------------------------------------------------ The marker

    [Fact]
    public void Rebasing_updates_the_base_revision_and_keeps_everything_else()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        Assert.True(DraftWorkbookStore.SetBaseRevision(
            this.DraftsRoot, DraftWorkbookStoreTests.SystemKey, "2026-09-20"));

        DraftMarker? marker = DraftMarkerStore.Load(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));

        Assert.Equal("2026-09-20", marker!.BaseRevision);
        Assert.Equal(DraftWorkbookStoreTests.SystemKey, marker.SystemKey);
    }

    [Fact]
    public void Rebasing_a_system_with_NO_draft_answers_false()
    {
        Assert.False(DraftWorkbookStore.SetBaseRevision(
            this.DraftsRoot, DraftWorkbookStoreTests.SystemKey, "2026-09-20"));
    }

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
            DraftFolderLayout.GetSystemFolder(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey),
            "Sheet1.png");

        File.WriteAllText(image, "image bytes");

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));

        Assert.False(File.Exists(image));
        Assert.False(File.Exists(this.WorkbookPath));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
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
            Assert.False(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
        }

        Assert.True(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
        Assert.Empty(DraftFolderImport.ImportUnmarkedFolders(
            this.DraftsRoot, [new KnownDraftSystem(DraftWorkbookStoreTests.SystemKey, IsPublished: true)], DateTimeOffset.UtcNow));

        // Excel closed: discarding again finishes it.
        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
    }

    [Fact]
    public void Discarding_a_system_with_NO_draft_answers_false_rather_than_throwing()
    {
        Assert.False(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
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

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));

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

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));

        Assert.False(Directory.Exists(DraftFolderLayout.GetSystemFolder(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey)));
        Assert.True(Directory.Exists(otherBoard));
        Assert.True(Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore")));
    }

    [Fact]
    public void A_manufacturer_folder_holding_a_shared_files_folder_is_kept()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string shared = Path.Combine(this.DraftsRoot, "Commodore", "Shared files");
        Directory.CreateDirectory(shared);

        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));

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
        Assert.Equal(string.Empty, DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));

        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));

        string first = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);
        string second = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        Assert.Equal(64, first.Length);
        Assert.Equal(first, second);
    }

    [Fact]
    public void The_fingerprint_changes_when_the_workbooks_content_does()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string before = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU 6510"));

        Assert.NotEqual(before, DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey));
    }

    [Fact]
    public void EditIfUnchanged_writes_when_the_fingerprint_still_matches()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string fingerprint = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        DraftWorkbookEditOutcome outcome = DraftWorkbookStore.EditIfUnchanged(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
            fingerprint,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "CPU 6510"));

        Assert.Equal(DraftWorkbookEditOutcome.Saved, outcome);
        Assert.Equal(
            "CPU 6510",
            DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey)!.Components.Single().FriendlyName);
    }

    [Fact]
    public void EditIfUnchanged_REFUSES_a_stale_fingerprint_and_writes_nothing()
    {
        this.CreateDraft(DraftWorkbookStoreTests.Component("U8", "CPU"));
        string stale = DraftWorkbookStore.Fingerprint(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey);

        // Edited behind the caller's back.
        DraftWorkbookStore.Edit(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
            board => DraftWorkbookStoreTests.WithFriendlyName(board, "U8", "Edited in Excel"));

        bool transformRan = false;
        DraftWorkbookEditOutcome outcome = DraftWorkbookStore.EditIfUnchanged(
            this.DraftsRoot,
            DraftWorkbookStoreTests.SystemKey,
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
            DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey)!.Components.Single().FriendlyName);
    }

    [Fact]
    public void EditIfUnchanged_says_NoDraft_when_there_is_nothing_to_edit()
    {
        Assert.Equal(
            DraftWorkbookEditOutcome.NoDraft,
            DraftWorkbookStore.EditIfUnchanged(this.DraftsRoot, DraftWorkbookStoreTests.SystemKey, "anything", board => board));
    }
}
