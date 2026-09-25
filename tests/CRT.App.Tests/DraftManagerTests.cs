using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for DraftManager - resolving the "Drafts/" root (NewContributeStrategy.md
// Phase 2, session 2a) and mapping a system's ExcelDataFile to its folder under it. Mirrors
// WorklogManagerTests' own coverage of WorklogManager.ResolveExplicitWorkbookRoot and
// DataManagerTests' coverage of DataManager.ResolveDataRoot - same switch-parsing rules, so the
// same cases: the argument itself, quote-stripping, case-insensitivity, first-match-wins,
// unrelated arguments ignored, and no argument given at all.
//
// Its own collection because LoadFrom mutates DraftManager's static _draftsRoot, the same reason
// WorklogManagerTests and DataManagerTests each have their own collection - see CLAUDE.md's Test
// seams section. Never call DraftManager.Load() from a test; LoadFrom is the seam for that.
[Collection("Drafts")]
public sealed class DraftManagerTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    // ------------------------------------------------------------------------ ResolveDraftsRoot

    [Fact]
    public void ResolveDraftsRoot_uses_the_drafts_root_command_line_argument()
    {
        Assert.Equal(
            "somewhere/Drafts",
            DraftManager.ResolveDraftsRoot(new[] { "--drafts-root=somewhere/Drafts" }));
    }

    [Fact]
    public void ResolveDraftsRoot_strips_surrounding_quotes()
    {
        Assert.Equal(
            "my data/Drafts",
            DraftManager.ResolveDraftsRoot(new[] { "--drafts-root=\"my data/Drafts\"" }));
    }

    [Fact]
    public void ResolveDraftsRoot_matches_the_argument_case_insensitively()
    {
        Assert.Equal("X", DraftManager.ResolveDraftsRoot(new[] { "--DRAFTS-ROOT=X" }));
    }

    [Fact]
    public void ResolveDraftsRoot_takes_the_first_matching_argument()
    {
        Assert.Equal(
            "first",
            DraftManager.ResolveDraftsRoot(new[] { "--drafts-root=first", "--drafts-root=second" }));
    }

    [Fact]
    public void ResolveDraftsRoot_ignores_unrelated_arguments()
    {
        Assert.Equal(
            "X",
            DraftManager.ResolveDraftsRoot(new[] { "--verbose", "--drafts-root=X", "--data-root=Y" }));
    }

    [Fact]
    public void ResolveDraftsRoot_falls_back_to_an_appdata_default_with_no_matching_argument()
    {
        string resolved = DraftManager.ResolveDraftsRoot(new[] { "--data-root=X" });

        Assert.Contains(AppConfig.AppFolderName, resolved);
        Assert.Contains(AppConfig.DraftsFolderName, resolved);
    }

    [Fact]
    public void ResolveDraftsRoot_falls_back_to_an_appdata_default_with_no_arguments_at_all()
    {
        Assert.Contains(AppConfig.DraftsFolderName, DraftManager.ResolveDraftsRoot(null));
        Assert.Contains(AppConfig.DraftsFolderName, DraftManager.ResolveDraftsRoot(Array.Empty<string>()));
    }

    // ------------------------------------------------------------------------ LoadFrom

    [Fact]
    public void LoadFrom_creates_the_root_folder_when_it_does_not_exist_yet()
    {
        string root = Path.Combine(this.thisWorkspace.Root, "Drafts");
        Assert.False(Directory.Exists(root));

        DraftManager.LoadFrom(root);

        Assert.True(Directory.Exists(root));
        Assert.Equal(root, DraftManager.DraftsRoot);
    }

    [Fact]
    public void LoadFrom_with_a_blank_path_resets_to_the_unloaded_state_without_touching_the_filesystem()
    {
        string root = Path.Combine(this.thisWorkspace.Root, "Drafts");
        DraftManager.LoadFrom(root);
        Assert.Equal(root, DraftManager.DraftsRoot);

        DraftManager.LoadFrom(string.Empty);

        Assert.Equal(string.Empty, DraftManager.DraftsRoot);
    }

    // ------------------------------------------------------------------------ GetSystemFolder

    [Fact]
    public void GetSystemFolder_maps_an_ExcelDataFile_to_its_manufacturer_hardware_board_folder()
    {
        string root = Path.Combine(this.thisWorkspace.Root, "Drafts");
        DraftManager.LoadFrom(root);

        string folder = DraftManager.GetSystemFolder("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx");

        Assert.Equal(Path.Combine(root, "Commodore", "C64", "250407"), folder);
    }

    [Fact]
    public void GetSystemFolder_returns_empty_when_the_root_has_not_loaded()
    {
        // DraftManager is a static singleton, so an earlier test in this collection may already
        // have loaded a real root - reset to the unloaded state explicitly first (LoadFrom(""))
        // rather than relying on run order.
        DraftManager.LoadFrom(string.Empty);

        Assert.Equal(string.Empty, DraftManager.GetSystemFolder("Commodore/C64/250407/Data.xlsx"));
    }

    [Fact]
    public void GetSystemFolder_returns_empty_for_a_blank_excel_data_file()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        Assert.Equal(string.Empty, DraftManager.GetSystemFolder(string.Empty));
    }

    [Fact]
    public void GetSystemFolder_returns_empty_for_a_path_with_no_folder_segments_to_take()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        // Just a bare file name, no "manufacturer/hardware/board/" prefix to build a folder from.
        Assert.Equal(string.Empty, DraftManager.GetSystemFolder("Data.xlsx"));
    }

    // ###########################################################################################
    // *** THE LoadDraftFor TESTS ARE GONE WITH THE METHOD (Phase 6, 2026-09-23). ***
    //
    // They pinned reading one system's draft.json - null when there was none, the saved draft
    // when there was, null when the drafts root had not loaded. A draft is now a board folder,
    // so there is no draft.json and nothing to read back; the equivalent questions are answered
    // by DraftBoardSource.HasDraft and DraftStatusReader.Resolve, which have their own tests.
    //
    // Deleted rather than rewritten because the method they covered no longer exists - this
    // note is here so their absence reads as a decision.
    // ###########################################################################################

    // ------------------------------------------------------------------------ EnumerateDraftedSystems

    private static HardwareBoardEntry BoardEntry(string excelDataFile) => new() { ExcelDataFile = excelDataFile };

    // ###########################################################################################
    // *** A DRAFT IS A MARKER NOW (Phase 6, 2026-09-23). *** EnumerateDraftedSystems asks whether
    // one exists, not whether a BoardDraft holds rows - see its own header for why the "is it
    // empty" test could not survive the new model.
    // ###########################################################################################
    private static void WriteDraftMarker(string excelDataFile) =>
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { SystemKey = excelDataFile });

    [Fact]
    public void EnumerateDraftedSystems_returns_only_systems_with_a_saved_draft()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string withDraft = "Commodore/C64/250407/Data.xlsx";
        string withoutDraft = "Commodore/C64/250425/Data.xlsx";
        WriteDraftMarker(withDraft);

        var result = DraftManager.EnumerateDraftedSystems(new[] { BoardEntry(withDraft), BoardEntry(withoutDraft) });

        Assert.Single(result);
        Assert.Equal(withDraft, result[0].ExcelDataFile);
    }

    // ###########################################################################################
    // *** THIS ASSERTION WAS INVERTED DELIBERATELY (Phase 6, 2026-09-23). ***
    //
    // It used to assert that a draft holding NO rows was excluded - "an empty BoardDraft is not a
    // system with local changes". That test cannot survive the new model, and the behaviour it
    // described is one we no longer want:
    //
    //   - a draft workbook is a full copy of the published board, so it is never empty;
    //   - asking whether it DIFFERS would mean parsing two workbooks per system just to decide
    //     whether to list a row;
    //   - and the answer would be wrong for the case that matters most - a contributor who has
    //     seeded a draft and not yet edited it still HAS one, and hiding it would leave them no
    //     way to discard it from inside the application.
    //
    // So a draft is listed because it EXISTS. Whether anything in it has changed is the row's
    // own summary to answer.
    // ###########################################################################################
    [Fact]
    public void EnumerateDraftedSystems_lists_a_draft_that_has_not_been_edited_yet()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/250407/Data.xlsx";
        WriteDraftMarker(excelDataFile);

        var result = DraftManager.EnumerateDraftedSystems(new[] { BoardEntry(excelDataFile) });

        Assert.Single(result);
    }

    [Fact]
    public void EnumerateDraftedSystems_returns_empty_for_an_empty_candidate_list()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        Assert.Empty(DraftManager.EnumerateDraftedSystems(Array.Empty<HardwareBoardEntry>()));
    }

    [Fact]
    public void EnumerateDraftedSystems_returns_every_drafted_system_when_several_have_drafts()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string first = "Commodore/C64/250407/Data.xlsx";
        string second = "Commodore/C128/A/Data.xlsx";
        WriteDraftMarker(first);
        WriteDraftMarker(second);

        var result = DraftManager.EnumerateDraftedSystems(new[] { BoardEntry(first), BoardEntry(second) });

        Assert.Equal(2, result.Count);
        Assert.Contains(result, e => e.ExcelDataFile == first);
        Assert.Contains(result, e => e.ExcelDataFile == second);
    }

    // ------------------------------------------------------------------------ DiscardDraft

    [Fact]
    public void DiscardDraft_removes_the_whole_system_folder()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/250407/Data.xlsx";
        string folder = DraftManager.GetSystemFolder(excelDataFile);
        WriteDraftMarker(excelDataFile);
        Assert.True(Directory.Exists(folder));

        DraftManager.DiscardDraft(excelDataFile);

        Assert.False(Directory.Exists(folder));
    }

    [Fact]
    public void DiscardDraft_also_removes_a_file_that_is_not_draft_json_in_the_same_folder()
    {
        // A draft folder holds the board workbook, its sidecar and every copied image - discard
        // must take the WHOLE folder, or megabytes of a contributor's imported files are left
        // behind with nothing pointing at them. "One folder is the whole draft."
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/250407/Data.xlsx";
        string folder = DraftManager.GetSystemFolder(excelDataFile);
        WriteDraftMarker(excelDataFile);
        string otherFile = Path.Combine(folder, "imported-schematic.png");
        File.WriteAllBytes(otherFile, new byte[] { 1, 2, 3 });

        DraftManager.DiscardDraft(excelDataFile);

        Assert.False(File.Exists(otherFile));
    }

    [Fact]
    public void DiscardDraft_on_a_system_with_no_draft_is_harmless()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        // Nothing there means nothing left behind, which is what "true" promises.
        Assert.True(DraftManager.DiscardDraft("Commodore/C64/250407/Data.xlsx"));
    }

    [Fact]
    public void DiscardDraft_reports_success_when_the_folder_is_gone()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/250407/Data.xlsx";
        WriteDraftMarker(excelDataFile);

        Assert.True(DraftManager.DiscardDraft(excelDataFile));
    }

    // ###########################################################################################
    // *** A LOCKED FILE MUST NOT THROW OUT OF A DISCARD (code review, 2026-09-25). ***
    //
    // DiscardDraft used to run its own bare recursive Directory.Delete. On Windows a draft
    // workbook open in Excel is locked, so the delete removed what it could, then threw an
    // IOException out of a fire-and-forget command - unobserved, with the Drafts tab never
    // refreshed and the folder half gone. It now goes through DraftWorkbookStore.Discard, which
    // catches that and answers false so the caller can say so.
    //
    // Windows only: file locks are advisory elsewhere, so there the delete simply succeeds.
    // ###########################################################################################
    [Fact]
    public void DiscardDraft_with_a_LOCKED_file_answers_false_instead_of_throwing()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/250407/Data.xlsx";
        string folder = DraftManager.GetSystemFolder(excelDataFile);
        WriteDraftMarker(excelDataFile);

        string lockedFile = Path.Combine(folder, "Data.xlsx");
        File.WriteAllBytes(lockedFile, new byte[] { 1, 2, 3 });

        bool discarded;

        using (new FileStream(lockedFile, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            discarded = DraftManager.DiscardDraft(excelDataFile);
        }

        Assert.False(discarded);
        Assert.True(File.Exists(lockedFile), "The locked workbook cannot have been deleted.");
    }

    [Fact]
    public void DiscardDraft_with_no_root_loaded_is_harmless()
    {
        DraftManager.LoadFrom(string.Empty);

        DraftManager.DiscardDraft("Commodore/C64/250407/Data.xlsx");
    }

    // ------------------------------------------------- EnumerateDraftOnlySystems (task 9)

    // The one place in the app that walks Drafts/ BACKWARDS - every other path maps a known
    // ExcelDataFile to its folder. It has to, because a brand-new system's identity exists nowhere
    // else yet, and it is what makes a separate registry file unnecessary.
    // A system that exists ONLY as a draft. Since Phase 6 its registration lives in the MARKER,
    // which is also the file that makes the folder a draft at all.
    private static void WriteNewSystemMarker(string hardware, string board, string excelDataFile) =>
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker
            {
                SystemKey = excelDataFile,
                NewSystem = new NewSystemRegistration
                {
                    HardwareName = hardware,
                    BoardName = board,
                    HardwareNotes = "some notes",
                    ExcelDataFile = excelDataFile,
                },
            });

    [Fact]
    public void EnumerateDraftOnlySystems_finds_a_registered_system_and_rebuilds_its_entry()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
        WriteNewSystemMarker("Commodore 64", "MyBoard", excelDataFile);

        var systems = DraftManager.EnumerateDraftOnlySystems();

        var entry = Assert.Single(systems);
        Assert.Equal("Commodore 64", entry.HardwareName);
        Assert.Equal("MyBoard", entry.BoardName);
        Assert.Equal(excelDataFile, entry.ExcelDataFile);
        Assert.Equal("some notes", entry.HardwareNotes);
    }

    // The flag every consumer branches on - DataManager's sync bookkeeping and DataValidator both
    // have to skip these, since their ExcelDataFile names a file that is never on disk.
    [Fact]
    public void EnumerateDraftOnlySystems_marks_what_it_finds_as_draft_only()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
        WriteNewSystemMarker("Commodore 64", "MyBoard", excelDataFile);

        Assert.True(Assert.Single(DraftManager.EnumerateDraftOnlySystems()).IsDraftOnly);
    }

    // An ordinary draft over an already-synced system carries no registration and must NOT produce
    // a drop-down entry - DataManager already lists that system from the main workbook, and a
    // second entry would be a duplicate of it.
    [Fact]
    public void EnumerateDraftOnlySystems_ignores_an_ordinary_draft_with_no_registration()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/250407/Data.xlsx";
        // An ordinary draft over an already-known system: a marker with NO registration.
        WriteDraftMarker(excelDataFile);

        Assert.Empty(DraftManager.EnumerateDraftOnlySystems());
    }

    [Fact]
    public void EnumerateDraftOnlySystems_ignores_a_folder_with_no_draft_file_at_all()
    {
        string draftsRoot = Path.Combine(this.thisWorkspace.Root, "Drafts");
        DraftManager.LoadFrom(draftsRoot);
        Directory.CreateDirectory(Path.Combine(draftsRoot, "Commodore", "C64", "250407"));

        Assert.Empty(DraftManager.EnumerateDraftOnlySystems());
    }

    [Fact]
    public void EnumerateDraftOnlySystems_finds_several_systems_across_different_manufacturers()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        string first = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
        string second = "Amstrad/CPC464/MyOther/Data CPC464 MyOther.xlsx";
        WriteNewSystemMarker("Commodore 64", "MyBoard", first);
        WriteNewSystemMarker("Amstrad CPC464", "MyOther", second);

        var systems = DraftManager.EnumerateDraftOnlySystems();

        Assert.Equal(2, systems.Count);
        Assert.Contains(systems, s => s.BoardName == "MyBoard");
        Assert.Contains(systems, s => s.BoardName == "MyOther");
    }

    [Fact]
    public void EnumerateDraftOnlySystems_is_empty_with_no_root_loaded()
    {
        DraftManager.LoadFrom(string.Empty);

        Assert.Empty(DraftManager.EnumerateDraftOnlySystems());
    }

    [Fact]
    public void EnumerateDraftOnlySystems_is_empty_when_the_drafts_root_does_not_exist_yet()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        Directory.Delete(Path.Combine(this.thisWorkspace.Root, "Drafts"), recursive: true);

        Assert.Empty(DraftManager.EnumerateDraftOnlySystems());
    }

    // THE decisive property behind storing the registration inside draft.json rather than in a
    // separate registry file: DiscardDraft deletes the whole folder, so the registration is retired
    // with it and nothing has to be kept in step. A registry file would have been left pointing at
    // a folder that no longer exists.
    [Fact]
    public void A_discarded_new_system_stops_being_enumerated_with_no_second_file_to_update()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
        WriteNewSystemMarker("Commodore 64", "MyBoard", excelDataFile);
        Assert.Single(DraftManager.EnumerateDraftOnlySystems());

        DraftManager.DiscardDraft(excelDataFile);

        Assert.Empty(DraftManager.EnumerateDraftOnlySystems());
    }

    // A registration with no identity key could not be mapped back to its own folder, so it would
    // be an entry pointing at nothing - skipped rather than surfaced as a broken drop-down row.
    [Fact]
    public void EnumerateDraftOnlySystems_skips_a_registration_with_no_identity_key()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
        string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
        // A registration with no ExcelDataFile names no system, so there is nothing to add to
        // the hardware/board lists - skipped rather than producing an entry that resolves to
        // nothing.
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker
            {
                SystemKey = excelDataFile,
                NewSystem = new NewSystemRegistration { HardwareName = "Commodore 64", BoardName = "MyBoard" },
            });

        Assert.Empty(DraftManager.EnumerateDraftOnlySystems());
    }
}
