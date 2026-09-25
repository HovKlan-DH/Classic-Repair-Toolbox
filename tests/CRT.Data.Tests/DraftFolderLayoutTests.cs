using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where everything lives inside a draft folder (NewContributeStrategy.md Phase 6).
//
// The whole point of the layout is that a draft folder is INDISTINGUISHABLE from a published
// board folder, so most of these tests are about the draft using the published names and
// positions rather than draft-specific ones.
//
// Pure path logic - no filesystem access, which is why none of these needs a TempWorkspace.
// ###########################################################################################
public sealed class DraftFolderLayoutTests
{
    private const string Root = @"C:\Drafts";
    private const string System = "Commodore/C64/250407/Data C64 250407.xlsx";

    [Fact]
    public void The_system_folder_mirrors_the_published_tree_layout()
    {
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407"),
            DraftFolderLayout.GetSystemFolder(Root, System));
    }

    // ###########################################################################################
    // *** THE SAME FILE NAME THE PUBLISHED SYSTEM USES. ***
    //
    // Not "Data C64 250407 (draft).xlsx". The request was that the folder be indistinguishable
    // from a real board, and a renamed workbook would announce itself as something else - it would
    // also break BoardDataReader's cache key, which IS the ExcelDataFile identity.
    // ###########################################################################################
    [Fact]
    public void The_workbook_keeps_the_PUBLISHED_file_name()
    {
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", "Data C64 250407.xlsx"),
            DraftFolderLayout.GetWorkbookPath(Root, System));
    }

    [Fact]
    public void The_sidecar_sits_beside_the_workbook_with_a_json_extension()
    {
        // Derived through BoardComponentHighlightStorage.GetJsonPath, so there is one definition
        // of "the sidecar beside this workbook" rather than a second extension swap here.
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", "Data C64 250407.json"),
            DraftFolderLayout.GetSidecarPath(Root, System));
    }

    // Beside the workbook, with the exact name published boards use - the Wiki's
    // Scope-baseline-folder page names it, so a different spelling here would be a folder no
    // published board has.
    [Fact]
    public void The_scope_baseline_folder_sits_beside_the_workbook_with_the_published_name()
    {
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", "Scope baseline"),
            DraftFolderLayout.GetScopeBaselineFolder(Root, System));

        Assert.Equal(
            Path.GetDirectoryName(DraftFolderLayout.GetWorkbookPath(Root, System)),
            Path.GetDirectoryName(DraftFolderLayout.GetScopeBaselineFolder(Root, System)));
    }

    // ###########################################################################################
    // The folder created up front must be the SAME folder a row's baseline image resolves into.
    // Asserted through GetReferencedFilePath rather than a second hand-built path, so the created
    // folder and the file resolution cannot drift apart without this failing.
    // ###########################################################################################
    [Fact]
    public void A_baseline_image_row_resolves_INTO_the_scope_baseline_folder()
    {
        string resolved = DraftFolderLayout.GetReferencedFilePath(
            Root,
            System,
            "Commodore/C64/250407/Scope baseline/U1_1_PAL.png");

        Assert.Equal(
            DraftFolderLayout.GetScopeBaselineFolder(Root, System),
            Path.GetDirectoryName(resolved));
    }

    [Fact]
    public void The_scope_baseline_folder_is_empty_when_the_system_cannot_be_resolved()
    {
        Assert.Equal(string.Empty, DraftFolderLayout.GetScopeBaselineFolder(string.Empty, System));
        Assert.Equal(string.Empty, DraftFolderLayout.GetScopeBaselineFolder(Root, "no-folder.xlsx"));
    }

    [Fact]
    public void The_marker_sits_at_the_system_folder_root()
    {
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", ".crt-draft.json"),
            DraftFolderLayout.GetMarkerPath(Root, System));
    }

    // ###########################################################################################
    // *** THE ONE GENUINELY SUBTLE RULE IN THIS CLASS. ***
    //
    // A BoardData row stores a file as a path from the DATA ROOT
    // ("Commodore/C64/250407/Sheet1.png"), because that is what the published tree needs. Inside a
    // draft folder the system's own three segments ARE the folder, so they have to be stripped -
    // otherwise the bytes land at "<draft>/Commodore/C64/250407/Commodore/C64/250407/Sheet1.png"
    // and every image in the draft points at nothing.
    // ###########################################################################################
    [Fact]
    public void A_referenced_file_has_the_systems_OWN_prefix_stripped()
    {
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", "Sheet1.png"),
            DraftFolderLayout.GetReferencedFilePath(Root, System, "Commodore/C64/250407/Sheet1.png"));
    }

    [Fact]
    public void A_referenced_file_in_a_SUBFOLDER_keeps_its_subfolder()
    {
        // "KiCad data/" and similar. Stripping the system prefix must not flatten what is below it.
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", "KiCad data", "board.kicad_pcb"),
            DraftFolderLayout.GetReferencedFilePath(
                Root, System, "Commodore/C64/250407/KiCad data/board.kicad_pcb"));
    }

    // ###########################################################################################
    // *** A SHARED FILE IS DELIBERATELY NOT GIVEN A DRAFT LOCATION. ***
    //
    // A manufacturer "Shared files" image is referenced by many boards. Copying one into a draft
    // would FORK it, and a later edit to the draft's copy would silently fail to reach the boards
    // that actually share it. Empty means "leave it where it is" - it keeps resolving against
    // Data/ exactly as it always did.
    // ###########################################################################################
    [Fact]
    public void A_file_belonging_to_ANOTHER_board_gets_no_draft_location()
    {
        Assert.Equal(
            string.Empty,
            DraftFolderLayout.GetReferencedFilePath(Root, System, "Commodore/Shared files/7805.jpg"));

        Assert.Equal(
            string.Empty,
            DraftFolderLayout.GetReferencedFilePath(Root, System, "Generic shared files/74LS00.png"));
    }

    [Fact]
    public void A_file_under_a_DIFFERENT_board_of_the_same_hardware_is_also_shared()
    {
        // Same manufacturer and hardware, different board - still not this system's to own.
        Assert.Equal(
            string.Empty,
            DraftFolderLayout.GetReferencedFilePath(Root, System, "Commodore/C64/250425/Sheet1.png"));
    }

    // ###########################################################################################
    // Case-insensitive, matching how the rest of the app compares these paths. A workbook
    // hand-edited to say "commodore/..." names the same folder on Windows, and refusing to
    // recognise it would silently treat every one of that board's own files as shared - so they
    // would all stop being copied, with no error anywhere.
    // ###########################################################################################
    [Fact]
    public void The_system_prefix_is_matched_case_insensitively()
    {
        Assert.Equal(
            Path.Combine(Root, "Commodore", "C64", "250407", "Sheet1.png"),
            DraftFolderLayout.GetReferencedFilePath(Root, System, "commodore/c64/250407/Sheet1.png"));
    }

    [Fact]
    public void An_unusable_input_resolves_to_EMPTY_rather_than_a_bogus_path()
    {
        // Callers treat empty as "no draft folder available". A bogus path would be far worse:
        // it would be created, and files would be written somewhere nobody looks.
        Assert.Equal(string.Empty, DraftFolderLayout.GetSystemFolder(string.Empty, System));
        Assert.Equal(string.Empty, DraftFolderLayout.GetSystemFolder(Root, string.Empty));
        Assert.Equal(string.Empty, DraftFolderLayout.GetSystemFolder(Root, "loose-file.xlsx"));
        Assert.Equal(string.Empty, DraftFolderLayout.GetReferencedFilePath(Root, System, null));
    }

    // ###########################################################################################
    // The marker must never be treated as board content - see DraftFolderLayout's own header for
    // the two independent mechanisms keeping it out of a submission. This is the second one.
    // ###########################################################################################
    [Fact]
    public void The_MARKER_is_recognised_as_draft_only_wherever_it_is_found()
    {
        Assert.True(DraftFolderLayout.IsDraftOnlyFile(".crt-draft.json"));

        // A full path in THIS machine's own form, as the app hands it over. Built with Path.Combine,
        // never a literal @"C:\..." - on the Linux CI runner a backslash is not a separator, so the
        // whole literal read as one file name and the test failed there alone.
        Assert.True(DraftFolderLayout.IsDraftOnlyFile(
            Path.Combine(Path.GetTempPath(), "Drafts", "Commodore", "C64", "250407", ".crt-draft.json")));

        // Case-insensitively, since the filesystem is.
        Assert.True(DraftFolderLayout.IsDraftOnlyFile(".CRT-Draft.JSON"));
    }

    [Fact]
    public void Ordinary_board_files_are_NOT_draft_only()
    {
        Assert.False(DraftFolderLayout.IsDraftOnlyFile("Data C64 250407.xlsx"));
        Assert.False(DraftFolderLayout.IsDraftOnlyFile("Data C64 250407.json"));
        Assert.False(DraftFolderLayout.IsDraftOnlyFile("Sheet1.png"));
        Assert.False(DraftFolderLayout.IsDraftOnlyFile(string.Empty));
    }
}
