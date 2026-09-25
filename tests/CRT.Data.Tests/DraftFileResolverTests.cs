using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which of the two roots a BoardData entry's File value resolves against.
//
// *** PHASE 6 CHANGED BOTH THE PATH AND THE ORDER (2026-09-23). *** A drafted file used to sit
// under a "Files" subfolder with the stored path appended whole, and the official copy was
// tried FIRST because a draft was an overlay - the published bytes were the real file and the
// draft folder held only newly attached extras.
//
// A draft is now a complete board folder, so:
//   - the file sits where a published board keeps it, which means the system's own
//     "Manufacturer/Hardware/Board/" prefix is STRIPPED (the draft folder already is those
//     segments);
//   - the DRAFT is tried first, because when one exists its files ARE the board's files.
//
// Rewritten rather than deleted: every behaviour pinned here still matters, but the paths and
// the precedence have deliberately moved.
// ###########################################################################################
public sealed class DraftFileResolverTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");
    private string DraftSystemFolder => Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C64", "250407");

    // ------------------------------------------------------------------------ Resolve

    [Fact]
    public void A_blank_relative_file_resolves_to_null()
    {
        Assert.Null(DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, string.Empty));
        Assert.Null(DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, "   "));
        Assert.Null(DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, null));
    }

    [Fact]
    public void An_official_file_with_no_drafted_copy_is_found_under_the_data_root()
    {
        string officialPath = Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(officialPath)!);
        File.WriteAllText(officialPath, "official");

        var resolved = DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/U8.pdf");

        Assert.Equal(officialPath, resolved);
    }

    [Fact]
    public void A_drafted_only_file_is_found_at_the_draft_folders_own_root()
    {
        string draftedPath = Path.Combine(this.DraftSystemFolder, "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(draftedPath)!);
        File.WriteAllText(draftedPath, "drafted");

        var resolved = DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/U8.pdf");

        Assert.Equal(draftedPath, resolved);
    }

    // ###########################################################################################
    // *** THE DRAFTED COPY WINS, AND THIS ASSERTION WAS FLIPPED DELIBERATELY. ***
    //
    // It used to assert the official copy won, which was right when a draft was an overlay.
    // Now a draft is a full copy of the board, so a drafted file is a REPLACEMENT for the
    // published one - returning the published bytes would show the contributor the old picture
    // for an image they have already replaced, with nothing to say why.
    // ###########################################################################################
    [Fact]
    public void The_DRAFTED_copy_wins_when_both_exist()
    {
        string officialPath = Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(officialPath)!);
        File.WriteAllText(officialPath, "official");

        string draftedPath = Path.Combine(this.DraftSystemFolder, "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(draftedPath)!);
        File.WriteAllText(draftedPath, "drafted");

        var resolved = DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/U8.pdf");

        Assert.Equal(draftedPath, resolved);
    }

    [Fact]
    public void Neither_copy_existing_resolves_to_null()
    {
        Assert.Null(DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/Missing.pdf"));
    }

    [Fact]
    public void A_blank_data_root_still_finds_the_drafted_copy()
    {
        string draftedPath = Path.Combine(this.DraftSystemFolder, "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(draftedPath)!);
        File.WriteAllText(draftedPath, "drafted");

        var resolved = DraftFileResolver.Resolve(
            string.Empty, this.DraftSystemFolder, "Commodore/C64/250407/U8.pdf");

        Assert.Equal(draftedPath, resolved);
    }

    // ###########################################################################################
    // A file belonging to ANOTHER board - a manufacturer "Shared files" image - has no drafted
    // location at all, so it resolves against Data/ even when a draft exists. Copying one into
    // a draft would fork it, and a later edit would silently not reach the boards sharing it.
    // ###########################################################################################
    [Fact]
    public void A_SHARED_file_resolves_against_the_data_root_even_with_a_draft()
    {
        string sharedPath = Path.Combine(this.DataRoot, "Commodore", "Shared files", "7805.jpg");
        Directory.CreateDirectory(Path.GetDirectoryName(sharedPath)!);
        File.WriteAllText(sharedPath, "shared");

        var resolved = DraftFileResolver.Resolve(
            this.DataRoot, this.DraftSystemFolder, "Commodore/Shared files/7805.jpg");

        Assert.Equal(sharedPath, resolved);
    }

    [Fact]
    public void A_blank_draft_folder_never_throws_and_still_finds_the_official_copy()
    {
        string officialPath = Path.Combine(this.DataRoot, "U8.pdf");
        Directory.CreateDirectory(this.DataRoot);
        File.WriteAllText(officialPath, "official");

        var resolved = DraftFileResolver.Resolve(this.DataRoot, string.Empty, "U8.pdf");

        Assert.Equal(officialPath, resolved);
    }

    [Fact]
    public void Forward_slashes_in_the_relative_path_are_normalized()
    {
        string officialPath = Path.Combine(this.DataRoot, "Sub", "Folder", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(officialPath)!);
        File.WriteAllText(officialPath, "official");

        var resolved = DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, "Sub/Folder/U8.pdf");

        Assert.Equal(officialPath, resolved);
    }

    // ------------------------------------------------------------------------ ResolveWithSource

    // ###########################################################################################
    // *** THE ROOT IS THE ONE THE PATH REALLY CAME FROM - including when BOTH copies exist. ***
    //
    // A caller that opens the file hands this root to ExternalTargetLauncher, which refuses a
    // path outside it. The component popup used to work the root out for itself, "drafted only
    // when the official copy is missing" - the old official-first order - so with both copies
    // present it paired the DRAFT path with the DATA root and the open was refused. Every file a
    // seeded draft copies is in both trees, so that was every datasheet on every drafted board.
    // ###########################################################################################
    [Fact]
    public void With_both_copies_the_resolution_names_the_DRAFT_FOLDER_as_its_root()
    {
        string officialPath = Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(officialPath)!);
        File.WriteAllText(officialPath, "official");

        string draftedPath = Path.Combine(this.DraftSystemFolder, "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(draftedPath)!);
        File.WriteAllText(draftedPath, "drafted");

        DraftFileResolution? resolved = DraftFileResolver.ResolveWithSource(
            this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/U8.pdf");

        Assert.NotNull(resolved);
        Assert.Equal(draftedPath, resolved!.FullPath);
        Assert.Equal(this.DraftSystemFolder, resolved.Root);
        Assert.True(resolved.IsDrafted);
    }

    [Fact]
    public void An_official_only_file_names_the_DATA_ROOT_as_its_root()
    {
        string officialPath = Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Datasheets", "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(officialPath)!);
        File.WriteAllText(officialPath, "official");

        DraftFileResolution? resolved = DraftFileResolver.ResolveWithSource(
            this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/U8.pdf");

        Assert.NotNull(resolved);
        Assert.Equal(officialPath, resolved!.FullPath);
        Assert.Equal(this.DataRoot, resolved.Root);
        Assert.False(resolved.IsDrafted);
    }

    [Fact]
    public void Resolve_and_ResolveWithSource_always_agree_on_the_path()
    {
        // Resolve is now a projection of ResolveWithSource. Pinned so a later edit cannot give
        // the two their own copies of the precedence rule again.
        string draftedPath = Path.Combine(this.DraftSystemFolder, "U8.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(draftedPath)!);
        File.WriteAllText(draftedPath, "drafted");

        const string stored = "Commodore/C64/250407/U8.pdf";

        Assert.Equal(
            DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, stored),
            DraftFileResolver.ResolveWithSource(this.DataRoot, this.DraftSystemFolder, stored)!.FullPath);

        Assert.Null(DraftFileResolver.ResolveWithSource(this.DataRoot, this.DraftSystemFolder, "Commodore/C64/250407/None.pdf"));
    }

    // ------------------------------------------------------------------------ BuildDraftFileDestination

    [Fact]
    public void BuildDraftFileDestination_strips_the_systems_own_prefix()
    {
        // The draft folder already IS "Commodore/C64/250407", so an unstripped combine would
        // write to "<draft>/Commodore/C64/250407/Commodore/C64/250407/..." - somewhere Resolve
        // never looks, so the file would be written and then never found again.
        string destination = DraftFileResolver.BuildDraftFileDestination(
            this.DraftSystemFolder, "Commodore/C64/250407/Datasheets/U8.pdf");

        Assert.Equal(Path.Combine(this.DraftSystemFolder, "Datasheets", "U8.pdf"), destination);
    }

    // The write and read sides must agree, which is the only thing that actually matters here.
    [Fact]
    public void BuildDraftFileDestination_and_Resolve_agree_on_where_a_drafted_file_lives()
    {
        const string stored = "Commodore/C64/250407/Datasheets/U8.pdf";

        string destination = DraftFileResolver.BuildDraftFileDestination(this.DraftSystemFolder, stored);
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        File.WriteAllText(destination, "drafted");

        Assert.Equal(destination, DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, stored));
    }

    // ###########################################################################################
    // *** A SHARED FILE THE CONTRIBUTOR ATTACHED is found where it was written (maintainer report,
    // 2026-09-25). *** The component editor lets a new image be filed in a shared folder
    // ("Commodore/Shared files/Component images"), and BuildDraftFileDestination writes it inside
    // the draft under that whole path. Resolve only ever looked in the draft for the board's OWN
    // files, so the new image was on disk and nothing could find it - the submission then refused
    // it as missing. A published shared file with no drafted copy still comes from Data/ (above).
    // ###########################################################################################
    [Fact]
    public void A_SHARED_file_the_contributor_attached_is_found_where_it_was_written()
    {
        const string stored = "Commodore/Shared files/Component images/HotCPU.png";

        string destination = DraftFileResolver.BuildDraftFileDestination(this.DraftSystemFolder, stored);
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        File.WriteAllText(destination, "drafted");

        Assert.Equal(destination, DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, stored));
    }

    // Attaching a file under the name of a published shared file REPLACES it - the contributor's
    // bytes are the ones shown and submitted, exactly as for the board's own files.
    [Fact]
    public void A_SHARED_file_the_contributor_attached_wins_over_the_published_copy()
    {
        const string stored = "Commodore/Shared files/7805.jpg";
        string published = Path.Combine(this.DataRoot, "Commodore", "Shared files", "7805.jpg");
        Directory.CreateDirectory(Path.GetDirectoryName(published)!);
        File.WriteAllText(published, "published");

        string destination = DraftFileResolver.BuildDraftFileDestination(this.DraftSystemFolder, stored);
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        File.WriteAllText(destination, "drafted");

        Assert.Equal(destination, DraftFileResolver.Resolve(this.DataRoot, this.DraftSystemFolder, stored));
    }

    [Fact]
    public void BuildDraftFileDestination_does_not_create_anything_on_disk()
    {
        string destination = DraftFileResolver.BuildDraftFileDestination(this.DraftSystemFolder, "U8.pdf");

        Assert.False(File.Exists(destination));
        Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
    }

    [Fact]
    public void BuildDraftFileDestination_trims_and_normalizes_slashes()
    {
        string destination = DraftFileResolver.BuildDraftFileDestination(
            this.DraftSystemFolder, "  Commodore/C64/250407/Sub/Folder/U8.pdf  ");

        Assert.Equal(Path.Combine(this.DraftSystemFolder, "Sub", "Folder", "U8.pdf"), destination);
    }
}
