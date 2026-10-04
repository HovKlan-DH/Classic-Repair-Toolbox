using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Whether a file a row names is there - the server's view (SuppliedFileLookup) and this
// computer's (DiskFileLookup), which the table's checks use (owner request, 2026-10-02).
//
// The disk lookup finds a file the way a submit finds it (SubmissionFileLocator: the draft's own
// copy first, then the downloaded data), and holds a file in the DOWNLOADED data to its exact
// spelling - the server, on Linux, refuses a path differing only by capitalisation from a published
// one, while Windows finds it either way.
// ###########################################################################################
public sealed class BoardFileLookupTests
{
    private const string Board = "Commodore/C64/250407";

    [Fact]
    public void The_servers_view_compares_exactly_and_says_how_a_near_miss_is_spelled()
    {
        var lookup = new SuppliedFileLookup([$"{Board}/U8.png"]);

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/U8.png").State);
        Assert.Equal(BoardFileState.Missing, lookup.Check($"{Board}/U9.png").State);

        BoardFileLookupResult near = lookup.Check($"{Board}/u8.PNG");
        Assert.Equal(BoardFileState.CaseDiffers, near.State);
        Assert.Equal($"{Board}/U8.png", near.Actual);

        Assert.Equal("in the submission", lookup.Where);
    }

    [Fact]
    public void A_file_in_the_downloaded_data_is_found()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile($"Data/{Board}/U8.png", "x");

        var lookup = new DiskFileLookup(workspace.Path_("Data"), draftSystemFolder: string.Empty);

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/U8.png").State);
        Assert.Equal(BoardFileState.Missing, lookup.Check($"{Board}/U9.png").State);
    }

    // A draft's own file sits in its board folder, without the board's own path in front.
    [Fact]
    public void A_file_in_the_drafts_own_folder_is_found()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile($"Drafts/{Board}/new.png", "x");
        Directory.CreateDirectory(workspace.Path_("Data"));

        var lookup = new DiskFileLookup(workspace.Path_("Data"), workspace.Path_("Drafts", "Commodore", "C64", "250407"));

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/new.png").State);
    }

    // ###########################################################################################
    // *** SPELLED DIFFERENTLY IN THE DOWNLOADED DATA. *** Found by Windows either way, refused by
    // the server - and on a case-sensitive disk not found at all by File.Exists, which is why the
    // real names are read: both say "spelled differently", with the real spelling.
    // ###########################################################################################
    [Fact]
    public void A_file_in_the_downloaded_data_spelled_differently_says_how_it_is_spelled()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile($"Data/{Board}/U8.png", "x");

        var lookup = new DiskFileLookup(workspace.Path_("Data"), draftSystemFolder: string.Empty);
        BoardFileLookupResult result = lookup.Check($"{Board}/u8.png");

        Assert.Equal(BoardFileState.CaseDiffers, result.State);
        Assert.Equal($"{Board}/U8.png", result.Actual);
        Assert.Equal("on this computer", lookup.Where);
    }

    // ###########################################################################################
    // *** A MISSING FILE IS LOOKED FOR AGAIN (code review, 2026-10-04). *** Every answer used to be
    // kept for the lookup's life - the table's - so a "file missing" error stayed on the open table
    // after the picture was put in place, while the draft's row and Submit found it. A file not
    // found is resolved afresh on the next ask, in the downloaded data and in the draft folder.
    // ###########################################################################################
    [Fact]
    public void A_missing_file_is_found_on_the_next_ask_once_it_arrives_in_the_downloaded_data()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile($"Data/{Board}/U8.png", "x");
        var lookup = new DiskFileLookup(workspace.Path_("Data"), draftSystemFolder: string.Empty);

        Assert.Equal(BoardFileState.Missing, lookup.Check($"{Board}/late.png").State);

        workspace.WriteFile($"Data/{Board}/late.png", "x");

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/late.png").State);
    }

    [Fact]
    public void A_missing_file_is_found_on_the_next_ask_once_it_is_put_in_the_draft_folder()
    {
        using var workspace = new TempWorkspace();
        Directory.CreateDirectory(workspace.Path_("Data"));
        string draft = workspace.Path_("Drafts", "Commodore", "C64", "250407");
        Directory.CreateDirectory(draft);
        var lookup = new DiskFileLookup(workspace.Path_("Data"), draft);

        Assert.Equal(BoardFileState.Missing, lookup.Check($"{Board}/new.png").State);

        workspace.WriteFile($"Drafts/{Board}/new.png", "x");

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/new.png").State);
    }

    // ###########################################################################################
    // A file spelled differently in the downloaded data and then put right is seen as right on the
    // next ask: the folder's listing is read again once the folder has been written since.
    // ###########################################################################################
    [Fact]
    public void A_file_whose_spelling_is_put_right_is_found_on_the_next_ask()
    {
        using var workspace = new TempWorkspace();
        string wrong = workspace.WriteFile($"Data/{Board}/u8.png", "x");
        var lookup = new DiskFileLookup(workspace.Path_("Data"), draftSystemFolder: string.Empty);

        Assert.Equal(BoardFileState.CaseDiffers, lookup.Check($"{Board}/U8.png").State);

        // Renamed through a third name, so a case-insensitive disk sees a change too.
        string between = Path.Combine(Path.GetDirectoryName(wrong)!, "between.png");
        File.Move(wrong, between);
        File.Move(between, Path.Combine(Path.GetDirectoryName(wrong)!, "U8.png"));

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/U8.png").State);
    }

    // A FOUND answer is kept for the lookup's life - re-asking thousands of found files after every
    // edit is what the cache exists to avoid - so a file that vanishes is only seen by a new lookup.
    [Fact]
    public void A_found_file_stays_found_for_the_life_of_the_lookup()
    {
        using var workspace = new TempWorkspace();
        string file = workspace.WriteFile($"Data/{Board}/U8.png", "x");
        var lookup = new DiskFileLookup(workspace.Path_("Data"), draftSystemFolder: string.Empty);

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/U8.png").State);

        File.Delete(file);

        Assert.Equal(BoardFileState.Found, lookup.Check($"{Board}/U8.png").State);
        Assert.Equal(BoardFileState.Missing, new DiskFileLookup(workspace.Path_("Data"), string.Empty).Check($"{Board}/U8.png").State);
    }
}
