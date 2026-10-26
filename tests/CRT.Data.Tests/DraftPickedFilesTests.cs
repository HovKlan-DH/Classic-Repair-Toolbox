using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftPickedFiles - the files an editor picked, copied into a draft folder under their names
// before the change citing them is saved, and taken back when that save is not made (code review,
// 2026-10-26): left lying there, the next write by any editor took them in, so a picture replaced
// under a name a row already cites entered the draft after "could not be saved".
// ###########################################################################################
public sealed class DraftPickedFilesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DraftFolder => Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C64", "250407");

    private string InDraft(string name) => Path.Combine(this.DraftFolder, name.Replace('/', Path.DirectorySeparatorChar));

    private string Picked(string content) => this.thisWorkspace.WriteFile(Path.Combine("Pictures", $"{Guid.NewGuid():N}.png"), content);

    [Fact]
    public void A_picked_file_is_copied_to_where_the_draft_takes_it_in()
    {
        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(this.Picked("new picture"), "Component images/U8.png")]);

        Assert.Equal("new picture", File.ReadAllText(this.InDraft("Component images/U8.png")));
        Assert.True(copies.AnyLying);
    }

    [Fact]
    public void Taken_back_the_copy_goes_and_nothing_is_left_in_the_draft()
    {
        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(this.Picked("new picture"), "Component images/U8.png")]);

        copies.TakeBack();

        Assert.False(File.Exists(this.InDraft("Component images/U8.png")));
        Assert.False(copies.AnyLying);
        Assert.Empty(Directory.EnumerateFiles(this.DraftFolder, "*", SearchOption.AllDirectories));
    }

    // A file already lying under that name - the contributor's, picked before and not taken in yet -
    // comes back as it was.
    [Fact]
    public void Taken_back_a_file_the_copy_replaced_comes_back()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Component images", "U8.png"), "the earlier picture");

        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(this.Picked("new picture"), "Component images/U8.png")]);
        Assert.Equal("new picture", File.ReadAllText(this.InDraft("Component images/U8.png")));

        copies.TakeBack();

        Assert.Equal("the earlier picture", File.ReadAllText(this.InDraft("Component images/U8.png")));
        Assert.False(Directory.Exists(this.InDraft(DraftPickedFiles.SetAsideFolder)));
    }

    // Kept, what the copy replaced is let go of - the save took the new file in.
    [Fact]
    public void Kept_the_file_a_copy_replaced_is_let_go_of()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Component images", "U8.png"), "the earlier picture");

        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(this.Picked("new picture"), "Component images/U8.png")]);
        copies.Keep();

        Assert.Equal("new picture", File.ReadAllText(this.InDraft("Component images/U8.png")));
        Assert.False(Directory.Exists(this.InDraft(DraftPickedFiles.SetAsideFolder)));
    }

    // The set-aside folder is CRT's own bookkeeping, never a file of the contributor's.
    [Fact]
    public void The_set_aside_folder_is_bookkeeping()
    {
        Assert.True(DraftFolderLayout.IsInBookkeepingFolder($"{DraftPickedFiles.SetAsideFolder}/abc"));
    }

    // A file picked from where it would be copied to is no copy at all - and never deleted by a take-back.
    [Fact]
    public void A_file_picked_from_where_it_lies_is_left_alone()
    {
        string lying = this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Component images", "U8.png"), "already here");

        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(lying, "Component images/U8.png")]);
        copies.TakeBack();

        Assert.Equal("already here", File.ReadAllText(lying));
    }

    // ###########################################################################################
    // A save made, but a copy the write did NOT take in - still lying where it was put - goes as the
    // copy of a save not made does, what it replaced coming back; one taken in is kept (code review,
    // 2026-10-26: Keep let go of both, and the stray copy waited for the next write by any editor).
    // ###########################################################################################
    [Fact]
    public void Kept_after_a_save_only_the_copies_the_write_took_in_stay()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Component local files", "other.pdf"), "the earlier file");

        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [
            (this.Picked("taken in"), "Component images/U8.png"),
            (this.Picked("not taken in"), "Component local files/other.pdf"),
        ]);

        // The write took the picture in - and removed it from where it arrived.
        File.Delete(this.InDraft("Component images/U8.png"));

        Assert.False(copies.KeepTakenIn());

        Assert.Equal("the earlier file", File.ReadAllText(this.InDraft("Component local files/other.pdf")));
        Assert.False(Directory.Exists(Path.Combine(this.DraftFolder, DraftPickedFiles.SetAsideFolder)));
        Assert.False(copies.AnyLying);
    }

    [Fact]
    public void Kept_after_a_save_that_took_every_copy_in_says_so()
    {
        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(this.Picked("taken in"), "Component images/U8.png")]);
        File.Delete(this.InDraft("Component images/U8.png"));

        Assert.True(copies.KeepTakenIn());
    }

    // A picked file that is gone by the time it is copied is skipped, as before.
    [Fact]
    public void A_picked_file_that_is_gone_is_skipped()
    {
        DraftPickedFiles copies = DraftPickedFiles.CopyInto(this.DraftFolder, [(Path.Combine(this.thisWorkspace.Root, "gone.png"), "Component images/U8.png")]);

        Assert.False(copies.AnyLying);
        Assert.False(File.Exists(this.InDraft("Component images/U8.png")));
    }
}
