using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// OwnedDeviceStore - the user's own devices (CRT 4.0 step 4), one "device_{id}" folder each in the
// Workbooks folder beside the workbooks, ids never handed out again. WorklogManager is the static
// both share, so this is the "Worklog" collection; every test points it at a fresh folder.
// ###########################################################################################
[Collection("Worklog")]
public sealed class OwnedDeviceStoreTests : IDisposable
{
    private const string Board = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();
    private readonly string thisRoot;

    public OwnedDeviceStoreTests()
    {
        this.thisRoot = this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(this.thisRoot);
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Detached-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    private static OwnedDeviceDetails Details(string name = "Breadbin #2", string board = Board, string status = "") =>
        new(name, board, " PAL ", " CA123456 ", " Flea market ", status);

    private static OwnedDeviceRecord Add(string name = "Breadbin #2", string board = Board) =>
        OwnedDeviceStore.Create(Details(name, board))!;

    private string Picture(string folder, string name)
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Picked", folder, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, [1, 2, 3]);
        return path;
    }

    [Fact]
    public void A_device_added_is_its_own_folder_and_reads_back_trimmed_and_Working()
    {
        OwnedDeviceRecord device = Add();

        Assert.Equal(1, device.Id);
        Assert.True(File.Exists(Path.Combine(this.thisRoot, "device_1", "index.json")));

        OwnedDeviceRecord read = Assert.Single(OwnedDeviceStore.GetAll());
        Assert.Equal(("Breadbin #2", Board, "PAL", "CA123456", "Flea market"), (read.Name, read.BoardKey, read.Region, read.SerialNumber, read.Origin));
        Assert.Equal(OwnedDeviceStatus.Working, read.Status);
        Assert.Equal(DateTime.Now.Date, read.CreatedDate);
    }

    [Fact]
    public void A_device_with_no_name_is_not_added()
    {
        Assert.Null(OwnedDeviceStore.Create(Details(name: "   ")));
        Assert.Empty(OwnedDeviceStore.GetAll());
    }

    [Fact]
    public void Devices_are_listed_by_name_whatever_their_capitals()
    {
        Add("plus/4");
        Add("Breadbin #2");
        Add("C64C silver");

        Assert.Equal(["Breadbin #2", "C64C silver", "plus/4"], OwnedDeviceStore.GetAll().Select(device => device.Name));
    }

    // A deleted device's number is spent: a repair still naming it must never come to name another.
    [Fact]
    public void A_deleted_devices_id_is_never_handed_out_again()
    {
        Add("One");
        OwnedDeviceRecord two = Add("Two");

        Assert.True(OwnedDeviceStore.Delete(two.Id));

        Assert.Equal(3, Add("Three").Id);
    }

    // ###########################################################################################
    // The device counter lives in the root's index.json beside the workbook counter - each
    // allocation rewrites the file, and must keep the other's number.
    // ###########################################################################################
    [Fact]
    public void Workbook_and_device_ids_share_the_root_file_and_keep_each_other()
    {
        WorklogManager.CreateWorkbook(Board, "First", "");
        Add("One");
        WorklogManager.CreateWorkbook(Board, "Second", "");

        using (var root = JsonDocument.Parse(File.ReadAllText(Path.Combine(this.thisRoot, "index.json"))))
        {
            Assert.Equal(2, root.RootElement.GetProperty("lastWorkbookId").GetInt32());
            Assert.Equal(1, root.RootElement.GetProperty("lastDeviceId").GetInt32());
        }

        Assert.Equal(2, Add("Two").Id);
        Assert.Equal(3, WorklogManager.CreateWorkbook(Board, "Third", "")!.Id);
    }

    // ###########################################################################################
    // The counter lost - the root file rewritten without it, as an unreadable file or an older CRT
    // leaves it - never hands out a deleted device's id its repairs still name (code review,
    // 2026-10-28: they appeared on the new device).
    // ###########################################################################################
    [Fact]
    public void A_lost_counter_never_hands_out_an_id_a_repair_still_names()
    {
        Add("One");
        OwnedDeviceRecord two = Add("Two");
        WorklogManager.CreateWorkbook(Board, "No video", "", two.Id);
        OwnedDeviceStore.Delete(two.Id);

        File.WriteAllText(Path.Combine(this.thisRoot, "index.json"), "{ \"lastWorkbookId\": 1 }");

        Assert.Equal(3, Add("Three").Id);
    }

    // A counter a later CRT adds to the root file is written back by this one.
    [Fact]
    public void A_counter_this_CRT_does_not_know_is_kept_in_the_root_file()
    {
        File.WriteAllText(Path.Combine(this.thisRoot, "index.json"), "{ \"lastWorkbookId\": 0, \"lastDeviceId\": 0, \"lastShelfId\": 7 }");

        Add("One");
        WorklogManager.CreateWorkbook(Board, "No video", "");

        using var root = JsonDocument.Parse(File.ReadAllText(Path.Combine(this.thisRoot, "index.json")));
        Assert.Equal(7, root.RootElement.GetProperty("lastShelfId").GetInt32());
    }

    // A device folder on disk from before the counter knew it is walked past, never written over.
    [Fact]
    public void A_device_folder_already_on_disk_is_walked_past()
    {
        Directory.CreateDirectory(Path.Combine(this.thisRoot, "device_1"));

        Assert.Equal(2, Add().Id);
    }

    [Fact]
    public void Devices_are_never_read_as_workbooks_nor_workbooks_as_devices()
    {
        WorklogManager.CreateWorkbook(Board, "No video", "");
        Add();

        Assert.Single(WorklogManager.GetAllWorkbooks());
        Assert.Single(OwnedDeviceStore.GetAll());
    }

    // ###########################################################################################
    // The plan: "A machine whose board is swapped records it as a change." The other fields are
    // simply saved; nothing is recorded for them.
    // ###########################################################################################
    [Fact]
    public void A_board_swapped_is_recorded_as_a_change_of_the_device()
    {
        OwnedDeviceRecord device = Add();

        OwnedDeviceRecord updated = OwnedDeviceStore.Update(device.Id, Details(name: "Breadbin #2 (NTSC)", board: "Commodore 64|250425", status: "faulty"))!;

        Assert.Equal(("Breadbin #2 (NTSC)", "Commodore 64|250425", OwnedDeviceStatus.Faulty), (updated.Name, updated.BoardKey, updated.Status));
        OwnedDeviceChangeRecord change = Assert.Single(updated.Changes);
        Assert.Equal("Board changed from Commodore 64 - 250407 to Commodore 64 - 250425", change.Text);

        // The same board again: nothing more recorded.
        OwnedDeviceStore.Update(device.Id, Details(name: "Breadbin #2", board: "commodore 64|250425"));
        Assert.Single(OwnedDeviceStore.Get(device.Id)!.Changes);
    }

    [Fact]
    public void Changes_are_added_with_their_date_and_removed_and_a_blank_one_is_nothing()
    {
        OwnedDeviceRecord device = Add();

        Assert.Null(OwnedDeviceStore.AddChange(device.Id, DateTime.Now, "   "));

        OwnedDeviceChangeRecord rom = OwnedDeviceStore.AddChange(device.Id, new DateTime(2025, 6, 14, 15, 30, 0), " JiffyDOS ROM fitted ")!;
        OwnedDeviceChangeRecord psu = OwnedDeviceStore.AddChange(device.Id, new DateTime(2026, 1, 2), "New power supply")!;

        Assert.Equal((1, 2), (rom.Id, psu.Id));
        Assert.Equal(new DateTime(2025, 6, 14), OwnedDeviceStore.Get(device.Id)!.Changes[0].Date);
        Assert.Equal("JiffyDOS ROM fitted", OwnedDeviceStore.Get(device.Id)!.Changes[0].Text);

        Assert.NotNull(OwnedDeviceStore.RemoveChange(device.Id, rom.Id));
        Assert.Equal(["New power supply"], OwnedDeviceStore.Get(device.Id)!.Changes.Select(change => change.Text));
    }

    [Fact]
    public void Notes_are_saved_trimmed()
    {
        OwnedDeviceRecord device = Add();

        OwnedDeviceStore.SetNotes(device.Id, "  Recapped in 2024.  ");

        Assert.Equal("Recapped in 2024.", OwnedDeviceStore.Get(device.Id)!.Notes);
    }

    // ###########################################################################################
    // Photos are copied into the device's folder under their own names, numbered on a clash, and
    // only pictures - never a hidden file. Removing one deletes its file.
    // ###########################################################################################
    [Fact]
    public void Photos_are_copied_in_under_free_names_and_removed_with_their_files()
    {
        OwnedDeviceRecord device = Add();

        IReadOnlyList<WorklogAttachmentRecord> added = OwnedDeviceStore.AddPhotos(
            device.Id,
            [this.Picture("a", "front.jpg"), this.Picture("b", "front.jpg"), this.Picture("a", "notes.txt"), this.Picture("a", "._front.jpg")]);

        Assert.Equal(["front.jpg", "front (2).jpg"], added.Select(photo => photo.FileName));
        Assert.Equal([1, 2], added.Select(photo => photo.Id));
        Assert.Equal(2, OwnedDeviceStore.Get(device.Id)!.Photos.Count);

        string path = OwnedDeviceStore.PhotoPath(device.Id, "front (2).jpg")!;
        Assert.True(File.Exists(path));

        Assert.True(OwnedDeviceStore.RemovePhoto(device.Id, 2));

        Assert.False(File.Exists(path));
        Assert.Equal(["front.jpg"], OwnedDeviceStore.Get(device.Id)!.Photos.Select(photo => photo.FileName));
    }

    // A hand-edited index.json naming "../x.png" never makes a path outside the device's folder.
    [Fact]
    public void A_photo_name_that_climbs_out_of_the_folder_is_no_path()
    {
        OwnedDeviceRecord device = Add();
        File.WriteAllBytes(Path.Combine(this.thisRoot, "outside.png"), [1]);

        Assert.Null(OwnedDeviceStore.PhotoPath(device.Id, "../outside.png"));
        Assert.Null(OwnedDeviceStore.PhotoPath(device.Id, ".."));
    }

    // ###########################################################################################
    // A device deleted takes its folder - photos, notes, changes - and leaves its repairs, which are
    // workbooks of their own and from then on read as on no device.
    // ###########################################################################################
    [Fact]
    public void Deleting_a_device_leaves_its_repairs_on_no_device()
    {
        OwnedDeviceRecord device = Add();
        OwnedDeviceStore.AddPhotos(device.Id, [this.Picture("a", "front.jpg")]);
        WorkbookRecord repair = WorklogManager.CreateWorkbook(Board, "No video", "", device.Id)!;

        Assert.True(OwnedDeviceStore.Delete(device.Id));

        Assert.False(Directory.Exists(Path.Combine(this.thisRoot, "device_1")));
        WorkbookRecord kept = Assert.Single(WorklogManager.GetAllWorkbooks());
        Assert.Equal(repair.Id, kept.Id);
        Assert.Null(OwnedDeviceRepairs.DeviceOf(kept, OwnedDeviceStore.GetAll()));
    }

    // What a later CRT writes into a device is written back by this one, never dropped.
    [Fact]
    public void A_field_this_CRT_does_not_know_is_kept_through_a_write()
    {
        OwnedDeviceRecord device = Add();
        string index = Path.Combine(this.thisRoot, "device_1", "index.json");
        string json = File.ReadAllText(index).TrimEnd().TrimEnd('}') + ", \"caseColour\": \"beige\" }";
        File.WriteAllText(index, json);

        OwnedDeviceStore.SetNotes(device.Id, "Recapped");

        using var written = JsonDocument.Parse(File.ReadAllText(index));
        Assert.Equal("beige", written.RootElement.GetProperty("caseColour").GetString());
    }

    [Theory]
    [InlineData("", "Working")]
    [InlineData("  in REPAIR ", "In repair")]
    [InlineData("not mine any more", "Not mine any more")]
    [InlineData("Lent to a friend", "Lent to a friend")]
    public void A_status_reads_as_its_known_word_and_an_unknown_one_is_kept(string stored, string read)
    {
        Assert.Equal(read, OwnedDeviceStatus.Normalise(stored));
    }

    // Each status is drawn in the colour CRT's own pass, fail and warning lines use; the rest muted.
    [Theory]
    [InlineData("working", "Text_Success_Fg")]
    [InlineData("Faulty", "Text_Fail_Fg")]
    [InlineData("In repair", "Text_Warning_Fg")]
    [InlineData("For parts", "Workbooks_Muted_Fg")]
    [InlineData("Lent to a friend", "Workbooks_Muted_Fg")]
    public void A_status_has_its_colour(string status, string key)
    {
        Assert.Equal(key, OwnedDeviceStatus.ColourKeyOf(status));
    }
}
