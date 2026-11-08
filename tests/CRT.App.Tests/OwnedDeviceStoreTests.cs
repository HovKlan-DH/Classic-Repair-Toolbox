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

    // A device added with no status is "Unknown" (owner request, 2026-11-06: "Instead have a status of
    // 'Unknown'", the first offered - "Working" until then), and with no condition no stars.
    [Fact]
    public void A_device_added_is_its_own_folder_and_reads_back_trimmed_and_Unknown()
    {
        OwnedDeviceRecord device = Add();

        Assert.Equal(1, device.Id);
        Assert.True(File.Exists(Path.Combine(this.thisRoot, "device_1", "index.json")));

        OwnedDeviceRecord read = Assert.Single(OwnedDeviceStore.GetAll());
        Assert.Equal(("Breadbin #2", Board, "PAL", "CA123456", "Flea market"), (read.Name, read.BoardKey, read.Region, read.SerialNumber, read.Description));
        Assert.Equal(OwnedDeviceStatus.Unknown, read.Status);
        Assert.Equal(0, read.Condition);
        Assert.Empty(read.Comments);
        Assert.Equal(DateTime.Now.Date, read.CreatedDate);
    }

    // ###########################################################################################
    // The description is STORED under its first name, "origin" (a name already stored keeps its spelling):
    // a device an earlier pre-release wrote reads its "Where it came from" as its description, and its
    // missing condition and comments as none.
    // ###########################################################################################
    [Fact]
    public void The_description_is_stored_as_origin_and_an_earlier_device_reads_with_no_stars_or_comments()
    {
        OwnedDeviceRecord device = Add();
        string index = Path.Combine(this.thisRoot, "device_1", "index.json");

        using (var written = JsonDocument.Parse(File.ReadAllText(index)))
        {
            Assert.Equal("Flea market", written.RootElement.GetProperty("origin").GetString());
            Assert.False(written.RootElement.TryGetProperty("description", out _));
        }

        File.WriteAllText(index, $$"""{ "id": {{device.Id}}, "name": "Old one", "origin": "A friend", "status": "Not mine any more", "condition": 9 }""");

        OwnedDeviceRecord read = OwnedDeviceStore.Get(device.Id)!;
        Assert.Equal(("A friend", "Not mine any more", 5), (read.Description, read.Status, read.Condition));
        Assert.Empty(read.Comments);
        Assert.Empty(read.Photos);
    }

    // ###########################################################################################
    // Its condition, 0 to 5 stars (owner request, 2026-11-06) - saved as the form holds it, anything past
    // either end taken to it; a star clicked sets that many, the one set clicked again none.
    // ###########################################################################################
    [Fact]
    public void The_condition_is_saved_within_0_to_5_stars()
    {
        OwnedDeviceRecord device = Add();

        Assert.Equal(4, OwnedDeviceStore.Update(device.Id, Details() with { Condition = 4 })!.Condition);
        Assert.Equal(5, OwnedDeviceStore.Update(device.Id, Details() with { Condition = 12 })!.Condition);
        Assert.Equal(0, OwnedDeviceStore.Update(device.Id, Details() with { Condition = -1 })!.Condition);

        Assert.Equal(3, OwnedDeviceCondition.AfterClick(0, 3));
        Assert.Equal(5, OwnedDeviceCondition.AfterClick(3, 5));
        Assert.Equal(0, OwnedDeviceCondition.AfterClick(3, 3));
        Assert.Equal(1, OwnedDeviceCondition.AfterClick(3, 1));
    }

    // ###########################################################################################
    // A device's comments (owner request, 2026-11-06: "Comments", as a worklog has them) - added, changed
    // and deleted one at a time, each saved at once (owner decision: "At once, as a worklog does"), ids the
    // device's own, never a blank one; a new device takes those the form held for it in its first write.
    // ###########################################################################################
    [Fact]
    public void Comments_are_added_changed_and_deleted_one_at_a_time()
    {
        OwnedDeviceRecord device = Add();
        var when = new DateTime(2026, 11, 6, 10, 15, 0);

        Assert.Null(OwnedDeviceStore.AddComment(device.Id, "   ", when));
        Assert.Equal([1], OwnedDeviceStore.AddComment(device.Id, " Recapped ", when)!.Comments.Select(comment => comment.Id));
        OwnedDeviceRecord two = OwnedDeviceStore.AddComment(device.Id, "New PLA fitted", when.AddDays(1))!;

        Assert.Equal([("Recapped", when), ("New PLA fitted", when.AddDays(1))], two.Comments.Select(comment => (comment.Text, comment.Date)));

        Assert.Equal("PSU recapped", OwnedDeviceStore.UpdateComment(device.Id, 1, " PSU recapped ")!.Comments[0].Text);
        Assert.Null(OwnedDeviceStore.UpdateComment(device.Id, 9, "Nothing there"));
        Assert.Null(OwnedDeviceStore.UpdateComment(device.Id, 1, "  "));

        Assert.Equal([2], OwnedDeviceStore.RemoveComment(device.Id, 1)!.Comments.Select(comment => comment.Id));
        Assert.Null(OwnedDeviceStore.RemoveComment(device.Id, 1));

        // Ids are the device's own, one past the highest.
        Assert.Equal([2, 3], OwnedDeviceStore.AddComment(device.Id, "Case cleaned", when)!.Comments.Select(comment => comment.Id));

        OwnedDeviceRecord added = OwnedDeviceStore.Create(
            Details(name: "Plus/4"),
            [new WorklogCommentRecord { Id = 7, Text = " Bought ", Date = when }, new WorklogCommentRecord { Id = 8, Text = " ", Date = when }])!;
        Assert.Equal([(1, "Bought", when)], OwnedDeviceStore.Get(added.Id)!.Comments.Select(comment => (comment.Id, comment.Text, comment.Date)));
    }

    // ###########################################################################################
    // Its links of interest (owner request, 2026-11-06: "the same 'Links of interest' for the device") - a headline and
    // an address, added, changed and deleted one at a time, ids the device's own; no address is no link; a new device
    // takes those the form held for it in its first write.
    // ###########################################################################################
    [Fact]
    public void Links_are_added_changed_and_deleted_one_at_a_time()
    {
        OwnedDeviceRecord device = Add();

        Assert.Null(OwnedDeviceStore.AddLink(device.Id, "Nothing", "  "));
        Assert.Equal([(1, "Service manual", "https://zimmers.net")],
            OwnedDeviceStore.AddLink(device.Id, " Service manual ", " https://zimmers.net ")!.Links.Select(link => (link.Id, link.Headline, link.Url)));
        Assert.Equal(2, OwnedDeviceStore.AddLink(device.Id, "", "www.lemon64.com")!.Links.Count);

        Assert.Equal("Manual", OwnedDeviceStore.UpdateLink(device.Id, 1, "Manual", "https://zimmers.net/c64")!.Links[0].Headline);
        Assert.Null(OwnedDeviceStore.UpdateLink(device.Id, 9, "x", "https://x.org"));
        Assert.Null(OwnedDeviceStore.UpdateLink(device.Id, 1, "x", ""));

        Assert.Equal([2], OwnedDeviceStore.RemoveLink(device.Id, 1)!.Links.Select(link => link.Id));
        Assert.Null(OwnedDeviceStore.RemoveLink(device.Id, 1));

        OwnedDeviceRecord added = OwnedDeviceStore.Create(
            Details(name: "Plus/4"),
            comments: null,
            links: [new WorklogLinkRecord { Id = 5, Headline = " Wiki ", Url = " https://c64-wiki.com " }, new WorklogLinkRecord { Id = 6, Url = " " }])!;
        Assert.Equal([(1, "Wiki", "https://c64-wiki.com")], OwnedDeviceStore.Get(added.Id)!.Links.Select(link => (link.Id, link.Headline, link.Url)));
    }

    // ###########################################################################################
    // The photos put in another order (owner request, 2026-11-06: "rearrange the photos, just like it is possible to do
    // in the worklogs") - each its place; an id it has not passed over, one left out after those named.
    // ###########################################################################################
    [Fact]
    public void The_photos_are_put_in_the_order_given()
    {
        OwnedDeviceRecord device = Add();
        OwnedDeviceStore.AddPhotos(device.Id,
        [
            new OwnedDevicePhotoSource(this.Picture("a", "one.jpg"), string.Empty),
            new OwnedDevicePhotoSource(this.Picture("a", "two.jpg"), string.Empty),
            new OwnedDevicePhotoSource(this.Picture("a", "three.jpg"), string.Empty),
        ]);

        List<string> Ordered() => OwnedDeviceStore.Get(device.Id)!.Photos
            .OrderBy(photo => photo.DisplayOrder).ThenBy(photo => photo.Id).Select(photo => photo.FileName).ToList();

        Assert.Equal(["one.jpg", "two.jpg", "three.jpg"], Ordered());

        Assert.NotNull(OwnedDeviceStore.SetPhotoOrder(device.Id, [3, 1, 2]));
        Assert.Equal(["three.jpg", "one.jpg", "two.jpg"], Ordered());

        Assert.NotNull(OwnedDeviceStore.SetPhotoOrder(device.Id, [2, 99]));
        Assert.Equal(["two.jpg", "three.jpg", "one.jpg"], Ordered());
    }

    // ###########################################################################################
    // A DEVICE WITH WORKBOOKS KEEPS ITS BOARD (owner decision, 2026-11-06, asked: "Locked while it has workbooks" - a
    // workbook is associated only to a device of its own board): the rest of the form is saved, a board handed in is
    // not taken; with none, it is.
    // ###########################################################################################
    [Fact]
    public void A_device_with_workbooks_keeps_its_board_and_one_without_takes_another()
    {
        OwnedDeviceRecord device = Add();
        OwnedDeviceRecord bare = Add("Plus/4");
        WorklogManager.CreateWorkbook(Board, "No video", "", device.Id);

        OwnedDeviceRecord kept = OwnedDeviceStore.Update(device.Id, Details(name: "Renamed", board: "Commodore 64|250425"))!;
        Assert.Equal(("Renamed", Board), (kept.Name, kept.BoardKey));

        Assert.Equal("Commodore 64|250425", OwnedDeviceStore.Update(bare.Id, Details(name: "Plus/4", board: "Commodore 64|250425"))!.BoardKey);
    }

    // ###########################################################################################
    // A photo carries its comment, as a worklog's does; a photo changed keeps its place - its comment, and
    // its picture when another is picked: copied in under a name of its own, the old file deleted only
    // once the device no longer names it.
    // ###########################################################################################
    [Fact]
    public void A_photo_carries_its_comment_and_is_changed_in_place()
    {
        OwnedDeviceRecord device = Add();

        WorklogAttachmentRecord photo = Assert.Single(OwnedDeviceStore.AddPhotos(
            device.Id,
            [new OwnedDevicePhotoSource(this.Picture("a", "front.jpg"), " Before cleaning ")]));
        Assert.Equal("Before cleaning", photo.Comment);
        string oldPath = OwnedDeviceStore.PhotoPath(device.Id, "front.jpg")!;

        Assert.True(OwnedDeviceStore.UpdatePhoto(device.Id, photo.Id, replacement: null, " After cleaning "));
        Assert.Equal(("front.jpg", "After cleaning"), OwnedDeviceStore.Get(device.Id)!.Photos.Select(each => (each.FileName, each.Comment)).Single());

        Assert.True(OwnedDeviceStore.UpdatePhoto(device.Id, photo.Id, this.Picture("b", "after.png"), "After"));
        WorklogAttachmentRecord changed = Assert.Single(OwnedDeviceStore.Get(device.Id)!.Photos);
        Assert.Equal((photo.Id, "after.png", "After"), (changed.Id, changed.FileName, changed.Comment));
        Assert.False(File.Exists(oldPath));
        Assert.NotNull(OwnedDeviceStore.PhotoPath(device.Id, "after.png"));

        // No picture is no change - and a photo not there neither.
        Assert.False(OwnedDeviceStore.UpdatePhoto(device.Id, photo.Id, this.Picture("c", "notes.txt"), "x"));
        Assert.False(OwnedDeviceStore.UpdatePhoto(device.Id, 99, null, "x"));
        Assert.Equal("After", Assert.Single(OwnedDeviceStore.Get(device.Id)!.Photos).Comment);
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
    // "Edit device" saves the form's fields - a board swapped is simply the new board. Nothing is
    // recorded beside them (owner request, 2026-11-06: "remove the 'Changes made to it', as this is
    // exactly what the workbooks will cover").
    // ###########################################################################################
    [Fact]
    public void Edit_saves_the_forms_fields_and_records_nothing_beside_them()
    {
        OwnedDeviceRecord device = Add();

        OwnedDeviceRecord updated = OwnedDeviceStore.Update(device.Id, Details(name: "Breadbin #2 (NTSC)", board: "Commodore 64|250425", status: "faulty"))!;

        Assert.Equal(("Breadbin #2 (NTSC)", "Commodore 64|250425", OwnedDeviceStatus.Faulty), (updated.Name, updated.BoardKey, updated.Status));

        using var written = JsonDocument.Parse(File.ReadAllText(Path.Combine(this.thisRoot, "device_1", "index.json")));
        Assert.False(written.RootElement.TryGetProperty("changes", out _));
    }

    // A device with no name is no device: the edit is refused and the device stays as it was.
    [Fact]
    public void An_edit_with_no_name_is_refused()
    {
        OwnedDeviceRecord device = Add();

        Assert.Null(OwnedDeviceStore.Update(device.Id, Details(name: "  ")));
        Assert.Equal("Breadbin #2", OwnedDeviceStore.Get(device.Id)!.Name);
    }

    // ###########################################################################################
    // The changes an earlier CRT 4.0 pre-release recorded on a device - a list it no longer shows -
    // stay in its file, exactly as they were, through every write: nothing the user typed is thrown
    // away (OwnedDeviceRecord.Unknown).
    // ###########################################################################################
    [Fact]
    public void Changes_an_earlier_CRT_recorded_stay_in_the_file_untouched()
    {
        OwnedDeviceRecord device = Add();
        string index = Path.Combine(this.thisRoot, "device_1", "index.json");
        const string changes =
            "[ { \"id\": 1, \"date\": \"2025-06-14T00:00:00\", \"text\": \"JiffyDOS ROM fitted\" }, " +
            "{ \"id\": 2, \"date\": \"2026-01-02T00:00:00\", \"text\": \"Board changed from Commodore 64 - 250407 to Commodore 64 - 250425\" } ]";
        File.WriteAllText(index, File.ReadAllText(index).TrimEnd().TrimEnd('}') + ", \"changes\": " + changes + " }");

        OwnedDeviceStore.Update(device.Id, Details(name: "Breadbin #2", board: "Commodore 64|250425"));
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("a", "front.jpg"), string.Empty)]);

        using var written = JsonDocument.Parse(File.ReadAllText(index));
        Assert.True(System.Text.Json.Nodes.JsonNode.DeepEquals(
            System.Text.Json.Nodes.JsonNode.Parse(changes),
            System.Text.Json.Nodes.JsonNode.Parse(written.RootElement.GetProperty("changes").GetRawText())));
    }

    // The description keeps the lines it was typed in - only its ends are trimmed.
    [Fact]
    public void The_description_keeps_its_lines()
    {
        OwnedDeviceRecord device = OwnedDeviceStore.Create(new OwnedDeviceDetails(
            "Breadbin #2", Board, "", "", "  A flea market in Odense\nCame with its box\n  ", ""))!;

        Assert.Equal("A flea market in Odense\nCame with its box", OwnedDeviceStore.Get(device.Id)!.Description);
    }

    // ###########################################################################################
    // *** A NOTE'S FIRST LINE KEEPS THE INDENT THAT SAYS WHAT IT IS (code review, 2026-11-08). *** Trimmed whole, a
    // description, a comment or a photo's comment whose first line is a bullet nested once - "  - recapped" - came back
    // a bullet of the first level, and one starting with three spaces and "- " (no bullet at that indent) became one.
    // Blank lines around the note and white space at its end still go (RichNote.TrimMarkup).
    // ###########################################################################################
    [Fact]
    public void A_notes_first_line_keeps_the_indent_that_says_what_it_is()
    {
        OwnedDeviceRecord device = OwnedDeviceStore.Create(new OwnedDeviceDetails(
            "Breadbin #2", Board, "", "", "\n  - recapped\n- RAM\n  ", ""))!;

        Assert.Equal("  - recapped\n- RAM", OwnedDeviceStore.Get(device.Id)!.Description);

        OwnedDeviceStore.AddComment(device.Id, "   - not a bullet\n", new DateTime(2026, 11, 8, 10, 0, 0));
        Assert.Equal("   - not a bullet", Assert.Single(OwnedDeviceStore.Get(device.Id)!.Comments).Text);

        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("a", "front.jpg"), "  - the front\n")]);
        Assert.Equal("  - the front", Assert.Single(OwnedDeviceStore.Get(device.Id)!.Photos).Comment);
    }

    // ###########################################################################################
    // A device's notes - the page's own box until 2026-11-07 (owner request: 'Remove the "Notes" as we now have the
    // "Description"') - are FOLDED INTO ITS DESCRIPTION as it is read, a blank line between, so nothing typed there is
    // lost; the next write writes them as the description, and "notes" no more.
    // ###########################################################################################
    [Fact]
    public void Notes_a_device_had_are_read_as_its_description_and_written_as_it()
    {
        OwnedDeviceRecord device = OwnedDeviceStore.Create(new OwnedDeviceDetails("Breadbin #2", Board, "", "", "A flea market", ""))!;
        string index = Path.Combine(this.thisRoot, "device_1", "index.json");
        File.WriteAllText(index, File.ReadAllText(index).TrimEnd().TrimEnd('}') + ", \"notes\": \"  Recapped in 2024.  \" }");

        Assert.Equal("A flea market\n\nRecapped in 2024.", OwnedDeviceStore.Get(device.Id)!.Description);

        OwnedDeviceStore.AddComment(device.Id, "SID replaced", new DateTime(2026, 11, 7, 10, 0, 0));

        using var written = JsonDocument.Parse(File.ReadAllText(index));
        Assert.Equal("A flea market\n\nRecapped in 2024.", written.RootElement.GetProperty("origin").GetString());
        Assert.False(written.RootElement.TryGetProperty("notes", out _));
        Assert.Equal("A flea market\n\nRecapped in 2024.", OwnedDeviceStore.Get(device.Id)!.Description);
    }

    [Theory]
    [InlineData("A flea market", null, "A flea market")]
    [InlineData("A flea market", "   ", "A flea market")]
    [InlineData("", "Recapped", "Recapped")]
    [InlineData("  ", " Recapped ", "Recapped")]
    [InlineData("A flea market\n", "Recapped\nNew PSU", "A flea market\n\nRecapped\nNew PSU")]
    public void Notes_are_folded_under_the_description_with_a_blank_line_between(string description, string? notes, string expected)
    {
        Assert.Equal(expected, OwnedDeviceStore.FoldNotes(description, notes));
    }

    // ###########################################################################################
    // SORTED BY HAND (owner request, 2026-11-07: "Sort manually in list (drag up/down)") - each device's place is
    // written where it moved, and a new device comes last: after the last placed, or unplaced while none is.
    // ###########################################################################################
    [Fact]
    public void The_order_by_hand_is_written_where_it_moved_and_a_new_device_comes_last()
    {
        OwnedDeviceRecord first = Add();
        OwnedDeviceRecord second = OwnedDeviceStore.Create(Details(name: "Plus/4"))!;
        Assert.Equal(0, second.DisplayOrder);

        Assert.True(OwnedDeviceStore.SetOrder([second.Id, first.Id]));

        Assert.Equal(1, OwnedDeviceStore.Get(second.Id)!.DisplayOrder);
        Assert.Equal(2, OwnedDeviceStore.Get(first.Id)!.DisplayOrder);

        OwnedDeviceRecord third = OwnedDeviceStore.Create(Details(name: "C16"))!;
        Assert.Equal(3, third.DisplayOrder);

        Assert.Equal(
            ["Plus/4", "Breadbin #2", "C16"],
            OwnedDeviceSorting.Sort(OwnedDeviceStore.GetAll(), OwnedDeviceSorting.Manual, key => (key, string.Empty)).Select(device => device.Name));
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
            new[] { this.Picture("a", "front.jpg"), this.Picture("b", "front.jpg"), this.Picture("a", "notes.txt"), this.Picture("a", "._front.jpg") }
                .Select(path => new OwnedDevicePhotoSource(path, string.Empty))
                .ToList());

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
    // A device deleted takes its folder - photos and notes - and leaves its repairs, which are
    // workbooks of their own and from then on read as on no device.
    // ###########################################################################################
    [Fact]
    public void Deleting_a_device_leaves_its_repairs_on_no_device()
    {
        OwnedDeviceRecord device = Add();
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("a", "front.jpg"), string.Empty)]);
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

        OwnedDeviceStore.AddComment(device.Id, "Recapped", new DateTime(2026, 11, 7, 10, 0, 0));

        using var written = JsonDocument.Parse(File.ReadAllText(index));
        Assert.Equal("beige", written.RootElement.GetProperty("caseColour").GetString());
    }

    // ###########################################################################################
    // "Unknown" first, "Not mine any more" no longer offered (owner request, 2026-11-06: "this is different
    // data as I see it") - a device an earlier pre-release stored it on keeps it, as any word this CRT does
    // not know; a blank one is "Unknown".
    // ###########################################################################################
    [Fact]
    public void The_statuses_offered_start_with_Unknown_and_have_no_Not_mine_any_more()
    {
        Assert.Equal(["Unknown", "Working", "Faulty", "In repair", "For parts"], OwnedDeviceStatus.All);
    }

    [Theory]
    [InlineData("", "Unknown")]
    [InlineData(" unknown ", "Unknown")]
    [InlineData("  in REPAIR ", "In repair")]
    [InlineData("Not mine any more", "Not mine any more")]
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
    [InlineData("Unknown", "Workbooks_Muted_Fg")]
    [InlineData("Lent to a friend", "Workbooks_Muted_Fg")]
    public void A_status_has_its_colour(string status, string key)
    {
        Assert.Equal(key, OwnedDeviceStatus.ColourKeyOf(status));
    }
}
