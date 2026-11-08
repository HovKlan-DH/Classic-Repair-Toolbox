using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DeviceHistoryExportModel - what goes into a device's exported history (CRT 4.0 step 4, slice 4d -
// the plan: "Export history puts every repair into one PDF or ZIP"): what the device is, its
// notes and photos, every repair oldest first, and what they add up to by the device page's rule.
// ###########################################################################################
public sealed class DeviceHistoryExportModelTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static readonly DateTime Made = new(2026, 10, 29, 14, 5, 0);

    private static OwnedDeviceRecord Device() => new()
    {
        Id = 3,
        Name = " Breadbin #2 ",
        BoardKey = "Commodore 64|250407",
        Region = "PAL",
        SerialNumber = "CA123456",
        Description = "A flea market",
        Status = "in repair",
        Condition = 3,
    };

    private static WorkbookExportModel.Document Repair(int id, string status, params (double Hours, double Cost)[] work)
    {
        var workbook = new WorkbookRecord { Id = id, BoardKey = "Commodore 64|250407", Title = $"Repair {id}", Status = status };

        WorklogEntryRecord[] entries = work.Select((spent, index) => new WorklogEntryRecord
        {
            Id = index + 1,
            SchematicName = "Sheet 1",
            Title = "Checked",
            State = status,
            WorkDoneItems = [new WorklogWorkDoneRecord { Text = "Work", HoursSpent = spent.Hours, Cost = spent.Cost }],
        }).ToArray();

        return WorkbookExportModel.Build(workbook, entries, null, _ => null, Made, "DKK");
    }

    private static DeviceHistoryExportModel.Document Build(OwnedDeviceRecord device, params WorkbookExportModel.Document[] repairs) =>
        DeviceHistoryExportModel.Build(device, "Commodore 64 - 250407", repairs, _ => null, Made, "dkk");

    // ###########################################################################################
    // A history reads as it happened: the oldest repair first, whatever order they were handed in -
    // the device's page lists them newest first, to start from the one going on.
    // ###########################################################################################
    [Fact]
    public void The_repairs_come_oldest_first()
    {
        DeviceHistoryExportModel.Document history = Build(Device(), Repair(12, "Open"), Repair(4, "Closed"), Repair(9, "Closed"));

        Assert.Equal([4, 9, 12], history.Repairs.Select(repair => repair.WorkbookId));
    }

    // What they add up to is the device page's own count: every repair's worklogs in one summary,
    // the repairs and the open ones counted - so the history and the page cannot disagree.
    [Fact]
    public void What_the_repairs_add_up_to_is_the_device_pages_count()
    {
        DeviceHistoryExportModel.Document history = Build(
            Device(),
            Repair(1, "Open", (1.5, 12), (0.5, 3)),
            Repair(2, "Closed", (0.75, 33)));

        Assert.Equal(2, history.History.RepairCount);
        Assert.Equal(1, history.History.OpenRepairCount);
        Assert.Equal(3, history.History.Totals.WorklogCount);
        Assert.Equal(2.75, history.History.Totals.TotalHours, 3);
        Assert.Equal("DKK", history.CurrencyCode);
    }

    // What the device is, as its page says it - name trimmed, status as one of its words, the line
    // under the name, its condition in words (a PDF font may have no star) and its description.
    [Fact]
    public void The_device_is_described_as_its_page_describes_it()
    {
        DeviceHistoryExportModel.Document history = Build(Device());

        Assert.Equal(3, history.DeviceId);
        Assert.Equal("Breadbin #2", history.Name);
        Assert.Equal(OwnedDeviceStatus.InRepair, history.Status);
        Assert.Equal("Commodore 64 - 250407 - PAL - serial CA123456", history.DetailLine);
        Assert.Equal("Condition: 3 of 5 stars", history.ConditionLine);
        Assert.Equal("A flea market", history.Description);
        Assert.Empty(history.Comments);
        Assert.Empty(history.Repairs);
    }

    // ###########################################################################################
    // The description keeps the lines it was typed in (owner request, 2026-11-06: the description - "Where
    // it came from" until that day - is several lines) - each break a "\n", whatever the box typed, so the
    // PDF breaks where the user did. A device has no notes of its own since 2026-11-07 - its description
    // says it (OwnedDeviceStore.FoldNotes).
    // ###########################################################################################
    [Fact]
    public void The_description_keeps_its_lines()
    {
        OwnedDeviceRecord device = Device();
        device.Description = "A flea market in Odense\r\nCame with its box";

        DeviceHistoryExportModel.Document history = Build(device);

        Assert.Equal("A flea market in Odense\nCame with its box", history.Description);
    }

    // ###########################################################################################
    // Its comments go in too (owner request, 2026-11-06: "Comments" on a device), dated as the page dates
    // them - oldest first, as a history reads; a blank one is none.
    // ###########################################################################################
    [Fact]
    public void The_comments_come_oldest_first_and_dated()
    {
        OwnedDeviceRecord device = Device();
        device.Comments =
        [
            new WorklogCommentRecord { Id = 2, Text = "Recapped.\r\nRuns cool now.", Date = new DateTime(2026, 11, 2, 18, 30, 0) },
            new WorklogCommentRecord { Id = 1, Text = " Bought at a flea market ", Date = new DateTime(2026, 10, 1, 9, 5, 0) },
            new WorklogCommentRecord { Id = 3, Text = "  ", Date = new DateTime(2026, 11, 3) },
        ];

        DeviceHistoryExportModel.Document history = Build(device);

        Assert.Equal(
            [new DeviceHistoryExportModel.Comment("2026-10-01 09:05", "Bought at a flea market"), new DeviceHistoryExportModel.Comment("2026-11-02 18:30", "Recapped.\nRuns cool now.")],
            history.Comments);
    }

    // A comment's first line nested once is exported so, not trimmed to the first level (code review, 2026-11-08).
    [Fact]
    public void A_comments_first_line_keeps_the_indent_that_says_what_it_is()
    {
        OwnedDeviceRecord device = Device();
        device.Comments = [new WorklogCommentRecord { Id = 1, Text = "  - replaced C14\r\n- U8", Date = new DateTime(2026, 11, 8, 9, 0, 0) }];

        Assert.Equal("  - replaced C14\n- U8", Assert.Single(Build(device).Comments).Text);
    }

    // The photos in the order the page shows them, each where it lies - one not there left out.
    [Fact]
    public void The_photos_come_in_their_order_and_one_not_there_is_left_out()
    {
        OwnedDeviceRecord device = Device();
        device.Photos =
        [
            new WorklogAttachmentRecord { Id = 1, FileName = "back.jpg", DisplayOrder = 2 },
            new WorklogAttachmentRecord { Id = 2, FileName = "gone.jpg", DisplayOrder = 0 },
            new WorklogAttachmentRecord { Id = 3, FileName = "front.jpg", DisplayOrder = 1, Comment = "As bought" },
        ];

        string folder = this.thisWorkspace.Path_("device_3");
        DeviceHistoryExportModel.Document history = DeviceHistoryExportModel.Build(
            device,
            "Commodore 64 - 250407",
            [],
            name => name == "gone.jpg" ? null : Path.Combine(folder, name),
            Made,
            "DKK");

        Assert.Equal(["front.jpg", "back.jpg"], history.Photos.Select(photo => photo.FileName));
        Assert.Equal(Path.Combine(folder, "front.jpg"), history.Photos[0].FullPath);
        Assert.Equal("As bought", history.Photos[0].Comment);
    }

    // Photos of one order come by their id, as the device's page shows them (code review,
    // 2026-10-29: the history kept the stored order instead).
    [Fact]
    public void Photos_of_one_order_come_by_their_id_as_on_the_page()
    {
        OwnedDeviceRecord device = Device();
        device.Photos =
        [
            new WorklogAttachmentRecord { Id = 2, FileName = "second.jpg", DisplayOrder = 0 },
            new WorklogAttachmentRecord { Id = 1, FileName = "first.jpg", DisplayOrder = 0 },
        ];

        DeviceHistoryExportModel.Document history = DeviceHistoryExportModel.Build(
            device, "Commodore 64 - 250407", [], name => this.thisWorkspace.Path_(name), Made, "DKK");

        Assert.Equal(["first.jpg", "second.jpg"], history.Photos.Select(photo => photo.FileName));
    }

    // ###########################################################################################
    // The file is named for the device - its id, its name made safe, the day - and a repair's files
    // go under its workbook's folder name as on disk.
    // ###########################################################################################
    [Fact]
    public void The_file_is_named_for_the_device_and_a_repairs_folder_for_its_workbook()
    {
        Assert.Equal("Device_3_Breadbin-2_20261029", DeviceHistoryExportModel.BuildFileBaseName(Build(Device())));
        Assert.Equal("Device_3_20261029", DeviceHistoryExportModel.BuildFileBaseName(3, " #/ ", Made));
        Assert.Equal("workbook_12", DeviceHistoryExportModel.RepairFolderName(Repair(12, "Open")));
    }
}
