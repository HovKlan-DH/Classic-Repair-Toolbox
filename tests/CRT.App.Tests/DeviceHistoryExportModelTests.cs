using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DeviceHistoryExportModel - what goes into a device's exported history (CRT 4.0 step 4, slice 4d -
// the plan: "Export history puts every repair into one PDF or ZIP"): what the device is, its
// changes and photos, every repair oldest first, and what they add up to by the device page's rule.
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
        Origin = "A flea market",
        Notes = " Keyboard needs cleaning. ",
        Status = "in repair",
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
    // under the name and where it came from.
    [Fact]
    public void The_device_is_described_as_its_page_describes_it()
    {
        DeviceHistoryExportModel.Document history = Build(Device());

        Assert.Equal(3, history.DeviceId);
        Assert.Equal("Breadbin #2", history.Name);
        Assert.Equal(OwnedDeviceStatus.InRepair, history.Status);
        Assert.Equal("Commodore 64 - 250407 - PAL - serial CA123456", history.DetailLine);
        Assert.Equal("Where it came from: A flea market", history.OriginLine);
        Assert.Equal("Keyboard needs cleaning.", history.Notes);
        Assert.Empty(history.Repairs);
    }

    // The changes oldest first (one day's by the order they were added), a blank one left out.
    [Fact]
    public void The_changes_come_oldest_first_and_a_blank_one_is_left_out()
    {
        OwnedDeviceRecord device = Device();
        device.Changes =
        [
            new OwnedDeviceChangeRecord { Id = 1, Date = new DateTime(2026, 1, 2), Text = "New power supply " },
            new OwnedDeviceChangeRecord { Id = 2, Date = new DateTime(2025, 6, 14), Text = "JiffyDOS ROM fitted" },
            new OwnedDeviceChangeRecord { Id = 3, Date = new DateTime(2025, 6, 14), Text = "  " },
            new OwnedDeviceChangeRecord { Id = 4, Date = new DateTime(2025, 6, 14), Text = "Keycaps cleaned" },
        ];

        Assert.Equal(
            ["JiffyDOS ROM fitted", "Keycaps cleaned", "New power supply"],
            Build(device).Changes.Select(change => change.Text));
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
