using System.IO.Compression;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A device's history written out (WorkbookPdfExporter.History.cs, CRT 4.0 step 4 slice 4d). The
// PDF's layout stays untested, as a workbook's does (asserting on PDF bytes tests QuestPDF); what is
// tested is that the document is made at all, with every part filled, and what the ZIP holds and
// under which names - the device's photos, and every repair's files under its workbook's and
// worklog's folder names, as they lie on disk.
// ###########################################################################################
public sealed class DeviceHistoryZipExportTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    // The app's own QuestPDF settings - WorkbookZipExportTests says why.
    public DeviceHistoryZipExportTests() => WorkbookPdfExporter.ConfigureQuestPdf();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string WriteFile(string folder, string name)
    {
        string path = Path.Combine(this.thisWorkspace.Root, folder, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "contents of " + name);
        return path;
    }

    // A repair whose one worklog has `files` attached, written where its folder resolver says.
    private WorkbookExportModel.Document Repair(int workbookId, string status, params string[] files)
    {
        var entry = new WorklogEntryRecord { Id = 1, SchematicName = "Sheet 1", Title = "Bad cap", Category = "Issue", State = status };

        foreach (string file in files)
        {
            this.WriteFile($"workbook_{workbookId}/worklog_1", file);
            entry.Files.Add(new WorklogAttachmentRecord { Id = entry.Files.Count + 1, FileName = file });
        }

        return WorkbookExportModel.Build(
            new WorkbookRecord { Id = workbookId, BoardKey = "Commodore 64|250407", Title = $"Repair {workbookId}", Status = status, StartDate = new DateTime(2026, 9, 3) },
            [entry],
            null,
            entryId => Path.Combine(this.thisWorkspace.Root, $"workbook_{workbookId}", $"worklog_{entryId}"),
            new DateTime(2026, 10, 29),
            "DKK");
    }

    private DeviceHistoryExportModel.Document History(params WorkbookExportModel.Document[] repairs)
    {
        var device = new OwnedDeviceRecord
        {
            Id = 3,
            Name = "Breadbin #2",
            BoardKey = "Commodore 64|250407",
            Notes = "Came with the box - see https://example.org/c64",
            Status = OwnedDeviceStatus.InRepair,
            Photos = [new WorklogAttachmentRecord { Id = 1, FileName = "front.txt" }],
            Changes = [new OwnedDeviceChangeRecord { Id = 1, Date = new DateTime(2025, 6, 14), Text = "JiffyDOS ROM fitted" }],
        };

        string photo = this.WriteFile("device_3", "front.txt");

        return DeviceHistoryExportModel.Build(
            device, "Commodore 64 - 250407", repairs, name => name == "front.txt" ? photo : null, new DateTime(2026, 10, 29), "DKK");
    }

    private static List<string> EntriesOf(string zip)
    {
        using ZipArchive archive = ZipFile.OpenRead(zip);
        return archive.Entries.Select(entry => entry.FullName).ToList();
    }

    // Every part filled - notes, changes, photos, two repairs - and the PDF is made.
    [Fact]
    public void A_history_with_every_part_is_written_as_a_PDF()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "history.pdf");

        WorkbookPdfExporter.WriteHistoryPdf(this.History(this.Repair(4, "Closed", "invoice.txt"), this.Repair(9, "Open")), path);

        Assert.True(new FileInfo(path).Length > 0);
    }

    // A device with no repair yet still has a history - what it is, its changes and photos.
    [Fact]
    public void A_history_with_no_repair_is_still_written()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "history.pdf");

        WorkbookPdfExporter.WriteHistoryPdf(this.History(), path);

        Assert.True(new FileInfo(path).Length > 0);
    }

    // ###########################################################################################
    // The ZIP: the PDF named for the device, the device's photos in a folder of their own, and each
    // repair's files under "workbook_<id>/worklog_<id>/" - so two repairs' "invoice.txt" never meet.
    // ###########################################################################################
    [Fact]
    public void The_ZIP_holds_the_PDF_the_photos_and_every_repairs_files_in_its_own_folder()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "history.zip");

        WorkbookPdfExporter.WriteHistoryZip(
            this.History(this.Repair(4, "Closed", "invoice.txt"), this.Repair(9, "Open", "invoice.txt", "scope.txt")),
            path);

        Assert.Equal(
            [
                "Device_3_Breadbin-2_20261029.pdf",
                "device photos/front.txt",
                "workbook_4/worklog_1/invoice.txt",
                "workbook_9/worklog_1/invoice.txt",
                "workbook_9/worklog_1/scope.txt",
            ],
            EntriesOf(path));
    }
}
