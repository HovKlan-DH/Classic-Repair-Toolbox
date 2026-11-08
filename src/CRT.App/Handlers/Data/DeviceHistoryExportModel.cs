using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE CONTENT OF A DEVICE'S HISTORY EXPORT (CRT 4.0 step 4, slice 4d - the plan: "Export history
    // puts every repair into one PDF or ZIP"): what the device is, its notes, comments and photos, and every
    // repair as its own workbook document - with no idea how it is drawn, as
    // WorkbookExportModel is for one workbook. WorkbookPdfExporter.History.cs paints it.
    //
    // The repairs come OLDEST FIRST: a history reads as it happened (the device's page lists them
    // newest first, to start from the one going on). What they add up to is the device page's own
    // rule - one WorkbookSummary over every repair's worklogs, its repairs and open ones counted.
    //
    // Pure apart from the photo resolver it is handed (OwnedDeviceStore.PhotoPath in the app - a
    // plain file name in the device's folder, there on disk).
    // ###########################################################################################
    public static class DeviceHistoryExportModel
    {
        public sealed class Document
        {
            public int DeviceId { get; init; }

            public string Name { get; init; } = string.Empty;

            public string Status { get; init; } = string.Empty;

            // "Commodore 64 - 250407 - PAL - serial CA123456" - the device page's line under its name.
            public string DetailLine { get; init; } = string.Empty;

            // "Condition: 3 of 5 stars" - its stars in words, as a PDF font may have no star.
            public string ConditionLine { get; init; } = string.Empty;

            // Its description - empty when it has none. Its line breaks kept, as "\n".
            public string Description { get; init; } = string.Empty;

            // Its comments, oldest first - a history reads as it happened.
            public IReadOnlyList<Comment> Comments { get; init; } = Array.Empty<Comment>();

            // Its links of interest, in their order - one with no address left out.
            public IReadOnlyList<WorklogLinkRecord> Links { get; init; } = Array.Empty<WorklogLinkRecord>();

            public IReadOnlyList<WorkbookExportModel.Attachment> Photos { get; init; } = Array.Empty<WorkbookExportModel.Attachment>();

            public IReadOnlyList<WorkbookExportModel.Document> Repairs { get; init; } = Array.Empty<WorkbookExportModel.Document>();

            public OwnedDeviceRepairs.History History { get; init; } = new(0, 0, WorkbookSummary.Summarize(null));

            public DateTime GeneratedAt { get; init; }

            // Captured when the document is built, as a workbook export's is - the paint runs off
            // the UI thread, where a setting changed mid-write could print two currencies.
            public string CurrencyCode { get; init; } = WorklogCurrency.DefaultCode;
        }

        // One comment of the device: when, as the page shows it, and what - its line breaks as "\n".
        public sealed record Comment(string DateText, string Text);

        // ###########################################################################################
        // The history of `device`: `repairs` are its workbooks' documents (each built as a workbook
        // export builds it, on its own board's pictures), `boardText` its board as a person reads it,
        // `photoPath` where a photo's file name lies - null for one that is not there, left out.
        // ###########################################################################################
        public static Document Build(
            OwnedDeviceRecord device,
            string boardText,
            IEnumerable<WorkbookExportModel.Document> repairs,
            Func<string, string?> photoPath,
            DateTime generatedAt,
            string? currencyCode)
        {
            List<WorkbookExportModel.Document> oldestFirst = (repairs ?? [])
                .OrderBy(repair => repair.WorkbookId)
                .ToList();

            return new Document
            {
                DeviceId = device.Id,
                Name = device.Name.Trim(),
                Status = OwnedDeviceStatus.Normalise(device.Status),
                DetailLine = OwnedDeviceWording.DetailLine(device, boardText),
                ConditionLine = OwnedDeviceWording.ConditionLine(device.Condition),
                Description = OwnedDeviceWording.DescriptionText(device.Description),
                Links = (device.Links ?? []).Where(link => !string.IsNullOrWhiteSpace(link.Url)).ToList(),
                Comments = (device.Comments ?? [])
                    .Where(comment => !string.IsNullOrWhiteSpace(comment.Text))
                    .OrderBy(comment => comment.Date)
                    .ThenBy(comment => comment.Id)
                    .Select(comment => new Comment(OwnedDeviceWording.CommentDate(comment.Date), RichNote.TrimMarkup(comment.Text.ReplaceLineEndings("\n"))))
                    .ToList(),
                Photos = PhotosOf(device, photoPath),
                Repairs = oldestFirst,
                // The device page's own count, through its own method (code review, 2026-10-29).
                History = OwnedDeviceRepairs.Summarize(oldestFirst.Select(repair => new OwnedDeviceRepairs.Repair(
                    repair.Status,
                    repair.Sections.SelectMany(section => section.Entries).Select(entry => entry.Record)))),
                GeneratedAt = generatedAt,
                CurrencyCode = WorklogCurrency.NormalizeCode(currencyCode),
            };
        }

        // The device's photos in the order its page shows them - by their order, then their id (code
        // review, 2026-10-29) - each where it lies, one not there left out.
        private static List<WorkbookExportModel.Attachment> PhotosOf(OwnedDeviceRecord device, Func<string, string?> photoPath)
        {
            var photos = new List<WorkbookExportModel.Attachment>();

            foreach (WorklogAttachmentRecord photo in device.Photos.OrderBy(photo => photo.DisplayOrder).ThenBy(photo => photo.Id))
            {
                if (photoPath(photo.FileName) is { } path)
                {
                    photos.Add(new WorkbookExportModel.Attachment
                    {
                        FileName = photo.FileName,
                        Comment = photo.Comment ?? string.Empty,
                        FullPath = path,
                    });
                }
            }

            return photos;
        }

        // ###########################################################################################
        // The suggested file name, without an extension: "Device_3_Breadbin-2_20261029" - its id, its
        // name made safe for a file name, and the day it was made. The name is the user's own for
        // their own computer, so unlike a workbook's title it goes in: it is what tells two devices'
        // histories apart in a folder.
        // ###########################################################################################
        public static string BuildFileBaseName(Document document) =>
            BuildFileBaseName(document.DeviceId, document.Name, document.GeneratedAt);

        // The same, before there is a document - the save dialog's suggestion.
        public static string BuildFileBaseName(int deviceId, string? deviceName, DateTime generatedAt)
        {
            var segments = new List<string> { "Device", deviceId.ToString(CultureInfo.InvariantCulture) };

            if (WorkbookExportModel.SanitizeForFileName(deviceName) is { Length: > 0 } name)
            {
                segments.Add(name);
            }

            segments.Add(generatedAt.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
            return string.Join("_", segments);
        }

        // The folder a repair's files go under in the history's ZIP - the workbook's own folder name
        // on disk (WorklogManager), as each worklog's files keep theirs inside it.
        public static string RepairFolderName(WorkbookExportModel.Document repair) =>
            WorklogManager.BuildWorkbookFolderName(repair.WorkbookId);

        // The folder the device's own photos go under in the history's ZIP.
        public const string PhotosFolderName = "device photos";
    }
}
