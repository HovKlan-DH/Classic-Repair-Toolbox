using System;
using System.Linq;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WorkbookPdfExporter, part: A DEVICE'S HISTORY (CRT 4.0 step 4, slice 4d - the plan: "Export
    // history puts every repair into one PDF or ZIP"), painted from DeviceHistoryExportModel.
    //
    // The device heads every page. The document opens with what it is - where it came from, its
    // notes, what its repairs add up to (the workbook's own summary panel, under the device page's
    // headline), the list of its repairs, the changes made to it and its photos - then every repair,
    // oldest first, each from a fresh page and drawn EXACTLY as its own workbook export draws it
    // (ComposeHeader, ComposeContent): one drawing of a repair, never a second.
    //
    // The ZIP holds the PDF, the device's photos, and each repair's attachments under its workbook's
    // and worklog's folder names ("workbook_12/worklog_3/front.jpg"), as they lie on disk.
    // ###########################################################################################
    public static partial class WorkbookPdfExporter
    {
        // Writes the history as a PDF, overwriting the path. Throws as WritePdf does.
        public static void WriteHistoryPdf(DeviceHistoryExportModel.Document history, string outputPath)
        {
            ArgumentNullException.ThrowIfNull(history);

            Document.Create(container => container.Page(page => ComposePage(
                page,
                e => ComposeHistoryHeader(e, history),
                e => ComposeHistoryContent(e, history)))).GeneratePdf(outputPath);
        }

        // The history's PDF, the device's photos and every repair's attachments, as one ZIP.
        public static void WriteHistoryZip(DeviceHistoryExportModel.Document history, string outputPath)
        {
            ArgumentNullException.ThrowIfNull(history);

            WriteArchive(
                outputPath,
                $"{DeviceHistoryExportModel.BuildFileBaseName(history)}.pdf",
                pdfPath => WriteHistoryPdf(history, pdfPath),
                (archive, usedEntryNames) =>
                {
                    foreach (var photo in history.Photos)
                    {
                        AddFile(archive, usedEntryNames, photo.FullPath, $"{DeviceHistoryExportModel.PhotosFolderName}/{photo.FileName}", HistoryExportName);
                    }

                    foreach (var repair in history.Repairs)
                    {
                        AddAttachments(archive, usedEntryNames, repair, DeviceHistoryExportModel.RepairFolderName(repair) + "/", HistoryExportName);
                    }
                });
        }

        // The device on every page: its name, then what this is, what it is and when it was made.
        private static void ComposeHistoryHeader(IContainer container, DeviceHistoryExportModel.Document history)
        {
            container.Column(column =>
            {
                column.Item().Text(PillLabel(history.Name, "(unnamed device)")).FontSize(16).Bold().FontColor(HeadingColor);

                column.Item().PaddingTop(2).Text(text =>
                {
                    text.DefaultTextStyle(x => x.FontSize(9).FontColor(LabelColor));
                    text.Span(OwnedDeviceWording.HistoryLabel);

                    foreach (string part in new[] { history.DetailLine, history.Status }.Where(part => !string.IsNullOrWhiteSpace(part)))
                    {
                        text.Span($"  ·  {part}");
                    }

                    text.Span($"  ·  exported {FormatDate(history.GeneratedAt)}");
                });

                column.Item().PaddingTop(6).LineHorizontal(1).LineColor(RuleColor);
            });
        }

        private static void ComposeHistoryContent(IContainer container, DeviceHistoryExportModel.Document history)
        {
            container.PaddingTop(10).Column(column =>
            {
                column.Spacing(14);

                if (history.OriginLine.Length > 0)
                {
                    column.Item().Element(e => ComposeLinkedText(e, history.OriginLine, 10, HeadingColor));
                }

                // Notes are free text that may well carry a link, as a workbook's note does.
                if (history.Notes.Length > 0)
                {
                    column.Item().Element(e => ComposeLinkedText(e, history.Notes, 10, HeadingColor));
                }

                column.Item().Element(e => ComposeSummary(
                    e,
                    history.History.Totals,
                    history.CurrencyCode,
                    OwnedDeviceRepairs.HeadlineStats(history.History, history.CurrencyCode)));

                column.Item().Element(e => ComposeSubList(
                    e,
                    OwnedDeviceWording.RepairsHeading,
                    history.Repairs.Count == 0
                        ? [OwnedDeviceWording.NoRepairsRecorded]
                        : history.Repairs.Select(repair => OwnedDeviceWording.HistoryRepairLine(
                            repair.WorkbookId, repair.Title, repair.Status, WorkbookDateText(repair), repair.Totals.WorklogCount))));

                if (history.Changes.Count > 0)
                {
                    column.Item().Element(e => ComposeSubList(
                        e,
                        OwnedDeviceWording.ChangesHeading(history.Changes.Count),
                        history.Changes.Select(change => $"{OwnedDeviceWording.ChangeDate(change.Date)}  {change.Text}")));
                }

                if (history.Photos.Count > 0)
                {
                    column.Item().Element(e => ComposePhotoGrid(e, history.Photos));
                }

                // Every repair from a fresh page, as its own workbook export draws it.
                foreach (var repair in history.Repairs)
                {
                    column.Item().PageBreak();
                    column.Item().Element(e => ComposeHeader(e, repair, withExportDate: false));
                    column.Item().Element(e => ComposeContent(e, repair));
                }
            });
        }
    }
}
