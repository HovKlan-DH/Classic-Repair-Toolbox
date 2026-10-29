using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // TabWorkbooks, part: "EXPORT HISTORY" on a device's page (CRT 4.0 step 4, slice 4d - the plan:
    // "Export history puts every repair into one PDF or ZIP"). Two buttons, as a workbook's export
    // has, the format the BUTTON's (TabWorkbooks.Export.cs says why): the PDF, or the PDF with the
    // device's photos and every repair's attachments.
    //
    // Each repair is built as its own workbook export builds it (WorkbookExportModel), on the
    // pictures of ITS board - the board shown from what CRT has read, any other read from its board
    // file as it is viewed, off the UI thread (never DataManager's board load: Main reads its
    // statics after its own load). What goes in is DeviceHistoryExportModel's, the drawing
    // WorkbookPdfExporter's.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // A second export may not start while the save dialog or the write is up - as a workbook's.
        private bool thisIsExportingHistory;

        // Stands in for the save dialog in a test: the path chosen for (asZip, suggested name), or
        // null as if cancelled. Null in the app.
        internal Func<bool, string, Task<string?>>? PickHistoryPathOverrideForTests { get; set; }

        private async void OnExportHistoryClick(object? sender, RoutedEventArgs e)
        {
            // Nothing may leave an async void handler: App lets such an exception end CRT.
            try
            {
                await this.ExportDeviceHistoryAsync(asZip: false);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Exporting a device's history failed - [{ex}]");
            }
        }

        private async void OnExportHistoryZipClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.ExportDeviceHistoryAsync(asZip: true);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Exporting a device's history failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The history of the device whose page is shown, to the file the user picks. Built and
        // written under CRT's please-wait, off the UI thread - other boards are read and every
        // picture laid out. A failure is logged, as a workbook export's is: no file appears.
        // ###########################################################################################
        internal async Task ExportDeviceHistoryAsync(bool asZip)
        {
            if (this.thisIsExportingHistory || this.SelectedDevice is not { } device)
            {
                return;
            }

            this.thisIsExportingHistory = true;

            try
            {
                DateTime now = DateTime.Now;
                string extension = asZip ? "zip" : "pdf";

                if (await this.PickHistoryPathAsync(asZip, DeviceHistoryExportModel.BuildFileBaseName(device.Id, device.Name, now)) is not { } picked)
                {
                    return;
                }

                // The format is the button's - see WorkbookExportModel.EnsureFileExtension.
                string path = WorkbookExportModel.EnsureFileExtension(picked, extension);

                // Everything the UI thread holds, read here: the pool reads no control or setting.
                HistorySources sources = this.HistorySourcesOf(device, now);

                // The icon font is an Avalonia resource, read on the UI thread (TabWorkbooks.Export.cs).
                WorkbookPdfExporter.EnsureIconFontLoaded();

                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ExportingDeviceHistory(asZip), () => Task.Run(async () =>
                {
                    DeviceHistoryExportModel.Document history = await BuildDeviceHistoryAsync(sources);

                    if (asZip)
                    {
                        WorkbookPdfExporter.WriteHistoryZip(history, path);
                    }
                    else
                    {
                        WorkbookPdfExporter.WriteHistoryPdf(history, path);
                    }
                }));

                Logger.Info($"Exported the history of device [#{device.Id}] to [{path}]");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to export the history of device [#{device.Id}] - [{ex.Message}]");
            }
            finally
            {
                this.thisIsExportingHistory = false;
            }
        }

        private async Task<string?> PickHistoryPathAsync(bool asZip, string suggestedName)
        {
            if (this.PickHistoryPathOverrideForTests is { } pick)
            {
                return await pick(asZip, suggestedName);
            }

            // The workbook export's own dialog (TabWorkbooks.Export.cs).
            return await this.PickExportPathAsync(
                asZip,
                asZip ? "Export the device's history as ZIP" : "Export the device's history as PDF",
                suggestedName);
        }

        // ###########################################################################################
        // What a history is built from, read on the UI thread: the device, its repairs as they are on
        // disk, the board shown with its pictures (null while CRT has not worked them out - the board
        // still loading, or unreadable), how every other board is viewed, and the currency.
        // ###########################################################################################
        private sealed record HistorySources(
            OwnedDeviceRecord Device,
            string BoardText,
            List<WorkbookRecord> Repairs,
            string ShownBoardKey,
            IReadOnlyDictionary<string, string>? ShownBoardPictures,
            BoardRoots Roots,
            DateTime GeneratedAt,
            string CurrencyCode);

        private HistorySources HistorySourcesOf(OwnedDeviceRecord device, DateTime generatedAt) => new(
            device,
            BoardTextOf(device.BoardKey),
            OwnedDeviceRepairs.Of(device.Id, WorklogManager.GetAllWorkbooks()),
            this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey() ?? string.Empty,
            this.ShownBoardPicturePaths(),
            this.BoardRootsNow(),
            generatedAt,
            UserSettings.WorklogCurrencyCode);

        // Where a board is read from as it is viewed - read on the UI thread, used anywhere.
        private sealed record BoardRoots(string DataRoot, string DraftsRoot, bool ViewOfficialOnly);

        private BoardRoots BoardRootsNow() =>
            new(DataManager.DataRoot, DraftManager.DraftsRoot, UserSettings.ViewOfficialPublishedOnly);

        // Built for a test from the device's page as the export would build it - nothing written.
        internal async Task<DeviceHistoryExportModel.Document?> BuildDeviceHistoryForTestsAsync() =>
            this.SelectedDevice is { } device ? await BuildDeviceHistoryAsync(this.HistorySourcesOf(device, DateTime.Now)) : null;

        // ###########################################################################################
        // The history, every repair built as its workbook export builds it - each board's pictures
        // read once: the board shown's as CRT holds them, and read like any other's while CRT holds
        // none (code review, 2026-10-29: its repairs went out without their board). A workbook naming
        // no board ("boardKey": null, hand-edited) is on the board "" - its repair kept, without
        // pictures (the same review: the whole export threw).
        // ###########################################################################################
        private static async Task<DeviceHistoryExportModel.Document> BuildDeviceHistoryAsync(HistorySources sources)
        {
            // The boards read together, not one after another (code review, 2026-10-29).
            string[] boards = sources.Repairs.Select(BoardOf).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

            IReadOnlyDictionary<string, string>[] read = await Task.WhenAll(boards.Select(boardKey =>
                string.Equals(boardKey, sources.ShownBoardKey, StringComparison.OrdinalIgnoreCase) && sources.ShownBoardPictures is { } shown
                    ? Task.FromResult(shown)
                    : ReadPicturesOfBoardAsync(boardKey, sources.Roots)));

            var picturesByBoard = new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < boards.Length; index++)
            {
                picturesByBoard[boards[index]] = read[index];
            }

            List<WorkbookExportModel.Document> repairs = sources.Repairs
                .Select(repair => WorkbookExportModel.Build(
                    repair,
                    WorklogManager.GetEntries(repair.Id),
                    picturesByBoard[BoardOf(repair)],
                    entryId => WorklogManager.GetEntryAttachmentsFolderPath(repair.Id, entryId),
                    sources.GeneratedAt,
                    sources.CurrencyCode))
                .ToList();

            return DeviceHistoryExportModel.Build(
                sources.Device,
                sources.BoardText,
                repairs,
                fileName => OwnedDeviceStore.PhotoPath(sources.Device.Id, fileName),
                sources.GeneratedAt,
                sources.CurrencyCode);
        }

        // The board a repair names - "" for none.
        private static string BoardOf(WorkbookRecord repair) => repair.BoardKey ?? string.Empty;

        // ###########################################################################################
        // The pictures of a board, by name, read here - its board file as the board is VIEWED (the
        // draft's, unless "view boards as officially published"), as DataManager would choose it,
        // read without its board load. A board CRT no longer lists has none: its repairs print
        // without their board, as a missing picture always has. Both exports' (PicturesOfBoardAsync).
        // ###########################################################################################
        private static async Task<IReadOnlyDictionary<string, string>> ReadPicturesOfBoardAsync(string boardKey, BoardRoots roots)
        {
            if (Main.FindEntryForBoardKey(boardKey) is not { } entry || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            // The board file and its files' draft folder from ONE resolution, by CRT.Data's own rule
            // (code review, 2026-10-29: both rules were copied here).
            DraftBoardSource.ViewedBoard viewed = DraftBoardSource.ResolveViewed(
                roots.DataRoot, roots.DraftsRoot, entry.ExcelDataFile, roots.ViewOfficialOnly, entry.IsPublished);

            // Its rows and its list of files from ONE read of its board file, the list kept for the
            // pictures' lookups (code review, 2026-10-29: read twice) - off the UI thread. A new
            // board viewed as published has no file: no pictures.
            return await Task.Run(() =>
            {
                var listings = new BoardFileListings();

                if (listings.ReadDocument(viewed.BoardFile) is not { } document)
                {
                    return (IReadOnlyDictionary<string, string>)new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                }

                BoardData board = BoardDocumentMapping.ToContent(document).Board;
                return BoardPicturePaths.ByName(board, roots.DataRoot, viewed.DraftFolder, entry.ExcelDataFile, listings);
            });
        }
    }
}
