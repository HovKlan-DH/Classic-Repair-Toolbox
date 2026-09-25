using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // TabWorkbooks, part: EXPORTING a workbook as a document to send to someone else.
    //
    // The typical use is documentation for the customer whose machine was repaired: what was wrong,
    // where on the board, what was done, how long it took and what it cost. That is why the export
    // is a readable document rather than a data dump - see WorkbookPdfExporter, which decides how
    // it looks, and WorkbookExportModel, which decides what goes in it.
    //
    // TWO FORMATS, offered from one button:
    //   PDF - the document itself. One file, opens anywhere, prints.
    //   ZIP - the same PDF plus the workbook's original photos and attached files. For when the
    //         recipient needs the actual bytes: a PDF shows a photo at page resolution and cannot
    //         carry an attached datasheet or scope capture at all.
    //
    // Both go through the same model, so the archive's PDF and a bare PDF export of the same
    // workbook are the same document.
    //
    // The button lives beside Edit/Delete in WorkbookHeaderActionsPanel because all three act on
    // the workbook the top-line is showing (thisHeaderWorkbook), and it is first of the three as
    // the only non-destructive one.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // Guards against a second export starting while the file picker or the write is still up -
        // the same reasoning as thisIsOpeningEntryEditor. ShowDialog does not block the dispatcher,
        // so a second click during the picker would otherwise start a parallel export of the same
        // workbook to the same suggested file name.
        private bool thisIsExportingWorkbook;

        private void OnExportWorkbookClick(object? sender, RoutedEventArgs e) =>
            this.ExportWorkbook(asZip: false);

        private void OnExportWorkbookZipClick(object? sender, RoutedEventArgs e) =>
            this.ExportWorkbook(asZip: true);

        // ###########################################################################################
        // Exports the workbook the top-line is showing, in the given format.
        //
        // ONE implementation for both buttons - they differ only in the extension offered and which
        // writer runs, so a second copy would be two places for the picker, the guard, the
        // off-thread write and the error handling to drift apart.
        //
        // The format comes from the BUTTON, not from the save dialog's file-type list. It was the
        // other way round at first, with one button and both types in the dialog, and the ZIP was
        // then invisible from the tab - reported as "I do not see the Export to ZIP anywhere". A
        // format only discoverable by opening a dropdown inside a dialog the user opened for
        // another reason is not discoverable.
        // ###########################################################################################
        private async void ExportWorkbook(bool asZip)
        {
            if (this.thisIsExportingWorkbook)
                return;

            var workbook = this.thisHeaderWorkbook;
            if (workbook == null)
                return;

            if (TopLevel.GetTopLevel(this) is not { StorageProvider: { } storageProvider })
                return;

            this.thisIsExportingWorkbook = true;

            try
            {
                var document = this.BuildExportDocument(workbook);
                string extension = asZip ? "zip" : "pdf";

                var fileType = asZip
                    ? new FilePickerFileType("ZIP archive (PDF + photos and files)")
                    {
                        Patterns = new[] { "*.zip" },
                        MimeTypes = new[] { "application/zip" },
                        AppleUniformTypeIdentifiers = new[] { "public.zip-archive" }
                    }
                    : new FilePickerFileType("PDF document")
                    {
                        Patterns = new[] { "*.pdf" },
                        MimeTypes = new[] { "application/pdf" },
                        AppleUniformTypeIdentifiers = new[] { "com.adobe.pdf" }
                    };

                var file = await storageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
                {
                    Title = asZip ? "Export workbook as ZIP" : "Export workbook as PDF",
                    SuggestedFileName = WorkbookExportModel.BuildFileBaseName(document),
                    DefaultExtension = extension,
                    ShowOverwritePrompt = true,
                    FileTypeChoices = new[] { fileType }
                });

                if (file == null)
                    return;

                string path = file.Path.LocalPath;

                // Not every picker backend appends the default extension to a name typed without
                // one, and the format here is the button's rather than the file name's - so the
                // extension is enforced rather than read back, or an "Export to ZIP" could write a
                // ZIP to a name the OS then treats as extensionless.
                path = WorkbookExportModel.EnsureFileExtension(path, extension);

                // The PDF's icon font is an Avalonia resource, and Avalonia's AssetLoader cannot be
                // reached from an arbitrary background thread - so its bytes are read HERE, while
                // still on the UI thread, and the export below only registers what is already in
                // memory. Without this the writer silently produced documents with no icons at all.
                WorkbookPdfExporter.EnsureIconFontLoaded();

                // Off the UI thread: a workbook with a board scan and a dozen photos takes real
                // time to lay out and encode, and doing it inline freezes the window with no
                // indication that anything is happening.
                await Task.Run(() =>
                {
                    if (asZip)
                    {
                        WorkbookPdfExporter.WriteZip(document, path);
                    }
                    else
                    {
                        WorkbookPdfExporter.WritePdf(document, path);
                    }
                });

                Logger.Info($"Exported workbook [#{workbook.Id}] to [{path}]");

                // DELIBERATELY not opened afterwards. ExternalTargetLauncher - the only sanctioned
                // way this app hands a file to the OS shell - admits a local path only if it
                // resolves INSIDE the current data root, and an export is saved wherever the user
                // chose, which is essentially never in there. Calling it would refuse every export
                // and log a "rejected external target" warning about the file it had just written.
                //
                // Nor is it worth widening that rule for this: the user picked the location a
                // moment ago in a save dialog, so they already know where the file is.
            }
            catch (Exception ex)
            {
                // Logged rather than swallowed, and deliberately not turned into a dialog: the app
                // has no error-modal convention, and every other failure on this tab reports the
                // same way. The user sees no file appear, which with the log line is enough to
                // diagnose. A missing file is honest; a silent partial write would not be.
                Logger.Warning($"Failed to export workbook [#{workbook.Id}]: [{ex.Message}]");
            }
            finally
            {
                this.thisIsExportingWorkbook = false;
            }
        }

        // ###########################################################################################
        // Assembles the export document for one workbook.
        //
        // Reads the entries FRESH (WorklogManager.GetEntries, not the within-pass cache): an export
        // is a document a customer keeps, and it must reflect what is on disk at the moment the
        // button was pressed rather than whatever a refresh happened to have cached earlier in the
        // session. Unlike the summary strip this runs once per click, so the extra read costs
        // nothing worth saving.
        //
        // The schematic image paths are resolved here, from this tab's board data, because
        // WorkbookExportModel deliberately knows nothing about BoardData or the data root.
        // ###########################################################################################
        private WorkbookExportModel.Document BuildExportDocument(WorkbookRecord workbook)
        {
            var entries = WorklogManager.GetEntries(workbook.Id);

            return WorkbookExportModel.Build(
                workbook,
                entries,
                this.BuildSchematicImagePaths(),
                entryId => WorklogManager.GetEntryAttachmentsFolderPath(workbook.Id, entryId),
                DateTime.Now,

                // The currency is captured HERE, on the UI thread, and travels in the document -
                // the write itself runs on a background thread, where a settings read would race a
                // user changing the setting mid-export.
                UserSettings.WorklogCurrencyCode);
        }

        // ###########################################################################################
        // Maps each of the current board's schematic names to its image file on disk.
        //
        // GroupBy before ToDictionary for the same reason RefreshBoardPreviews does it: board Excel
        // files arrive from classic-repair-toolbox.dk independently of app releases and are not
        // checked for duplicate schematic names, so a bare ToDictionary throws on the first
        // duplicate - which here would take down the export of an otherwise perfectly good workbook.
        // First wins, matching the board pane.
        // ###########################################################################################
        private Dictionary<string, string> BuildSchematicImagePaths()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var schematics = this.CurrentBoardDataForPreviews?.Schematics;
            if (schematics == null)
                return result;

            foreach (var group in schematics
                .Where(s => !string.IsNullOrWhiteSpace(s.SchematicName) && !string.IsNullOrWhiteSpace(s.SchematicImageFile))
                .GroupBy(s => s.SchematicName, StringComparer.OrdinalIgnoreCase))
            {
                var schematic = group.First();
                result[group.Key] = Path.Combine(
                    DataManager.DataRoot,
                    schematic.SchematicImageFile.Replace('/', Path.DirectorySeparatorChar));
            }

            return result;
        }

        // Lets the headless tests build the very document an export would write, without a file
        // picker and without writing anything - the same override-then-real seam the rest of this
        // tab uses. The PDF/ZIP writing itself is not covered: it needs QuestPDF to produce bytes,
        // and asserting on those tests the library rather than this app.
        internal WorkbookExportModel.Document BuildExportDocumentForTests(WorkbookRecord workbook) =>
            this.BuildExportDocument(workbook);
    }
}
