using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // PICKING EACH SCHEMATIC IMAGE'S CAD NAME (owner request, 2026-10-10: "In the "Schematic
    // images" then have that mapping available, so I can select a KiCad name. Once selected, it
    // cannot be selected again").
    //
    // A CAD name is the KiCad VIEW an image's traces are drawn from. The names on offer are read
    // from the very files the Schematics tab draws this board's traces from - the published "KiCad
    // data" folder and the draft's (KiCadRawFileScanner.ScanBoard) - and WITHOUT parsing a PCB
    // (KiCadRawProjectLoader.LoadViewDisplayNamesAsync), so reading them is quick. They are read
    // alongside the images while the window opens, under its "please wait" (Opening), so every row
    // shows its choices from the start. Which name each image may pick is SchematicCadNameChoices'.
    //
    // *** THE DROP-DOWNS' LISTS ARE CHANGED IN PLACE, NEVER REPLACED. *** A ComboBox whose items
    // are swapped lets go of its selection and writes "nothing" back through its binding, so every
    // image would have lost its CAD name the moment the names arrived. Each row keeps one list and
    // BoardSchematicRow.ApplyCadNameChoices inserts and removes within it.
    //
    // Part of the BoardFilesWindow partial class - see BoardFilesWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        // The KiCad view names on offer - empty until read, and for a board with no KiCad data.
        private IReadOnlyList<string> thisCadViewNames = Array.Empty<string>();

        // Completes when the names have been read and handed to the rows - for tests.
        internal Task CadNamesLoading { get; private set; } = Task.CompletedTask;

        // ###########################################################################################
        // Reads the view names and hands them to every row. The line under the section's heading
        // says what is going on until then, and afterwards only when there is nothing to pick.
        // ###########################################################################################
        private async Task LoadCadViewNamesAsync()
        {
            this.ShowCadNamesState("Reading the board's KiCad data for its CAD names...");

            string publishedFolder = this.PublishedKiCadFolder;
            string draftFolder = this.KiCadImportFolder;

            (int fileCount, List<string> names) = await Task.Run(async () =>
            {
                List<string> files = KiCadRawFileScanner.ScanBoard(publishedFolder, draftFolder);

                return files.Count == 0
                    ? (0, new List<string>())
                    : (files.Count, await KiCadRawProjectLoader.LoadViewDisplayNamesAsync(files));
            });

            this.thisCadViewNames = names;

            this.ShowCadNamesState(
                fileCount == 0
                    ? "This board has no KiCad data yet, so there is no CAD name to pick. Add it with \"KiCad data\" in the Drafts tab."
                    : names.Count == 0
                        ? "The board's KiCad data has no views, so there is no CAD name to pick."
                        : string.Empty);

            this.ApplyCadNameChoices(this.Schematics);
        }

        private void ShowCadNamesState(string text)
        {
            this.CadNamesStateText.Text = text;
            this.CadNamesStateText.IsVisible = text.Length > 0;
        }

        // Every row's choices, each judged against all the others - one image's pick changes
        // what the rest may pick.
        private void ApplyCadNameChoices(IReadOnlyList<BoardSchematicRow> rows)
        {
            List<(string SchematicName, string CadName)> images = rows
                .Select(row => (row.SchematicName, row.CadName))
                .ToList();

            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].ApplyCadNameChoices(SchematicCadNameChoices.For(images, i, this.thisCadViewNames));
            }
        }

        // ###########################################################################################
        // A row's drop-down picked another CAD name. Queued AFTER the binding that reported it has
        // finished: a failed save rebuilds the list, which must not happen underneath a ComboBox
        // still in the middle of setting its selection.
        // ###########################################################################################
        private void OnCadNameChosen(BoardSchematicRow row, string cadName) =>
            Dispatcher.UIThread.Post(() => _ = this.SaveCadNameAsync(row, cadName));

        // ###########################################################################################
        // Writes one image's CAD name into the draft's "Board schematics" sheet, then lets every
        // drop-down catch up. The row is found by its name AND its image file (SchematicRows), so
        // of two images sharing a name the one picked for gets it. One of the window's draft
        // changes: queued, off the UI thread, under the "please wait" (BoardFilesWindow.DraftChanges.cs).
        // ###########################################################################################
        internal Task SaveCadNameAsync(BoardSchematicRow row, string cadName) =>
            this.ChangeDraftAsync(async () =>
            {
                if (!this.Schematics.Contains(row))
                {
                    // The list was read again meanwhile; the pick belongs to a row no longer shown.
                    return;
                }

                string draftsRoot = DraftManager.DraftsRoot;
                string excelDataFile = this.thisExcelDataFile;
                SchematicRowRef target = this.RefOf(row);
                bool found = false;

                bool saved = await this.WriteDraftAsync(CrtWaitWording.SavingCadName, () => DraftWorkbookStore.Edit(
                    draftsRoot,
                    excelDataFile,
                    board =>
                    {
                        (BoardData updated, bool wasFound) = SchematicRows.WithCadName(board, target, cadName);
                        found = wasFound;
                        return updated;
                    }));

                if (!saved || !found)
                {
                    this.ShowStatus(
                        saved
                            ? $"[{row.SchematicName}] is no longer in the draft - it changed outside this window, so the list has been read again."
                            : "The CAD name could not be saved. If this board's Excel file is open in Excel, close it there and try again.",
                        isError: true);

                    await this.ReloadSchematicsUnderOverlayAsync();
                    return;
                }

                this.ChangedTheDraft = true;

                row.CadName = cadName;
                this.ApplyCadNameChoices(this.Schematics);

                this.ShowStatus(cadName.Length == 0
                    ? $"[{row.SchematicName}] has no CAD name now."
                    : $"[{row.SchematicName}] now has the CAD name [{cadName}].");
            });
    }
}
