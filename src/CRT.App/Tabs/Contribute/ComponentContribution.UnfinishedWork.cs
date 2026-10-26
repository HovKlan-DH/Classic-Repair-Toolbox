using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // UNFINISHED WORK, AND TAKING A COMPONENT'S WORK BACK (CRT 4.0 step 1's last slice, 2026-10-25 -
    // the plan: "Unfinished work is fine. Only Submit asks for problems to be fixed ... Undo changes to
    // IC1 - one component, back to its published values").
    //
    // Owns:
    //   - the folder a picked file goes in when its row names none - its kind's usual one
    //     (ContributionFileLocations.DefaultFolderFor), so "Save to draft" no longer refuses it;
    //   - "Undo changes to [U8]": shown while the component differs from the published board
    //     (ComponentRevert.HasChanges), it puts every row of the component back as published - one
    //     step of the draft's history (DraftEdits.UndoComponentChangesAsync), so the draft's Undo on
    //     the Drafts tab takes it back - and shows the component again as it now is, or closes the
    //     window on a component the published board does not have.
    //
    // Part of the ComponentContributionWindow partial class - see ComponentContribution.axaml.cs for
    // the file map.
    // ###########################################################################################
    public partial class ComponentContributionWindow
    {
        // Lets a test stand in for the published board - the window reads the data folder's otherwise.
        internal Func<BoardData?>? PublishedBoardOverrideForTests { get; set; }

        // Completes when the button has been decided after a load - for tests.
        internal Task UndoChangesChecked { get; private set; } = Task.CompletedTask;

        // ---- the folder of a picked file ----------------------------------------------------------

        // The usual folder of this row's kind - a scope baseline is an image row naming a pin and a
        // scope setting (ComponentImageKinds, the one rule).
        private string DefaultFolderFor(IContributionFileRow row) => row switch
        {
            ContributionComponentImageRow image => ContributionFileLocations.DefaultFolderFor(
                ComponentImageKinds.IsScopeBaseline(new ComponentImageEntry
                {
                    Pin = image.Pin,
                    TimeDiv = image.TimeDiv,
                    VoltsDiv = image.VoltsDiv,
                    TriggerLevelVolts = image.TriggerLevelVolts,
                })
                    ? ContributionFileLocations.FileRowKind.ScopeBaseline
                    : ContributionFileLocations.FileRowKind.ComponentPicture),
            ContributionComponentLocalFileRow => ContributionFileLocations.DefaultFolderFor(ContributionFileLocations.FileRowKind.ComponentFile),
            _ => ContributionFileLocations.DefaultFolderFor(ContributionFileLocations.FileRowKind.BoardFile),
        };

        // ###########################################################################################
        // Every newly picked file whose row names no folder is filed in its kind's usual one, before
        // the save - so nothing is stored as a bare name at the top of the board (owner report,
        // 2026-09-25), and nothing is refused for it either.
        // ###########################################################################################
        private void FillDefaultFileLocations()
        {
            foreach (IContributionFileRow row in this.thisComponentImageRows.Cast<IContributionFileRow>()
                         .Concat(this.thisComponentLocalFileRows)
                         .Concat(this.thisBoardLocalFileRows))
            {
                if (!string.IsNullOrWhiteSpace(row.OriginalFilePath) &&
                    !string.IsNullOrWhiteSpace(row.File) &&
                    string.IsNullOrWhiteSpace(row.FileLocation))
                {
                    row.FileLocation = this.DefaultFolderFor(row);
                    this.SetAvailableFileLocations(row);
                    row.NotifyFileLocationChanged();
                }
            }
        }

        // ---- "Undo changes to <component>" --------------------------------------------------------

        // ###########################################################################################
        // Shows "Undo changes to [U8]" when the component differs from the published board - read off
        // the UI thread after the component is loaded (the draft as every editor shows it, against the
        // published board file). Hidden for a new component, which has nothing published to go back to.
        // ###########################################################################################
        private void CheckUndoChanges()
        {
            this.UndoChangesButton.IsVisible = false;

            string label = this.thisBoardLabel.Trim();

            if (this.thisIsNewComponent || label.Length == 0 || this.thisBoardExcelFile.Length == 0)
            {
                this.UndoChangesChecked = Task.CompletedTask;
                return;
            }

            this.UndoChangesChecked = this.CheckUndoChangesAsync(label);
        }

        private async Task CheckUndoChangesAsync(string label)
        {
            try
            {
                string dataRoot = this.thisDataRoot;
                string draftsRoot = DraftManager.DraftsRoot;
                string excelDataFile = this.thisBoardExcelFile;
                Func<BoardData?>? publishedOverride = this.PublishedBoardOverrideForTests;

                bool changed = await Task.Run(() =>
                {
                    DraftSession? session = DraftManager.Sessions.Find(draftsRoot, excelDataFile);
                    BoardData? draft = session?.Content.Board ?? DraftWorkbookStore.LoadDraftBoard(draftsRoot, excelDataFile);

                    if (draft is null)
                    {
                        return false;
                    }

                    BoardData? published = publishedOverride is not null
                        ? publishedOverride()
                        : DraftChangeSummaryReader.PublishedBoard(dataRoot, excelDataFile, draftOnly: false);

                    // A file of it replaced under its own name is a change too - the two lists of files.
                    return ComponentRevert.HasChanges(
                        draft,
                        published,
                        label,
                        DraftFiles.ListedIn(DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile)),
                        session?.ListedFiles ?? DraftFiles.ListedIn(DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile)));
                });

                // Still the component it was asked for - the window may have been loaded again meanwhile.
                if (!string.Equals(this.thisBoardLabel.Trim(), label, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                this.UndoChangesButton.Content = ReviewChangesWindow.UndoLabel(label);
                ToolTip.SetTip(
                    this.UndoChangesButton,
                    $"Put [{label}] back as it is published - every row of it, and anything typed here. The draft's Undo, on the Contribute tab, takes this back.");
                this.UndoChangesButton.IsVisible = changed;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not tell whether [{label}] differs from the published board - [{ex.Message}]");
            }
        }

        private async void OnUndoChangesClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.UndoChangesAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Undoing the changes to a component failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The component back to its published values, under this window's "please wait" - then the
        // board on screen read again, and the window showing the component as it now is (closed, for
        // a component the published board does not have: there is nothing left to show).
        // ###########################################################################################
        internal async Task UndoChangesAsync()
        {
            string label = this.thisBoardLabel.Trim();

            if (label.Length == 0 || this.thisBoardExcelFile.Length == 0)
            {
                return;
            }

            string dataRoot = this.thisDataRoot;
            string excelDataFile = this.thisBoardExcelFile;
            Func<BoardData?>? publishedOverride = this.PublishedBoardOverrideForTests;

            DraftWorkbookEditOutcome outcome = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.UndoingComponentChanges, async () =>
            {
                BoardData? published = publishedOverride is not null
                    ? publishedOverride()
                    : await Task.Run(() => DraftChangeSummaryReader.PublishedBoard(dataRoot, excelDataFile, draftOnly: false));

                return await DraftEdits.UndoComponentChangesAsync(excelDataFile, label, published, DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile));
            });

            if (outcome != DraftWorkbookEditOutcome.Saved)
            {
                this.ShowStatus($"[{label}] could not be put back. " + DraftSaveWording.EditNotSaved(outcome), true);
                return;
            }

            this.thisRefreshBoardAfterSave?.Invoke();

            BoardData? board = DraftManager.Sessions.Find(DraftManager.DraftsRoot, excelDataFile)?.Content.Board;

            if (board is null || !board.Components.Any(component => string.Equals(component.BoardLabel?.Trim(), label, StringComparison.OrdinalIgnoreCase)))
            {
                this.Close();
                return;
            }

            this.LoadComponent(board, dataRoot, this.thisHardwareName, this.thisBoardName, this.thisLocalRegion, label, excelDataFile);
            this.ShowStatus($"[{label}] is back as it is published. The draft's Undo, on the Contribute tab, takes this back.", false);
        }
    }
}
