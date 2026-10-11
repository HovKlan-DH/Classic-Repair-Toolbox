using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // RENAMING A SCHEMATIC IMAGE (owner request, 2026-10-10: "add management of the name of the
    // schematic image also" - the "Schematic name" column of the "Board schematics" sheet).
    //
    // Each row has a Name box beside its CAD name. What is typed is saved when the box is left or
    // Enter is pressed - and when the window closes with a name still typed - and Escape puts the
    // name back. A blank name, or one another image has (whatever its capitals), is refused with the
    // reason on the status line, and the box shows the name the image still has; a refusal while
    // closing keeps the window open, so the reason can be read (BoardFilesWindow.DraftChanges.cs). A
    // name the draft could not be written with (Excel holds the file) stays in its box, unsaved, to
    // be saved once the file is free.
    //
    // The save is DraftWorkbookStore.RenameSchematic, which carries the image's component
    // highlights and KiCad calibration over to the new name in the same write - both are filed
    // under the name in the board's JSON sidecar (SchematicRename has why). This machine's worklogs
    // and settings follow the image by its FILE when the board is next shown (LocalSchematicRename).
    // It is one of the window's draft changes: queued, off the UI thread, under the "please wait".
    //
    // Part of the BoardFilesWindow partial class - see BoardFilesWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        private void OnSchematicNameKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && sender is Control { DataContext: BoardSchematicRow row })
            {
                _ = this.CommitSchematicNameAsync(row);
                e.Handled = true;
            }
        }

        private void OnSchematicNameLostFocus(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { DataContext: BoardSchematicRow row })
            {
                _ = this.CommitSchematicNameAsync(row);
            }
        }

        // Escape in a Name box with an unsaved name: the name the image has, back in the box. True
        // when that is what Escape did (so it does not close the window too).
        private bool RevertNameEdit(object? source)
        {
            if (source is TextBox { DataContext: BoardSchematicRow row } && row.HasNameEdit)
            {
                row.NameDraft = row.SchematicName;
                return true;
            }

            return false;
        }

        // ###########################################################################################
        // Saves what the row's Name box holds as the image's name - or says why not and puts the
        // name back. True when the name was saved or there was nothing to save.
        //
        // Judged when it RUNS, after every change queued before it: a rename queued twice (Enter,
        // then the box losing focus) finds nothing left to save the second time.
        // ###########################################################################################
        internal Task<bool> CommitSchematicNameAsync(BoardSchematicRow row) =>
            this.ChangeDraftAsync(async () =>
            {
                if (!this.TryPlanRename(row, out PlannedRename? plan))
                {
                    return false;
                }

                if (plan is null)
                {
                    return true;
                }

                string draftsRoot = DraftManager.DraftsRoot;
                string excelDataFile = this.thisExcelDataFile;

                SchematicRenameOutcome outcome = await this.WriteDraftAsync(
                    CrtWaitWording.SavingSchematicName,
                    () => DraftWorkbookStore.RenameSchematic(draftsRoot, excelDataFile, plan.Row, plan.Wanted));

                bool renamed = this.ShowRenameOutcome(row, plan, outcome);

                // The draft changed outside this window: the list shows what it holds now.
                if (outcome is SchematicRenameOutcome.NameTaken or SchematicRenameOutcome.NotFound)
                {
                    await this.ReloadSchematicsUnderOverlayAsync();
                }

                return renamed;
            });

        // ###########################################################################################
        // Every typed name saved at once, on the UI thread - only for a shutdown of CRT or the
        // computer, which cannot wait for the queue (OnClosing).
        // ###########################################################################################
        private void SaveTypedNamesNow()
        {
            foreach (BoardSchematicRow row in this.Schematics.Where(row => row.HasNameEdit).ToList())
            {
                if (this.TryPlanRename(row, out PlannedRename? plan) && plan is not null)
                {
                    this.ShowRenameOutcome(
                        row,
                        plan,
                        DraftWorkbookStore.RenameSchematic(DraftManager.DraftsRoot, this.thisExcelDataFile, plan.Row, plan.Wanted));
                }
            }
        }

        // A rename to write: which row, from what, to what.
        private sealed record PlannedRename(SchematicRowRef Row, string OldName, string Wanted);

        // ###########################################################################################
        // What the row's Name box asks for. False, with the reason on the status line and the old
        // name back in the box, when it cannot be the image's name; true with no plan when there is
        // nothing to save. The other images are checked here first, so the reason comes at once; the
        // draft on disk is checked again by the save itself.
        // ###########################################################################################
        private bool TryPlanRename(BoardSchematicRow row, out PlannedRename? plan)
        {
            plan = null;

            if (!this.Schematics.Contains(row) || !row.HasNameEdit)
            {
                return true;
            }

            string oldName = row.SchematicName;
            string wanted = row.NameDraft?.Trim() ?? string.Empty;

            if (string.Equals(wanted, oldName.Trim(), System.StringComparison.Ordinal))
            {
                // Nothing to save - spaces typed round the name included.
                row.NameDraft = oldName;
                return true;
            }

            string? problem = SchematicRename.Problem(
                wanted,
                this.Schematics.Where(other => !ReferenceEquals(other, row)).Select(other => other.SchematicName));

            if (problem is not null)
            {
                row.NameDraft = oldName;
                this.ShowStatus(problem, isError: true);
                return false;
            }

            plan = new PlannedRename(this.RefOf(row), oldName, wanted);
            return true;
        }

        // What the save did, on screen. True when the image has the new name. The caller reads the
        // list again after NameTaken and NotFound.
        private bool ShowRenameOutcome(BoardSchematicRow row, PlannedRename plan, SchematicRenameOutcome outcome)
        {
            switch (outcome)
            {
                case SchematicRenameOutcome.Renamed:
                    this.ChangedTheDraft = true;
                    row.SchematicName = plan.Wanted;
                    row.NameDraft = plan.Wanted;
                    this.ShowStatus($"Renamed [{plan.OldName}] to [{plan.Wanted}].");
                    return true;

                case SchematicRenameOutcome.Unchanged:
                    row.NameDraft = plan.OldName;
                    return true;

                case SchematicRenameOutcome.Blank:
                    row.NameDraft = plan.OldName;
                    this.ShowStatus(SchematicRename.NeedsAName, isError: true);
                    return false;

                case SchematicRenameOutcome.NameTaken:
                    // Taken in the draft on disk though not in this list - changed outside it.
                    row.NameDraft = plan.OldName;
                    this.ShowStatus(SchematicRename.NameTaken(plan.Wanted), isError: true);
                    return false;

                case SchematicRenameOutcome.NotFound:
                    this.ShowStatus(
                        $"[{plan.OldName}] is no longer in the draft - it changed outside this window, so the list has been read again.",
                        isError: true);
                    return false;

                // ###########################################################################################
                // *** THE TYPED NAME STAYS IN ITS BOX (code review, 2026-10-11). *** The name is fine -
                // the file could not be written, most often because Excel holds it. The box used to show
                // the old name again, so after closing Excel the name had to be typed again, and a
                // second close of the window went through without it. It stays, unsaved: the next
                // Enter, leaving the box or closing the window saves it - and Escape gives it up.
                // ###########################################################################################
                case SchematicRenameOutcome.WriteFailed:
                    this.ShowStatus(BoardFilesWindow.NameNotWritten, isError: true);
                    return false;

                default:
                    // NoDraft: there is nothing left to save it into.
                    row.NameDraft = plan.OldName;
                    this.ShowStatus("The new name could not be saved - this board's draft is no longer there.", isError: true);
                    return false;
            }
        }

        internal const string NameNotWritten =
            "The new name could not be saved. If this board's Excel file is open in Excel, close it there and " +
            "try again - or press Escape in the name box to keep the old name.";
    }
}
