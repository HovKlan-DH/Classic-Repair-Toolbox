using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // EVERY CHANGE THIS WINDOW MAKES TO THE DRAFT: ONE AT A TIME, OFF THE UI THREAD, UNDER THE
    // "PLEASE WAIT" (code review, 2026-10-10).
    //
    // A rename, a CAD name, a new order and a removal each rewrite the whole draft workbook and its
    // sidecar - up to a couple of seconds on the largest board - and they did it on the UI thread:
    // the window froze with nothing to say why, against the rule that every wait the user watches
    // runs under BusyOverlay. Now each is saved on the pool while the overlay holds the window -
    // through the draft held in memory since 2026-10-25 (ApplyToDraftAsync), so the change itself is
    // applied on the UI thread and the save is what waits.
    //
    // *** ONE AT A TIME. *** The overlay takes every NEW click, but a click whose press came first
    // still lands - typing a name and then clicking another row's Remove starts the rename (the box
    // loses focus on the press) and the removal (on the release). Two writes of one workbook at
    // once would lose one of them, so each change waits for the one before it, and reads the rows
    // only once it runs - after a rename before it has given its row the new name.
    //
    // *** CLOSING WAITS FOR THEM, AND FOR A NAME STILL TYPED (OnClosing). *** The Drafts tab looks at
    // ChangedTheDraft once the window has closed, so a change still being written then would be
    // missed. And a typed name that cannot be saved (another image has it, or Excel holds the file)
    // keeps the window open with the reason on its status line - closing used to drop it without a
    // word.
    //
    // Part of the BoardFilesWindow partial class - see BoardFilesWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        // The last change queued; each runs once the one before it has finished.
        private Task thisDraftChanges = Task.CompletedTask;

        // Set once the window may close: everything queued is done and every typed name saved.
        private bool thisReadyToClose;

        // The save-then-close under way, if any - a second close waits for the same one.
        private Task? thisClosingAfterSaves;

        // Completes when every change queued so far has finished - for tests, which wait for it as
        // the user waits for the overlay.
        internal Task DraftChanging => this.thisDraftChanges;

        // Asks before removing an image with something drawn on it - for tests, which cannot answer a
        // dialog. Null is the real dialog (RemoveSchematicWindow).
        internal Func<string, SchematicRemovalLoss, Task<bool>>? ConfirmRemovalOverrideForTests { get; set; }

        // ###########################################################################################
        // Queues one change: it runs once every change queued before it has finished.
        // ###########################################################################################
        private Task<T> ChangeDraftAsync<T>(Func<Task<T>> change)
        {
            Task before = this.thisDraftChanges;
            Task<T> next = RunAfterAsync();
            this.thisDraftChanges = next;

            return next;

            async Task<T> RunAfterAsync()
            {
                try
                {
                    await before;
                }
                catch (Exception ex)
                {
                    // Already told to whoever queued it; it is no reason to drop this one.
                    Logger.Warning($"A change to the draft failed before the next one ran - [{ex.Message}]");
                }

                return await change();
            }
        }

        private Task ChangeDraftAsync(Func<Task> change) =>
            this.ChangeDraftAsync(async () =>
            {
                await change();
                return true;
            });

        // One piece of file work, on the pool, under the window's "please wait".
        private Task<T> WriteDraftAsync<T>(string message, Func<T> write) =>
            BusyOverlay.RunLocalAsync(this, message, () => Task.Run(write));

        // ###########################################################################################
        // One change of the draft's rows through the draft held in memory (DraftEdits - CRT 4.0 step
        // 1, 2026-10-25): applied on the UI thread, so the Drafts tab's table shows it and its Ctrl+Z
        // takes it back, and saved on the pool, under the window's "please wait". A save that cannot
        // be made is taken back out, so this window's "could not" lines still mean nothing changed.
        // ###########################################################################################
        private Task<DraftWorkbookEditOutcome> ApplyToDraftAsync(
            string message,
            string description,
            Func<BoardData, BoardData?> transform,
            IReadOnlyDictionary<string, string?>? follow = null)
        {
            string excelDataFile = this.thisExcelDataFile;

            return BusyOverlay.RunLocalAsync(this, message, () => DraftEdits.ApplyAndSaveAsync(excelDataFile, description, transform, follow));
        }

        // ###########################################################################################
        // Which row of the draft a row on screen is: its name, its image file, and its place among
        // the rows sharing both (SchematicRows) - so of two images sharing a name, the change goes
        // to the one it was made on.
        // ###########################################################################################
        private SchematicRowRef RefOf(BoardSchematicRow row)
        {
            List<(string SchematicName, string ImageFile)> rows = this.Schematics
                .Select(other => (other.SchematicName, other.ImageFile))
                .ToList();

            return SchematicRows.RefOf(rows, this.Schematics.IndexOf(row));
        }

        // ###########################################################################################
        // Closing: once every change has been written and every typed name saved (see the header).
        // Not on a shutdown of CRT or the computer, which nothing should hold up: a typed name is
        // then saved at once, as before - unless a change is still being written, which it must not
        // run beside.
        // ###########################################################################################
        protected override void OnClosing(WindowClosingEventArgs e)
        {
            bool somethingToFinish = !this.thisDraftChanges.IsCompleted || this.Schematics.Any(row => row.HasNameEdit);

            if (somethingToFinish && !this.thisReadyToClose)
            {
                if (e.CloseReason is WindowCloseReason.WindowClosing or WindowCloseReason.Undefined)
                {
                    e.Cancel = true;

                    if (this.thisClosingAfterSaves is null || this.thisClosingAfterSaves.IsCompleted)
                    {
                        this.thisClosingAfterSaves = this.SaveThenCloseAsync();
                    }
                }
                else if (this.thisDraftChanges.IsCompleted)
                {
                    this.SaveTypedNamesNow();
                }
            }

            base.OnClosing(e);
        }

        private async Task SaveThenCloseAsync()
        {
            bool allSaved = true;

            foreach (BoardSchematicRow row in this.Schematics.Where(row => row.HasNameEdit).ToList())
            {
                allSaved &= await this.CommitSchematicNameAsync(row);
            }

            try
            {
                await this.thisDraftChanges;
            }
            catch (Exception ex)
            {
                Logger.Warning($"A change to the draft failed while the window was closing - [{ex.Message}]");
            }

            if (!allSaved)
            {
                // The status line says why; the window stays so it can be read.
                return;
            }

            this.thisReadyToClose = true;
            this.Close();
        }
    }
}
