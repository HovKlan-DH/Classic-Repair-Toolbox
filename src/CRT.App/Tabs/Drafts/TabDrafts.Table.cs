using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // TABLE MODE on the Drafts tab - "Edit in table format" (owner request, 2026-09-24).
    //
    // Owns: which draft's table is open, switching the tab between its list layout and its table
    // layout, and every prompt about unsaved table edits. The table itself is BoardTableEditor;
    // this part only decides when it opens, closes and saves.
    //
    // *** ONE RULE RUNS THROUGH ALL OF IT: the table's unsaved edits are never silently lost, and
    // never silently left out. *** Closing asks. Discarding the draft is covered by the discard
    // confirmation itself. Managing files and submitting "settle" the edits first (save, or throw
    // away and reload), because both act on the draft FILE - an import would make the table
    // unsaveable, and a submission would go without what is on screen. Quitting asks too, from
    // Main.OnWindowClosing.
    // ###########################################################################################
    public partial class TabDrafts
    {
        internal bool IsTableOpen => this.thisTableEntry is not null;

        internal bool HasUnsavedTableEdits => this.IsTableOpen && this.TableEditor.HasUnsavedChanges;

        // ###########################################################################################
        // Whether the table is open on THIS board's draft and holds edits not saved yet - the
        // question the Contribute tab's component editor and the label editor ask before they write
        // the same draft (owner request, 2026-09-24). A save of theirs would make the table's
        // own save refused, and its edits lost; so they hold back and say why instead. A table open
        // on another board, or with nothing unsaved (it catches up by itself), is no obstacle.
        // ###########################################################################################
        internal bool HasUnsavedTableEditsFor(string? excelDataFile) =>
            this.HasUnsavedTableEdits &&
            this.thisTableEntry is not null &&
            string.Equals(this.thisTableEntry.ExcelDataFile, excelDataFile?.Trim(), StringComparison.OrdinalIgnoreCase);

        private bool IsTableOpenFor(HardwareBoardEntry entry) =>
            this.thisTableEntry is not null &&
            string.Equals(this.thisTableEntry.ExcelDataFile, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase);

        // The row's "Edit in table format" / "Close table" button.
        internal async Task ToggleTableAsync(HardwareBoardEntry entry)
        {
            if (this.IsTableOpenFor(entry))
            {
                await this.CloseTableAsync();
            }
            else
            {
                await this.OpenTableAsync(entry);
            }
        }

        // ###########################################################################################
        // Opens one draft's table. The published board is what the colours compare against - none
        // for a system that exists only as a draft, where nothing is coloured.
        // ###########################################################################################
        internal async Task OpenTableAsync(HardwareBoardEntry entry)
        {
            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            DraftStatus? status = DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile);

            BoardData? published = this.PublishedBoardOverrideForTests is not null
                ? this.PublishedBoardOverrideForTests(entry)
                : await TabDrafts.LoadPublishedBoardAsync(entry, status);

            // Resting on a file cell shows the file - the published copy and the draft's.
            this.TableEditor.FileSource = new DraftTableFileSource(
                DataManager.DataRoot,
                DraftFolderLayout.GetSystemFolder(DraftManager.DraftsRoot, entry.ExcelDataFile));

            if (!this.TableEditor.Load(DraftManager.DraftsRoot, entry.ExcelDataFile, published))
            {
                Logger.Warning($"Could not open the table for the draft of [{entry.ExcelDataFile}] - its workbook could not be read");
                return;
            }

            this.thisTableEntry = entry;
            this.RefreshDrafts();
            this.FocusTableIfOpen();
        }

        // ###########################################################################################
        // Puts keyboard focus in the open table, so typing lands in a cell rather than in the
        // always-on component filter on the left (see Main.ShouldReturnFocusToComponentSearch, which
        // keeps it there once it is). Called when a table opens and when the tab is shown again.
        //
        // Posted, not immediate: a table just opened, or a tab just switched to, has not been laid
        // out yet, and a control that is not yet visible cannot take focus.
        // ###########################################################################################
        internal void FocusTableIfOpen()
        {
            if (this.IsTableOpen)
            {
                Dispatcher.UIThread.Post(this.TableEditor.FocusGrid, DispatcherPriority.Background);
            }
        }

        // ###########################################################################################
        // Closes the table, asking first when it holds unsaved edits. False when the contributor
        // chose to stay (Cancel), or chose Save and the save was refused - the table stays open
        // with the reason on screen.
        // ###########################################################################################
        internal async Task<bool> CloseTableAsync()
        {
            if (!this.IsTableOpen)
            {
                return true;
            }

            if (!await this.ConfirmLeavingTableAsync())
            {
                return false;
            }

            this.CloseTableWithoutAsking();
            return true;
        }

        // ###########################################################################################
        // Closes the table, without asking, if it is open on this system's draft - for a draft that
        // is about to be deleted by something other than this tab (retirement after publishing).
        // The caller has already made sure nothing unsaved is lost: see HasUnsavedTableEditsFor.
        // ###########################################################################################
        internal void CloseTableIfOpenFor(string? excelDataFile)
        {
            if (this.thisTableEntry is not null &&
                string.Equals(this.thisTableEntry.ExcelDataFile, excelDataFile?.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                this.CloseTableWithoutAsking();
            }
        }

        private void CloseTableWithoutAsking()
        {
            this.TableEditor.Clear();
            this.thisTableEntry = null;
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // True when it is fine to leave the table: nothing unsaved, or the contributor saved (and
        // the save went through) or chose to discard. `owner` is for Main, which asks while this
        // tab may not be the one on screen.
        // ###########################################################################################
        internal async Task<bool> ConfirmLeavingTableAsync(Window? owner = null)
        {
            if (!this.HasUnsavedTableEdits)
            {
                return true;
            }

            return await this.AskAboutUnsavedTableEditsAsync(owner) switch
            {
                UnsavedTableEditsChoice.Save => this.TableEditor.Save() == DraftWorkbookEditOutcome.Saved,
                UnsavedTableEditsChoice.Discard => true,
                _ => false,
            };
        }

        // ###########################################################################################
        // Before an action on the draft FILE (importing files, submitting): the table's unsaved
        // edits are saved, or thrown away by reloading - the table itself stays open either way.
        // True when the action may go ahead.
        // ###########################################################################################
        private async Task<bool> SettleTableEditsAsync(HardwareBoardEntry entry)
        {
            if (!this.IsTableOpenFor(entry) || !this.TableEditor.HasUnsavedChanges)
            {
                return true;
            }

            switch (await this.AskAboutUnsavedTableEditsAsync(owner: null))
            {
                case UnsavedTableEditsChoice.Save:
                    return this.TableEditor.Save() == DraftWorkbookEditOutcome.Saved;

                case UnsavedTableEditsChoice.Discard:
                    this.TableEditor.Reload();
                    return true;

                default:
                    return false;
            }
        }

        // ###########################################################################################
        // Asks what to do with the table's unsaved edits. When the draft file has changed since the
        // table read it, a save would be REFUSED, so the question offers none and says why -
        // offering it made "Close table" a loop (reported, 2026-09-24): Save, refused, the table
        // stays open, and Close asks the same again.
        // ###########################################################################################
        private async Task<UnsavedTableEditsChoice> AskAboutUnsavedTableEditsAsync(Window? owner)
        {
            // Save is offered only when it can land: not over a draft changed since the table read
            // it (refused), and not into one Excel holds open (it would fail).
            UnsavedTableEditsPrompt prompt =
                this.TableEditor.HasDraftChangedOnDisk() ? UnsavedTableEditsPrompt.DraftChangedOnDisk
                : this.TableEditor.IsDraftHeldOpenElsewhere() ? UnsavedTableEditsPrompt.DraftOpenElsewhere
                : UnsavedTableEditsPrompt.Leaving;

            if (this.UnsavedTableEditsAnswerForTests is not null)
            {
                return this.UnsavedTableEditsAnswerForTests(prompt);
            }

            Window? ownerWindow = owner ?? TopLevel.GetTopLevel(this) as Window;
            if (ownerWindow is null)
            {
                return UnsavedTableEditsChoice.Cancel;
            }

            var window = new UnsavedTableEditsWindow();
            window.Initialize(prompt);

            return await window.ShowDialog<UnsavedTableEditsChoice?>(ownerWindow) ?? UnsavedTableEditsChoice.Cancel;
        }

        // ###########################################################################################
        // Switches the tab between its two layouts - see the markup's comment on DraftsBodyGrid for
        // why the table is a sibling of the list's ScrollViewer and never inside it.
        //
        // Runs at the end of every RefreshDrafts. If the open draft is no longer listed (discarded
        // from elsewhere, or published and retired), the table closes: there is nothing left for
        // it to save into.
        // ###########################################################################################
        private void ApplyTableMode()
        {
            DraftListItem? openItem = this.thisTableEntry is null
                ? null
                : this.Drafts.FirstOrDefault(item =>
                    string.Equals(item.ExcelDataFile, this.thisTableEntry.ExcelDataFile, StringComparison.OrdinalIgnoreCase));

            if (this.thisTableEntry is not null && openItem is null)
            {
                this.TableEditor.Clear();
                this.thisTableEntry = null;
            }

            bool tableMode = openItem is not null;

            // The list gives way to the open row's own host (see the markup's comment on
            // DraftsBodyGrid). The list's own visibility is set by RefreshDrafts, which also knows
            // whether there are any drafts.
            this.OpenDraftRow.Content = openItem;
            this.OpenDraftRow.IsVisible = tableMode;
            this.TableEditor.IsVisible = tableMode;

            // *** THE ROWS ARE CHANGED IN PLACE, NEVER REPLACED. *** Assigning a new RowDefinitions
            // collection left the Grid measuring its cells against the OLD row kinds (it caches
            // which cells sit in star rows), so the now-Auto row was measured with no height at all
            // and the open draft's row drew nothing. Setting each definition's Height is a change
            // the Grid does track.
            this.DraftsBodyGrid.RowDefinitions[0].Height = tableMode ? GridLength.Auto : GridLength.Star;
            this.DraftsBodyGrid.RowDefinitions[1].Height = tableMode ? GridLength.Star : GridLength.Auto;
        }

        // A save in the table changes the draft: its row count here, and - when it is the board on
        // screen - the component list and schematics too. The same after the table reloaded because
        // something outside it (Excel) changed the draft: those were read before the change.
        private void OnTableSaved(object? sender, EventArgs e)
        {
            if (this.thisMainWindow is not null)
            {
                this.thisMainWindow.ApplyDraftsTabVisibility();
                this.thisMainWindow.ReloadCurrentBoardAfterDraftFileImport(this.TableEditor.ExcelDataFile);
            }
            else
            {
                this.RefreshDrafts();
            }
        }

        // The published board to colour against, or null when there is none (a draft-only system,
        // or a published file that is missing - the table then simply shows no colours rather
        // than calling every row an addition).
        private static async Task<BoardData?> LoadPublishedBoardAsync(HardwareBoardEntry entry, DraftStatus? status)
        {
            if (entry.IsDraftOnly || status?.IsNewSystem == true)
            {
                return null;
            }

            string publishedPath = DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile);
            if (publishedPath.Length == 0 || !File.Exists(publishedPath))
            {
                return null;
            }

            return await BoardDataReader.LoadAsync(publishedPath, publishedPath);
        }
    }
}
