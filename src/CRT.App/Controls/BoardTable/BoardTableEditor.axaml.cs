using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Data.Converters;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "EDIT IN TABLE FORMAT" - a draft's workbook as editable sheets, inline on the Drafts tab
    // (owner request, 2026-09-24).
    //
    // *** THIS CONTROL ONLY PAINTS. *** What a cell is, whether it differs from the published data,
    // where a deleted row goes and what a save writes are all decided in CRT.Data -
    // BoardTableDocument, BoardTableSheet and DraftTableSession - because the project owner wants the
    // same table in the Maintainer tab later. Keep it that way: logic added here is logic the
    // maintainer's copy would have to duplicate, and logic no test without a display can reach.
    //
    // It deliberately knows nothing of Main or DataManager either: it is handed a drafts folder, a
    // system and the published board to compare with, and raises Saved. TabDrafts does the wiring.
    //
    // THE GRID is ProDataGrid, an MIT-licensed fork of Avalonia's own DataGrid (which is deprecated
    // from Avalonia 12, while the maintained alternative, TreeDataGrid, needs a paid licence that
    // an open-source GPL project cannot use). It keeps the original's API, so going back to the
    // original package is a package swap rather than a rewrite.
    //
    // *** THE COLOURS COME FROM A PER-COLUMN CELL THEME, not from code poking cells. *** Each data
    // column gets its own ControlTheme whose Background is bound to that column's cell state. The
    // grid recycles cell containers as it scrolls, and a binding follows the recycled row by
    // itself - anything that set a cell's colour directly would paint the wrong row the moment a
    // container was reused.
    //
    // Editing is one cell at a time, as the project owner asked. Copy and paste work inside a cell
    // being edited (it is an ordinary text box) and on the current cell via Ctrl+C / Ctrl+V
    // (BoardTableClipboard). Pasting blocks of cells is a later feature. Several ROWS can be selected
    // and deleted at once (2026-10-02). Ctrl+Z and Ctrl+Y undo and redo through the model's own
    // history (BoardTableHistory).
    //
    // FILE MAP: this file (loading, sheets, toolbar, keys, columns),
    // BoardTableEditor.Filter.cs (the colour key's pills as the filter),
    // BoardTableEditor.CellColours.cs (each cell's background and tooltip - its state's wash and
    // its problem's corner mark),
    // BoardTableEditor.Selection.cs (several rows selected, "Delete row" on all of them),
    // BoardTableEditor.Search.cs (the search box, and its marks through BoardTableSearchAdapter),
    // BoardTableEditor.TextWrap.cs (cells that always wrap, columns dragged wider than they size
    // themselves, and a double-click on a heading's edge fitting a column to its text),
    // BoardTableEditor.RowDrag.cs (dragging a row by its grip, with the worklog-style placeholder),
    // BoardTableEditor.FileWatch.cs (noticing the draft being changed, or open, in Excel) and
    // BoardTableEditor.FilePreview.cs (the hover card on a file cell - its content is
    // BoardTableFilePreview, its bytes the host's IBoardTableFileSource).
    // ###########################################################################################
    public partial class BoardTableEditor : UserControl
    {
        // The table on screen, and - in the Drafts tab only - the draft FILE behind it. In document
        // mode (the Maintainer tab, Open(BoardTableDocument)) there is no session: Save raises
        // SaveRequested for the host to save, and Reload and the file notices do not apply.
        private BoardTableDocument? thisDocument;
        private DraftTableSession? thisSession;
        private BoardTableSheet? thisCurrentSheet;
        private BoardData? thisPublished;
        private string thisDraftsRoot = string.Empty;
        private string thisExcelDataFile = string.Empty;

        // The downloaded data, so the checks can look for the files a row names (DraftTableSession).
        private string thisDataRoot = string.Empty;

        // The table's text size - smaller than the grid's default so a sheet shows more rows at
        // once (owner request, 2026-09-24). The grid, its headers and every column use it.
        internal const double CellFontSize = 12;

        // Coalesces a burst of cell edits (a paste, a fast typist) into one refresh.
        private bool thisRefreshPosted;

        // What the grid shows: the current sheet's rows through a view, so the colour key's picked
        // kinds can filter them without touching the sheet's own list.
        private DataGridCollectionView? thisView;

        // True while code - not the user - is moving the sheet tabs' selection, so the selection
        // handler does not treat it as a click and steal keyboard focus into the grid.
        private bool thisSyncingSheetTabs;

        public BoardTableEditor()
        {
            this.InitializeComponent();

            this.thisCellToBrush = new FuncMultiValueConverter<object?, IBrush?>(values =>
            {
                object?[] both = values.ToArray();

                BoardTableCellState state = both.Length > 0 && both[0] is BoardTableCellState s ? s : BoardTableCellState.Unchanged;
                BoardProblemLevel problem = both.Length > 1 && both[1] is BoardProblemLevel p ? p : BoardProblemLevel.None;

                return this.CachedCellBrush(state, problem);
            });

            // *** EXCEL'S EDITING GESTURES, NOT THE GRID'S DEFAULT. *** The grid's default starts
            // editing on a SINGLE click, which would make "select a cell, press Ctrl+V" impossible
            // - the click would already have opened the cell's editor. So, as in Excel: one click
            // selects, and double-click, F2 or simply typing starts editing.
            this.TableGrid.EditTriggers =
                DataGridEditTriggers.CellDoubleClick | DataGridEditTriggers.TextInput | DataGridEditTriggers.F2;

            // *** ROWS ARE DRAGGED BY THEIR GRIP, and the TABLE moves them, not the grid. *** The
            // row header (the narrow column at the far left) is the handle, so a drag can never
            // start from a click meant to select a cell. See BoardTableEditor.RowDrag.cs - the
            // grid's own row drag is left OFF, since it cannot show the travelling placeholder.
            this.WireRowDragging();

            // Resting on a file cell shows the file - see BoardTableEditor.FilePreview.cs.
            this.WireFilePreview();

            this.TableGrid.LoadingRow += this.OnLoadingRow;
            this.TableGrid.BeginningEdit += BoardTableEditor.OnBeginningEdit;
            this.WireDraftFileWatching();
            this.TableGrid.CurrentCellChanged += (_, _) => this.UpdateToolbar();

            // Several rows selected and deleted at once (BoardTableEditor.Selection.cs), and the
            // search box (BoardTableEditor.Search.cs).
            this.WireSelection();
            this.WireSearch();

            // Columns dragged wider than they size themselves, and fitted to their text by a
            // double-click on a heading's edge (BoardTableEditor.TextWrap.cs).
            this.WireColumnWidths();

            // Tunnel, so a selected-but-not-editing cell sees Ctrl+C/Ctrl+V before the grid's own
            // row-copy handling does. While a cell IS being edited, the text box inside it is the
            // source and handles copy and paste itself - see OnGridKeyDown.
            this.TableGrid.AddHandler(KeyDownEvent, this.OnGridKeyDown, RoutingStrategies.Tunnel);

            // Undo and redo on the whole control rather than the grid, so they still work right
            // after a toolbar button took the focus ("Delete row", then Ctrl+Z).
            this.AddHandler(KeyDownEvent, this.OnEditorKeyDown, RoutingStrategies.Tunnel);

            this.ActualThemeVariantChanged += (_, _) => this.RebuildColumns();

            this.SheetTabs.SelectionChanged += this.OnSheetTabSelectionChanged;

            this.UpdateToolbar();
        }

        // Raised after a save has written the draft - TabDrafts refreshes its row count and the
        // board on screen from it.
        public event EventHandler? Saved;

        public bool HasUnsavedChanges => this.thisDocument?.HasUnsavedChanges == true;

        public bool HasTable => this.thisDocument is not null;

        // ###########################################################################################
        // DOCUMENT MODE - the Maintainer tab's table (2026-09-25). Shows `document` with no
        // draft file behind it: "Save changes" raises SaveRequested, and the host saves it however
        // it saves (the Maintainer tab sends it to the server as an amendment), then calls
        // Open again with the saved state, or ShowMessage to say why not. Reload and the notices
        // about a draft file being changed or open in Excel do not apply and stay hidden.
        // ###########################################################################################
        public void Open(BoardTableDocument document, string? message = null, string? preferredSheet = null)
        {
            ArgumentNullException.ThrowIfNull(document);

            // The sheet already on screen when the host re-opens after a save; otherwise the one the
            // host asks for - the Maintainer tab's last-visited sheet, so moving through the
            // queue stays on it (owner request, 2026-09-26) - and failing that the FIRST SHEET WITH A
            // CHANGE: a maintainer opening a submission wants what changed, not "Board schematics"
            // every time. A sheet whose tab "Show changes only" hides gives way - see Attach.
            string? keepSheet = this.thisCurrentSheet?.Name
                ?? preferredSheet
                ?? document.Sheets.FirstOrDefault(sheet => sheet.ChangeCount > 0)?.Name;

            this.thisSession = null;
            this.thisDraftsRoot = string.Empty;
            this.thisExcelDataFile = string.Empty;
            this.thisPublished = null;
            this.thisDataRoot = string.Empty;

            this.Attach(document, session: null, keepSheet);
            this.ShowStatus(message ?? string.Empty);
        }

        // Raised by "Save changes" in document mode - see Open.
        public event EventHandler? SaveRequested;

        // ###########################################################################################
        // READ-ONLY (2026-10-03): the Maintainer tab's Systems screen shows the board of a system the
        // account may not change. The table is there to look at - sheets, search, the colour key,
        // copying a cell, the file cards - but nothing in it can be typed, pasted, moved, inserted
        // or deleted, and there is no "Save changes": an edit that could never be sent is worse than
        // none. The host says why (ShowMessage). Off by default; the Drafts tab never sets it.
        // ###########################################################################################
        public bool IsReadOnly
        {
            get => this.thisIsReadOnly;
            set
            {
                this.thisIsReadOnly = value;
                this.TableGrid.CommitEdit();
                this.TableGrid.IsReadOnly = value;
                this.UpdateRowsDraggable();
                this.UpdateToolbar();
            }
        }

        private bool thisIsReadOnly;

        // Whether a draft FILE is behind the table (the Drafts tab) or only a document (review).
        public bool IsFileBacked => this.thisSession is not null;

        // The document on screen, with any cell still in its editor committed into it first - what
        // the host reads when it saves in document mode.
        public BoardTableDocument? CommitAndGetDocument()
        {
            this.TableGrid.CommitEdit();
            return this.thisDocument;
        }

        // A message under the toolbar, for the host in document mode ("Saved.", "Not saved: ...").
        public void ShowMessage(string message) => this.ShowStatus(message ?? string.Empty);

        // Whether the draft file changed after this table read it - in which case a save would be
        // refused, so nothing should offer one. Reads the file, so it is a method, not a property.
        internal bool HasDraftChangedOnDisk() => this.thisSession?.HasChangedOnDisk() == true;

        public string ExcelDataFile => this.thisExcelDataFile;

        internal DraftTableSession? SessionForTests => this.thisSession;

        internal BoardTableSheet? CurrentSheet => this.thisCurrentSheet;

        // ###########################################################################################
        // Opens a draft as a table. `published` is what differences are coloured against - null
        // for a system with nothing published. Returns false when the draft cannot be read.
        // ###########################################################################################
        //
        // `dataRoot` is the downloaded data, where the checks look for the files a row names after
        // the draft's own folder (owner request, 2026-10-02) - the host's to hand over, since this
        // control knows nothing of DataManager. Empty: files are not looked for.
        public bool Load(string draftsRoot, string excelDataFile, BoardData? published, string dataRoot = "")
        {
            DraftTableSession? session = DraftTableSession.Open(draftsRoot, excelDataFile, published, dataRoot);
            if (session is null)
            {
                return false;
            }

            // Another draft starts with an empty search (agreed, 2026-10-02); the same one - read
            // again - keeps it.
            if (!string.Equals(excelDataFile, this.thisExcelDataFile, StringComparison.OrdinalIgnoreCase))
            {
                this.thisSearchText = string.Empty;
                this.SearchBox.Text = string.Empty;
            }

            this.thisDraftsRoot = draftsRoot;
            this.thisExcelDataFile = excelDataFile;
            this.thisPublished = published;
            this.thisDataRoot = dataRoot ?? string.Empty;

            this.Attach(session.Document, session, keepSheet: null);
            this.ShowStatus(session.Document.HasBaseline
                ? string.Empty
                : "Nothing of this system is published yet, so no row is marked as added, changed or deleted - every row is your own.");

            return true;
        }

        // Closes the table. Unsaved edits are dropped - the caller asks first (TabDrafts).
        public void Clear()
        {
            this.Detach();

            this.thisDocument = null;
            this.thisSession = null;
            this.thisCurrentSheet = null;
            this.thisPublished = null;
            this.thisExcelDataFile = string.Empty;

            this.TableGrid.ItemsSource = Array.Empty<BoardTableRow>();
            this.TableGrid.Columns.Clear();
            this.ClearSheetTabs();
            this.ShowStatus(string.Empty);

            // A closed table's search goes with it (agreed, 2026-10-02).
            this.SearchText = string.Empty;
            this.OpenElsewhereBar.IsVisible = false;
            this.SetChangedOnDisk(false);
            this.UpdateFileWatching();
        }

        // ###########################################################################################
        // Writes the table into the draft and re-reads it, so what is on screen afterwards is
        // exactly what the file now holds (a blank row is gone, for instance). Refused - and said
        // why - when the file changed since it was read; see DraftTableSession for why the table may
        // never write over such a change.
        //
        // This one is synchronous, all on the calling thread - what the tests drive. "Save changes"
        // and the unsaved-edits prompts use SaveAsync.
        // ###########################################################################################
        public DraftWorkbookEditOutcome Save()
        {
            if (this.thisSession is null)
            {
                return DraftWorkbookEditOutcome.NoDraft;
            }

            // A cell still in its editor has not reached the model yet.
            this.TableGrid.CommitEdit();

            // Asked BEFORE saving: a save reloads the table, and the reloaded rows are no longer
            // "inserted this sitting".
            bool placesNewComponents = this.thisSession.Document.Sheets.Any(sheet => sheet.HasRowsToPlaceOnSave);

            DraftWorkbookEditOutcome outcome = this.thisSession.Save();
            this.AfterSave(outcome, placesNewComponents, reread: null);

            return outcome;
        }

        // ###########################################################################################
        // *** THE SAME SAVE, UNDER "PLEASE WAIT" (owner question, 2026-09-30: "When in "Draft" and I
        // have changed something in the table and I click the button "Save changes", shouldn't it
        // then show that standard "Wait" thing?"). *** Measured: the largest shipped board (C128
        // 310378) takes about two seconds the first time and 0.6-1 s after - read, write, read
        // back - and ran ON the UI thread, so the window froze with nothing drawn. The overlay
        // could not have shown then even had it been asked to.
        //
        // So the table's rows are taken on the UI thread (DraftTableSession.PrepareSave), and the
        // workbook work - and the re-read the table reopens on - runs on the pool under the window's
        // overlay (BusyOverlay.RunLocalAsync: local work that cannot be stopped halfway). The file
        // watch leaves the file alone meanwhile (thisSaveInFlight): the save's own write would
        // read as a change from outside.
        // ###########################################################################################
        public async Task<DraftWorkbookEditOutcome> SaveAsync()
        {
            if (this.thisSession is not DraftTableSession session)
            {
                return DraftWorkbookEditOutcome.NoDraft;
            }

            this.TableGrid.CommitEdit();

            bool placesNewComponents = session.Document.Sheets.Any(sheet => sheet.HasRowsToPlaceOnSave);

            Func<DraftWorkbookEditOutcome> write = session.PrepareSave();
            string draftsRoot = this.thisDraftsRoot;
            string excelDataFile = this.thisExcelDataFile;
            BoardData? published = this.thisPublished;
            string dataRoot = this.thisDataRoot;

            (DraftWorkbookEditOutcome Outcome, DraftTableSession? Reread) result;
            this.thisSaveInFlight = true;

            // ###########################################################################################
            // *** PAST THE LIMIT THE TABLE IS LOCKED, NOT LEFT LIVE (code review, 2026-10-01). ***
            // RunLocalAsync lifts the overlay when the write is still going after two minutes and
            // carries on waiting - which unblocks the keyboard and mouse over a grid that still holds
            // the pre-save document. A contributor who then typed had those edits thrown away
            // without a word when the save finished and the table reopened on the re-read file.
            // `stillRunning` is told at exactly that moment, so the editor is disabled until the
            // write ends: nothing can be typed that the save would then discard.
            // ###########################################################################################
            bool wasEnabled = this.IsEnabled;
            bool lockedForSave = false;

            try
            {
                result = await BusyOverlay.RunLocalAsync(this, BoardTableEditor.SavingWait, () => Task.Run(() =>
                {
                    DraftWorkbookEditOutcome outcome = write();

                    return (outcome, outcome == DraftWorkbookEditOutcome.Saved
                        ? DraftTableSession.Open(draftsRoot, excelDataFile, published, dataRoot)
                        : null);
                }), stillRunning: () =>
                {
                    lockedForSave = true;
                    this.IsEnabled = false;
                    this.ShowStatus(BoardTableEditor.StillSaving);
                    this.LockedForSaveForTests?.Invoke();
                });
            }
            finally
            {
                this.thisSaveInFlight = false;

                if (lockedForSave)
                    this.IsEnabled = wasEnabled;
            }

            // Another table opened, or this one closed, while it saved: nothing of it to update.
            if (!ReferenceEquals(this.thisSession, session))
            {
                return result.Outcome;
            }

            // The re-read already hashed the saved file on the pool - no second hash here, on the UI
            // thread (DraftTableSession.CompleteSave).
            session.CompleteSave(result.Outcome, result.Reread?.Fingerprint);
            this.AfterSave(result.Outcome, placesNewComponents, result.Reread);

            return result.Outcome;
        }

        public const string SavingWait = "Saving the table into your draft...";

        // Said while the table is locked past the wait limit - see SaveAsync.
        public const string StillSaving =
            "Still saving - the table is locked until the write finishes, so nothing you type is lost.";

        // Told, for a test, at the moment the table is locked (the editor is disabled by then).
        internal Action? LockedForSaveForTests { get; set; }

        // True while SaveAsync is writing the draft.
        private bool thisSaveInFlight;

        // What a save leaves on screen - the table read again, and a sentence saying how it went.
        private void AfterSave(DraftWorkbookEditOutcome outcome, bool placesNewComponents, DraftTableSession? reread)
        {
            switch (outcome)
            {
                case DraftWorkbookEditOutcome.Saved:
                    this.ReloadKeepingPlace(reread);
                    this.ShowStatus(placesNewComponents
                        ? "Saved. New components were put into their category, in label order - move them if you want them elsewhere."
                        : "Saved.");
                    this.Saved?.Invoke(this, EventArgs.Empty);
                    break;

                case DraftWorkbookEditOutcome.ChangedOnDisk:
                    this.SetChangedOnDisk(true);
                    this.ShowStatus(
                        "Not saved: the draft was changed outside this table after it was opened - in Excel, " +
                        "or elsewhere in the application. Press \"Reload\" to see the current version. " +
                        "Edits made here since the last save will be lost.");
                    break;

                case DraftWorkbookEditOutcome.WriteFailed:
                    this.ShowStatus(
                        "Not saved: the draft file could not be written. If it is open in Excel, close it there " +
                        "and try again.");
                    break;

                default:
                    this.ShowStatus("Not saved: this draft no longer exists.");
                    break;
            }

            this.UpdateToolbar();
        }

        // ###########################################################################################
        // Re-reads the draft from its file, dropping any unsaved edits. The caller confirms first
        // when there are some.
        // ###########################################################################################
        public bool Reload()
        {
            if (this.thisExcelDataFile.Length == 0)
            {
                return false;
            }

            bool reloaded = this.ReloadKeepingPlace();
            this.ShowStatus(reloaded ? string.Empty : "The draft could not be read again.");

            return reloaded;
        }

        internal void SelectSheet(BoardTableSheet sheet)
        {
            if (this.thisDocument is null || !this.thisDocument.Sheets.Contains(sheet))
            {
                return;
            }

            this.TableGrid.CommitEdit();
            this.EndRowDrag();
            this.HideFilePreview();

            this.thisCurrentSheet = sheet;
            this.RebuildColumns();

            // The filter is (re)applied once the grid has the view - see ApplyViewFilter.
            this.thisView = new DataGridCollectionView(sheet.Rows);
            this.TableGrid.ItemsSource = this.thisView;
            this.ApplyViewFilter();

            this.UpdateSheetTabs();
            this.UpdateToolbar();
            this.UpdateSearchMarks();
        }

        // The cell the grid's cursor is on, or null (no cell, or the marker column).
        internal BoardTableCell? CurrentCell
        {
            get
            {
                DataGridCellInfo current = this.TableGrid.CurrentCell;

                if (!current.IsValid || current.Item is not BoardTableRow row || current.Column?.Tag is not int column)
                {
                    return null;
                }

                return column >= 0 && column < row.Cells.Count ? row.Cells[column] : null;
            }
        }

        internal BoardTableRow? CurrentRow =>
            this.TableGrid.CurrentCell is { IsValid: true, Item: BoardTableRow row } ? row : null;


        // Puts the grid's cursor on one cell and scrolls to it.
        internal void SelectCell(BoardTableRow row, int columnIndex)
        {
            if (this.thisCurrentSheet is null)
            {
                return;
            }

            DataGridColumn? column = this.TableGrid.Columns.FirstOrDefault(c => c.Tag is int index && index == columnIndex);

            // The index among the rows the grid SHOWS - with "Show changes only" on, that is not the
            // sheet's own index, and a row filtered out cannot be the current one.
            int rowIndex = this.thisView?.IndexOf(row) ?? this.thisCurrentSheet.Rows.IndexOf(row);

            if (column is null || rowIndex < 0)
            {
                return;
            }

            var cell = new DataGridCellInfo(row, column, rowIndex, column.DisplayIndex, true);
            this.TableGrid.CurrentCell = cell;

            // And that one cell is the selection: a selection of several rows left behind where the
            // cursor was (an undo, a delete, Tab) is not what "Delete row" should act on next
            // (BoardTableEditor.Selection.cs). Moving the cursor does not change the grid's
            // selection by itself.
            this.TableGrid.SelectedCells = [cell];
            this.TableGrid.ScrollIntoView(row, column);
            this.UpdateToolbar();
        }

        // ###########################################################################################
        // Toolbar actions. Each is a thin call into BoardTableSheet, which owns the rule.
        // ###########################################################################################
        // "Insert row": an empty row below the selected one, with the cursor on it ready for typing.
        // There was an "Insert row above" beside it until 2026-10-02 (owner request: "We can of
        // course remove one of the "Insert row" buttons, as that can be moved afterwards") - a row
        // is moved by its grip. BoardTableSheet.InsertRowAbove stays in the model.
        internal void InsertRowBelow()
        {
            if (this.thisCurrentSheet is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            BoardTableRow inserted = this.thisCurrentSheet.InsertRow(this.CurrentRow);
            this.SelectCell(inserted, 0);
        }

        // "Delete row" - every selected row: BoardTableEditor.Selection.cs.

        // ###########################################################################################
        // Tab / Shift+Tab: one cell right or left, wrapping to the next or previous row at the end
        // of a row, the way Excel does. Moves among the rows the grid SHOWS, so with "Show changes
        // only" on it goes from change to change. Stays put at the very first or last cell.
        // ###########################################################################################
        internal void MoveToNeighbourCell(bool forward)
        {
            if (this.thisCurrentSheet is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            if (this.CurrentRow is not { } row)
            {
                this.FocusGrid();
                return;
            }

            List<BoardTableRow> shown = this.ShownRows();
            int rowIndex = shown.IndexOf(row);
            int column = this.CurrentCell?.ColumnIndex ?? -1;
            int lastColumn = this.thisCurrentSheet.Columns.Count - 1;

            if (forward)
            {
                if (column < lastColumn)
                {
                    column++;
                }
                else if (rowIndex >= 0 && rowIndex + 1 < shown.Count)
                {
                    rowIndex++;
                    column = 0;
                }
            }
            else
            {
                if (column > 0)
                {
                    column--;
                }
                else if (rowIndex > 0)
                {
                    rowIndex--;
                    column = lastColumn;
                }
                else
                {
                    column = Math.Max(column, 0);
                }
            }

            if (rowIndex >= 0)
            {
                this.SelectCell(shown[rowIndex], column);
            }

            this.TableGrid.Focus();
        }

        private List<BoardTableRow> ShownRows() =>
            this.thisView?.Cast<BoardTableRow>().ToList() ?? this.thisCurrentSheet?.Rows.ToList() ?? [];

        internal void MoveRowUp() => this.MoveCurrentRow(up: true);

        internal void MoveRowDown() => this.MoveCurrentRow(up: false);

        // Moves the row under the cursor and keeps the cursor on it, in the same column.
        private void MoveCurrentRow(bool up)
        {
            if (this.IsNarrowed || this.thisIsReadOnly || this.thisCurrentSheet is null || this.CurrentRow is not { IsDeleted: false } row)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            int column = this.CurrentCell?.ColumnIndex ?? 0;
            bool moved = up ? this.thisCurrentSheet.MoveRowUp(row) : this.thisCurrentSheet.MoveRowDown(row);

            if (moved)
            {
                this.SelectCell(row, column);
            }

            this.UpdateToolbar();
        }

        // ###########################################################################################
        // Ctrl+V on a selected cell: the clipboard's text, if it is ONE cell, replaces the cell's.
        // A block of cells is refused with a word of explanation rather than squeezed into one
        // cell - see BoardTableClipboard.
        // ###########################################################################################
        internal bool PasteIntoCurrentCell(string? clipboardText)
        {
            if (this.thisIsReadOnly || this.CurrentCell is not { } cell || cell.Row.IsDeleted)
            {
                return false;
            }

            if (!BoardTableClipboard.TryReadSingleCell(clipboardText, out string text))
            {
                if (!string.IsNullOrEmpty(clipboardText))
                {
                    this.ShowStatus("The clipboard holds several cells. For now the table pastes one cell at a time - copy a single cell and try again.");
                }

                return false;
            }

            cell.Text = text;
            this.ShowStatus(string.Empty);

            return true;
        }

        internal string? CopyTextOfCurrentCell() =>
            this.CurrentCell is { } cell ? BoardTableClipboard.ForSingleCell(cell.Text) : null;

        private async void OnGridKeyDown(object? sender, KeyEventArgs e)
        {
            // *** TAB MOVES ACROSS THE ROW, AS IN EXCEL (owner request, 2026-09-24). *** Left
            // alone it is the window's focus navigation, which took the cursor out of the table to
            // the hardware drop-down. Handled here on the TUNNEL route, before that happens, and
            // whether or not a cell is being edited - the edit is committed first.
            if (e.Key == Key.Tab && e.KeyModifiers is KeyModifiers.None or KeyModifiers.Shift)
            {
                e.Handled = true;
                this.MoveToNeighbourCell(forward: e.KeyModifiers == KeyModifiers.None);
                return;
            }

            // A cell being edited: its own text box copies and pastes.
            if (e.Source is TextBox)
            {
                return;
            }

            // Alt+Up / Alt+Down move the selected row, the keyboard twin of dragging it.
            if (e.KeyModifiers == KeyModifiers.Alt && e.Key is Key.Up or Key.Down)
            {
                e.Handled = true;
                this.MoveCurrentRow(up: e.Key == Key.Up);
                return;
            }

            bool isCopy = BoardTableEditor.Matches(e, hotkeys => hotkeys.Copy);
            bool isPaste = BoardTableEditor.Matches(e, hotkeys => hotkeys.Paste);

            if (!isCopy && !isPaste)
            {
                return;
            }

            IClipboard? clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard is null || this.CurrentCell is null)
            {
                return;
            }

            e.Handled = true;

            try
            {
                if (isCopy)
                {
                    await clipboard.SetTextAsync(this.CopyTextOfCurrentCell() ?? string.Empty);
                }
                else
                {
                    this.PasteIntoCurrentCell(await clipboard.TryGetTextAsync());
                }
            }
            catch (Exception ex)
            {
                // A clipboard another program holds open throws on Windows. Worth a word to the
                // contributor, never worth an unhandled exception from an async void handler.
                CrtLog.Warning($"Table editor clipboard {(isCopy ? "copy" : "paste")} failed - [{ex.Message}]");
                this.ShowStatus("The clipboard could not be used just now. Try again.");
            }
        }

        // The platform's own gestures (Cmd on macOS, Ctrl elsewhere). Nothing matches when the
        // platform does not say - a headless or unusual platform then simply lacks the shortcut.
        private static bool Matches(KeyEventArgs e, Func<PlatformHotkeyConfiguration, List<KeyGesture>> gestures)
        {
            PlatformHotkeyConfiguration? hotkeys = Application.Current?.PlatformSettings?.HotkeyConfiguration;
            if (hotkeys is not null)
            {
                return gestures(hotkeys).Any(gesture => gesture.Matches(e));
            }

            return false;
        }

        // ###########################################################################################
        // Ctrl+Z / Ctrl+Y - and Ctrl+Shift+Z, and Cmd on macOS: the platform's own undo and redo
        // gestures (owner request, 2026-09-24). Tunnel, so the grid never sees them first.
        //
        // NOT while a cell is being edited: there the cell's own text box undoes the typing inside
        // it, as Excel does, and the table's history takes over once the cell is committed.
        // ###########################################################################################
        private void OnEditorKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Source is TextBox)
            {
                return;
            }

            if (BoardTableEditor.Matches(e, hotkeys => hotkeys.Undo))
            {
                e.Handled = true;
                this.Undo();
            }
            else if (BoardTableEditor.Matches(e, hotkeys => hotkeys.Redo))
            {
                e.Handled = true;
                this.Redo();
            }
        }

        internal void Undo() => this.StepHistory(history => history.Undo());

        internal void Redo() => this.StepHistory(history => history.Redo());

        // Takes one step and puts the cursor where it happened - switching sheet first when the
        // change was on another one - so the contributor sees what was just undone.
        private void StepHistory(Func<BoardTableHistory, BoardTableHistoryResult?> step)
        {
            if (this.thisDocument is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();
            int column = this.CurrentCell?.ColumnIndex ?? 0;

            BoardTableHistoryResult? result = step(this.thisDocument.History);
            if (result is null)
            {
                return;
            }

            if (!ReferenceEquals(result.Sheet, this.thisCurrentSheet))
            {
                this.SelectSheet(result.Sheet);
            }

            if (result.Row is not null)
            {
                int target = result.Column >= 0 ? result.Column : column;
                this.SelectCell(result.Row, Math.Clamp(target, 0, result.Sheet.Columns.Count - 1));
            }

            this.UpdateToolbar();
            this.FocusGrid();
        }

        // ###########################################################################################
        // The row and cell buttons hand the keyboard straight back to the table, so the next key -
        // typing, Tab, Ctrl+Z - goes where the contributor is working. Left on the button it was
        // worse than inconvenient: a button that disables itself by acting ("Delete row" leaves
        // the cursor on the red ghost, which cannot be deleted) drops the focus to NOTHING, and
        // Ctrl+Z straight after then reached no handler at all (caught by
        // BoardTableEditorTests.Ctrl_Z_works_right_after_a_toolbar_button_was_used).
        // ###########################################################################################
        private void OnInsertRowBelowClick(object? sender, RoutedEventArgs e) => this.ThenFocusGrid(this.InsertRowBelow);

        private void OnDeleteRowClick(object? sender, RoutedEventArgs e) => this.ThenFocusGrid(this.DeleteRow);

        private void ThenFocusGrid(Action action)
        {
            action();
            this.FocusGrid();
        }

        private async void OnSaveClick(object? sender, RoutedEventArgs e)
        {
            // Document mode: the host saves - see Open.
            if (this.thisSession is null && this.thisDocument is not null)
            {
                this.TableGrid.CommitEdit();
                this.SaveRequested?.Invoke(this, EventArgs.Empty);
                return;
            }

            await this.SaveAsync();
        }

        private async void OnReloadClick(object? sender, RoutedEventArgs e)
        {
            if (this.HasUnsavedChanges && !await this.ConfirmDiscardAsync())
            {
                return;
            }

            this.Reload();
        }

        // Reload throws away what is on screen, so it asks first - with no "Save" choice, since the
        // usual reason to reload is that saving has just been refused.
        private async Task<bool> ConfirmDiscardAsync()
        {
            if (TopLevel.GetTopLevel(this) is not Window owner)
            {
                return false;
            }

            var window = new UnsavedTableEditsWindow();
            window.Initialize(UnsavedTableEditsPrompt.Reloading);

            return await window.ShowDialog<UnsavedTableEditsChoice?>(owner) == UnsavedTableEditsChoice.Discard;
        }

        // A deleted row is a picture of what was published, not something to type into.
        private static void OnBeginningEdit(object? sender, DataGridBeginningEditEventArgs e)
        {
            if (e.Row?.DataContext is BoardTableRow { IsDeleted: true })
            {
                e.Cancel = true;
            }
        }

        // Rows are recycled as the grid scrolls, so the deleted class is set every time a row is
        // (re)loaded rather than once. A ghost never turns live in place - restoring replaces it -
        // so this is the only moment the class can need to change.
        private void OnLoadingRow(object? sender, DataGridRowEventArgs e)
        {
            e.Row.Classes.Set("BoardTableDeleted", e.Row.DataContext is BoardTableRow { IsDeleted: true });
            e.Row.Classes.Set(BoardTableEditor.DropPlaceholderClass, this.IsDropPlaceholder(e.Row.DataContext));
        }

        private void Attach(BoardTableDocument document, DraftTableSession? session, string? keepSheet)
        {
            this.Detach();

            this.thisDocument = document;
            this.thisSession = session;
            document.Changed += this.OnDocumentChanged;

            foreach (BoardTableSheet sheet in document.Sheets)
            {
                sheet.CellEdited += this.OnCellEdited;
            }

            // *** A PICK THAT SHOWS NOTHING HERE IS NOT APPLIED. *** A filter left on from another
            // table (this editor is reused) that matches no row of this one hid EVERY row with no
            // visible reason: reported, with "Show changes only", as a "Board schematics" sheet
            // showing empty although the draft had three schematic images - a table with nothing
            // published, where nothing is added, changed or deleted. The PICK is kept, and applies
            // again to the next table with such rows.
            BoardTableRowKinds filter = document.HasRowsShownBy(this.thisFilterWanted)
                ? this.thisFilterWanted
                : BoardTableRowKinds.None;

            // The search box's text, worked out on this document's rows (BoardTableEditor.Search.cs).
            this.ReapplySearchTo(document);

            this.ApplyFilter(filter);
            this.BuildSheetTabs();

            // The sheet asked for, unless the filter or the search hides its tab
            // (BoardTableDocument.SheetToShow).
            this.SelectSheet(document.SheetToShow(keepSheet, filter, this.thisSearch));

            // A fresh read of the file: whatever the warning bar said is no longer true.
            this.SetChangedOnDisk(false);

            if (session is not null)
            {
                this.ShowWhetherOpenElsewhere(session);
            }
            else
            {
                this.OpenElsewhereBar.IsVisible = false;
            }

            this.UpdateFileWatching();

            // With nothing published there is nothing to be added, changed or deleted against, so
            // only the Errors and Warnings pills stay - a duplicate is a duplicate either way.
            bool hasBaseline = document.HasBaseline;
            this.AddedPill.IsVisible = hasBaseline;
            this.ModifiedPill.IsVisible = hasBaseline;
            this.DeletedPill.IsVisible = hasBaseline;
        }

        private void Detach()
        {
            this.EndRowDrag();
            this.HideFilePreview();

            if (this.thisDocument is null)
            {
                return;
            }

            this.thisDocument.Changed -= this.OnDocumentChanged;

            foreach (BoardTableSheet sheet in this.thisDocument.Sheets)
            {
                sheet.CellEdited -= this.OnCellEdited;
            }
        }

        // Re-opens the draft on the same sheet, with the cursor back on (roughly) the same row.
        // `reread` is the draft already read again off the UI thread (SaveAsync); without one it is
        // read here.
        private bool ReloadKeepingPlace(DraftTableSession? reread = null)
        {
            string? sheetName = this.thisCurrentSheet?.Name;
            int rowIndex = this.thisCurrentSheet is null || this.CurrentRow is null
                ? -1
                : this.thisCurrentSheet.Rows.IndexOf(this.CurrentRow);
            int column = this.CurrentCell?.ColumnIndex ?? 0;

            // The row itself, by content: a save can MOVE a new component into its category, and
            // the cursor should follow it there - so the contributor sees where it went rather than
            // finding a different row under the cursor.
            List<string>? currentValues = this.CurrentRow?.Cells.Select(cell => cell.Text).ToList();

            DraftTableSession? session = reread ?? DraftTableSession.Open(this.thisDraftsRoot, this.thisExcelDataFile, this.thisPublished, this.thisDataRoot);
            if (session is null)
            {
                return false;
            }

            this.Attach(session.Document, session, sheetName);

            if (rowIndex >= 0 && this.thisCurrentSheet is { Rows.Count: > 0 } sheet)
            {
                BoardTableRow? same = currentValues is null
                    ? null
                    : sheet.Rows.FirstOrDefault(candidate =>
                        !candidate.IsDeleted && candidate.Cells.Select(cell => cell.Text).SequenceEqual(currentValues));

                this.SelectCell(same ?? sheet.Rows[Math.Min(rowIndex, sheet.Rows.Count - 1)], column);
            }

            return true;
        }

        private void OnCellEdited(object? sender, BoardTableCell cell)
        {
            if (this.thisRefreshPosted)
            {
                return;
            }

            this.thisRefreshPosted = true;

            // After the grid has finished committing the edit - see BoardTableCell.Text.
            Dispatcher.UIThread.Post(() =>
            {
                this.thisRefreshPosted = false;

                if (this.thisDocument is null)
                {
                    return;
                }

                foreach (BoardTableSheet sheet in this.thisDocument.Sheets.Where(s => s.NeedsRefresh))
                {
                    sheet.Refresh();
                }

                this.UpdateToolbar();
            }, DispatcherPriority.Background);
        }

        // Test seam: runs the refresh a cell edit schedules, without waiting for the dispatcher.
        internal void RefreshPendingForTests()
        {
            foreach (BoardTableSheet sheet in this.thisDocument?.Sheets.Where(s => s.NeedsRefresh) ?? [])
            {
                sheet.Refresh();
            }

            this.UpdateToolbar();
        }

        private void OnDocumentChanged(object? sender, EventArgs e)
        {
            this.UpdateSheetTabs();
            this.UpdateToolbar();

            // A row whose last change was just reverted - or whose last error was just fixed - drops
            // out of the filter, and a newly changed one joins it: the view only re-reads states
            // when asked.
            if (this.thisFilter != BoardTableRowKinds.None && this.thisView is not null && !this.thisView.IsEditingItem)
            {
                this.thisView.Refresh();
            }

            // An edit moves the runs the search found, though not which rows it shows.
            if (this.thisSearch.IsActive)
            {
                this.UpdateSearchMarks();
            }
        }

        // ###########################################################################################
        // One tab per sheet, labelled with its change count when it has any: "Components (3)". The
        // tabs carry no content of their own - see the markup's comment on SheetTabs.
        // ###########################################################################################
        private void BuildSheetTabs()
        {
            this.ClearSheetTabs();

            if (this.thisDocument is null)
            {
                return;
            }

            this.thisSyncingSheetTabs = true;

            try
            {
                foreach (BoardTableSheet sheet in this.thisDocument.Sheets)
                {
                    this.SheetTabs.Items.Add(new TabItem { Tag = sheet });
                }
            }
            finally
            {
                this.thisSyncingSheetTabs = false;
            }

            this.UpdateSheetTabs();
        }

        private void ClearSheetTabs()
        {
            this.thisSyncingSheetTabs = true;

            try
            {
                this.SheetTabs.Items.Clear();
            }
            finally
            {
                this.thisSyncingSheetTabs = false;
            }
        }

        private void UpdateSheetTabs()
        {
            this.thisSyncingSheetTabs = true;

            try
            {
                IReadOnlyList<BoardTableSheet> shown =
                    this.thisDocument?.SheetsShown(this.thisFilter, this.thisSearch, this.thisCurrentSheet) ?? [];

                foreach (TabItem tab in this.SheetTabs.Items.OfType<TabItem>())
                {
                    if (tab.Tag is not BoardTableSheet sheet)
                    {
                        continue;
                    }

                    // *** ITS NAME AND CHANGE COUNT, NOTHING MORE (owner request, 2026-10-02: "the
                    // tabs ... should not show "2 flagged" and "8 error" - the tabs will be obvious
                    // when you click those badges"). *** A picked pill hides every tab with none
                    // of its rows, so the tabs left ARE the answer. The flagged, error and warning
                    // pills the tabs carried from 2026-09-26 and 2026-10-02 are gone.
                    tab.Header = BoardTableEditor.SheetTabText(sheet);
                    tab.IsVisible = shown.Contains(sheet);

                    if (ReferenceEquals(sheet, this.thisCurrentSheet) && !ReferenceEquals(this.SheetTabs.SelectedItem, tab))
                    {
                        this.SheetTabs.SelectedItem = tab;
                    }
                }
            }
            finally
            {
                this.thisSyncingSheetTabs = false;
            }
        }

        // ###########################################################################################
        // A sheet tab picked by the user. Switches the grid to that sheet and hands it the keyboard,
        // so the contributor can start typing straight away rather than clicking into it first.
        //
        // Marked handled: SelectionChanged BUBBLES, and this TabControl sits inside the main window's
        // own - whose handler does guard on the source, but has no reason to see this at all.
        // ###########################################################################################
        private void OnSheetTabSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(e.Source, this.SheetTabs))
            {
                return;
            }

            e.Handled = true;

            if (this.thisSyncingSheetTabs || this.SheetTabs.SelectedItem is not TabItem { Tag: BoardTableSheet sheet })
            {
                return;
            }

            if (!ReferenceEquals(sheet, this.thisCurrentSheet))
            {
                this.SelectSheet(sheet);
            }

            Dispatcher.UIThread.Post(this.FocusGrid, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Gives the grid keyboard focus, putting the cursor on the first cell when there is none
        // yet - so typing goes into the table rather than into the always-on component filter on
        // the left. That filter pulls focus back after every click elsewhere in the window; it
        // leaves the Drafts tab alone while a table is open (Main.ShouldReturnFocusToComponentSearch),
        // and this is what puts focus here in the first place.
        // ###########################################################################################
        internal void FocusGrid()
        {
            if (this.thisCurrentSheet is null)
            {
                return;
            }

            if (this.CurrentCell is null)
            {
                BoardTableRow? first = this.thisCurrentSheet.Rows.FirstOrDefault(row => !row.IsDeleted)
                    ?? this.thisCurrentSheet.Rows.FirstOrDefault();

                if (first is not null)
                {
                    this.SelectCell(first, 0);
                }
            }

            this.TableGrid.Focus();
        }

        internal static string SheetTabText(BoardTableSheet sheet) =>
            sheet.ChangeCount > 0
                ? $"{sheet.Name} ({sheet.ChangeCount.ToString(CultureInfo.InvariantCulture)})"
                : sheet.Name;

        private void UpdateToolbar()
        {
            this.InsertRowBelowButton.IsEnabled = this.thisCurrentSheet is not null;
            this.UpdateDeleteButton();

            // A read-only table offers nothing that changes it (see IsReadOnly).
            this.InsertRowBelowButton.IsVisible = !this.thisIsReadOnly;
            this.DeleteRowButton.IsVisible = !this.thisIsReadOnly;
            this.SaveButton.IsVisible = !this.thisIsReadOnly;

            // *** THE COLOUR KEY COUNTS THE WHOLE DRAFT (owner decision, 2026-10-02: "Added,
            // Modified and Deleted should work per system like Error and Warning"). *** Every
            // sheet together, whichever is on screen - BoardTableDocument's counts. It counted the
            // sheet on screen, which read "0 Errors" while the draft's row said "8 errors"; picking
            // a pill takes the table to the sheets that have them. A sheet's own change count is on
            // its tab.
            int added = this.thisDocument?.AddedCount ?? 0;
            int modified = this.thisDocument?.ModifiedCount ?? 0;
            int deleted = this.thisDocument?.DeletedCount ?? 0;
            int errors = this.thisDocument?.ErrorRowCount ?? 0;
            int warnings = this.thisDocument?.WarningRowCount ?? 0;

            this.AddedCountText.Text = added.ToString(CultureInfo.InvariantCulture);
            this.ModifiedCountText.Text = modified.ToString(CultureInfo.InvariantCulture);
            this.DeletedCountText.Text = deleted.ToString(CultureInfo.InvariantCulture);
            this.ErrorsCountText.Text = errors.ToString(CultureInfo.InvariantCulture);
            this.WarningsCountText.Text = warnings.ToString(CultureInfo.InvariantCulture);

            // A pill with nothing to count is outlined, with no fill (owner request, 2026-10-03 -
            // it faded back instead from 2026-09-24); every pill not picked is dimmed. See the
            // LegendPill styles.
            this.AddedPill.Classes.Set("Empty", added == 0);
            this.ModifiedPill.Classes.Set("Empty", modified == 0);
            this.DeletedPill.Classes.Set("Empty", deleted == 0);
            this.ErrorsPill.Classes.Set("Empty", errors == 0);
            this.WarningsPill.Classes.Set("Empty", warnings == 0);

            // A picked pill is at full strength with a firm outline: the filter that is on.
            this.AddedPill.Classes.Set("Selected", this.thisFilter.HasFlag(BoardTableRowKinds.Added));
            this.ModifiedPill.Classes.Set("Selected", this.thisFilter.HasFlag(BoardTableRowKinds.Modified));
            this.DeletedPill.Classes.Set("Selected", this.thisFilter.HasFlag(BoardTableRowKinds.Deleted));
            this.ErrorsPill.Classes.Set("Selected", this.thisFilter.HasFlag(BoardTableRowKinds.Errors));
            this.WarningsPill.Classes.Set("Selected", this.thisFilter.HasFlag(BoardTableRowKinds.Warnings));

            // The problems no cell can carry - a highlight, drawn on a schematic - in the colour of
            // the worst of them.
            IReadOnlyList<BoardDataProblem> outside = this.thisDocument?.ProblemsOutsideSheets ?? [];
            string? outsideText = BoardTableProblemWording.OutsideSheets(outside);

            this.OutsideProblemsText.Text = outsideText ?? string.Empty;
            this.OutsideProblemsText.IsVisible = outsideText is not null;
            this.OutsideProblemsText.Foreground = BoardTableEditor.MarkBrushFor(BoardDataChecks.Worst(outside));

            // A search finding nothing in any sheet: why the one sheet left is empty (2026-10-03).
            string? nothingFound = this.thisDocument is null ? null : this.thisSearch.NothingFoundLine(this.thisDocument, this.thisFilter);
            this.NoSearchMatchText.Text = nothingFound ?? string.Empty;
            this.NoSearchMatchText.IsVisible = nothingFound is not null;
            // Reload re-reads the draft FILE, so it exists only when there is one.
            this.ReloadButton.IsEnabled = this.thisSession is not null;
            this.ReloadButton.IsVisible = this.thisSession is not null || this.thisDocument is null;
            // Off while the draft has changed underneath the table (that save would be refused) or is
            // held open in Excel (it would fail) - see BoardTableEditor.FileWatch.cs.
            this.SaveButton.IsEnabled = this.HasUnsavedChanges && !this.thisDraftChangedOnDisk && !this.thisDraftHeldElsewhere;
        }

        private void ShowStatus(string message)
        {
            this.StatusText.Text = message;
            this.StatusText.IsVisible = message.Length > 0;
        }

        // ###########################################################################################
        // The grid's columns for the current sheet: a narrow read-only marker column (+ ~ - !),
        // then one column per schema column.
        // ###########################################################################################
        private void RebuildColumns()
        {
            this.TableGrid.Columns.Clear();
            this.thisCellBrushes.Clear();

            // Built capped again - see BoardTableEditor.TextWrap.cs.
            this.thisColumnWidthsFreed = false;

            if (this.thisCurrentSheet is null)
            {
                return;
            }

            var markerTheme = new ControlTheme(typeof(DataGridCell)) { BasedOn = BoardTableEditor.DefaultCellTheme() };
            markerTheme.Setters.Add(new Setter(DataGridCell.MinHeightProperty, 0d));
            markerTheme.Setters.Add(new Setter(ToolTip.TipProperty, new Binding(nameof(BoardTableRow.MarkerToolTip))));
            BoardTableEditor.AddCellToolTipSetters(markerTheme);

            var markerSelected = new Style(selector => selector.Nesting().Class(":selected"));
            markerSelected.Setters.Add(new Setter(DataGridCell.BackgroundProperty, Brushes.Transparent));
            markerTheme.Children.Add(markerSelected);

            this.TableGrid.Columns.Add(new DataGridTextColumn
            {
                Header = string.Empty,
                Binding = new Binding(nameof(BoardTableRow.Marker)),
                IsReadOnly = true,
                FontSize = BoardTableEditor.CellFontSize,

                // The grid's cell text carries a 12px margin on EACH side, so anything narrower
                // than about 36px leaves the marker a few pixels wide - a speck, not a sign.
                Width = new DataGridLength(36),
                CanUserResize = false,
                CellTheme = markerTheme,
                Tag = -1,
            });

            for (int i = 0; i < this.thisCurrentSheet.Columns.Count; i++)
            {
                this.TableGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = this.thisCurrentSheet.Columns[i],
                    Binding = new Binding($"Cells[{i}].Text") { Mode = BindingMode.TwoWay },

                    // *** EXPLICITLY EDITABLE, or NOTHING can be typed (reported 2026-09-24). ***
                    // Left unset, the grid works out read-only-ness from the binding path, and
                    // "Cells[i]" indexes an IReadOnlyList - no setter - so it declared every data
                    // column read-only. Double-click, F2 and typing then all did nothing, although
                    // Text itself is settable. A deleted row is still refused, in OnBeginningEdit.
                    IsReadOnly = false,

                    // Per column: a text column carries its OWN font size for its cells, so the
                    // grid's FontSize alone shrank only the headers (caught by rendering it).
                    FontSize = BoardTableEditor.CellFontSize,

                    // Sized to the text, up to a limit a drag or a double-click on the heading's
                    // edge can go past (BoardTableEditor.TextWrap.cs).
                    Width = DataGridLength.Auto,
                    MinWidth = 60,
                    MaxWidth = BoardTableEditor.MaxColumnWidth,
                    CellTheme = this.BuildCellTheme(i),
                    Tag = i,
                });
            }
        }

    }
}
