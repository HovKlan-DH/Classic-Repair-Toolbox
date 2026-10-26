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
    // "EDIT IN TABLE FORMAT" - a draft's workbook as editable sheets, inline on the Contribute tab
    // (owner request, 2026-09-24).
    //
    // *** THIS CONTROL ONLY PAINTS. *** What a cell is, whether it differs from the published data,
    // where a deleted row goes and what a save writes are all decided in CRT.Data -
    // BoardTableDocument, BoardTableSheet and DraftSession - because the project owner wants the
    // same table in the Maintainer tab later. Keep it that way: logic added here is logic the
    // maintainer's copy would have to duplicate, and logic no test without a display can reach.
    //
    // It deliberately knows nothing of Main or DataManager either: it is handed a draft's session (or
    // a document) and the published board to compare with, and raises Saved. TabDrafts does the wiring.
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
    // *** IN THE DRAFTS TAB THE TABLE IS A VIEW OF THE DRAFT HELD IN MEMORY (CRT 4.0 step 1,
    // 2026-10-25). *** Its changes go into the draft's DraftSession, which saves itself; undo and redo
    // are the session's; there is no Save and no Reload - BoardTableEditor.DraftSession.cs. The
    // Maintainer tab's tables are documents (Open) and keep their explicit Save.
    //
    // A cell is typed into one at a time, as the project owner asked. Copy and paste work inside a
    // cell being edited (it is an ordinary text box), and on the selection via Ctrl+C / Ctrl+V -
    // blocks of cells since CRT 4.0, with fill down and find & replace (2026-10-25,
    // BoardTableEditor.BlockEdits.cs). Several ROWS can be selected and deleted at once (2026-10-02).
    // Ctrl+Z and Ctrl+Y undo and redo through the draft's session, or - a document - the model's own
    // history (BoardTableHistory).
    //
    // FILE MAP: this file (loading, sheets, toolbar, keys, columns),
    // BoardTableEditor.Filter.cs (the colour key's pills as the filter),
    // BoardTableEditor.CellColours.cs (each cell's background and tooltip - its state's wash and
    // its problem's corner mark),
    // BoardTableEditor.Selection.cs (several rows selected, "Delete row" on all of them),
    // BoardTableEditor.BlockEdits.cs (copying and pasting blocks, fill down, the replace bar),
    // BoardTableEditor.Search.cs (the search box, and the search handed to BoardTableCellText),
    // BoardTableEditor.TextWrap.cs (cells that always wrap, columns dragged wider than they size
    // themselves, and a double-click on a heading's edge fitting a column to its text),
    // BoardTableEditor.RowDrag.cs (dragging a row by its grip, with the worklog-style placeholder),
    // BoardTableEditor.DraftSession.cs (the Contribute tab's table on the draft held in memory: its
    // changes handed over, saving by itself, undo and redo, the file changed outside it) and
    // BoardTableEditor.FilePreview.cs (the hover card on a file cell - its content is
    // BoardTableFilePreview, its bytes the host's IBoardTableFileSource).
    // ###########################################################################################
    public partial class BoardTableEditor : UserControl
    {
        // The table on screen, and - in the Contribute tab only - the draft behind it (thisDraftSession,
        // BoardTableEditor.DraftSession.cs). In document mode (the Maintainer tab,
        // Open(BoardTableDocument)) there is no session: Save raises SaveRequested for the host.
        private BoardTableDocument? thisDocument;
        private BoardTableSheet? thisCurrentSheet;
        private BoardData? thisPublished;
        private string thisExcelDataFile = string.Empty;

        // What a changed cell's tooltip calls the published value - the host's, kept for re-reads.
        private string? thisBaselineLabel;

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
            this.WireDraftSession();
            this.TableGrid.CurrentCellChanged += (_, _) => this.UpdateToolbar();

            // Several rows selected and deleted at once (BoardTableEditor.Selection.cs), and the
            // search box (BoardTableEditor.Search.cs).
            this.WireSelection();
            this.WireSearch();

            // Block copy and paste, fill down, find & replace (BoardTableEditor.BlockEdits.cs).
            this.WireBlockEdits();

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

        // Raised after the draft saved itself with a change of the table's in it - TabDrafts
        // refreshes its row count and the board on screen from it.
        public event EventHandler? Saved;

        // ###########################################################################################
        // Whether something on screen is not in the file yet: a change the table has not handed to
        // its draft yet, or the draft's own save still to come or failed (DraftSession). A document's
        // edits until the host saves them.
        // ###########################################################################################
        public bool HasUnsavedChanges =>
            this.thisDocument?.HasUnsavedChanges == true || this.thisDraftSession?.HasUnsavedChanges == true;

        // ###########################################################################################
        // Raised when HasUnsavedChanges turns true or false - an edit, an undo back to the save, a
        // save, a table opened or closed (2026-10-09: the Boards screen's "Compare sources" is off
        // while BETA's table holds a change, since comparing opens the table again). Checked
        // wherever the Save button is (UpdateToolbar), so the two never disagree.
        // ###########################################################################################
        public event EventHandler? UnsavedChangesChanged;

        private bool thisLastUnsaved;

        private void NoteUnsavedState()
        {
            bool unsaved = this.HasUnsavedChanges;

            if (unsaved == this.thisLastUnsaved)
                return;

            this.thisLastUnsaved = unsaved;
            this.UnsavedChangesChanged?.Invoke(this, EventArgs.Empty);
        }

        public bool HasTable => this.thisDocument is not null;

        // ###########################################################################################
        // DOCUMENT MODE - the Maintainer tab's table (2026-09-25). Shows `document` with no
        // draft file behind it: "Save changes" raises SaveRequested, and the host saves it however
        // it saves (the Maintainer tab sends it to the server as an amendment), then calls
        // Open again with the saved state, or ShowMessage to say why not. A draft's save state and
        // its notices do not apply and stay hidden.
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

            this.thisExcelDataFile = string.Empty;
            this.thisPublished = null;
            this.thisBaselineLabel = null;

            this.Attach(document, session: null, keepSheet);
            this.ShowStatus(message ?? string.Empty);
        }

        // Raised by "Save changes" in document mode - see Open.
        public event EventHandler? SaveRequested;

        // ###########################################################################################
        // READ-ONLY (2026-10-03): the Maintainer tab's Boards screen shows the data of a board the
        // account may not change. The table is there to look at - sheets, search, the colour key,
        // copying a cell, the file cards - but nothing in it can be typed, pasted, moved, inserted
        // or deleted, and there is no "Save changes": an edit that could never be sent is worse than
        // none. The host says why (ShowMessage). Off by default; the Contribute tab never sets it.
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

        // Whether a draft FILE is behind the table (the Contribute tab) or only a document (review).
        public bool IsFileBacked => this.thisDraftSession is not null;

        // The document on screen, with any cell still in its editor committed into it first - what
        // the host reads when it saves in document mode.
        public BoardTableDocument? CommitAndGetDocument()
        {
            this.TableGrid.CommitEdit();
            return this.thisDocument;
        }

        // ###########################################################################################
        // A cell still in its editor, committed into the document - for a host acting on the table
        // from a control of its own (the Boards screen's "Compare sources", code review 2026-10-09).
        // The grid does not commit when focus moves to such a control, and every button of the
        // table's own commits first for the same reason.
        // ###########################################################################################
        public void CommitCellEdit() => this.TableGrid.CommitEdit();

        // Whether a cell is in its editor now - nothing may build the table again under it.
        public bool IsEditingCell => this.thisIsEditingCell;

        // Raised when a cell's editor closes, committed or cancelled.
        public event EventHandler? CellEditEnded;

        // A message under the toolbar, for the host in document mode ("Saved.", "Not saved: ...").
        public void ShowMessage(string message) => this.ShowStatus(message ?? string.Empty);

        public string ExcelDataFile => this.thisExcelDataFile;

        internal BoardTableSheet? CurrentSheet => this.thisCurrentSheet;

        // Closes the table. What is not saved is dropped - the caller saves first (TabDrafts,
        // SaveNowAsync). The draft's session stays with its host, saved.
        public void Clear()
        {
            this.Detach();

            this.thisDocument = null;
            this.UseDraftSession(null);
            this.thisCurrentSheet = null;
            this.thisPublished = null;
            this.thisBaselineLabel = null;
            this.thisExcelDataFile = string.Empty;

            this.TableGrid.ItemsSource = Array.Empty<BoardTableRow>();
            this.TableGrid.Columns.Clear();
            this.ClearSheetTabs();
            this.ShowStatus(string.Empty);

            // A closed table's search goes with it (agreed, 2026-10-02).
            this.SearchText = string.Empty;
            this.NoteUnsavedState();
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

        // The sheet of that name shown - for a host that opens the table on one of the board's parts
        // (the Contribute tab's left column). False when the table holds no such sheet.
        internal bool SelectSheetNamed(string sheetName)
        {
            if (this.thisDocument?.FindSheet(sheetName) is not BoardTableSheet sheet)
            {
                return false;
            }

            this.SelectSheet(sheet);
            return true;
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

        // Ctrl+C, Ctrl+V and Ctrl+D act on the selection - BoardTableEditor.BlockEdits.cs.

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

            // Ctrl+D fills down, as in a spreadsheet.
            if (BoardTableEditor.IsCommand(e, Key.D))
            {
                e.Handled = true;
                this.FillDown();
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
                    await clipboard.SetTextAsync(this.CopySelectionText() ?? string.Empty);
                }
                else
                {
                    this.Paste(await clipboard.TryGetTextAsync());
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
            // Ctrl+F to the search box and Ctrl+H to the replace bar from anywhere in the table, a
            // cell being typed in too - it is committed first (BoardTableEditor.BlockEdits.cs).
            if (BoardTableEditor.IsCommand(e, Key.F))
            {
                e.Handled = true;
                this.TableGrid.CommitEdit();
                this.SearchBox.Focus();
                this.SearchBox.SelectAll();
                return;
            }

            if (BoardTableEditor.IsCommand(e, Key.H) && !this.thisIsReadOnly && this.thisDocument is not null)
            {
                e.Handled = true;
                this.OpenReplaceBar();
                return;
            }

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

        // A draft's undo and redo are its session's, across every editor of it; a document's, the
        // table's own history.
        internal void Undo()
        {
            if (this.thisDraftSession is not null)
            {
                this.StepDraftSession(undo: true);
                return;
            }

            this.StepHistory(history => history.Undo());
        }

        internal void Redo()
        {
            if (this.thisDraftSession is not null)
            {
                this.StepDraftSession(undo: false);
                return;
            }

            this.StepHistory(history => history.Redo());
        }

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

        // "Save changes" is a document's only (the Maintainer tab) - a draft saves itself. The host
        // saves - see Open.
        private void OnSaveClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisDraftSession is null && this.thisDocument is not null)
            {
                this.TableGrid.CommitEdit();
                this.SaveRequested?.Invoke(this, EventArgs.Empty);
            }
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

        private void Attach(BoardTableDocument document, DraftSession? session, string? keepSheet)
        {
            this.Detach();

            this.thisDocument = document;
            this.UseDraftSession(session);
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
            this.ShowSaveState();

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

            // A draft's change goes into its session as soon as it has settled.
            this.SchedulePush();
            this.ShowSaveState();

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
        // leaves the Contribute tab alone (Main.ShouldReturnFocusToComponentSearch),
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
            this.UpdateBlockEditButtons();

            // A read-only table offers nothing that changes it (see IsReadOnly).
            this.InsertRowBelowButton.IsVisible = !this.thisIsReadOnly;
            this.DeleteRowButton.IsVisible = !this.thisIsReadOnly;
            // A draft saves itself and says how that goes instead (SaveStatePanel); a document is
            // saved by its host.
            this.SaveButton.IsVisible = !this.thisIsReadOnly && this.thisDraftSession is null;

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

            // The problems in the sheets, as text, while their pill is picked (code review, 2026-10-13).
            string? problemList = BoardTableProblemWording.InSheets(
                this.thisDocument?.Problems ?? [],
                errors: this.thisFilter.HasFlag(BoardTableRowKinds.Errors),
                warnings: this.thisFilter.HasFlag(BoardTableRowKinds.Warnings));

            this.ProblemListText.Text = problemList ?? string.Empty;
            this.ProblemListText.IsVisible = problemList is not null;

            // A search finding nothing in any sheet: why the one sheet left is empty (2026-10-03).
            string? nothingFound = this.thisDocument is null ? null : this.thisSearch.NothingFoundLine(this.thisDocument, this.thisFilter);
            this.NoSearchMatchText.Text = nothingFound ?? string.Empty;
            this.NoSearchMatchText.IsVisible = nothingFound is not null;
            this.SaveButton.IsEnabled = this.HasUnsavedChanges;
            this.NoteUnsavedState();
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
                this.TableGrid.Columns.Add(new BoardTableTextColumn
                {
                    Header = this.thisCurrentSheet.Columns[i],
                    Binding = new Binding($"Cells[{i}].Text") { Mode = BindingMode.TwoWay },
                    DisplayPath = $"Cells[{i}].Text",

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
