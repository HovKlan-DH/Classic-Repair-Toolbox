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
    // (maintainer request, 2026-09-24).
    //
    // *** THIS CONTROL ONLY PAINTS. *** What a cell is, whether it differs from the published data,
    // where a deleted row goes and what a save writes are all decided in CRT.Data -
    // BoardTableDocument, BoardTableSheet and DraftTableSession - because the maintainer wants the
    // same table in the review application later. Keep it that way: logic added here is logic the
    // reviewer's copy would have to duplicate, and logic no test without a display can reach.
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
    // Deliberately simple for now, as the maintainer asked: one cell at a time. Copy and paste
    // work inside a cell being edited (it is an ordinary text box) and on a selected cell via
    // Ctrl+C / Ctrl+V (BoardTableClipboard). Pasting blocks of cells is a later feature. Ctrl+Z and
    // Ctrl+Y undo and redo through the model's own history (BoardTableHistory).
    //
    // FILE MAP: this file (loading, sheets, toolbar, keys, colours, the filter),
    // BoardTableEditor.RowDrag.cs (dragging a row by its grip, with the worklog-style placeholder) and
    // BoardTableEditor.FileWatch.cs (noticing the draft being changed, or open, in Excel).
    // ###########################################################################################
    public partial class BoardTableEditor : UserControl
    {
        private DraftTableSession? thisSession;
        private BoardTableSheet? thisCurrentSheet;
        private BoardData? thisPublished;
        private string thisDraftsRoot = string.Empty;
        private string thisExcelDataFile = string.Empty;

        // The table's text size - smaller than the grid's default so a sheet shows more rows at
        // once (maintainer request, 2026-09-24). The grid, its headers and every column use it.
        internal const double CellFontSize = 12;

        // Coalesces a burst of cell edits (a paste, a fast typist) into one refresh.
        private bool thisRefreshPosted;

        // What the grid shows: the current sheet's rows through a view, so "Show changes only" can
        // filter them without touching the sheet's own list.
        private DataGridCollectionView? thisView;

        private bool thisOnlyChanges;

        // True while code - not the user - is moving the sheet tabs' selection, so the selection
        // handler does not treat it as a click and steal keyboard focus into the grid.
        private bool thisSyncingSheetTabs;

        // Maps a cell state to its wash. Rebuilt with the columns, so a theme switch repaints.
        private readonly IValueConverter thisStateToBrush;

        public BoardTableEditor()
        {
            this.InitializeComponent();

            this.thisStateToBrush = new FuncValueConverter<BoardTableCellState, IBrush?>(BoardTableEditor.BrushFor);

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

            this.TableGrid.LoadingRow += this.OnLoadingRow;
            this.TableGrid.BeginningEdit += BoardTableEditor.OnBeginningEdit;
            this.WireDraftFileWatching();
            this.TableGrid.CurrentCellChanged += (_, _) => this.UpdateToolbar();

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

        public bool HasUnsavedChanges => this.thisSession?.Document.HasUnsavedChanges == true;

        public bool HasTable => this.thisSession is not null;

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
        public bool Load(string draftsRoot, string excelDataFile, BoardData? published)
        {
            DraftTableSession? session = DraftTableSession.Open(draftsRoot, excelDataFile, published);
            if (session is null)
            {
                return false;
            }

            this.thisDraftsRoot = draftsRoot;
            this.thisExcelDataFile = excelDataFile;
            this.thisPublished = published;

            this.Attach(session, keepSheet: null);
            this.ShowStatus(session.Document.HasBaseline
                ? string.Empty
                : "Nothing of this system is published yet, so no row is marked as added, changed or deleted - every row is your own.");

            return true;
        }

        // Closes the table. Unsaved edits are dropped - the caller asks first (TabDrafts).
        public void Clear()
        {
            this.Detach();

            this.thisSession = null;
            this.thisCurrentSheet = null;
            this.thisPublished = null;
            this.thisExcelDataFile = string.Empty;

            this.TableGrid.ItemsSource = Array.Empty<BoardTableRow>();
            this.TableGrid.Columns.Clear();
            this.ClearSheetTabs();
            this.ShowStatus(string.Empty);
            this.OpenElsewhereBar.IsVisible = false;
            this.SetChangedOnDisk(false);
            this.UpdateFileWatching();
        }

        // ###########################################################################################
        // Writes the table into the draft and re-reads it, so what is on screen afterwards is
        // exactly what the file now holds (a blank row is gone, for instance).
        //
        // Synchronous on purpose: it is one workbook write, and the contributor has just asked for
        // it. Refused - and said why - when the file changed since it was read; see
        // DraftTableSession for why the table may never write over such a change.
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

            switch (outcome)
            {
                case DraftWorkbookEditOutcome.Saved:
                    this.ReloadKeepingPlace();
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

            return outcome;
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
            if (this.thisSession is null || !this.thisSession.Document.Sheets.Contains(sheet))
            {
                return;
            }

            this.TableGrid.CommitEdit();
            this.EndRowDrag();

            this.thisCurrentSheet = sheet;
            this.RebuildColumns();

            // The filter is (re)applied once the grid has the view - see ApplyOnlyChangesFilter.
            this.thisView = new DataGridCollectionView(sheet.Rows);
            this.TableGrid.ItemsSource = this.thisView;
            this.ApplyOnlyChangesFilter();

            this.UpdateSheetTabs();
            this.UpdateToolbar();
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

            this.TableGrid.CurrentCell = new DataGridCellInfo(row, column, rowIndex, column.DisplayIndex, true);
            this.TableGrid.ScrollIntoView(row, column);
            this.UpdateToolbar();
        }

        // ###########################################################################################
        // Toolbar actions. Each is a thin call into BoardTableSheet, which owns the rule.
        // ###########################################################################################
        // "Insert row below" and "Insert row above" (maintainer request, 2026-09-24): an empty row
        // beside the selected one, with the cursor on it ready for typing.
        internal void InsertRowBelow() => this.InsertRowBeside(above: false);

        internal void InsertRowAbove() => this.InsertRowBeside(above: true);

        private void InsertRowBeside(bool above)
        {
            if (this.thisCurrentSheet is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            BoardTableRow inserted = above
                ? this.thisCurrentSheet.InsertRowAbove(this.CurrentRow)
                : this.thisCurrentSheet.InsertRow(this.CurrentRow);
            this.SelectCell(inserted, 0);
        }

        internal void DeleteRow()
        {
            if (this.thisCurrentSheet is null || this.CurrentRow is not { IsDeleted: false } row)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            int index = this.thisCurrentSheet.Rows.IndexOf(row);
            int column = this.CurrentCell?.ColumnIndex ?? 0;

            if (this.thisCurrentSheet.DeleteRow(row) && this.thisCurrentSheet.Rows.Count > 0)
            {
                this.SelectCell(this.thisCurrentSheet.Rows[Math.Min(index, this.thisCurrentSheet.Rows.Count - 1)], column);
            }

            this.UpdateToolbar();
        }

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

        // ###########################################################################################
        // "Show changes only": hide every plainly unchanged row (maintainer request, 2026-09-24).
        // Moving rows is switched off meanwhile - a position among rows that cannot be seen means
        // nothing.
        // ###########################################################################################
        internal bool OnlyChanges
        {
            get => this.thisOnlyChanges;
            set
            {
                if (this.thisOnlyChanges == value)
                {
                    return;
                }

                this.thisOnlyChanges = value;
                this.OnlyChangesCheckBox.IsChecked = value;
                this.TableGrid.CommitEdit();
                this.UpdateRowsDraggable();

                this.ApplyOnlyChangesFilter();

                this.UpdateToolbar();
            }
        }

        // ###########################################################################################
        // *** THE GRID MUST BE TOLD THE VIEW'S FILTER IS OURS. *** Its column-filtering model
        // (FilteringModel) owns a view's Filter by default and writes its own - empty - predicate
        // over it whenever it takes the view or is attached again. So "Show changes only" was lost,
        // with the box still ticked, first on every sheet switch and then on every switch of the
        // MAIN tabs (Drafts -> Contribute -> Drafts), both reported. OwnsViewFilter = false makes
        // it leave the filter alone; the table uses none of its column filtering.
        // ###########################################################################################
        private void ApplyOnlyChangesFilter()
        {
            if (this.TableGrid.FilteringModel is { } model)
            {
                model.OwnsViewFilter = false;
            }

            if (this.thisView is not null)
            {
                this.thisView.Filter = this.thisOnlyChanges ? BoardTableEditor.ShowsInOnlyChanges : null;
            }
        }

        private void OnOnlyChangesChanged(object? sender, RoutedEventArgs e) =>
            this.OnlyChanges = this.OnlyChangesCheckBox.IsChecked == true;

        private static bool ShowsInOnlyChanges(object item) =>
            item is BoardTableRow row && BoardTableSheet.IsChangeRow(row);

        internal void MoveRowUp() => this.MoveCurrentRow(up: true);

        internal void MoveRowDown() => this.MoveCurrentRow(up: false);

        // Moves the row under the cursor and keeps the cursor on it, in the same column.
        private void MoveCurrentRow(bool up)
        {
            if (this.thisOnlyChanges || this.thisCurrentSheet is null || this.CurrentRow is not { IsDeleted: false } row)
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
            if (this.CurrentCell is not { } cell || cell.Row.IsDeleted)
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
            // *** TAB MOVES ACROSS THE ROW, AS IN EXCEL (maintainer request, 2026-09-24). *** Left
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
                Logger.Warning($"Table editor clipboard {(isCopy ? "copy" : "paste")} failed - [{ex.Message}]");
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
        // gestures (maintainer request, 2026-09-24). Tunnel, so the grid never sees them first.
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
            if (this.thisSession is null)
            {
                return;
            }

            this.TableGrid.CommitEdit();
            int column = this.CurrentCell?.ColumnIndex ?? 0;

            BoardTableHistoryResult? result = step(this.thisSession.Document.History);
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
        private void OnInsertRowAboveClick(object? sender, RoutedEventArgs e) => this.ThenFocusGrid(this.InsertRowAbove);

        private void OnInsertRowBelowClick(object? sender, RoutedEventArgs e) => this.ThenFocusGrid(this.InsertRowBelow);

        private void OnDeleteRowClick(object? sender, RoutedEventArgs e) => this.ThenFocusGrid(this.DeleteRow);

        private void ThenFocusGrid(Action action)
        {
            action();
            this.FocusGrid();
        }

        private void OnSaveClick(object? sender, RoutedEventArgs e) => this.Save();

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

        private void Attach(DraftTableSession session, string? keepSheet)
        {
            this.Detach();

            this.thisSession = session;
            session.Document.Changed += this.OnDocumentChanged;

            foreach (BoardTableSheet sheet in session.Document.Sheets)
            {
                sheet.CellEdited += this.OnCellEdited;
            }

            this.BuildSheetTabs();

            BoardTableSheet first = (keepSheet is null ? null : session.Document.FindSheet(keepSheet))
                ?? session.Document.Sheets[0];

            // *** NOTHING PUBLISHED, NO FILTER. *** The box is hidden then, and every row is
            // "unchanged" - so a filter left on from another draft's table (this editor is reused)
            // hid EVERY row with no visible reason: reported as a "Board schematics" sheet showing
            // empty although the draft had three schematic images.
            if (!session.Document.HasBaseline)
            {
                this.OnlyChanges = false;
            }

            this.SelectSheet(first);

            // A fresh read of the file: whatever the warning bar said is no longer true.
            this.SetChangedOnDisk(false);
            this.ShowWhetherOpenElsewhere(session);
            this.UpdateFileWatching();

            // With nothing published there is nothing to be added, changed or deleted against, so
            // only the flagged pill stays - a duplicate is a duplicate either way.
            bool hasBaseline = session.Document.HasBaseline;
            this.AddedPill.IsVisible = hasBaseline;
            this.ModifiedPill.IsVisible = hasBaseline;
            this.DeletedPill.IsVisible = hasBaseline;
            this.OnlyChangesCheckBox.IsVisible = hasBaseline;
        }

        private void Detach()
        {
            this.EndRowDrag();

            if (this.thisSession is null)
            {
                return;
            }

            this.thisSession.Document.Changed -= this.OnDocumentChanged;

            foreach (BoardTableSheet sheet in this.thisSession.Document.Sheets)
            {
                sheet.CellEdited -= this.OnCellEdited;
            }
        }

        // Re-opens the draft on the same sheet, with the cursor back on (roughly) the same row.
        private bool ReloadKeepingPlace()
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

            DraftTableSession? session = DraftTableSession.Open(this.thisDraftsRoot, this.thisExcelDataFile, this.thisPublished);
            if (session is null)
            {
                return false;
            }

            this.Attach(session, sheetName);

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

                if (this.thisSession is null)
                {
                    return;
                }

                foreach (BoardTableSheet sheet in this.thisSession.Document.Sheets.Where(s => s.NeedsRefresh))
                {
                    sheet.Refresh();
                }

                this.UpdateToolbar();
            }, DispatcherPriority.Background);
        }

        // Test seam: runs the refresh a cell edit schedules, without waiting for the dispatcher.
        internal void RefreshPendingForTests()
        {
            foreach (BoardTableSheet sheet in this.thisSession?.Document.Sheets.Where(s => s.NeedsRefresh) ?? [])
            {
                sheet.Refresh();
            }

            this.UpdateToolbar();
        }

        private void OnDocumentChanged(object? sender, EventArgs e)
        {
            this.UpdateSheetTabs();
            this.UpdateToolbar();

            // A row whose last change was just reverted drops out of "Show changes only", and a newly
            // changed one joins it - the filter only re-reads states when asked.
            if (this.thisOnlyChanges && this.thisView is not null && !this.thisView.IsEditingItem)
            {
                this.thisView.Refresh();
            }
        }

        // ###########################################################################################
        // One tab per sheet, labelled with its change count when it has any: "Components (3)". The
        // tabs carry no content of their own - see the markup's comment on SheetTabs.
        // ###########################################################################################
        private void BuildSheetTabs()
        {
            this.ClearSheetTabs();

            if (this.thisSession is null)
            {
                return;
            }

            this.thisSyncingSheetTabs = true;

            try
            {
                foreach (BoardTableSheet sheet in this.thisSession.Document.Sheets)
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
                foreach (TabItem tab in this.SheetTabs.Items.OfType<TabItem>())
                {
                    if (tab.Tag is not BoardTableSheet sheet)
                    {
                        continue;
                    }

                    tab.Header = BoardTableEditor.SheetTabText(sheet);

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
            BoardTableRow? row = this.CurrentRow;

            this.InsertRowAboveButton.IsEnabled = this.thisCurrentSheet is not null;
            this.InsertRowBelowButton.IsEnabled = this.thisCurrentSheet is not null;
            this.DeleteRowButton.IsEnabled = row is { IsDeleted: false };

            // The colour key's pills count the sheet on screen. The first three add up to its tab's
            // number; the flagged one is counted apart, as the tab's number does not include it.
            this.AddedCountText.Text = (this.thisCurrentSheet?.AddedCount ?? 0).ToString(CultureInfo.InvariantCulture);
            this.ModifiedCountText.Text = (this.thisCurrentSheet?.ModifiedCount ?? 0).ToString(CultureInfo.InvariantCulture);
            this.DeletedCountText.Text = (this.thisCurrentSheet?.DeletedCount ?? 0).ToString(CultureInfo.InvariantCulture);
            this.FlaggedCountText.Text = (this.thisCurrentSheet?.FlaggedCount ?? 0).ToString(CultureInfo.InvariantCulture);

            // A pill with nothing to count fades back, like a disabled control (maintainer request,
            // 2026-09-24), so the eye goes to the kinds that are actually there.
            this.AddedPill.Classes.Set("Empty", (this.thisCurrentSheet?.AddedCount ?? 0) == 0);
            this.ModifiedPill.Classes.Set("Empty", (this.thisCurrentSheet?.ModifiedCount ?? 0) == 0);
            this.DeletedPill.Classes.Set("Empty", (this.thisCurrentSheet?.DeletedCount ?? 0) == 0);
            this.FlaggedPill.Classes.Set("Empty", (this.thisCurrentSheet?.FlaggedCount ?? 0) == 0);
            this.ReloadButton.IsEnabled = this.thisSession is not null;
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

            if (this.thisCurrentSheet is null)
            {
                return;
            }

            var markerTheme = new ControlTheme(typeof(DataGridCell)) { BasedOn = BoardTableEditor.DefaultCellTheme() };
            markerTheme.Setters.Add(new Setter(DataGridCell.MinHeightProperty, 0d));
            markerTheme.Setters.Add(new Setter(ToolTip.TipProperty, new Binding(nameof(BoardTableRow.MarkerToolTip))));

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

                    Width = DataGridLength.Auto,
                    MinWidth = 60,
                    MaxWidth = 420,
                    CellTheme = this.BuildCellTheme(i),
                    Tag = i,
                });
            }
        }

        private ControlTheme BuildCellTheme(int columnIndex)
        {
            var theme = new ControlTheme(typeof(DataGridCell)) { BasedOn = BoardTableEditor.DefaultCellTheme() };

            // The grid theme's cells are taller than the table's denser rows (RowHeight in the
            // markup), which clipped the current cell's frame to its left and right edges.
            theme.Setters.Add(new Setter(DataGridCell.MinHeightProperty, 0d));

            theme.Setters.Add(new Setter(
                DataGridCell.BackgroundProperty,
                new Binding($"Cells[{columnIndex}].State") { Converter = this.thisStateToBrush }));

            // *** A SELECTED CELL KEEPS ITS OWN COLOUR (maintainer request, 2026-09-24). *** The
            // grid theme fills a selected cell with the accent colour, which read as one more
            // state - a green that could be taken for "added" - and hid the orange, green or red
            // of the cell underneath. The current cell is marked by its dashed frame instead (see
            // the markup). A trigger of our own, added after the grid theme's, outranks it.
            var selected = new Style(selector => selector.Nesting().Class(":selected"));
            selected.Setters.Add(new Setter(
                DataGridCell.BackgroundProperty,
                new Binding($"Cells[{columnIndex}].State") { Converter = this.thisStateToBrush }));
            theme.Children.Add(selected);

            theme.Setters.Add(new Setter(
                ToolTip.TipProperty,
                new Binding($"Cells[{columnIndex}].ToolTip")));

            return theme;
        }

        // The grid theme's own cell theme, so ours only ADDS the background and tooltip rather than
        // replacing the cell's whole template.
        private static ControlTheme? DefaultCellTheme() =>
            Application.Current?.TryGetResource(typeof(DataGridCell), Application.Current.ActualThemeVariant, out object? theme) == true
                ? theme as ControlTheme
                : null;

        internal static IBrush? BrushFor(BoardTableCellState state) => state switch
        {
            BoardTableCellState.Added => ThemeResources.Resolve<IBrush>("BoardTable_Added_Bg", Brushes.LightGreen),
            BoardTableCellState.Modified => ThemeResources.Resolve<IBrush>("BoardTable_Modified_Bg", Brushes.Orange),
            BoardTableCellState.Deleted => ThemeResources.Resolve<IBrush>("BoardTable_Deleted_Bg", Brushes.LightPink),
            BoardTableCellState.Flagged => ThemeResources.Resolve<IBrush>("BoardTable_Flagged_Bg", Brushes.Lavender),
            _ => Brushes.Transparent,
        };
    }
}
