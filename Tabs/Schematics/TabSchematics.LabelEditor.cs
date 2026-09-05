using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.GestureRecognizers;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// Component label editor mode: entering/leaving the mode, the editor menu, loading the
// working copy, applying or cancelling changes, validation and save dialogs, the search
// filter, undo/redo, and the pooled editor label visuals.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private readonly List<Border> thisEditorLabelContainers = new();

    private readonly List<TextBlock> thisEditorLabelTextBlocks = new();

    private readonly List<ScaleTransform> thisEditorLabelScaleTransforms = new();

    private string thisLastEditorLabelVisualSignature = string.Empty;

    private string thisLastCreatedLabelEditorCategory = string.Empty;

    private string thisLabelEditorSearchText = string.Empty;

    private readonly Stack<LabelEditorUndoState> thisLabelEditorUndoStack = new();

    private readonly Stack<LabelEditorUndoState> thisLabelEditorRedoStack = new();

    // Label editor
    private bool thisIsLabelEditorMode;

    private bool thisIsShowingLabelEditorMenu;

    private Point thisLastLabelEditorMenuPoint;

    private string thisLabelEditorSchematicName = string.Empty;

    private readonly List<EditableComponentHighlight> thisLabelEditorWorkingHighlights = new();

    private EditableComponentHighlight? thisPendingNewLabelEditorHighlight;

    // ###########################################################################################
    // Returns true when the pointer is currently inside the floating label-editor menu bounds.
    // ###########################################################################################
    private bool IsPointerInsideLabelEditorMenu(Point containerPoint)
    {
        if (!this.SchematicsLabelEditorMenuBorder.IsVisible)
        {
            return false;
        }

        Point? translatedTopLeft = this.SchematicsLabelEditorMenuBorder.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
        if (!translatedTopLeft.HasValue)
        {
            return false;
        }

        var menuRect = new Rect(translatedTopLeft.Value, this.SchematicsLabelEditorMenuBorder.Bounds.Size);
        return menuRect.Contains(containerPoint);
    }

    // ###########################################################################################
    // Shows the floating schematic action menu at the requested schematic container location.
    // The menu adapts its height to contributor mode, label editor mode, and the new KiCad trace
    // calibration workflow.
    // ###########################################################################################
    private void ShowLabelEditorMenu(Point containerPoint)
    {
        if (!this.CanShowSchematicsActionsMenu())
        {
            return;
        }

        this.thisLastLabelEditorMenuPoint = containerPoint;
        this.UpdateLabelEditorMenuButtons();

        double estimatedWidth = 250.0;
        double estimatedHeight =
            this.thisIsLabelEditorMode ? 110.0 :
            this.thisIsKiCadTraceCalibrationMode ? 105.0 :
            195.0;

        double x = Math.Clamp(
            containerPoint.X,
            6.0,
            Math.Max(6.0, this.SchematicsContainer.Bounds.Width - estimatedWidth));

        double y = Math.Clamp(
            containerPoint.Y,
            6.0,
            Math.Max(6.0, this.SchematicsContainer.Bounds.Height - estimatedHeight));

        this.SchematicsLabelEditorMenuBorder.Margin = new Thickness(x, y, 0, 0);
        this.SchematicsLabelEditorMenuBorder.IsVisible = true;
        this.thisIsShowingLabelEditorMenu = true;
    }

    // ###########################################################################################
    // Hides the floating label-editor action menu.
    // ###########################################################################################
    private void HideLabelEditorMenu()
    {
        this.SchematicsLabelEditorMenuBorder.IsVisible = false;
        this.thisIsShowingLabelEditorMenu = false;
    }

    // ###########################################################################################
    // Updates the menu text and button visibility according to the current editor and calibration
    // state, including the new KiCad trace calibration start/apply/discard workflow.
    // ###########################################################################################
    private void UpdateLabelEditorMenuButtons()
    {
        this.SchematicsLabelEditorMenuStateTextBlock.Text =
            this.thisIsLabelEditorMode
                ? "Component label editor mode"
                : this.thisIsKiCadTraceCalibrationMode
                    ? "KiCad trace calibration"
                    : "Contributor mode actions";

        this.EnableLabelEditorButton.IsVisible =
            !this.thisIsLabelEditorMode &&
            !this.thisIsKiCadTraceCalibrationMode;

        this.BeginKiCadTraceCalibrationButton.IsVisible =
            !this.thisIsLabelEditorMode &&
            !this.thisIsKiCadTraceCalibrationMode &&
            this.HasCurrentSchematicKiCadTraces();

        this.ApplyLabelEditorChangesButton.IsVisible = this.thisIsLabelEditorMode;
        this.CancelLabelEditorChangesButton.IsVisible = this.thisIsLabelEditorMode;

        this.ApplyKiCadTraceCalibrationButton.IsVisible = this.thisIsKiCadTraceCalibrationMode;
        this.CancelKiCadTraceCalibrationButton.IsVisible = this.thisIsKiCadTraceCalibrationMode;
    }

    // ###########################################################################################
    // Loads a fresh working-copy snapshot of highlight rectangles for the current schematic.
    // ###########################################################################################
    private void LoadLabelEditorWorkingCopyForCurrentSchematic()
    {
        this.thisLabelEditorWorkingHighlights.Clear();
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlight = null;
        this.thisLabelEditorOriginalDragRectangles.Clear();
        this.thisLabelEditorSchematicName = this.GetCurrentSchematicName();

        // Via the property rather than MainWindow directly, so a headless test can supply board
        // data without constructing Main - see TabSchematics.LabelEditor.TestSeams.cs. In the
        // running app the override is null and this is exactly MainWindow.CurrentBoardData.
        var boardData = this.CurrentBoardDataForLabelEditor;
        if (boardData == null || string.IsNullOrWhiteSpace(this.thisLabelEditorSchematicName))
        {
            return;
        }

        foreach (var row in boardData.ComponentHighlights.Where(h =>
                     string.Equals(h.SchematicName, this.thisLabelEditorSchematicName, StringComparison.OrdinalIgnoreCase)))
        {
            if (!RectGeometry.TryParseDouble(row.X, out var x) ||
                !RectGeometry.TryParseDouble(row.Y, out var y) ||
                !RectGeometry.TryParseDouble(row.Width, out var width) ||
                !RectGeometry.TryParseDouble(row.Height, out var height))
            {
                continue;
            }

            if (width <= 0 || height <= 0)
            {
                continue;
            }

            string boardLabel = row.BoardLabel?.Trim() ?? string.Empty;
            string category = boardData.Components
                .FirstOrDefault(component => string.Equals(component.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                ?.Category?.Trim() ?? string.Empty;

            this.thisLabelEditorWorkingHighlights.Add(new EditableComponentHighlight
            {
                SchematicName = row.SchematicName?.Trim() ?? string.Empty,
                BoardLabel = boardLabel,
                Category = category,
                X = x,
                Y = y,
                Width = width,
                Height = height
            });
        }
    }

    // ###########################################################################################
    // Enters label-editor mode and captures the current schematic highlights into a working copy.
    // Closes the action menu immediately after enabling the editor.
    // ###########################################################################################
    private void BeginLabelEditorMode()
    {
        // Worklog entry mode intercepts the pointer handlers ahead of this one, so leaving it
        // active would make the label editor appear frozen - it would be on, but unable to receive
        // any input. BeginWorklogEntryMode refuses to start over this mode; this is the other half
        // of that exclusion.
        this.CancelWorklogEntryMode();

        this.thisIsLabelEditorMode = true;
        this.thisLastCreatedLabelEditorCategory = string.Empty;

        this.LoadLabelEditorWorkingCopyForCurrentSchematic();
        this.UpdateLabelEditorLockState();
        this.RefreshLabelEditorOverlay();
        this.HideLabelEditorMenu();
        this.SchematicsContainer.Focus();

        Logger.Info($"Label editor enabled for schematic [{this.thisLabelEditorSchematicName}] with [{this.thisLabelEditorWorkingHighlights.Count}] rectangles loaded");
    }

    // ###########################################################################################
    // Cancels the current label-editor session and discards all in-memory editor changes.
    // Clears undo and redo because the editor session is ending.
    // ###########################################################################################
    private void CancelLabelEditorChanges()
    {
        this.thisIsLabelEditorMode = false;
        this.thisLabelEditorSchematicName = string.Empty;
        this.thisLastCreatedLabelEditorCategory = string.Empty;
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlight = null;
        this.thisPendingNewLabelEditorHighlight = null;
        this.thisIsDrawingLabelEditorRectangle = false;
        this.thisLabelEditorDraftRectangle = null;
        this.thisLabelEditorDragMode = LabelEditorDragMode.None;
        this.thisLabelEditorOriginalSelectionBounds = default;
        this.thisLabelEditorOriginalDragRectangles.Clear();
        this.thisLabelEditorWorkingHighlights.Clear();
        this.thisLabelEditorUndoStack.Clear();
        this.thisLabelEditorRedoStack.Clear();

        this.UpdateLabelEditorLockState();
        this.HideLabelEditorMenu();
        this.HideNewLabelEditorPrompt();
        this.RefreshLabelEditorOverlay();
        this.SchematicsContainer.Focus();

        Logger.Info("Label editor changes canceled");
    }

    // ###########################################################################################
    // Validates and saves the current editor session to the board Excel file, then reloads the
    // board from disk so the runtime state reflects the persisted workbook content.
    // Clears undo and redo because the editor session is ending.
    // ###########################################################################################
    private async void ApplyLabelEditorChanges()
    {
        if (this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            this.ConfirmNewLabelEditorPrompt();

            if (this.SchematicsNewLabelPromptBorder.IsVisible)
            {
                Logger.Info("Label editor save aborted because the new-label prompt is still visible after confirmation attempt");
                return;
            }
        }

        if (!this.TryValidateLabelEditorSave(out var validationError))
        {
            Logger.Warning($"Label editor validation failed - [{validationError}]");
            await this.ShowLabelEditorValidationFailedDialogAsync(validationError);
            return;
        }

        string schematicName = this.GetCurrentSchematicName();
        string excelPath = this.MainWindow?.GetCurrentBoardExcelPath() ?? string.Empty;
        string cacheKey = this.MainWindow?.GetCurrentBoardEntry()?.ExcelDataFile?.Trim() ?? string.Empty;
        string region = this.MainWindow?.LocalRegion?.Trim() ?? string.Empty;

        Logger.Debug($"Label editor resolved Excel save path: [{excelPath}]");
        Logger.Debug($"Label editor resolved cache key: [{cacheKey}]");
        Logger.Debug($"Label editor resolved schematic name: [{schematicName}]");
        Logger.Debug($"Label editor resolved region: [{region}]");

        if (string.IsNullOrWhiteSpace(excelPath))
        {
            Logger.Warning("Label editor save failed - could not resolve current board Excel path");
            return;
        }

        var saveRows = this.BuildLabelEditorSaveRowsForCurrentSchematic();

        Logger.Debug($"Label editor built save rows count: [{saveRows.Count}]");

        var saveResult = await BoardDataWriter.SaveLabelEditorChangesAsync(
            excelPath,
            schematicName,
            saveRows,
            region);

        if (!saveResult.Success)
        {
            Logger.Warning($"Label editor save failed - [{saveResult.ErrorMessage}]");
            await this.ShowLabelEditorSaveFailedDialogAsync(saveResult.ErrorMessage);
            return;
        }

        Logger.Info($"Label editor save succeeded for Excel path: [{excelPath}]");

        BoardDataReader.ClearCache(cacheKey);

        this.thisIsLabelEditorMode = false;
        this.thisLastCreatedLabelEditorCategory = string.Empty;
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlight = null;
        this.thisPendingNewLabelEditorHighlight = null;
        this.thisIsDrawingLabelEditorRectangle = false;
        this.thisLabelEditorDraftRectangle = null;
        this.thisLabelEditorDragMode = LabelEditorDragMode.None;
        this.thisLabelEditorOriginalSelectionBounds = default;
        this.thisLabelEditorOriginalDragRectangles.Clear();
        this.thisLabelEditorWorkingHighlights.Clear();
        this.thisLabelEditorUndoStack.Clear();
        this.thisLabelEditorRedoStack.Clear();

        this.UpdateLabelEditorLockState();
        this.HideLabelEditorMenu();
        this.HideNewLabelEditorPrompt();
        this.RefreshLabelEditorOverlay();
        this.SchematicsContainer.Focus();

        if (this.MainWindow != null)
        {
            await this.MainWindow.DisableLaunchDataSyncAfterLocalBoardEditAsync();
            this.MainWindow.ReloadCurrentBoardFromDisk(schematicName);
        }

        Logger.Info($"Label editor changes saved and reload requested for schematic [{schematicName}]");
    }

    // ###########################################################################################
    // Refreshes the label-editor overlay from the current working-copy rectangles and selected item.
    // Uses the same main highlight color and opacity as the normal schematic highlight.
    // Applies overlay state in one batch to avoid repeated redraws during drag operations.
    // ###########################################################################################
    private void RefreshLabelEditorOverlay(IReadOnlyList<(Point Start, Point End)>? snapGuides = null)
    {
        if (!this.thisIsLabelEditorMode || this.currentFullResBitmap == null)
        {
            this.SchematicsLabelEditorOverlay.ApplyState(
                rectangles: Array.Empty<Rect>(),
                selectedIndex: -1,
                selectedIndices: Array.Empty<int>(),
                selectionBounds: null,
                hoveredIndex: -1,
                draftRectangle: null,
                snapGuides: Array.Empty<(Point Start, Point End)>(),
                bitmapPixelSize: this.currentFullResBitmap?.PixelSize ?? new PixelSize(0, 0),
                viewMatrix: this.schematicsMatrix,
                highlightColor: Colors.IndianRed,
                highlightOpacity: 0.20,
                isVisible: false);

            this.UpdateComponentLabels();
            return;
        }

        string schematicName = this.GetCurrentSchematicName();

        var itemsForCurrentSchematic = this.thisLabelEditorWorkingHighlights
            .Where(row => string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var rects = itemsForCurrentSchematic
            .Select(row => new Rect(row.X, row.Y, row.Width, row.Height))
            .ToList();

        int selectedIndex = -1;
        if (this.thisSelectedLabelEditorHighlight != null)
        {
            selectedIndex = itemsForCurrentSchematic.IndexOf(this.thisSelectedLabelEditorHighlight);
        }

        var selectedIndices = itemsForCurrentSchematic
            .Select((row, index) => new { row, index })
            .Where(x => this.thisSelectedLabelEditorHighlights.Contains(x.row))
            .Select(x => x.index)
            .ToList();

        Color highlightColor = Colors.IndianRed;
        double highlightOpacity = 0.20;

        if (!string.IsNullOrWhiteSpace(schematicName) &&
            this.schematicByName.TryGetValue(schematicName, out var schematic))
        {
            highlightColor = RectGeometry.ParseColorOrDefault(schematic.SchematicHighlightColor, Colors.IndianRed);
            highlightOpacity = RectGeometry.ParseOpacityOrDefault(schematic.SchematicHighlightOpacity, 0.20);
        }

        this.SchematicsLabelEditorOverlay.ApplyState(
            rectangles: rects,
            selectedIndex: selectedIndex,
            selectedIndices: selectedIndices,
            selectionBounds: null,
            hoveredIndex: -1,
            draftRectangle: this.thisLabelEditorDraftRectangle,
            snapGuides: snapGuides ?? Array.Empty<(Point Start, Point End)>(),
            bitmapPixelSize: this.currentFullResBitmap.PixelSize,
            viewMatrix: this.schematicsMatrix,
            highlightColor: highlightColor,
            highlightOpacity: highlightOpacity,
            isVisible: true);

        this.UpdateComponentLabels();
    }

    // ###########################################################################################
    // Shows the prompt used for entering the board label and category of a newly drawn rectangle.
    // Reuses the last confirmed new-component category when available.
    // ###########################################################################################
    private void ShowNewLabelEditorPrompt(Point containerPoint)
    {
        double estimatedWidth = 280.0;
        double estimatedHeight = 170.0;

        double x = Math.Clamp(containerPoint.X, 6.0, Math.Max(6.0, this.SchematicsContainer.Bounds.Width - estimatedWidth));
        double y = Math.Clamp(containerPoint.Y, 6.0, Math.Max(6.0, this.SchematicsContainer.Bounds.Height - estimatedHeight));

        var categories = this.GetAvailableLabelEditorCategories();

        string preferredCategory = categories
            .FirstOrDefault(category => string.Equals(
                category,
                this.thisLastCreatedLabelEditorCategory,
                StringComparison.OrdinalIgnoreCase))
            ?? this.MainWindow?.CategoryFilterListBox.SelectedItems?
                .Cast<string>()
                .FirstOrDefault(category => !string.IsNullOrWhiteSpace(category))
            ?? categories.FirstOrDefault()
            ?? "General";

        this.NewLabelEditorBoardLabelTextBox.Text = string.Empty;
        this.NewLabelEditorCategoryComboBox.ItemsSource = categories;
        this.NewLabelEditorCategoryComboBox.SelectedItem = preferredCategory;

        this.SchematicsNewLabelPromptBorder.Margin = new Thickness(x, y, 0, 0);
        this.SchematicsNewLabelPromptBorder.IsVisible = true;

        Dispatcher.UIThread.Post(() => this.NewLabelEditorBoardLabelTextBox.Focus(), DispatcherPriority.Background);
    }

    // ###########################################################################################
    // Hides the prompt used for entering the board label of a newly drawn rectangle.
    // ###########################################################################################
    private void HideNewLabelEditorPrompt()
    {
        this.SchematicsNewLabelPromptBorder.IsVisible = false;
        this.NewLabelEditorBoardLabelTextBox.Text = string.Empty;
        this.NewLabelEditorCategoryComboBox.ItemsSource = null;
        this.NewLabelEditorCategoryComboBox.SelectedItem = null;
    }

    // ###########################################################################################
    // Commits the entered board label and category onto the newly created rectangle and keeps it selected.
    // Remembers the chosen category so the next new component defaults to the same category.
    // ###########################################################################################
    private void ConfirmNewLabelEditorPrompt()
    {
        if (this.thisPendingNewLabelEditorHighlight == null)
        {
            this.HideNewLabelEditorPrompt();
            this.SchematicsContainer.Focus();
            return;
        }

        string boardLabel = this.NewLabelEditorBoardLabelTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(boardLabel))
        {
            this.NewLabelEditorBoardLabelTextBox.Focus();
            return;
        }

        string category = this.NewLabelEditorCategoryComboBox.SelectedItem as string ?? string.Empty;
        if (string.IsNullOrWhiteSpace(category))
        {
            this.NewLabelEditorCategoryComboBox.Focus();
            return;
        }

        this.thisPendingNewLabelEditorHighlight.BoardLabel = boardLabel;
        this.thisPendingNewLabelEditorHighlight.Category = category;
        this.thisLastCreatedLabelEditorCategory = category;
        this.thisPendingNewLabelEditorHighlight = null;

        this.HideNewLabelEditorPrompt();
        this.RefreshLabelEditorOverlay();
        this.SchematicsContainer.Focus();
    }

    // ###########################################################################################
    // Cancels the newly created rectangle naming prompt and removes the pending rectangle.
    // ###########################################################################################
    private void CancelNewLabelEditorPrompt()
    {
        if (this.thisPendingNewLabelEditorHighlight != null)
        {
            this.thisLabelEditorWorkingHighlights.Remove(this.thisPendingNewLabelEditorHighlight);
            this.thisSelectedLabelEditorHighlights.Remove(this.thisPendingNewLabelEditorHighlight);
            this.thisPendingNewLabelEditorHighlight = null;
            this.thisSelectedLabelEditorHighlight = null;
        }

        this.thisLabelEditorDraftRectangle = null;
        this.thisIsDrawingLabelEditorRectangle = false;

        this.HideNewLabelEditorPrompt();
        this.RefreshLabelEditorOverlay();
        this.SchematicsContainer.Focus();
    }

    // ###########################################################################################
    // Builds the available category list for the new-label prompt and ensures at least one item
    // exists so a newly created component can always be categorized immediately.
    // ###########################################################################################
    private List<string> GetAvailableLabelEditorCategories()
    {
        var boardData = this.MainWindow?.CurrentBoardData;

        var categories = boardData?.Components
            .Select(component => component.Category?.Trim() ?? string.Empty)
            .Where(category => !string.IsNullOrWhiteSpace(category))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(category => category, StringComparer.OrdinalIgnoreCase)
            .ToList()
            ?? new List<string>();

        if (categories.Count == 0)
        {
            categories.Add("General");
        }

        return categories;
    }

    // ###########################################################################################
    // Handles Enter/Escape while the new-label prompt is open so the prompt can be confirmed
    // directly from the keyboard without clicking the buttons.
    // ###########################################################################################
    private void OnNewLabelEditorPromptKeyDown(object? sender, KeyEventArgs e)
    {
        if (!this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            return;
        }

        if (e.Key == Key.Enter)
        {
            this.ConfirmNewLabelEditorPrompt();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Escape)
        {
            this.CancelNewLabelEditorPrompt();
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Validates the current schematic editor rows before anything is written to disk.
    // ###########################################################################################
    private bool TryValidateLabelEditorSave(out string validationError)
    {
        validationError = string.Empty;

        string schematicName = this.GetCurrentSchematicName();
        if (string.IsNullOrWhiteSpace(schematicName))
        {
            validationError = "No schematic is currently selected";
            return false;
        }

        var rows = this.thisLabelEditorWorkingHighlights
            .Where(row => string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.BoardLabel))
            {
                validationError = "One or more rectangles are missing a board label";
                return false;
            }

            if (string.IsNullOrWhiteSpace(row.Category))
            {
                validationError = $"The label [{row.BoardLabel}] is missing a category";
                return false;
            }

            if (row.Width <= 0 || row.Height <= 0)
            {
                validationError = $"The label [{row.BoardLabel}] has invalid rectangle dimensions";
                return false;
            }
        }

        var conflictingCategoryGroup = rows
            .Where(row => !string.IsNullOrWhiteSpace(row.BoardLabel))
            .GroupBy(row => row.BoardLabel.Trim(), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group
                .Select(row => row.Category?.Trim() ?? string.Empty)
                .Where(category => !string.IsNullOrWhiteSpace(category))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count() > 1);

        if (conflictingCategoryGroup != null)
        {
            validationError = $"The board label [{conflictingCategoryGroup.Key}] has conflicting categories in the editor";
            return false;
        }

        return true;
    }

    // ###########################################################################################
    // Converts the current schematic editor rows into the workbook save DTO used by the writer.
    // Newly created component rows should be inserted with an empty Region value so they become
    // shared/global entries instead of inheriting the currently selected PAL or NTSC region.
    // ###########################################################################################
    private List<LabelEditorSaveRow> BuildLabelEditorSaveRowsForCurrentSchematic()
    {
        string schematicName = this.GetCurrentSchematicName();

        return this.thisLabelEditorWorkingHighlights
            .Where(row => string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            .Select(row => new LabelEditorSaveRow
            {
                SchematicName = row.SchematicName.Trim(),
                BoardLabel = row.BoardLabel.Trim(),
                Category = row.Category.Trim(),
                Region = string.Empty,
                X = row.X,
                Y = row.Y,
                Width = row.Width,
                Height = row.Height
            })
            .ToList();
    }

    // ###########################################################################################
    // Shows a modal error dialog when label-editor changes cannot be written to the Excel file
    // so save failures are visible immediately instead of only appearing in the logfile.
    // ###########################################################################################
    private Task ShowLabelEditorSaveFailedDialogAsync(string errorMessage) =>
        ShowDialogAsync(this.BuildLabelEditorSaveFailedDialog(errorMessage));

    private Window BuildLabelEditorSaveFailedDialog(string errorMessage) =>
        this.BuildLabelEditorErrorDialog(
            windowTitle: "Label editor save failed",
            headline: "Unable to save label editor changes",
            errorMessage: errorMessage,
            fallbackMessage: "The label editor changes could not be saved.",
            hint: "If the Excel workbook is open in another program, close it and try again. Check the logfile for technical details.");

    // ###########################################################################################
    // Shows a modal error dialog when the label editor contains invalid data so validation
    // problems are visible immediately instead of only appearing in the logfile.
    // ###########################################################################################
    private Task ShowLabelEditorValidationFailedDialogAsync(string errorMessage) =>
        ShowDialogAsync(this.BuildLabelEditorValidationFailedDialog(errorMessage));

    private Window BuildLabelEditorValidationFailedDialog(string errorMessage) =>
        this.BuildLabelEditorErrorDialog(
            windowTitle: "Label editor validation failed",
            headline: "Unable to apply component label editor changes",
            errorMessage: errorMessage,
            fallbackMessage: "The label editor contains invalid data.",
            hint: null);

    // ###########################################################################################
    // Shows an already-built dialog, modally when there is a window to own it.
    //
    // The fallback matters: TopLevel.GetTopLevel returns null when this control is not attached to
    // a window, and ShowDialog with no owner throws. A non-modal Show is a degraded but working
    // outcome - the user still sees the error - which beats an exception raised while reporting
    // one.
    // ###########################################################################################
    private async Task ShowDialogAsync(Window dialog)
    {
        if (TopLevel.GetTopLevel(this) is Window owner)
        {
            await dialog.ShowDialog(owner);
        }
        else
        {
            dialog.Show();
        }
    }

    // ###########################################################################################
    // Builds the label editor's error dialog, without showing it.
    //
    // The save and validation dialogs were written out separately and were identical for 48
    // straight lines - same sizing, the same four theme brush lookups, the same red banner and the
    // same nested layout - differing only in their title, headline, fallback text and whether a
    // trailing hint paragraph is shown. Those four are now the parameters, so the two dialogs
    // cannot drift apart visually.
    //
    // "errorMessage" is what actually went wrong; "fallbackMessage" stands in when it arrives
    // blank, so the dialog never opens with an empty body. "hint" is null for a dialog that has
    // nothing further to add, and that paragraph is then omitted entirely rather than emitted
    // blank - an empty TextBlock still takes a line's worth of height.
    //
    // Kept separate from ShowDialogAsync above purely so it can be tested: ShowDialog blocks on a
    // real window with nothing to dismiss it in a headless run, so a test that called the shipped
    // entry points directly would hang rather than fail. The tests assert on the returned Window -
    // the SAME one the app displays, since the caller does nothing to it but show it.
    // ###########################################################################################
    private Window BuildLabelEditorErrorDialog(
        string windowTitle,
        string headline,
        string errorMessage,
        string fallbackMessage,
        string? hint)
    {
        string message = string.IsNullOrWhiteSpace(errorMessage)
            ? fallbackMessage
            : errorMessage;

        var closeButton = new Button
        {
            Content = "Close",
            MinWidth = 110,
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
            HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center
        };

        var dialog = new Window
        {
            Title = windowTitle,
            Width = 540,
            MinWidth = 460,
            CanResize = false,
            ShowInTaskbar = false,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };

        closeButton.Click += (_, _) => dialog.Close();

        var errorAccentBrush = new SolidColorBrush(Color.Parse("#C62828"));
        var panelBackgroundBrush = this.ResolveThemeBrush("Schematics_Panels_Bg", new SolidColorBrush(Color.Parse("#FFF8F8")));
        var panelBorderBrush = this.ResolveThemeBrush("Schematics_Panels_Border", new SolidColorBrush(Color.Parse("#E0B4B4")));
        var foregroundBrush = this.ResolveThemeBrush("Schematics_Panels_Fg", Brushes.Black);

        var contentPanel = new StackPanel
        {
            Spacing = 14,
            Children =
            {
                new Border
                {
                    Background = errorAccentBrush,
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(12, 10),
                    Child = new StackPanel
                    {
                        Orientation = Avalonia.Layout.Orientation.Horizontal,
                        Spacing = 10,
                        Children =
                        {
                            new TextBlock
                            {
                                Text = "⚠",
                                FontSize = 22,
                                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                                Foreground = Brushes.White
                            },
                            new TextBlock
                            {
                                Text = headline,
                                FontSize = 14,
                                FontWeight = FontWeight.Bold,
                                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                                Foreground = Brushes.White,
                                TextWrapping = TextWrapping.Wrap
                            }
                        }
                    }
                },
                new TextBlock
                {
                    Text = message,
                    Foreground = foregroundBrush,
                    TextWrapping = TextWrapping.Wrap
                }
            }
        };

        if (!string.IsNullOrWhiteSpace(hint))
        {
            contentPanel.Children.Add(new TextBlock
            {
                Text = hint,
                Foreground = foregroundBrush,
                Opacity = 0.85,
                TextWrapping = TextWrapping.Wrap
            });
        }

        contentPanel.Children.Add(closeButton);

        dialog.Content = new Border
        {
            Background = panelBackgroundBrush,
            BorderBrush = panelBorderBrush,
            BorderThickness = new Thickness(1),
            Padding = new Thickness(18),
            Child = contentPanel
        };

        return dialog;
    }

    // ###########################################################################################
    // Captures the current label-editor working state so it can be restored by undo or redo.
    // ###########################################################################################
    private LabelEditorUndoState CreateLabelEditorUndoState()
    {
        var state = new LabelEditorUndoState
        {
            PrimarySelectedIndex = this.thisSelectedLabelEditorHighlight != null
                ? this.thisLabelEditorWorkingHighlights.IndexOf(this.thisSelectedLabelEditorHighlight)
                : -1
        };

        foreach (var row in this.thisLabelEditorWorkingHighlights)
        {
            state.Highlights.Add(new LabelEditorUndoHighlightState
            {
                SchematicName = row.SchematicName,
                BoardLabel = row.BoardLabel,
                Category = row.Category,
                X = row.X,
                Y = row.Y,
                Width = row.Width,
                Height = row.Height,
                IsSelected = this.thisSelectedLabelEditorHighlights.Contains(row)
            });
        }

        return state;
    }

    // ###########################################################################################
    // Captures the label-editor state as it existed before the active drag started.
    // ###########################################################################################
    private LabelEditorUndoState CreateLabelEditorUndoStateFromOriginalDragState()
    {
        var state = new LabelEditorUndoState
        {
            PrimarySelectedIndex = this.thisSelectedLabelEditorHighlight != null
                ? this.thisLabelEditorWorkingHighlights.IndexOf(this.thisSelectedLabelEditorHighlight)
                : -1
        };

        foreach (var row in this.thisLabelEditorWorkingHighlights)
        {
            Rect rect = this.thisLabelEditorOriginalDragRectangles.TryGetValue(row, out var originalRect)
                ? originalRect
                : new Rect(row.X, row.Y, row.Width, row.Height);

            state.Highlights.Add(new LabelEditorUndoHighlightState
            {
                SchematicName = row.SchematicName,
                BoardLabel = row.BoardLabel,
                Category = row.Category,
                X = rect.X,
                Y = rect.Y,
                Width = rect.Width,
                Height = rect.Height,
                IsSelected = this.thisSelectedLabelEditorHighlights.Contains(row)
            });
        }

        return state;
    }

    // ###########################################################################################
    // Compares two label-editor snapshots so duplicate undo and redo entries can be skipped.
    // ###########################################################################################
    private static bool AreLabelEditorUndoStatesEqual(LabelEditorUndoState leftState, LabelEditorUndoState rightState)
    {
        const double epsilon = 0.0001;

        if (ReferenceEquals(leftState, rightState))
        {
            return true;
        }

        if (leftState.Highlights.Count != rightState.Highlights.Count ||
            leftState.PrimarySelectedIndex != rightState.PrimarySelectedIndex)
        {
            return false;
        }

        for (int i = 0; i < leftState.Highlights.Count; i++)
        {
            var left = leftState.Highlights[i];
            var right = rightState.Highlights[i];

            if (!string.Equals(left.SchematicName, right.SchematicName, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(left.BoardLabel, right.BoardLabel, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(left.Category, right.Category, StringComparison.OrdinalIgnoreCase) ||
                Math.Abs(left.X - right.X) > epsilon ||
                Math.Abs(left.Y - right.Y) > epsilon ||
                Math.Abs(left.Width - right.Width) > epsilon ||
                Math.Abs(left.Height - right.Height) > epsilon ||
                left.IsSelected != right.IsSelected)
            {
                return false;
            }
        }

        return true;
    }

    // ###########################################################################################
    // Pushes one undo snapshot and clears redo because a new forward edit has been committed.
    // ###########################################################################################
    private void PushLabelEditorUndoState(LabelEditorUndoState state)
    {
        if (this.thisLabelEditorUndoStack.Count > 0 &&
            AreLabelEditorUndoStatesEqual(this.thisLabelEditorUndoStack.Peek(), state))
        {
            return;
        }

        this.thisLabelEditorUndoStack.Push(state);
        this.thisLabelEditorRedoStack.Clear();
    }

    // ###########################################################################################
    // Restores one label-editor snapshot back into the working state and refreshes the overlay.
    // ###########################################################################################
    private void RestoreLabelEditorUndoState(LabelEditorUndoState state)
    {
        this.thisLabelEditorWorkingHighlights.Clear();
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlight = null;
        this.thisPendingNewLabelEditorHighlight = null;
        this.thisIsDrawingLabelEditorRectangle = false;
        this.thisLabelEditorDraftRectangle = null;
        this.thisLabelEditorDragMode = LabelEditorDragMode.None;
        this.thisLabelEditorOriginalSelectionBounds = default;
        this.thisLabelEditorOriginalDragRectangles.Clear();

        this.HideNewLabelEditorPrompt();

        foreach (var snapshot in state.Highlights)
        {
            var row = new EditableComponentHighlight
            {
                SchematicName = snapshot.SchematicName,
                BoardLabel = snapshot.BoardLabel,
                Category = snapshot.Category,
                X = snapshot.X,
                Y = snapshot.Y,
                Width = snapshot.Width,
                Height = snapshot.Height
            };

            this.thisLabelEditorWorkingHighlights.Add(row);

            if (snapshot.IsSelected)
            {
                this.thisSelectedLabelEditorHighlights.Add(row);
            }
        }

        if (state.PrimarySelectedIndex >= 0 &&
            state.PrimarySelectedIndex < this.thisLabelEditorWorkingHighlights.Count)
        {
            var primary = this.thisLabelEditorWorkingHighlights[state.PrimarySelectedIndex];
            if (this.thisSelectedLabelEditorHighlights.Contains(primary))
            {
                this.thisSelectedLabelEditorHighlight = primary;
            }
        }

        if (this.thisSelectedLabelEditorHighlight == null)
        {
            this.thisSelectedLabelEditorHighlight = this.GetFirstSelectedLabelEditorHighlightForCurrentSchematic();
        }

        this.RefreshLabelEditorOverlay();
        this.SchematicsContainer.Focus();
    }

    // ###########################################################################################
    // Restores the previous label-editor snapshot and moves the current state onto the redo stack.
    // ###########################################################################################
    private bool TryUndoLabelEditorChange()
    {
        if (!this.thisIsLabelEditorMode || this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            return false;
        }

        var currentState = this.CreateLabelEditorUndoState();

        while (this.thisLabelEditorUndoStack.Count > 0 &&
               AreLabelEditorUndoStatesEqual(this.thisLabelEditorUndoStack.Peek(), currentState))
        {
            this.thisLabelEditorUndoStack.Pop();
        }

        if (this.thisLabelEditorUndoStack.Count == 0)
        {
            return false;
        }

        var previousState = this.thisLabelEditorUndoStack.Pop();

        if (this.thisLabelEditorRedoStack.Count == 0 ||
            !AreLabelEditorUndoStatesEqual(this.thisLabelEditorRedoStack.Peek(), currentState))
        {
            this.thisLabelEditorRedoStack.Push(currentState);
        }

        this.RestoreLabelEditorUndoState(previousState);
        return true;
    }

    // ###########################################################################################
    // Restores the next label-editor snapshot and moves the current state back onto the undo stack.
    // ###########################################################################################
    private bool TryRedoLabelEditorChange()
    {
        if (!this.thisIsLabelEditorMode || this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            return false;
        }

        var currentState = this.CreateLabelEditorUndoState();

        while (this.thisLabelEditorRedoStack.Count > 0 &&
               AreLabelEditorUndoStatesEqual(this.thisLabelEditorRedoStack.Peek(), currentState))
        {
            this.thisLabelEditorRedoStack.Pop();
        }

        if (this.thisLabelEditorRedoStack.Count == 0)
        {
            return false;
        }

        var nextState = this.thisLabelEditorRedoStack.Pop();

        if (this.thisLabelEditorUndoStack.Count == 0 ||
            !AreLabelEditorUndoStatesEqual(this.thisLabelEditorUndoStack.Peek(), currentState))
        {
            this.thisLabelEditorUndoStack.Push(currentState);
        }

        this.RestoreLabelEditorUndoState(nextState);
        return true;
    }

    // ###########################################################################################
    // Clears the cached editor label visual pool so stale controls are not retained when the
    // component label editor is inactive.
    // ###########################################################################################
    private void ResetEditorComponentLabelVisualCache()
    {
        this.thisEditorLabelContainers.Clear();
        this.thisEditorLabelTextBlocks.Clear();
        this.thisEditorLabelScaleTransforms.Clear();
        this.thisLastEditorLabelVisualSignature = string.Empty;
    }

    // ###########################################################################################
    // Ensures the reusable editor label visual pool contains at least the requested number of
    // controls. Labels are created once and then reused during drag/resize operations.
    // ###########################################################################################
    private void EnsureEditorComponentLabelVisualPoolSize(int requiredCount)
    {
        while (this.thisEditorLabelContainers.Count < requiredCount)
        {
            var textBlock = new TextBlock
            {
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                TextAlignment = TextAlignment.Center
            };
            textBlock.Bind(TextBlock.ForegroundProperty, this.GetResourceObservable("Schematics_ComponentLabel_Fg"));

            var innerBorder = new Border
            {
                BorderThickness = new Avalonia.Thickness(1),
                CornerRadius = new Avalonia.CornerRadius(4),
                Padding = new Avalonia.Thickness(6, 4),
                Child = textBlock
            };
            innerBorder.Bind(Border.BackgroundProperty, this.GetResourceObservable("Schematics_ComponentLabel_Bg"));
            innerBorder.Bind(Border.BorderBrushProperty, this.GetResourceObservable("Schematics_ComponentLabel_Border"));

            var scaleTransform = new ScaleTransform(1.0, 1.0);
            var transformGroup = new TransformGroup();
            transformGroup.Children.Add(scaleTransform);

            var container = new Border
            {
                IsHitTestVisible = false,
                IsVisible = false,
                Child = innerBorder,
                RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative),
                RenderTransform = transformGroup
            };

            this.thisEditorLabelContainers.Add(container);
            this.thisEditorLabelTextBlocks.Add(textBlock);
            this.thisEditorLabelScaleTransforms.Add(scaleTransform);
            this.SchematicsLabelsCanvas.Children.Add(container);
        }
    }

    // ###########################################################################################
    // Builds a lightweight signature for the currently visible editor labels so controls are only
    // re-bound when the count or label text actually changes.
    // ###########################################################################################
    private static string BuildEditorComponentLabelVisualSignature(IReadOnlyList<EditableComponentHighlight> rows)
    {
        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var parts = new string[rows.Count];

        for (int i = 0; i < rows.Count; i++)
        {
            parts[i] = rows[i].BoardLabel?.Trim() ?? string.Empty;
        }

        return string.Join("\u001F", parts);
    }

    // ###########################################################################################
    // Updates the reusable editor label controls without clearing and rebuilding the entire canvas
    // on every pointer move. While label-editor search is active, only matching labels are shown.
    // ###########################################################################################
    private void UpdateEditorComponentLabels(
        IReadOnlyList<EditableComponentHighlight> rows,
        Rect contentRect,
        double imgWidth,
        double imgHeight,
        double inverseScale)
    {
        var visibleRows = string.IsNullOrWhiteSpace(this.thisLabelEditorSearchText)
            ? rows.ToList()
            : rows.Where(this.DoesLabelEditorHighlightMatchSearch).ToList();

        if (this.thisEditorLabelContainers.Count == 0 && this.SchematicsLabelsCanvas.Children.Count > 0)
        {
            this.SchematicsLabelsCanvas.Children.Clear();
        }

        this.EnsureEditorComponentLabelVisualPoolSize(visibleRows.Count);

        string newSignature = BuildEditorComponentLabelVisualSignature(visibleRows);
        bool textChanged = !string.Equals(
            this.thisLastEditorLabelVisualSignature,
            newSignature,
            StringComparison.Ordinal);

        if (textChanged)
        {
            for (int i = 0; i < visibleRows.Count; i++)
            {
                this.thisEditorLabelTextBlocks[i].Text = visibleRows[i].BoardLabel;
            }

            this.thisLastEditorLabelVisualSignature = newSignature;
        }

        for (int i = 0; i < visibleRows.Count; i++)
        {
            var row = visibleRows[i];
            var container = this.thisEditorLabelContainers[i];
            var scaleTransform = this.thisEditorLabelScaleTransforms[i];

            double centerX = row.X + (row.Width / 2.0);
            double centerY = row.Y + (row.Height / 2.0);

            double localX = contentRect.X + (centerX / imgWidth) * contentRect.Width;
            double localY = contentRect.Y + (centerY / imgHeight) * contentRect.Height;

            scaleTransform.ScaleX = inverseScale;
            scaleTransform.ScaleY = inverseScale;

            bool needsMeasure =
                textChanged ||
                !container.IsVisible ||
                container.DesiredSize.Width <= 0 ||
                container.DesiredSize.Height <= 0;

            container.IsVisible = true;

            if (needsMeasure)
            {
                container.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            }

            Size desiredSize = container.DesiredSize;

            Canvas.SetLeft(container, localX - (desiredSize.Width / 2.0));
            Canvas.SetTop(container, localY - (desiredSize.Height / 2.0));
        }

        for (int i = visibleRows.Count; i < this.thisEditorLabelContainers.Count; i++)
        {
            this.thisEditorLabelContainers[i].IsVisible = false;
        }
    }

    // ###########################################################################################
    // Updates transient editor-overlay state such as hovered handles or snap guides without
    // rebuilding the entire editor model or touching the reusable label pool.
    // ###########################################################################################
    private void SetLabelEditorOverlayTransientState(
        int? hoveredIndex = null,
        IReadOnlyList<(Point Start, Point End)>? snapGuides = null)
    {
        this.SchematicsLabelEditorOverlay.ApplyState(
            rectangles: this.SchematicsLabelEditorOverlay.Rectangles,
            selectedIndex: this.SchematicsLabelEditorOverlay.SelectedIndex,
            selectedIndices: this.SchematicsLabelEditorOverlay.SelectedIndices,
            selectionBounds: this.SchematicsLabelEditorOverlay.SelectionBounds,
            hoveredIndex: hoveredIndex ?? this.SchematicsLabelEditorOverlay.HoveredIndex,
            draftRectangle: this.SchematicsLabelEditorOverlay.DraftRectangle,
            snapGuides: snapGuides ?? this.SchematicsLabelEditorOverlay.SnapGuides,
            bitmapPixelSize: this.SchematicsLabelEditorOverlay.BitmapPixelSize,
            viewMatrix: this.SchematicsLabelEditorOverlay.ViewMatrix,
            highlightColor: this.SchematicsLabelEditorOverlay.HighlightColor,
            highlightOpacity: this.SchematicsLabelEditorOverlay.HighlightOpacity,
            isVisible: this.SchematicsLabelEditorOverlay.IsVisible);
    }

    // ###########################################################################################
    // Enables or disables navigation controls that would invalidate the current label-editor session.
    // While the editor is active, schematic thumbnails and the main hardware/board selectors are locked.
    // Also updates thumbnail relevance dimming, the main-window special-mode banner, and the
    // shared component search box mode.
    // ###########################################################################################
    private void UpdateLabelEditorLockState()
    {
        bool isNavigationEnabled = !this.thisIsLabelEditorMode;

        if (!isNavigationEnabled)
        {
            this.thisIsDraggingThumbnail = false;
            this.thisThumbnailDragStartEventArgs = null;
            this.ClearThumbnailDropPlaceholder();
            this.HideThumbnailDragGhost();
        }

        this.SchematicsThumbnailList.IsEnabled = isNavigationEnabled;
        this.MainWindow?.SetSchematicsEditorNavigationEnabled(isNavigationEnabled);
        this.MainWindow?.SetSchematicsLabelEditorModeBannerVisible(this.thisIsLabelEditorMode);
        this.MainWindow?.UpdateComponentSearchTextBoxMode();

        if (this.MainWindow?.CurrentBoardData != null)
        {
            this.MainWindow.OnComponentSearchTextChanged(null, null!);
        }

        bool hasComponentSelection = this.highlightIndexBySchematic.Count > 0;
        bool hasBlinkEligibleSelection = this.HasBlinkEligibleSelection();
        double blinkFactor = this.MainWindow?.GetCurrentBlinkFactor(hasBlinkEligibleSelection) ?? 1.0;

        this.ApplyHighlightVisuals(hasComponentSelection, blinkFactor);
    }

    // ###########################################################################################
    // Applies search text to the label editor so the search box can find editor components by
    // board label or category while the editor is active.
    // ###########################################################################################
    internal void ApplyLabelEditorSearchFilter(string searchTerm)
    {
        this.thisLabelEditorSearchText = searchTerm?.Trim() ?? string.Empty;

        if (!this.thisIsLabelEditorMode)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(this.thisLabelEditorSearchText))
        {
            this.RefreshLabelEditorOverlay();
            return;
        }

        var matches = this.GetMatchingLabelEditorHighlightsForCurrentSchematic();
        if (matches.Count == 0)
        {
            this.RefreshLabelEditorOverlay();
            return;
        }

        if (this.thisSelectedLabelEditorHighlight == null ||
            !matches.Contains(this.thisSelectedLabelEditorHighlight))
        {
            this.SetSingleSelectedLabelEditorHighlight(matches[0], refresh: false);
        }

        this.RefreshLabelEditorOverlay();
    }

    // ###########################################################################################
    // Returns the current-schematic label editor highlights that match the active search text.
    // ###########################################################################################
    private List<EditableComponentHighlight> GetMatchingLabelEditorHighlightsForCurrentSchematic()
    {
        string schematicName = this.GetCurrentSchematicName();

        return this.thisLabelEditorWorkingHighlights
            .Where(row => string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            .Where(this.DoesLabelEditorHighlightMatchSearch)
            .ToList();
    }

    // ###########################################################################################
    // Returns true when the given label editor highlight matches the active search text.
    // Matches board label and category using the same multi-term AND behavior as normal search.
    // ###########################################################################################
    private bool DoesLabelEditorHighlightMatchSearch(EditableComponentHighlight highlight)
    {
        if (string.IsNullOrWhiteSpace(this.thisLabelEditorSearchText))
        {
            return true;
        }

        var searchTerms = this.thisLabelEditorSearchText
            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        if (searchTerms.Length == 0)
        {
            return true;
        }

        string searchableText = string.Join(
            " | ",
            new[]
            {
                highlight.BoardLabel?.Trim() ?? string.Empty,
                highlight.Category?.Trim() ?? string.Empty
            }.Where(value => !string.IsNullOrWhiteSpace(value)));

        foreach (string term in searchTerms)
        {
            if (searchableText.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }
        }

        return true;
    }
}