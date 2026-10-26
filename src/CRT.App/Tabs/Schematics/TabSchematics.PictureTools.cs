using Avalonia;
using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Linq;

namespace CRT;

// ###########################################################################################
// THE PICTURE TOOLS AND THE COMPONENT MENU (CRT 4.0 step 2, slice 6 - the plan: "Contributor mode
// goes. Mark components and Align KiCad open the Schematics tab in that tool, with a bar and a Done
// button. Right-clicking a component offers Edit this component to everyone.").
//
// A tool is started on a PICTURE from the Contribute tab (StartPictureTool, through Main). The
// picture's image and the board's KiCad data are read in the background, so the request waits until
// what the tool works on is there - PictureToolStart decides - and what can never come is refused
// with the reason. While a tool is on, the bar over the picture says what it is on, how it is used,
// and offers Done (the label editor's or the calibration's own save into the draft) and Cancel.
// The bar FOLLOWS the tool (FollowPictureTool, called wherever a tool starts or ends), so Esc, a
// save and a board change all leave it right; when a tool started from the Contribute tab ends,
// PictureToolEnded tells Main, which goes back there. A board load ends everything silently.
//
// The floating menu that was contributor mode's (over an empty spot: "Enable component label
// editor", "Calibrate KiCad traces") is now the menu a right-click on a COMPONENT opens for everyone:
// highlight it or take the highlight away - what that right-click did by itself before - and "Edit
// this component", which opens it in the Contribute tab's pane (EditComponentRequested).
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private bool thisIsShowingSchematicsMenu;

    // The component the menu is open on.
    private string thisMenuBoardLabel = string.Empty;

    // A tool asked for on a picture, waiting for it - with the board it was asked on.
    private (PictureTool Tool, string Picture, string BoardKey)? thisPendingTool;

    // A tool that cannot start on its picture, and why - shown on the bar until closed.
    private (PictureTool Tool, string Picture, string Reason)? thisRefusedTool;

    // The bar is up for a tool the Contribute tab asked for: its end is told to Main.
    private bool thisToolFromWorkspace;

    // The picture whose image has been read, and whether it could be.
    private string? thisImageSettledFor;
    private bool thisImageMissing;

    // The board's KiCad data as last settled - null while not known yet (a board being loaded).
    private BoardKiCad? thisKiCadSettled;

    // A tool the Contribute tab started has ended - Done, Cancel or Close.
    internal event EventHandler? PictureToolEnded;

    // The label editor or the calibration went off, however it was started - Main reads the board
    // again then if that was asked for while it was on (code review, 2026-10-26).
    internal event EventHandler? ToolOff;

    private bool thisToolWasOn;

    // "Edit this component" on the component menu, with the component's label.
    internal event EventHandler<string>? EditComponentRequested;

    private void WirePictureTools()
    {
        this.PictureToolDoneButton.Click += (_, _) => this.OnPictureToolDone();
        this.PictureToolCancelButton.Click += (_, _) => this.OnPictureToolCancel();

        this.MenuHighlightButton.Click += (_, _) => this.OnMenuHighlightClick();
        this.MenuEditComponentButton.Click += (_, _) => this.OnMenuEditComponentClick();
        ToolTip.SetTip(this.MenuEditComponentButton, PictureToolWording.EditComponentTip);
    }

    // ###########################################################################################
    // A tool on a picture, asked for by the Contribute tab: the picture selected, then the tool
    // started as soon as what it works on is there - or refused, the reason on the bar. Whatever tool
    // is already on is left alone (the label editor holds the board and the picture).
    // ###########################################################################################
    internal void StartPictureTool(PictureTool tool, string picture)
    {
        if (this.thisIsLabelEditorMode || this.thisIsKiCadTraceCalibrationMode)
        {
            return;
        }

        this.CancelWorklogEntryMode();
        this.HideSchematicsMenu();

        this.thisRefusedTool = null;
        this.thisPendingTool = (tool, picture?.Trim() ?? string.Empty, this.MainWindow?.GetCurrentBoardKey() ?? string.Empty);
        this.thisToolFromWorkspace = true;

        this.SelectSchematicByName(picture ?? string.Empty);
        this.TryStartPendingTool();
    }

    // ###########################################################################################
    // The waiting tool started, refused or left waiting, as PictureToolStart decides on what the tab
    // has now. Asked whenever that may have changed: the request, a picture's image read, the board's
    // KiCad data settled.
    // ###########################################################################################
    private void TryStartPendingTool()
    {
        if (this.thisPendingTool is not { } pending)
        {
            this.FollowPictureTool();
            return;
        }

        string boardKey = this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;

        if (!string.Equals(boardKey, pending.BoardKey, StringComparison.OrdinalIgnoreCase))
        {
            // Another board came in meanwhile: the request was for a picture of the one before.
            this.thisPendingTool = null;
            this.FollowPictureTool();
            return;
        }

        bool onBoard = this.currentThumbnails.Any(thumbnail => string.Equals(thumbnail.Name, pending.Picture, StringComparison.OrdinalIgnoreCase));
        bool selected = string.Equals(this.GetCurrentSchematicName(), pending.Picture, StringComparison.OrdinalIgnoreCase);
        bool settled = selected && string.Equals(this.thisImageSettledFor, pending.Picture, StringComparison.OrdinalIgnoreCase);

        PictureImage image = !settled ? PictureImage.Loading : this.thisImageMissing ? PictureImage.Missing : PictureImage.Shown;
        BoardKiCad kiCad = this.thisKiCadSettled ?? BoardKiCad.Loading;
        bool hasView = kiCad == BoardKiCad.Loaded && selected && this.ResolveKiCadViewForCurrentSchematic() is not null;

        PictureToolStart decision = PictureToolStart.Decide(pending.Tool, pending.Picture, onBoard, selected, image, kiCad, hasView);

        switch (decision.Step)
        {
            case PictureToolStep.Wait:
                this.FollowPictureTool();
                return;

            case PictureToolStep.Refuse:
                this.thisPendingTool = null;
                this.thisRefusedTool = (pending.Tool, pending.Picture, decision.Reason);
                Logger.Info($"Picture tool [{pending.Tool}] not started on [{pending.Picture}] - [{decision.Reason}]");
                this.FollowPictureTool();
                return;

            default:
                this.thisPendingTool = null;

                if (pending.Tool == PictureTool.MarkComponents)
                    this.BeginLabelEditorMode();
                else
                    this.BeginKiCadTraceCalibrationMode();

                this.FollowPictureTool();
                return;
        }
    }

    // ###########################################################################################
    // The bar, as the tools are: a tool on - what it is on, how it is used, Cancel and Done; a request
    // waiting - "Opening", Cancel; a request refused - why, Close; nothing - hidden, and a tool the
    // Contribute tab started is told to have ended.
    // ###########################################################################################
    private void FollowPictureTool()
    {
        PictureTool? onNow = this.thisIsLabelEditorMode ? PictureTool.MarkComponents
            : this.thisIsKiCadTraceCalibrationMode ? PictureTool.AlignKiCad
            : null;

        if (onNow is { } tool)
        {
            this.thisToolWasOn = true;
            this.thisRefusedTool = null;
            this.ShowPictureToolBar(
                PictureToolWording.Title(tool, this.GetCurrentSchematicName()),
                PictureToolWording.Help(tool),
                PictureToolWording.Cancel,
                withDone: true);
            return;
        }

        if (this.thisToolWasOn)
        {
            this.thisToolWasOn = false;
            this.ToolOff?.Invoke(this, EventArgs.Empty);
        }

        if (this.thisPendingTool is { } pending)
        {
            this.ShowPictureToolBar(PictureToolWording.Opening(pending.Picture), string.Empty, PictureToolWording.Cancel, withDone: false);
            return;
        }

        if (this.thisRefusedTool is { } refused)
        {
            this.ShowPictureToolBar(PictureToolWording.CannotStart(refused.Tool, refused.Picture), refused.Reason, PictureToolWording.Close, withDone: false);
            return;
        }

        this.PictureToolBarBorder.IsVisible = false;

        if (this.thisToolFromWorkspace)
        {
            this.thisToolFromWorkspace = false;
            this.PictureToolEnded?.Invoke(this, EventArgs.Empty);
        }
    }

    private void ShowPictureToolBar(string title, string help, string cancel, bool withDone)
    {
        this.PictureToolTitleText.Text = title;
        this.PictureToolHelpText.Text = help;
        this.PictureToolHelpText.IsVisible = help.Length > 0;
        this.PictureToolCancelButton.Content = cancel;
        this.PictureToolDoneButton.IsVisible = withDone;
        this.PictureToolBarBorder.IsVisible = true;
    }

    // Done: the tool's own save into the draft - the bar follows once it has ended.
    private void OnPictureToolDone()
    {
        if (this.thisIsLabelEditorMode)
            this.ApplyLabelEditorChanges();
        else if (this.thisIsKiCadTraceCalibrationMode)
            this.ApplyKiCadTraceCalibration();
    }

    // Cancel, or Close: the tool's own cancel, or the request waiting or refused let go.
    private void OnPictureToolCancel()
    {
        if (this.thisIsLabelEditorMode)
        {
            this.CancelLabelEditorChanges();
            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            this.CancelKiCadTraceCalibrationMode();
            return;
        }

        this.thisPendingTool = null;
        this.thisRefusedTool = null;
        this.FollowPictureTool();
    }

    // ###########################################################################################
    // The picture's image has been read - shown, or not readable - and a tool waiting for it may start.
    // ###########################################################################################
    private void ImageSettled(string? picture, bool shown)
    {
        this.thisImageSettledFor = picture?.Trim();
        this.thisImageMissing = !shown;
        this.TryStartPendingTool();
    }

    // ###########################################################################################
    // The board's KiCad data has settled - read, or none - and an "Align KiCad" waiting for it may
    // start. Called at the end of LoadKiCadProjectForCurrentBoardAsync, and by Main for a board that
    // has no KiCad files at all (it does not ask for a load then).
    // ###########################################################################################
    internal void KiCadSettled(bool hasProject)
    {
        this.thisKiCadSettled = hasProject ? BoardKiCad.Loaded : BoardKiCad.None;
        this.TryStartPendingTool();
    }

    // A board load: whatever was asked or refused is for the board before, and its end is no
    // reason to leave this tab. The tools themselves end with the viewer's reset.
    private void ForgetPictureToolsForNewBoard()
    {
        this.thisPendingTool = null;
        this.thisRefusedTool = null;
        this.thisToolFromWorkspace = false;
        this.thisImageSettledFor = null;
        this.thisKiCadSettled = null;
        this.FollowPictureTool();
    }

    // ###########################################################################################
    // THE COMPONENT MENU: a right-click on a component - its name, its highlight on or off, and
    // "Edit this component". At the pointer, kept inside the view.
    // ###########################################################################################
    private void ShowComponentMenu(Point containerPoint, string boardLabel)
    {
        string label = boardLabel.Trim();

        if (label.Length == 0 || this.thisIsWorklogEntryMode)
        {
            return;
        }

        this.thisMenuBoardLabel = label;

        string friendlyName = this.CurrentBoardDataForLabelEditor?.Components
            .FirstOrDefault(component => string.Equals(component.BoardLabel?.Trim(), label, StringComparison.OrdinalIgnoreCase))
            ?.FriendlyName?.Trim() ?? string.Empty;

        this.SchematicsMenuTitleText.Text = friendlyName.Length == 0 ? label : $"{label} - {friendlyName}";
        this.MenuHighlightButton.Content = this.IsComponentBoardLabelSelected(label)
            ? PictureToolWording.TakeHighlightAway(label)
            : PictureToolWording.Highlight(label);
        this.MenuEditComponentButton.Content = PictureToolWording.EditComponent;

        const double estimatedWidth = 250.0;
        const double estimatedHeight = 120.0;

        double x = Math.Clamp(containerPoint.X, 6.0, Math.Max(6.0, this.SchematicsContainer.Bounds.Width - estimatedWidth));
        double y = Math.Clamp(containerPoint.Y, 6.0, Math.Max(6.0, this.SchematicsContainer.Bounds.Height - estimatedHeight));

        this.SchematicsMenuBorder.Margin = new Thickness(x, y, 0, 0);
        this.SchematicsMenuBorder.IsVisible = true;
        this.thisIsShowingSchematicsMenu = true;
    }

    private void HideSchematicsMenu()
    {
        this.SchematicsMenuBorder.IsVisible = false;
        this.thisIsShowingSchematicsMenu = false;
    }

    // Whether a press landed on the menu - it is the menu's, not the picture's.
    private bool IsPointerInsideSchematicsMenu(Point containerPoint)
    {
        if (!this.SchematicsMenuBorder.IsVisible)
        {
            return false;
        }

        Point? topLeft = this.SchematicsMenuBorder.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
        return topLeft.HasValue && new Rect(topLeft.Value, this.SchematicsMenuBorder.Bounds.Size).Contains(containerPoint);
    }

    // What a right-click on a component did by itself before the menu: its highlight on or off.
    private void OnMenuHighlightClick()
    {
        string label = this.thisMenuBoardLabel;
        this.HideSchematicsMenu();

        if (label.Length > 0)
        {
            this.ToggleComponentSelectionByBoardLabel(label);
        }
    }

    private void OnMenuEditComponentClick()
    {
        string label = this.thisMenuBoardLabel;
        this.HideSchematicsMenu();

        if (label.Length > 0)
        {
            this.EditComponentRequested?.Invoke(this, label);
        }
    }

    // ---- for tests ------------------------------------------------------------------------------------

    internal void ShowComponentMenuForTests(string boardLabel) => this.ShowComponentMenu(new Point(10, 10), boardLabel);

    internal void ImageSettledForTests(string picture, bool shown = true) => this.ImageSettled(picture, shown);

    internal bool IsComponentSelectedForTests(string boardLabel) => this.IsComponentBoardLabelSelected(boardLabel);

    internal bool IsCalibrationModeForTests => this.thisIsKiCadTraceCalibrationMode;
}
