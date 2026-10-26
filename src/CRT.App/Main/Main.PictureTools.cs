using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE PICTURE TOOLS BETWEEN THE TWO TABS (CRT 4.0 step 2, slice 6 - "Contributor mode goes").
    //
    // "Mark components" and "Align KiCad" on a picture of the Contribute tab's Pictures open the
    // Schematics tab in that tool (TabSchematics.PictureTools.cs, with its bar) - on the board as its
    // draft now is: read again first when the Contribute tab's saves left it behind, the picture
    // selected by that load (OnBoardShownUpToDate). When the tool ends - Done, Cancel or Close - CRT
    // goes back to the Contribute tab it came from.
    //
    // "Edit this component" on the menu a right-click on a component opens - for everyone - opens it
    // in the Contribute tab's pane.
    // ###########################################################################################
    public partial class Main
    {
        // The tool on the Schematics tab was started from the Contribute tab: its end goes back there.
        private bool thisPictureToolFromContribute;

        // The board read again was asked for while a tool was on: the schematic to restore, done
        // when the tool ends.
        private string? thisReloadWhenToolEnds;

        internal bool IsReloadWaitingForToolForTests => this.thisReloadWhenToolEnds is not null;

        private void WirePictureTools()
        {
            this.TabSchematicsControl.PictureToolEnded += (_, _) => this.OnPictureToolEnded();
            this.TabSchematicsControl.EditComponentRequested += (sender, label) => _ = this.EditComponentAsync(label);

            // After the tool's own end has run: the label editor's Done reads the board again itself,
            // which serves a reload that waited (ReloadCurrentBoardFromDisk), so nothing is read twice.
            this.TabSchematicsControl.ToolOff += (_, _) => Dispatcher.UIThread.Post(this.ReloadWhatWaitedForTool, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // *** A BOARD READ AGAIN WAITS WHILE A PICTURE TOOL IS ON (code review, 2026-10-26). *** A reload
        // resets the Schematics tab, and the label editor's rectangles not yet saved went with it -
        // asked for by the Contribute tab (a pane's change caught up, "Show on board", a tool asked for
        // on a board left behind) while the tool was still on. True when it waits.
        // ###########################################################################################
        private bool WaitsForTool(string schematicNameToRestore)
        {
            if (!this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.thisReloadWhenToolEnds = null;
                return false;
            }

            this.thisReloadWhenToolEnds = schematicNameToRestore ?? string.Empty;
            Logger.Info("The board is read again once the picture tool on the Schematics tab ends");
            return true;
        }

        private void ReloadWhatWaitedForTool()
        {
            if (this.thisReloadWhenToolEnds is { } schematic && !this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.ReloadCurrentBoardFromDisk(schematic);
            }
        }

        // ###########################################################################################
        // A tool on a picture, from the Contribute tab: the Schematics tab, the picture, the tool. With
        // a tool already on there, that one is shown - it holds the board and the picture until it ends.
        // ###########################################################################################
        internal void StartPictureTool(PictureTool tool, string picture, bool readBoardFirst)
        {
            if (this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.SwitchToSchematicsTab();
                return;
            }

            this.thisPictureToolFromContribute = true;
            this.OnBoardShownUpToDate(readBoardFirst, picture, () => this.TabSchematicsControl.StartPictureTool(tool, picture));
            this.SwitchToSchematicsTab();
        }

        private void OnPictureToolEnded()
        {
            if (!this.thisPictureToolFromContribute)
            {
                return;
            }

            this.thisPictureToolFromContribute = false;
            this.SwitchToContributeTab();
        }

        // ###########################################################################################
        // "Edit this component": the Contribute tab, the component open in its pane - the board's
        // draft made as it opens, as for every other way in, and let go of again when the pane closes
        // with nothing in it (TabDrafts.LetGoOfDraftMadeOnlyToLook).
        // ###########################################################################################
        internal async Task EditComponentAsync(string boardLabel)
        {
            try
            {
                this.SwitchToContributeTab();
                await this.TabDrafts.OpenComponentAsync(boardLabel);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Opening [{boardLabel}] on the Contribute tab failed - [{ex}]");
            }
        }
    }
}
