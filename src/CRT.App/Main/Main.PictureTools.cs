using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
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
    // goes back to the Contribute tab it came from; "Done - next picture" (step 3, slice 3d) goes on
    // to the next picture's alignment instead (AlignNext).
    //
    // "Edit component" on a component's popup - for everyone - opens it in the Contribute tab's pane
    // (owner request, 2026-11-05: on a menu a right-click opened until then).
    //
    // *** EVERY PICTURE ALIGNED WITH DONE IS MARKED ONCE THE TOOL HANDS BACK (owner request,
    // 2026-11-01: "it is not clear that I must first do the Align KiCad, before I can do the Mark all
    // automatically from KiCad" - Done marks without asking, Undo takes it back). *** Each alignment
    // saved with Done or "Done - next picture" is remembered with its board (AlignedFromWorkspace); when
    // the tool ends and the Contribute tab is back, each of them with no mark on it yet gets every
    // component's rectangle from KiCad (TabDrafts.MarkAfterAligningAsync) - never while a tool is on,
    // whose Done writes its picture's rectangles as it holds them. Aligned AGAIN, the marks KiCad drew
    // move with it (owner request, 2026-11-01), so each is remembered with the alignment it had.
    //
    // "BACK TO CONTRIBUTE" (owner request, 2026-10-29): after the Contribute tab's "Show on board" - a
    // component, or a net lit up - the Schematics tab's bar says what was shown and goes back to the
    // Contribute tab, which still has the pane open where it was. The bar goes once the Contribute tab
    // is shown again, however, once another board is shown, and when a picture tool starts (its own
    // end goes back).
    // ###########################################################################################
    public partial class Main
    {
        // The tool on the Schematics tab was started from the Contribute tab: its end goes back there.
        private bool thisPictureToolFromContribute;

        // The pictures aligned with Done since the tool started, with the board each is of - marked once
        // the tool hands back (MarkPicturesAligned).
        private readonly List<(string BoardKey, string Picture, string ImageFile, KiCadCalibrationEntry? AlignedFrom)> thisAlignedToMark = new();

        // The marking that followed the last tool's end - for a test to wait for.
        internal Task PicturesAlignedMarkedForTests { get; private set; } = Task.CompletedTask;

        // Stands in for the Contribute tab's marking of one picture (its name and image file) - a test's
        // Main has no draft to mark.
        internal Func<string, string, Task>? MarkAfterAligningOverrideForTests { get; set; }

        // The board "Show on board" showed something of while "Back to Contribute" is offered - null
        // while it is not.
        private string? thisBackToContributeBoard;

        internal bool IsBackToContributeOfferedForTests => this.thisBackToContributeBoard is not null;

        // ###########################################################################################
        // `what` ("[U8]", "the net [PHI2]") is being shown on the Schematics tab for the Contribute tab
        // - the way back offered, for the board shown now.
        // ###########################################################################################
        private void OfferBackToContribute(string what)
        {
            this.thisBackToContributeBoard = this.GetCurrentBoardKey();
            this.TabSchematicsControl.ShowBackToContribute(what);
        }

        private void TakeBackToContributeAway()
        {
            if (this.thisBackToContributeBoard is null)
            {
                return;
            }

            this.thisBackToContributeBoard = null;
            this.TabSchematicsControl.ShowBackToContribute(null);
        }

        // A board shown: the bar was about the one before - unless that board was only read again.
        private void FollowBackToContributeAfterLoad()
        {
            if (this.thisBackToContributeBoard is { } offered &&
                !string.Equals(offered, this.GetCurrentBoardKey(), StringComparison.OrdinalIgnoreCase))
            {
                this.TakeBackToContributeAway();
            }
        }

        // The board read again was asked for while a tool was on: the schematic to restore, done
        // when the tool ends.
        private string? thisReloadWhenToolEnds;

        internal bool IsReloadWaitingForToolForTests => this.thisReloadWhenToolEnds is not null;

        private void WirePictureTools()
        {
            this.TabSchematicsControl.PictureToolEnded += (_, _) => this.OnPictureToolEnded();
            this.TabSchematicsControl.AlignNextRequested += (_, picture) => this.AlignNext(picture);
            this.TabSchematicsControl.AlignedFromWorkspace += (_, aligned) => this.RememberAligned(aligned.Picture, aligned.ImageFile, aligned.AlignedFrom);
            this.TabSchematicsControl.BackToContributeRequested += (_, _) => this.SwitchToContributeTab();

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

        // Whether "Mark components" or "Align KiCad" is on on the Schematics tab - its Done writes what
        // it holds, so the Contribute tab's "Mark all automatically" waits for it.
        internal bool IsPictureToolOn => this.TabSchematicsControl.IsLabelEditorActive;

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

            // The tool goes back by itself once it ends.
            this.TakeBackToContributeAway();

            // Pictures aligned under an earlier tool that never handed back (another board shown
            // meanwhile) are that tool's - not marked by this one's end.
            this.thisAlignedToMark.Clear();

            this.thisPictureToolFromContribute = true;

            this.SwitchToSchematicsTab();
            this.OnBoardShownUpToDate(readBoardFirst, picture, () => this.TabSchematicsControl.StartPictureTool(tool, picture));
        }

        // ###########################################################################################
        // "Done - next picture" (step 3, slice 3d): the alignment saved, the next picture's. A board
        // read again that waited for the tool is read NOW, with that picture selected, and the tool
        // started once it is in - run as the tool went off, it reset the tab and forgot the request.
        // Still the tool the Contribute tab asked for: the last one's end goes back there.
        // ###########################################################################################
        private void AlignNext(string picture)
        {
            bool readBoardFirst = this.thisReloadWhenToolEnds is not null;
            this.thisReloadWhenToolEnds = null;

            this.OnBoardShownUpToDate(readBoardFirst, picture, () => this.TabSchematicsControl.StartPictureTool(PictureTool.AlignKiCad, picture));
        }

        private void OnPictureToolEnded()
        {
            if (!this.thisPictureToolFromContribute)
            {
                return;
            }

            this.thisPictureToolFromContribute = false;
            this.SwitchToContributeTab();

            // After the tool has fully let go (its end is told from inside its own save).
            Dispatcher.UIThread.Post(() => this.PicturesAlignedMarkedForTests = this.MarkPicturesAlignedAsync(), DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Every picture aligned with Done since the tool started, marked from KiCad - of the board shown
        // only (one another board's load took the tool from is not), one after another. The board is
        // asked before EACH one, since each is marked on whatever board is shown then (code review,
        // 2026-11-01: another board picked as the overlay let go between two marks).
        // ###########################################################################################
        internal void RememberAligned(string picture, string imageFile, KiCadCalibrationEntry? alignedFrom = null) =>
            this.thisAlignedToMark.Add((this.GetCurrentBoardKey(), picture, imageFile, alignedFrom));

        internal async Task MarkPicturesAlignedAsync()
        {
            var aligned = this.thisAlignedToMark.ToList();
            this.thisAlignedToMark.Clear();

            try
            {
                foreach ((string board, string picture, string imageFile, KiCadCalibrationEntry? alignedFrom) in aligned)
                {
                    if (string.Equals(board, this.GetCurrentBoardKey(), StringComparison.OrdinalIgnoreCase))
                    {
                        await (this.MarkAfterAligningOverrideForTests?.Invoke(picture, imageFile) ?? this.TabDrafts.MarkAfterAligningAsync(picture, imageFile, alignedFrom));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Marking the pictures just aligned with KiCad failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // "Edit component": the Contribute tab, the component open in its pane - the board's draft made as
        // it opens, as for every other way in, and let go of again when the pane closes with nothing in it
        // (TabDrafts.LetGoOfDraftMadeOnlyToLook). With a picture tool on, that tool is shown, as for a tool
        // asked for meanwhile: it holds the board and the picture until it ends.
        // ###########################################################################################
        internal async Task EditComponentAsync(string boardLabel)
        {
            if (this.IsPictureToolOn)
            {
                this.SwitchToSchematicsTab();
                return;
            }

            try
            {
                this.SwitchToContributeTab();
                await this.TabDrafts.EditComponentAsync(boardLabel);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Opening [{boardLabel}] on the Contribute tab failed - [{ex}]");
            }
        }
    }
}
