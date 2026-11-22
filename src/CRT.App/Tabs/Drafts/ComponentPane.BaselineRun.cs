using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT'S BASELINES WITH THE OSCILLOSCOPE (owner requests, 2026-11-18: "how I easily can take an entire oscilloscope
    // baseline in one go - e.g. for updating it", and 2026-11-20: "Remove the button named "Capture from my oscilloscope" ... Rename
    // "Capture pin by pin" to "Capture from oscilloscope pin-by-pin" ... open this in a new window"). See ComponentPane.axaml.cs for
    // the file map.
    //
    //   - "CAPTURE FROM OSCILLOSCOPE PIN-BY-PIN" opens BaselineCaptureWindow on the component shown - ONE window, brought to the
    //     front when it is open already - and the pane closes it as another component or region is shown, or as it lets go; it
    //     tells it when the oscilloscope comes or goes (ShowCaptureOffer) and when the baselines are read again (FollowRows).
    //   - EVERY CAPTURE IS THE PANE'S, IN TWO HALVES: TakeCaptureAsync - the screen and what the oscilloscope says it is set to, the
    //     screen kept in the user's OWN app data folder (AppConfig.ScopeCaptureWaitFolder, never the shared temporary one) while the
    //     window shows it, not yet accepted - and PutCaptureAsync, once it is ACCEPTED: into the baseline - or a new one - as ONE step
    //     of the draft, saved, the waiting screen gone (owner request, 2026-11-20: "this must be accepted before it will be used").
    //     Under the "please wait" of the window that asked (`anchor`). The settings read back are what the baseline stores
    //     (BaselineSettingWords), a value not answered leaving its own as it was.
    //   - "CAPTURE AGAIN" on a baseline opens the window on its pin and captures there, to be checked and accepted (owner request,
    //     the same day: "better than just do an instant-capture, without having good possibility to review").
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The oscilloscope - the "Oscilloscope" tab's, from the host (TabDrafts.UseScope). Null: no capture is offered.
        internal IBaselineScope? Scope { get; set; }

        // Where "Capture from oscilloscope pin-by-pin"'s window was left, and how to remember it - the host's (UserSettings, set by
        // TabDrafts.Initialize); a test's pane keeps none.
        internal Func<BaselineCaptureWindowPlace?>? CaptureWindowPlace { get; set; }

        internal Action<BaselineCaptureWindowPlace>? RememberCaptureWindowPlace { get; set; }

        // For a test: the screen of CRT's window the capture window opens over - its working area and scaling - in place of the
        // headless platform's one large screen.
        internal (PixelRect WorkingArea, double Scaling)? CaptureWindowScreenForTests { get; set; }

        private BaselineCaptureWindow? thisCaptureWindow;

        // The window to follow the baselines once the files being added are in the draft (FollowRowsInBaselineRun).
        private bool thisBaselineRunFollowWanted;

        internal BaselineCaptureWindow? CaptureWindowForTests => this.thisCaptureWindow;

        // Called with where a capture waits, as it is written - for a test.
        internal Action<string>? CaptureWrittenForTests { get; set; }

        // ---- what the window reads of the component shown -------------------------------------

        internal string ShownLabel => this.thisLabel;

        internal string ShownRegion => this.thisRegion;

        internal string ShownTitle => this.TitleText.Text ?? this.thisLabel;

        internal bool BoardHasRegions => this.thisBoardHasRegions;

        // The region a new baseline goes in - what the window shows a capture of a pin with none as, before it is accepted.
        internal string RegionForCapturedBaseline => this.RegionForNewBaseline();

        internal IReadOnlyList<ComponentPaneBaselineRow> BaselineRows => this.thisBaselines;

        // The pins a run can go round - the pane's pins, the datasheet's CRT holds of the chip, and the baselines.
        internal IReadOnlyList<BaselineRunPin> RunPins()
        {
            List<ComponentPinEntry> pins = this.thisPins.Select(row => row.ToEntry()).ToList();
            ComponentEntry? shown = this.ShownComponent();
            DatasheetPinout? datasheet = shown is null ? null : DatasheetPins.PinoutOf(shown, pins);

            return BaselineRunPins.Of(this.PackageBox.Text, pins, datasheet?.Pins, this.thisBaselines.Select(row => row.Committed).ToList());
        }

        // The component's case drawn on `drawing`, as its "Pinout" page draws it.
        internal void ShowCaseOn(PinoutDrawing drawing)
        {
            ArgumentNullException.ThrowIfNull(drawing);

            List<ComponentPinEntry> pins = this.thisPins.Select(row => row.ToEntry()).ToList();
            ComponentEntry? shown = this.ShownComponent();

            drawing.Show(this.PackageBox.Text ?? string.Empty, pins, shown is null ? null : DatasheetPins.PinoutOf(shown, pins)?.Logic);
        }

        // Where a file of the board lies - the draft's, else the published board's - null when it is nowhere on this machine.
        internal string? ResolveFileOfBoard(string? file)
        {
            if (this.thisSession is not { } session || string.IsNullOrWhiteSpace(file))
            {
                return null;
            }

            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            return DraftFileResolver.Resolve(DataManager.DataRoot, draftFolder, file, session.ExcelDataFile);
        }

        // ---- the window -------------------------------------------------------------------------

        private void OnRunOpenClick(object? sender, RoutedEventArgs e) => this.OpenBaselineCaptureWindow();

        // ###########################################################################################
        // "Capture from oscilloscope pin-by-pin": the window on the component shown, in front of CRT's - the one open brought to the
        // front. Null when there is nothing to capture with or for (said under the title for a component with no pins).
        // ###########################################################################################
        internal BaselineCaptureWindow? OpenBaselineCaptureWindow()
        {
            if (this.Scope is null || this.IsSeveral || this.thisLabel.Length == 0)
            {
                return null;
            }

            if (this.thisCaptureWindow is { } open)
            {
                open.Activate();
                return open;
            }

            // A new baseline half filled in stays as it is (code review, 2026-11-21: "Capture again" threw it away): added, it is
            // a line of the window's like any other (FollowRowsInBaselineRun).
            this.CommitHeld();
            this.ShowMessage(null);

            IReadOnlyList<BaselineRunPin> pins = this.RunPins();

            if (pins.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.RunNoPins);
                return null;
            }

            var window = new BaselineCaptureWindow();

            window.Closed += (_, _) =>
            {
                if (ReferenceEquals(this.thisCaptureWindow, window))
                {
                    this.thisCaptureWindow = null;
                }
            };

            this.thisCaptureWindow = window;
            window.Begin(this, pins, this.IsCaptureOffered, this.CaptureWindowPlace?.Invoke(), this.RememberCaptureWindowPlace);

            if (TopLevel.GetTopLevel(this) is Window owner)
            {
                window.Show(owner);
            }
            else
            {
                window.Show();
            }

            return window;
        }

        // ###########################################################################################
        // *** CAPTURES NOT YET ACCEPTED HOLD THE PANE ON THEIR COMPONENT (code review, 2026-11-22). *** The window is one of its own, so
        // CRT's stays clickable - and another component, a part of the board, "Table" or a region picked closed it, every screen
        // captured deleted unasked: each pin to be probed again. While any waits the host is refused (KeepsItselfOpen) and the window,
        // brought to the front, asks what its own close asks - accept them and close, discard them and close, or keep capturing;
        // answered, the pick is made again. Another board shown, or the draft discarded, still let them go: neither can be called off.
        // ###########################################################################################
        // `region`: asked for another region to be shown - one the window's captures were taken in is no reason to hold.
        internal bool KeepsWaitingCaptures(string? region = null)
        {
            if (this.thisCaptureWindow is not { HoldsCaptures: true } window ||
                (region is not null && string.Equals(region.Trim(), window.RunRegion, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            window.AskAboutWaitingCaptures();
            return true;
        }

        // The window closed by the pane - another board shown, the draft let go of, the component gone: a capture there not yet
        // accepted goes with it, asked about by nobody (it is in no draft, and the component it was of is no longer shown).
        internal void CloseBaselineCaptureWindow()
        {
            if (this.thisCaptureWindow is { } window)
            {
                this.thisCaptureWindow = null;
                window.LetGo();
            }
        }

        // ###########################################################################################
        // The component's baselines read again (an Undo in the header, another editor's change) or changed in the pane (CommitBaselines -
        // one added, deleted, given another pin or image): the window finds its baselines again - or, the component shown in another
        // region now, it closes, its steps being of the region it opened in.
        // ###########################################################################################
        private void FollowRowsInBaselineRun()
        {
            if (this.thisCaptureWindow is not { } window)
            {
                return;
            }

            // A file on its way into the draft: followed once it is in (AddToDraftAsync), as no picture is read before then -
            // the window read a new baseline's picture while the save moved it (code review, 2026-11-21).
            if (this.thisFilesArriving)
            {
                this.thisBaselineRunFollowWanted = true;
                return;
            }

            if (!string.Equals(window.RunRegion, this.thisRegion, StringComparison.OrdinalIgnoreCase))
            {
                // ###########################################################################################
                // Captures not yet accepted are never let go of unasked (code review, 2026-11-22): a region picked is held while any
                // waits (KeepsRegionShown), so the region changed under them by an edit - "Board: Regions", or an "Undo" of it. The
                // window follows the region the pane shows now, where they would go in, and asks what to do with them.
                // ###########################################################################################
                if (window.HoldsCaptures)
                {
                    window.FollowRegion(this.thisRegion);
                    window.FollowRows();
                    window.AskAboutWaitingCaptures();
                    return;
                }

                this.CloseBaselineCaptureWindow();
                return;
            }

            window.FollowRows();
        }

        // The component renamed, or a name of it changed, in the pane: the window's title and heading say it (code review, 2026-11-22).
        private void ShowNamesInBaselineRun() => this.thisCaptureWindow?.ShowComponentNames();

        // ---- the capture, for the window ----------------------------------------------------------

        private async void OnCaptureBaselineAgainClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: ComponentPaneBaselineRow row })
                {
                    await this.CaptureAgainInWindowAsync(row);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Capturing an oscilloscope baseline again failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // "CAPTURE AGAIN" on a baseline (owner request, 2026-11-20: "this button should open the "Capture from oscilloscope pin-by-pin"
        // window and go to that pin and capture the image from scope and then show it in this view instead. Then its data can be
        // controlled by the user before it maybe will be accepted"): the window on that baseline's pin, the oscilloscope set as it was
        // measured, and a capture taken there - not yet accepted, so not in the draft until "Accept and continue to next pin".
        // ###########################################################################################
        internal async Task CaptureAgainInWindowAsync(ComponentPaneBaselineRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            this.CommitHeld();
            this.ShowMessage(null);

            if (this.OpenBaselineCaptureWindow() is not { } window)
            {
                return;
            }

            // A capture under way there: it is not moved to another pin meanwhile (BaselineCaptureWindow.IsBusy) - brought forward, and
            // why nothing more happens said (code review, 2026-11-22).
            if (window.IsBusy)
            {
                this.ShowMessage(ComponentPaneWording.RunWindowBusy);
                return;
            }

            // A row the window cannot find is one the pane has let go of since it was pressed - said, never a capture somewhere else.
            if (!window.GoToBaseline(row))
            {
                this.ShowMessage(ComponentPaneWording.RunBaselineNotFound);
                return;
            }

            await window.CaptureCurrentAsync();
        }

        // ###########################################################################################
        // The oscilloscope's screen - kept in the user's OWN app data folder until it is accepted into the draft or let go of
        // (AppConfig.ScopeCaptureWaitFolder, never the shared temporary one, where another account could swap it) - and what it is set
        // to, read right after it, under `anchor`'s "please wait". Null, with `sayProblem` told why, when nothing came.
        // ###########################################################################################
        internal async Task<ScopeCapture?> TakeCaptureAsync(StyledElement anchor, Action<string?> sayProblem)
        {
            ArgumentNullException.ThrowIfNull(anchor);
            ArgumentNullException.ThrowIfNull(sayProblem);

            if (this.Scope is not { } scope)
            {
                return null;
            }

            (byte[]? png, ScopeSettingsReading settings) = await BusyOverlay.RunLocalAsync(anchor, CrtWaitWording.CapturingOscilloscope, async () =>
            {
                byte[]? screen = await scope.CaptureAsync(CancellationToken.None);
                ScopeSettingsReading read = screen is { Length: > 0 }
                    ? await scope.ReadSettingsAsync(CancellationToken.None)
                    : new ScopeSettingsReading(null, null, null);

                return (screen, read);
            });

            if (png is not { Length: > 0 })
            {
                sayProblem(ComponentPaneWording.CaptureFailed);
                return null;
            }

            string path = Path.Combine(AppConfig.ScopeCaptureWaitFolder, $"{Guid.NewGuid():N}.png");

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                await File.WriteAllBytesAsync(path, png);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not keep the scope capture at [{path}] - [{ex.Message}]");
                sayProblem(ComponentPaneWording.CaptureFailed);
                return null;
            }

            this.CaptureWrittenForTests?.Invoke(path);
            return new ScopeCapture(path, settings);
        }

        // ###########################################################################################
        // A capture ACCEPTED into `row` - its image and the settings read with it, the rest as it was - or, with no row, a new baseline
        // of `pin` named `signal`, in the region a new baseline goes in. ONE step, the draft saved (AddToDraftAsync), under the "please
        // wait" of `anchor`'s window; the screen waiting in the user's folder goes once it is in the draft, and stays when it could not
        // be, to be accepted again. The baseline's file name, or null with `sayProblem` told why.
        //
        // *** ONLY WHILE ITS COMPONENT AND ITS BASELINE ARE STILL THE PANE'S (code review, 2026-11-21). *** The window's "please wait"
        // covers the window alone, so CRT's own stays clickable while the capture is copied in: another component shown, or the
        // baselines read again (an "Undo"), and the change - run after the copy - filed it under the component shown by then, or wrote
        // it into a row the pane no longer had. The copy is taken back then, and the capture waits on.
        // ###########################################################################################
        internal async Task<string?> PutCaptureAsync(
            StyledElement anchor,
            ScopeCapture capture,
            ComponentPaneBaselineRow? row,
            string pin,
            string signal,
            string machineState,
            Action<string?> sayProblem)
        {
            ArgumentNullException.ThrowIfNull(anchor);
            ArgumentNullException.ThrowIfNull(capture);
            ArgumentNullException.ThrowIfNull(sayProblem);

            string name = ComponentPaneFileNames.ForBaseline(this.thisLabel, pin, signal, capture.Path, this.TakenNames());
            ScopeSettingsReading settings = capture.Settings;

            // A baseline the pane holds back - its pin or signal name emptied - takes nothing into the draft, the capture included: said
            // before anything is copied (code review, 2026-11-21: it read "Accepted", the screen was deleted, and the draft kept the old).
            if (row is not null && row.ToEntry() is var held && ComponentPaneWording.BaselineProblem(held.Pin, held.Name, name) is { } problem)
            {
                sayProblem(ComponentPaneWording.BaselineHeldBack(problem));
                return null;
            }
            string label = this.thisLabel;
            string region = this.thisRegion;
            DraftSession? session = this.thisSession;

            // The component still the pane's one shown, and on the board - not deleted, nor several picked.
            bool StillThePanes() =>
                ReferenceEquals(this.thisSession, session) &&
                session is not null &&
                !this.IsSeveral &&
                string.Equals(this.thisLabel, label, StringComparison.Ordinal) &&
                string.Equals(this.thisRegion, region, StringComparison.Ordinal) &&
                ComponentPaneEdits.RowShown(session.Content.Board, label, region) is not null &&
                (row is null || this.thisBaselines.Contains(row));

            // What the window showed before it was accepted, by the same rule.
            ComponentImageEntry captured = BaselineSettingWords.Captured(row?.ToEntry(), settings, name, pin, signal, machineState, this.RegionForNewBaseline());

            bool added = await this.AddToDraftAsync([(capture.Path, name)], _ =>
            {
                if (row is not null)
                {
                    row.File = captured.File;
                    row.Thumbnail = null;
                    row.TimeDiv = captured.TimeDiv;
                    row.VoltsDiv = captured.VoltsDiv;
                    row.TriggerLevel = captured.TriggerLevelVolts;
                }
                else
                {
                    ComponentPaneBaselineRow added = ComponentPaneBaselineRow.From(captured);

                    added.CaptureOffered = this.IsCaptureOffered;
                    this.thisBaselines.Add(added);
                }

                this.CommitBaselines();
            }, anchor, sayProblem, StillThePanes);

            if (!added)
            {
                return null;
            }

            this.ShowEmptyLines();
            ComponentPane.DeleteCapture(capture.Path);
            return name;
        }

        // A capture waiting in the user's own folder, gone once it is in the draft or let go of.
        internal static void DeleteCapture(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not remove the scope capture [{path}] - [{ex.Message}]");
            }
        }

        // The screen captured, where it waits, and what the oscilloscope was set to.
        internal sealed record ScopeCapture(string Path, ScopeSettingsReading Settings);
    }
}
