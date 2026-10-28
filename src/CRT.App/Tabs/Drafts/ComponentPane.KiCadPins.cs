using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT'S PINS FROM THE BOARD'S KICAD DATA (CRT 4.0 step 2 - the plan: "On a KiCad board,
    // Fill from KiCad adds each pin's net and its name from the symbol").
    //
    // Once a component is shown, the host reads its pins from the board's KiCad data off the UI
    // thread and hands them over (UseKiCadPins - KiCadPinTable, the files the Schematics tab draws).
    // Then:
    //
    //   - "FILL FROM KICAD" adds every pin the KiCad data has and the pins lack, and fills in the names
    //     and directions still empty - what was typed is never changed - as ONE step
    //     (KiCadComponentPins.Fill). Offered while the KiCad data has pins for the component.
    //   - EACH PIN SHOWS ITS NET ON THIS BOARD, beside its direction. The net is the board's, read from
    //     its KiCad data as it is shown - never stored with the pins, which say what the chip is.
    //   - THE PINS ARE ASKED FOR AGAIN when the draft's KiCad files change while the component is shown
    //     (FollowKiCadFiles) - a project imported in the KiCad data window.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The pins of the component shown, as the board's KiCad data has them - none until the host
        // has read them, and for a board without KiCad data.
        private IReadOnlyList<KiCadPin> thisKiCadPins = [];

        // A component is shown (opened, or just added) - the host reads its KiCad pins for it.
        internal event EventHandler<string>? ComponentShown;

        internal IReadOnlyList<KiCadPin> KiCadPinsForTests => this.thisKiCadPins;

        internal IReadOnlyList<ComponentPanePinRow> PinsForTests => this.thisPins;

        // ###########################################################################################
        // The KiCad pins of `label` of the draft `session`, read by the host - taken only while the pane
        // still shows that component of that draft. `ask`: the KiCadPinsAsk the read was started for -
        // an older ask is never taken, whichever read ends last (code review, 2026-10-26: one started
        // before a KiCad import put the old pins back). Null takes the pins whatever was asked.
        // ###########################################################################################
        internal void UseKiCadPins(DraftSession session, string label, IReadOnlyList<KiCadPin> pins, int? ask = null)
        {
            if (!ReferenceEquals(session, this.thisSession) || this.thisIsNew || this.thisBoardPart is not null ||
                !string.Equals(label?.Trim(), this.thisLabel, StringComparison.OrdinalIgnoreCase) ||
                (ask is int asked && asked != this.thisKiCadPinsAsk))
            {
                return;
            }

            this.thisKiCadPins = pins ?? [];
            this.ShowKiCadPins();
        }

        // Another component, a board part or nothing shown: the KiCad pins are another's.
        private void ForgetKiCadPins()
        {
            this.thisKiCadPins = [];
            this.FillFromKiCadButton.IsVisible = false;
            this.PinNetHeading.IsVisible = false;
        }

        // The draft's KiCad files as they were when the component's pins were last asked for.
        private string thisKiCadFilesAskedFor = string.Empty;

        // Which time the pins were asked for - the host hands back the ask a read answers.
        private int thisKiCadPinsAsk;

        internal int KiCadPinsAsk => this.thisKiCadPinsAsk;

        private void RaiseComponentShown()
        {
            this.thisKiCadFilesAskedFor = this.thisSession is { } session ? ComponentPane.KiCadFilesOf(session) : string.Empty;
            this.thisKiCadPinsAsk++;
            this.ComponentShown?.Invoke(this, this.thisLabel);
        }

        private static string KiCadFilesOf(DraftSession session) =>
            string.Join('\n', session.ListedFiles.Values
                .Where(DraftFiles.IsKiCad)
                .Select(file => $"{file.Name}|{file.Sha256}")
                .Order(StringComparer.OrdinalIgnoreCase));

        // ###########################################################################################
        // The draft's KiCad data changed while the component is shown - a project imported or a file
        // removed in the KiCad data window: its pins asked for again (code review, 2026-10-26: "Fill
        // from KiCad" stayed hidden and the nets empty until the component was opened again). After
        // every change and every save of the draft - a file is listed once the draft is written.
        // ###########################################################################################
        private void FollowKiCadFiles()
        {
            // The Pictures part asks for the board's KiCad views again (code review, 2026-10-28: a
            // project imported while the part was open left "Shows KiCad view" empty until it was
            // opened again).
            if (this.thisBoardPart == PaneBoardPart.Pictures && this.thisSession is { } shown)
            {
                if (!string.Equals(ComponentPane.KiCadFilesOf(shown), this.thisKiCadFilesAskedFor, StringComparison.Ordinal))
                {
                    this.AskForKiCadViews();
                }

                return;
            }

            if (this.thisSession is not { } session || this.thisIsNew || this.thisBoardPart is not null || this.thisLabel.Length == 0)
            {
                return;
            }

            if (!string.Equals(ComponentPane.KiCadFilesOf(session), this.thisKiCadFilesAskedFor, StringComparison.Ordinal))
            {
                this.RaiseComponentShown();
            }
        }

        // A save of the draft runs off the UI thread and says so there - followed on the UI thread.
        private void OnSessionSaveStateChanged(object? sender, EventArgs e)
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.Post(() => this.OnSessionSaveStateChanged(sender, e));
                return;
            }

            if (ReferenceEquals(sender, this.thisSession))
            {
                this.FollowKiCadFiles();
            }
        }

        // ###########################################################################################
        // The offer, and each pin's net by its number as it now reads - after the pins are shown and
        // after every change of them (ShowPinout).
        // ###########################################################################################
        private void ShowKiCadPins()
        {
            this.FillFromKiCadButton.IsVisible = this.thisKiCadPins.Count > 0;
            ToolTip.SetTip(this.FillFromKiCadButton, ComponentPaneWording.FillFromKiCadTip(this.thisLabel));
            ToolTip.SetTip(this.PinNetHeading, ComponentPaneWording.NetHeadingTip);

            Dictionary<string, string> nets = KiCadComponentPins.NetsByPin(this.thisKiCadPins);

            foreach (ComponentPanePinRow row in this.thisPins)
            {
                row.UseNet(nets.TryGetValue(row.Pin.Trim(), out string? net) ? net : string.Empty);
            }

            this.PinNetHeading.IsVisible = this.thisPins.Any(row => row.HasNet);
        }

        private void OnFillFromKiCadClick(object? sender, RoutedEventArgs e) => this.FillFromKiCad();

        // ###########################################################################################
        // Every pin the KiCad data has and the pins lack added, and the empty names and directions
        // filled in, as ONE step - what came of it said under the title. False when there was nothing
        // to add or fill in.
        // ###########################################################################################
        internal bool FillFromKiCad()
        {
            if (this.thisSession is null || this.thisKiCadPins.Count == 0)
            {
                return false;
            }

            this.CommitHeld();
            this.ShowMessage(null);

            List<ComponentPinEntry> pins = KiCadComponentPins.Fill(
                this.thisPins.Select(row => row.ToEntry()).ToList(),
                this.thisKiCadPins,
                this.thisRowRegion,
                out int added,
                out int filled);

            if (added + filled == 0)
            {
                this.ShowMessage(ComponentPaneWording.NothingToFillFromKiCad);
                return false;
            }

            ComponentPane.Fill(this.thisPins, pins.Select(ComponentPanePinRow.From));

            this.Apply(
                DraftStepWording.ComponentPinsFromKiCad(this.thisLabel),
                board => ComponentPaneEdits.WithPins(board, this.thisLabel, this.thisRegion, pins));

            this.ShowPinout();
            this.ShowEmptyLines();
            this.ShowMessage(ComponentPaneWording.FilledFromKiCad(added, filled));

            return true;
        }
    }
}
