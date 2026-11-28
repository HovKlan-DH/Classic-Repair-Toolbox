using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Oscilloscope;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "CAPTURE FROM OSCILLOSCOPE PIN-BY-PIN" (owner requests, 2026-11-18 - "how I easily can take an entire oscilloscope baseline in one
    // go" - and 2026-11-20: "open this in a new window (remember size and position), where it will display everything related with
    // capturing the image ... support the same keyboard options as available in the component info window, when having a scope
    // connected ... without needing to doublecheck from multiple windows or loosing focus from the oscilloscope"). Opened by the
    // component pane for the component it shows (ComponentPane.BaselineRun.cs) - from its button, or from a baseline's "Capture
    // again" on that baseline's pin - and working through it: the pins, the baselines, the oscilloscope and every capture into the
    // draft are the pane's; every rule of pins and steps is CRT.Data's BaselineRunPins.
    //
    //   - ONE VIEW, NO START: every pin listed - each baseline of a pin a line, a pin with none one - checked to capture (the quick
    //     picks, which go to the first pin they pick; a checkbox; a pin clicked on the case); the pin shown is the list's selection.
    //   - CAPTURE, CHECK, ACCEPT - AND ONLY THEN THE DRAFT (owner requests, 2026-11-20: "it must stay on the pin, so I can check that it
    //     is good and its data looks good ... "Accept and continue to next pin". The "Previous pin" should not "unmark" it as
    //     accepted", then "New capture - not yet accepted", "so it is clear this must be accepted before it will be used"): ENTER
    //     captures the pin shown - its screen and what the oscilloscope says it is set to - and STAYS on it, "not yet accepted": the
    //     capture waits in the user's own folder, in no draft; Enter again takes it again. "ACCEPT AND CONTINUE TO NEXT PIN" (Right)
    //     puts it into its baseline or a new one, ONE step of the draft, saved, and shows the next picked pin not accepted, the
    //     oscilloscope set as its baseline was measured - with nothing captured, it only goes on. Going back accepts and un-accepts
    //     nothing. A new pin's signal name may be typed until it is accepted.
    //   - CLOSING WITH CAPTURES NOT YET ACCEPTED ASKS FIRST - "Close capturing", Escape, the window's own close - in a panel over the
    //     buttons: accept them, discard them, or keep capturing. Closed by the pane (another component or region shown, the pane
    //     letting go - LetGo) or with CRT, they are let go of: in no draft, the component they were of no longer shown. Asked to close
    //     while a capture goes into the draft, it closes once that capture is in.
    //   - THE OLD BASELINE FRAMED BESIDE THE NEW CAPTURE (owner requests, the same day: "I want also to see the old baseline image",
    //     "have it more clear/visible what is the old and new image", "a red'ish background color"): a pin captured again shows the
    //     baseline the run started from on the left, red, the capture on the right - amber while not yet accepted, green once accepted -
    //     and each changed value says what it was. The old picture's bytes are read before the capture can replace it - a draft's own
    //     file is listed no more after.
    //   - THE PICTURES TAKE THE HEIGHT THE REST LEAVES (owner request, the same day: "use the full height"), the fields under them in a
    //     table whose values line up ("align the values").
    //   - EVERY PICTURE IS CLICKED TO SEE IT LARGE (owner request, the same day: "click any image or thumbnail I see in this window and
    //     then it will do the same image view as e.g. when working in worklogs"): WorklogPhotoViewerWindow, which zooms, its settings
    //     each named and bold, and which a click outside closes. A thumbnail's press is the picture's alone - the line is not picked by
    //     it, nor the oscilloscope set. "Click to see it larger" FOLLOWS THE POINTER over a picture ("have the hover/helper text follow
    //     the mouse") - drawn in the window, taking no click, never a tooltip: a tooltip's own window under the pointer takes the next
    //     click (ToolTipPlacement, PointerHintGeometry).
    //   - THE BUTTONS NEVER MOVE (owner request, the same day: "it can jump up and down, depending if it has image or not"): a bar of
    //     their own at the foot of the column, "Close capturing" at its far end.
    //   - THE COMPONENT INFO WINDOW'S KEYS (Controlling-oscilloscope-with-keyboard), the numpad's always the oscilloscope's here: Left
    //     the picked pin before, Right "Accept and continue to next pin", the number keys to a pin, Space to the first; Up and Down
    //     the trigger level, numpad + and - the time per division, 1 and 2 the volts, "." clear statistics, "*" single, "/" run;
    //     Escape closes.
    //   - THE OSCILLOSCOPE GONE, the window stays and says so, nothing captured.
    //   - ITS SIZE AND PLACE ARE KEPT (UserSettings.BaselineCaptureWindowPlace, handed in through the pane - a test's window keeps none),
    //     made to fit the screen it opens on (DialogPlacement.KeptPlaceOn).
    //   - ITS LINES FOLLOW THE DRAFT (code review, 2026-11-21): the pane's baselines read again or changed, every line finds its
    //     baseline by CRT.Data's BaselineRunPins.Follow - an "Undo" of an accepted capture takes its line back, a baseline added in
    //     the pane gets one - and a capture goes only into a baseline the pane still has.
    //
    // The file map:
    //   - BaselineCaptureWindow.axaml.cs - opening, the lines following the draft, closing, picking, the pin shown, its fields, the
    //     oscilloscope and capturing.
    //   - BaselineCaptureWindow.Pictures.cs - the pictures, and every picture seen large.
    //   - BaselineCaptureWindow.Keys.cs - the keys.
    //   - BaselineCaptureTarget.cs - a line of the list (BaselineCaptureTarget) and a key along the foot (BaselineCaptureKey).
    // ###########################################################################################
    public partial class BaselineCaptureWindow : Window
    {
        private ComponentPane? thisPane;
        private readonly ObservableCollection<BaselineCaptureTarget> thisTargets = new();
        private int thisCurrentIndex = -1;
        private bool thisConnected;
        private bool thisWorking;
        private bool thisShowingSelection;
        private string thisRunRegion = string.Empty;

        // Closing without asking about captures not yet accepted - the pane let go, or "Discard and close"/"Accept and close" decided.
        private bool thisLettingGo;

        // Closed: a capture that comes in after it goes at once (code review, 2026-11-21: one ending after Escape stayed in the
        // user's folder for good).
        private bool thisClosed;

        // A capture going into the draft: the pane's baselines change under it, and the lines follow once it is in, its own
        // accepted with them (FollowRows) - followed before, its new baseline was taken for one added in the pane.
        private bool thisPutting;
        private bool thisFollowWanted;

        // The line whose capture is going into the draft - closed meanwhile, it is let go of once the put ends (AcceptCaptureAsync),
        // never deleted while it is being copied in.
        private BaselineCaptureTarget? thisPuttingTarget;

        // The window's own close asked for while a capture went into the draft: held until it is in, then closed - asking first about
        // any other capture waiting (code review, 2026-11-21: the close was dropped without a word, and had to be asked again).
        private bool thisCloseWanted;

        // The oscilloscope being set for the pin shown - a capture waits for it.
        private Task thisScopeSet = Task.CompletedTask;

        // The old baselines' and the waiting captures' pictures being read, off the UI thread.
        private readonly List<Task> thisPictureReads = new();

        // The fields' value blocks, by the field's name - their labels beside them (ShowFacts).
        private readonly Dictionary<string, (TextBlock Label, TextBlock Value)> thisFactCells = new(StringComparer.Ordinal);

        // Where each field stands: the pin's down the left, how the oscilloscope was set down the right.
        private static readonly (string Name, int Row, int Column)[] FactPlaces =
        [
            (ComponentPaneWording.FactPin, 0, 0),
            (ComponentPaneWording.FactSignalName, 1, 0),
            (ComponentPaneWording.FactExpectedReading, 2, 0),
            (ComponentPaneWording.FactMachineState, 3, 0),
            (ComponentPaneWording.FactNote, 4, 0),
            (ComponentPaneWording.FactTimeDiv, 0, 4),
            (ComponentPaneWording.FactVoltsDiv, 1, 4),
            (ComponentPaneWording.FactTriggerLevel, 2, 4),
            (ComponentPaneWording.FactAppliesTo, 3, 4),
        ];

        // The window's place (UserSettings through the pane) - and its Normal size and position, as it is moved and sized
        // (KeptWindowPlace, the collection window's too since 2026-11-28).
        private Action<WindowPlace>? thisRememberPlace;
        private readonly KeptWindowPlace thisPlace;

        public BaselineCaptureWindow()
        {
            this.InitializeComponent();

            this.TargetsList.ItemsSource = this.thisTargets;
            this.TargetsList.SelectionChanged += this.OnTargetSelected;
            this.CasePinout.PinPicked += this.OnCasePinPicked;
            this.KeysList.ItemsSource = ComponentPaneWording.RunKeys.Select(key => new BaselineCaptureKey(key.Keys, key.What)).ToList();
            this.IntroText.Text = ComponentPaneWording.CaptureWindowIntro;
            this.ScopeConnectedText.Text = ComponentPaneWording.ScopeConnected;
            this.NoScopeText.Text = ComponentPaneWording.ScopeNotConnected;
            this.NoOldImageText.Text = ComponentPaneWording.PictureOldUnreadable;
            this.PictureHintText.Text = ComponentPaneWording.PictureHint;
            this.CreateFactCells();

            this.AddHandler(InputElement.KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);

            this.thisPlace = new KeptWindowPlace(this);
            this.Closing += this.OnClosing;
            this.Closed += this.OnClosed;
        }

        internal IReadOnlyList<BaselineCaptureTarget> TargetsForTests => this.thisTargets;

        internal int CurrentIndexForTests => this.thisCurrentIndex;

        internal BaselineCaptureTarget? CurrentForTests => this.thisCurrentIndex >= 0 ? this.thisTargets[this.thisCurrentIndex] : null;

        // A field of the baseline shown as it reads - null while it is not shown.
        internal string? FactShownForTests(string name) =>
            this.thisFactCells.TryGetValue(name, out (TextBlock Label, TextBlock Value) cell) && cell.Value.IsVisible && this.FactsGrid.IsVisible
                ? BoldRuns.TextOf(cell.Value)
                : null;

        internal TextBlock FactValueBlockForTests(string name) => this.thisFactCells[name].Value;

        // Whether closing is asking about captures not yet accepted.
        internal bool IsAskingAboutWaitingForTests => this.NotAcceptedPanel.IsVisible;

        // How many times the lines followed the pane's baselines (FollowRows) - a commit that changed nothing follows nothing.
        internal int RowsFollowedForTests { get; private set; }

        // Every big picture being read - a test waits for them before its files go (Windows refuses a file being read).
        internal Task ImagesReadForTests => Task.WhenAll(this.thisImagesBeingRead.Values.Concat(this.thisPictureReads).ToList());

        // Stands in for opening the picture viewer - a test sees the viewer made; null opens it over this window.
        internal Action<WorklogPhotoViewerWindow>? ShowPictureOverrideForTests { get; set; }

        // ---- opening ----------------------------------------------------------------------------

        // ###########################################################################################
        // The window on `pane`'s component: its pins as the steps to go round, the pins with a baseline picked (else every pin), the
        // first picked one shown - and where it was left, when it was (`place`).
        // ###########################################################################################
        internal void Begin(
            ComponentPane pane,
            IReadOnlyList<BaselineRunPin> pins,
            bool connected,
            WindowPlace? place,
            Action<WindowPlace>? rememberPlace)
        {
            ArgumentNullException.ThrowIfNull(pane);
            ArgumentNullException.ThrowIfNull(pins);

            this.thisPane = pane;
            this.thisRememberPlace = rememberPlace;
            this.thisRunRegion = pane.ShownRegion;
            this.ShowComponentNames();
            this.MachineStateBox.Text = pane.BaselineRows.Select(row => row.MachineState.Trim()).LastOrDefault(state => state.Length > 0) ?? string.Empty;

            // Fitted to its screen whether or not a place is kept (code review, 2026-11-22: the first time, at its 1280 x 860, a
            // laptop's screen had its button bar and keys under the screen's edge) - the screen of CRT's window it opens over, when
            // no place is kept or the place kept is on no screen.
            (PixelRect WorkingArea, double Scaling)? ownerScreen = this.OwnerScreenOf(pane);

            if (place is not null)
            {
                this.thisPlace.Use(place, ownerScreen);
            }
            else if (ownerScreen is { } screen)
            {
                this.thisPlace.FitTo(screen);
            }

            var picked = BaselineRunPins.FirstChoice(pins).ToHashSet();
            IReadOnlyList<ComponentPaneBaselineRow> rows = pane.BaselineRows;

            foreach (BaselineRunLine line in BaselineRunPins.Lines(pins, rows.Select(row => row.Committed).ToList()))
            {
                line.Step.IsPicked = picked.Contains(line.Step.Pin);

                // The baseline as the run started - what a capture of it is shown beside.
                ComponentPaneBaselineRow? row = line.Row >= 0 ? rows[line.Row] : null;
                this.thisTargets.Add(new BaselineCaptureTarget(line.Step, this.OnPickedChanged) { BeforeEntry = row?.Committed, Row = row });
            }

            pane.ShowCaseOn(this.CasePinout);
            this.FollowRows();
            this.ShowScopeConnected(connected);

            int first = this.thisTargets.ToList().FindIndex(target => target.Step.IsPicked);
            this.GoTo(first >= 0 ? first : 0, setScope: true);
        }

        // The region the run's steps are of - another one shown, the pane closes the window, or, captures waiting, has it follow.
        internal string RunRegion => this.thisRunRegion;

        // The region the pane shows now, its captures waiting - an edit changed it under them (ComponentPane.FollowRowsInBaselineRun).
        internal void FollowRegion(string region) => this.thisRunRegion = region ?? string.Empty;

        // ###########################################################################################
        // The component's label and names as the pane shows them now - in the title and over the pins - read as it opens, as its lines
        // follow the pane, and as the pane renames it or changes a name (code review, 2026-11-22: renamed U27A in the pane, the
        // window still said U27 while its captures went in under U27A).
        // ###########################################################################################
        internal void ShowComponentNames()
        {
            if (this.thisPane is not { } pane)
            {
                return;
            }

            this.Title = ComponentPaneWording.CaptureWindowTitle(pane.ShownLabel);
            this.ComponentText.Text = pane.ShownTitle;
        }

        // ###########################################################################################
        // A capture being taken - from the moment it waits for the oscilloscope to be set for its pin, before any "please wait" is up -
        // or a capture going into the draft: NOTHING MOVES TO ANOTHER PIN MEANWHILE, by key or by mouse (code review, 2026-11-21:
        // "Previous pin" clicked then set the oscilloscope for another pin, and the screen was stored with that pin's settings).
        // ###########################################################################################
        internal bool IsBusy => this.thisWorking || this.thisPutting;

        // Where a new baseline goes - what a capture of a pin with none is shown as, before it is accepted (WaitingEntry).
        private string CapturedBaselineRegion => this.thisPane?.RegionForCapturedBaseline ?? string.Empty;

        // The screen of CRT's window it opens over - its working area and scaling - null when none is known.
        private (PixelRect WorkingArea, double Scaling)? OwnerScreenOf(ComponentPane pane)
        {
            if (pane.CaptureWindowScreenForTests is { } given)
            {
                return given;
            }

            return TopLevel.GetTopLevel(pane) is Window owner && this.Screens.ScreenFromWindow(owner) is { Scaling: > 0 } screen
                ? (screen.WorkingArea, screen.Scaling)
                : null;
        }

        // Where it is left (KeptWindowPlace) - its Normal size and position as they are now, or, maximized, as they were before.
        private void RememberPlace() => this.thisRememberPlace?.Invoke(this.thisPlace.Current());

        // ###########################################################################################
        // The oscilloscope there or gone, from the pane: gone, nothing is captured or sent to it, and the window says so - what was
        // captured stays, and it goes on once the oscilloscope is back.
        // ###########################################################################################
        internal void ShowScopeConnected(bool connected)
        {
            bool cameBack = connected && !this.thisConnected;

            this.thisConnected = connected;
            this.ScopeConnectedPanel.IsVisible = connected;
            this.NoScopeNotice.IsVisible = !connected;
            this.ShowButtons();

            // Back: set as the pin shown was measured - it was not, while it was gone (code review, 2026-11-21: a pin moved to
            // meanwhile was captured with the oscilloscope as it was set for the pin before).
            if (cameBack && this.CurrentForTests?.Row is { } row)
            {
                this.SetScopeFor(row);
            }
        }

        // ###########################################################################################
        // THE PANE'S BASELINES READ AGAIN OR CHANGED (an "Undo", another editor's change, a baseline added, deleted or given another pin
        // in the pane): the lines as they stand now - every rule CRT.Data's BaselineRunPins.Follow (code review, 2026-11-21: an Undo of
        // an accepted capture left its line "Accepted", and a baseline the pane added had no line). Each line kept is the very line it
        // was, the pin shown stays shown, and a line that went takes nothing with it: a capture waiting is never dropped.
        // ###########################################################################################
        internal void FollowRows()
        {
            if (this.thisPane is not { } pane)
            {
                return;
            }

            if (this.thisPutting)
            {
                this.thisFollowWanted = true;
                return;
            }

            this.RowsFollowedForTests++;
            this.CommitName();

            IReadOnlyList<ComponentPaneBaselineRow> rows = pane.BaselineRows;
            BaselineCaptureTarget? shown = this.CurrentForTests;
            Dictionary<BaselineRunStep, BaselineCaptureTarget> byStep = this.thisTargets.ToDictionary(target => target.Step);
            IReadOnlyList<BaselineRunLine> lines = BaselineRunPins.Follow(this.Steps(), pane.RunPins(), rows.Select(row => row.Committed).ToList());
            var targets = new List<BaselineCaptureTarget>(lines.Count);

            foreach (BaselineRunLine line in lines)
            {
                ComponentPaneBaselineRow? row = line.Row >= 0 ? rows[line.Row] : null;
                BaselineCaptureTarget target = byStep.TryGetValue(line.Step, out BaselineCaptureTarget? kept)
                    ? kept
                    : new BaselineCaptureTarget(line.Step, this.OnPickedChanged) { BeforeEntry = row?.Committed };

                target.Row = row;
                target.ShowPictures();
                targets.Add(target);
            }

            foreach (BaselineCaptureTarget gone in this.thisTargets.Except(targets).ToList())
            {
                gone.LetGoOfCapture();
            }

            if (!targets.SequenceEqual(this.thisTargets))
            {
                int current = shown is null ? -1 : targets.IndexOf(shown);

                if (current < 0 && targets.Count > 0)
                {
                    current = Math.Max(0, shown is null ? 0 : targets.FindIndex(target => ComponentVocabulary.IsSamePin(target.Step.Pin.Pin, shown.Step.Pin.Pin)));
                }

                this.thisShowingSelection = true;

                try
                {
                    this.thisTargets.Clear();

                    foreach (BaselineCaptureTarget target in targets)
                    {
                        this.thisTargets.Add(target);
                    }

                    this.thisCurrentIndex = current;
                    this.TargetsList.SelectedIndex = current;
                }
                finally
                {
                    this.thisShowingSelection = false;
                }
            }

            // Each line's label too: a baseline's pin text or signal name changed in the pane is the line's now (BaselineRunPins.Follow).
            foreach (BaselineCaptureTarget target in this.thisTargets)
            {
                target.ShowState(this.IsNextTarget(target));
                target.ShowLabel();
            }

            this.ShowComponentNames();

            // The case drawn as the pane's package and pins now read - a pin added there is one to click (ShowCase frames it again).
            pane.ShowCaseOn(this.CasePinout);
            this.ShowCurrent();
            this.ShowProgress();

            // Another line shown in place of one gone - its baseline deleted, or given another pin, in the pane: the oscilloscope set as
            // ITS baseline was measured (code review, 2026-11-22: Enter captured it with the settings of the line before).
            if (this.CurrentForTests is { Row: { } shownRow } nowShown && !ReferenceEquals(nowShown, shown))
            {
                this.SetScopeFor(shownRow);
            }
        }

        // ---- closing ----------------------------------------------------------------------------

        // "Close capturing" (owner request, 2026-11-20) - as Escape: captures not yet accepted are asked about first.
        private void OnCloseCaptureClick(object? sender, RoutedEventArgs e) => this.Close();

        // ###########################################################################################
        // Closed by the pane - another component or region shown, or the pane letting go: captures not yet accepted are let go of,
        // asked about by nobody (they are in no draft, and the component they were of is no longer shown).
        // ###########################################################################################
        internal void LetGo()
        {
            this.thisLettingGo = true;
            this.Close();
        }

        // The close asked for while a capture went into the draft, now that it is in - unless the window closed since, or another
        // capture is going in (which holds it again).
        private void CloseAsWanted()
        {
            if (!this.thisCloseWanted || this.thisClosed || this.thisPutting)
            {
                return;
            }

            this.thisCloseWanted = false;
            this.Close();
        }

        // ###########################################################################################
        // Closing - "Close capturing", Escape, the window's own close, AND CRT QUITTING - with captures not yet accepted is held, and
        // the panel over the buttons asks: accept them, discard them, or keep capturing. CRT asks this window first as it quits
        // (Avalonia asks an owned window before its owner's Closing), so held here, the quitting is called off (code review,
        // 2026-11-22: quitting deleted them without a word, every pin to be probed again). Closed by the pane - another board, the
        // draft let go of, the component gone - they are let go of, once CLOSED (OnClosed), never as it is closing (code review,
        // 2026-11-21: a close called off left the window open with its captures gone).
        // ###########################################################################################
        private void OnClosing(object? sender, WindowClosingEventArgs e)
        {
            bool asksFirst = !this.thisLettingGo && e.CloseReason is WindowCloseReason.WindowClosing or WindowCloseReason.OwnerWindowClosing;

            // A capture going into the draft finishes first - the window's own close is covered by no "please wait" (code review,
            // 2026-11-21) - and the window closes once it is in (thisCloseWanted, AcceptCaptureAsync). The pane letting go closes it
            // at once: the capture is let go of as the put ends.
            if (this.thisPutting && asksFirst)
            {
                e.Cancel = true;
                this.thisCloseWanted = true;
                return;
            }

            int waiting = this.WaitingTargets().Count;

            if (asksFirst && waiting > 0)
            {
                e.Cancel = true;

                // Asked as CRT quits: brought before CRT's window, which had the front.
                if (e.CloseReason == WindowCloseReason.OwnerWindowClosing)
                {
                    this.Activate();
                }

                this.ShowCloseQuestion(waiting);
                return;
            }

            this.RememberPlace();
        }

        // The panel over the buttons asking about `waiting` captures not yet accepted - accept, discard, or keep capturing.
        private void ShowCloseQuestion(int waiting)
        {
            BoldRuns.ShowMarked(this.NotAcceptedText, ComponentPaneWording.RunNotAcceptedOnClose(waiting));
            this.NotAcceptedPanel.IsVisible = true;

            // The keyboard on the safe answer: Enter - the footswitch too - keeps capturing.
            this.KeepCapturingButton.Focus();
        }

        // Whether captures not yet accepted wait in it - the pane stays on their component while any does (ComponentPane
        // .KeepsWaitingCaptures).
        internal bool HoldsCaptures => !this.thisClosed && this.thisTargets.Any(target => target.Waiting is not null);

        // ###########################################################################################
        // Asked before the pane leaves their component (code review, 2026-11-22): brought to the front, asking what its own close asks.
        // A capture going into the draft finishes first, under its "please wait" - asked then, the question would go with it.
        // ###########################################################################################
        internal void AskAboutWaitingCaptures()
        {
            this.Activate();

            int waiting = this.WaitingTargets().Count;

            if (!this.thisPutting && waiting > 0)
            {
                this.ShowCloseQuestion(waiting);
            }
        }

        // Closed: every capture not accepted goes from the user's folder, and one still being taken goes as it comes (CaptureCurrentAsync).
        private void OnClosed(object? sender, EventArgs e)
        {
            this.thisClosed = true;

            // The one being copied into the draft is let go of once that ends (AcceptCaptureAsync).
            foreach (BaselineCaptureTarget target in this.WaitingTargets().Where(target => !ReferenceEquals(target, this.thisPuttingTarget)))
            {
                target.LetGoOfCapture();
            }
        }

        private List<BaselineCaptureTarget> WaitingTargets() => this.thisTargets.Where(target => target.Waiting is not null).ToList();

        private void OnKeepCapturingClick(object? sender, RoutedEventArgs e) => this.NotAcceptedPanel.IsVisible = false;

        private void OnDiscardAndCloseClick(object? sender, RoutedEventArgs e) => this.LetGo();

        private async void OnAcceptAndCloseClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AcceptWaitingAndCloseAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Accepting the captures before closing failed - [{ex}]");
            }
        }

        // "Accept and close": every capture not yet accepted into the draft, one step each - the first that cannot be shown, with
        // why, and the window kept open.
        internal async Task AcceptWaitingAndCloseAsync()
        {
            // One at a time: a key or a click meanwhile starts no other save or capture (code review, 2026-11-21).
            if (this.thisWorking)
            {
                return;
            }

            this.NotAcceptedPanel.IsVisible = false;
            this.thisWorking = true;

            try
            {
                foreach (BaselineCaptureTarget target in this.WaitingTargets())
                {
                    if (!await this.AcceptCaptureAsync(target))
                    {
                        // Shown on that pin WITH why (code review, 2026-11-21: going to it hid the reason, and the button seemed
                        // to do nothing).
                        string? why = this.ProblemText.IsVisible ? this.ProblemText.Text : null;
                        this.GoTo(this.thisTargets.IndexOf(target), setScope: false);
                        this.SayProblem(why);
                        return;
                    }
                }
            }
            finally
            {
                this.thisWorking = false;
            }

            this.LetGo();
        }

        // ---- picking ----------------------------------------------------------------------------

        private void OnPickWithBaselineClick(object? sender, RoutedEventArgs e) => this.PickPins(BaselineRunPins.WithBaseline);

        private void OnPickWithoutBaselineClick(object? sender, RoutedEventArgs e) => this.PickPins(BaselineRunPins.WithoutBaseline);

        private void OnPickEveryPinClick(object? sender, RoutedEventArgs e) => this.PickPins(BaselineRunPins.EveryPin);

        private void OnPickNoPinsClick(object? sender, RoutedEventArgs e) => this.PickPins(_ => []);

        // ###########################################################################################
        // A quick pick: exactly the pins it names - and the first of them shown, the oscilloscope set for it (owner request, 2026-11-20:
        // "it should automatically jump to the first entry in that scope"); with none picked, the pin shown stays. Whether a pin has a
        // baseline is asked of its lines as they are NOW (code review, 2026-11-21: a pin given one in this run was still one "without").
        // ###########################################################################################
        internal void PickPins(Func<IEnumerable<BaselineRunPin>, IReadOnlyList<BaselineRunPin>> pick)
        {
            if (this.IsBusy)
            {
                return;
            }

            List<BaselineRunPin> pins = this.thisTargets
                .GroupBy(target => target.Step.Pin.Pin, StringComparer.OrdinalIgnoreCase)
                .Select(lines => lines.First().Step.Pin with { Baselines = lines.Count(target => target.Row is not null) })
                .ToList();
            var picked = pick(pins).Select(pin => pin.Pin).ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (BaselineCaptureTarget target in this.thisTargets)
            {
                target.SetPicked(picked.Contains(target.Step.Pin.Pin));
            }

            this.OnPickedChanged();

            int first = this.thisTargets.ToList().FindIndex(target => target.Step.IsPicked);

            if (first >= 0)
            {
                this.GoTo(first, setScope: true);
            }
        }

        private void OnPickedChanged()
        {
            foreach (BaselineCaptureTarget target in this.thisTargets)
            {
                target.ShowState(this.IsNextTarget(target));
            }

            this.ShowCase();
            this.ShowProgress();
            this.ShowNext();
            this.ShowButtons();
        }

        // A pin clicked on the case: shown, and picked.
        private void OnCasePinPicked(object? sender, PinPickedEventArgs e)
        {
            e.Handled = true;

            if (!this.IsBusy)
            {
                this.GoToPin(e.Pin.ToString(System.Globalization.CultureInfo.InvariantCulture), pickIt: true);
            }
        }

        // The pin `typed` names shown - false when no line is of it. Picked (a pin clicked on the case), EVERY line of the pin is
        // picked - a pin measured twice has two (code review, 2026-11-21: the case framed the pin and the second was never visited).
        internal bool GoToPin(string typed, bool pickIt)
        {
            int index = BaselineRunPins.IndexOfPin(this.Steps(), typed);

            if (index < 0)
            {
                return false;
            }

            if (pickIt)
            {
                string pin = this.thisTargets[index].Step.Pin.Pin;
                bool changed = false;

                foreach (BaselineCaptureTarget target in this.thisTargets.Where(target => ComponentVocabulary.IsSamePin(target.Step.Pin.Pin, pin)))
                {
                    changed |= target.SetPicked(true);
                }

                if (changed)
                {
                    this.OnPickedChanged();
                }
            }

            this.GoTo(index, setScope: true);
            return true;
        }

        // ###########################################################################################
        // The line of `row` shown and picked, the oscilloscope set as it was measured - the pane's "Capture again" on it (owner request,
        // 2026-11-20). The lines follow the pane's baselines first, so every baseline it has - one with no image too - has its line
        // (code review, 2026-11-21). False when no line is of it: a row the pane has let go of.
        // ###########################################################################################
        internal bool GoToBaseline(ComponentPaneBaselineRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            this.FollowRows();

            int index = this.thisTargets.ToList().FindIndex(target => ReferenceEquals(target.Row, row));

            if (index < 0)
            {
                return false;
            }

            if (!this.thisTargets[index].Step.IsPicked)
            {
                this.thisTargets[index].SetPicked(true);
                this.OnPickedChanged();
            }

            this.GoTo(index, setScope: true);
            this.Activate();
            return true;
        }

        // ---- the pin shown ----------------------------------------------------------------------

        private List<BaselineRunStep> Steps() => this.thisTargets.Select(target => target.Step).ToList();

        private void OnTargetSelected(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisShowingSelection || this.TargetsList.SelectedIndex < 0 || this.TargetsList.SelectedIndex == this.thisCurrentIndex)
            {
                return;
            }

            // Busy: the line shown stays the list's selection.
            if (this.IsBusy)
            {
                this.thisShowingSelection = true;

                try
                {
                    this.TargetsList.SelectedIndex = this.thisCurrentIndex;
                }
                finally
                {
                    this.thisShowingSelection = false;
                }

                return;
            }

            this.GoTo(this.TargetsList.SelectedIndex, setScope: true);
        }

        private void OnPreviousClick(object? sender, RoutedEventArgs e)
        {
            if (!this.IsBusy)
            {
                this.GoBy(-1);
            }
        }

        private async void OnAcceptClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AcceptAndContinueAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Accepting a pin's oscilloscope baseline failed - [{ex}]");
            }
        }

        private async Task AcceptFromKeyAsync()
        {
            try
            {
                await this.AcceptAndContinueAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Accepting a pin's oscilloscope baseline failed - [{ex}]");
            }
        }

        // The picked line before (-1) or after (+1) the one shown - accepting and un-accepting nothing.
        internal void GoBy(int direction)
        {
            int index = BaselineRunPins.Neighbour(this.Steps(), this.thisCurrentIndex, direction);

            if (index >= 0)
            {
                this.GoTo(index, setScope: true);
            }
        }

        // ###########################################################################################
        // "Accept and continue to next pin": the capture of the pin shown accepted - into the draft, ONE step, saved - and the next
        // picked pin not accepted yet shown, the oscilloscope set for it; with none left it stays, saying so. With nothing captured it
        // only goes on; a capture that cannot be accepted (no signal name, a save refused) stays, with why.
        // ###########################################################################################
        internal async Task AcceptAndContinueAsync()
        {
            if (this.thisWorking || this.thisCurrentIndex < 0 || this.thisCurrentIndex >= this.thisTargets.Count)
            {
                return;
            }

            this.thisWorking = true;

            try
            {
                this.CommitName();

                BaselineCaptureTarget accepting = this.thisTargets[this.thisCurrentIndex];

                if (!await this.AcceptCaptureAsync(accepting))
                {
                    return;
                }

                // From where its line stands now - the lines followed the draft as it went in, and may have moved (code review,
                // 2026-11-21).
                int index = this.thisTargets.IndexOf(accepting);
                int next = BaselineRunPins.NextToAccept(this.Steps(), index >= 0 ? index : this.thisCurrentIndex);

                this.OnPickedChanged();

                // The window closing as asked while the capture went in: the oscilloscope is not set for a pin nobody will see.
                if (next >= 0 && !this.thisCloseWanted)
                {
                    this.GoTo(next, setScope: true);
                }
                else
                {
                    this.ShowCurrent();
                }
            }
            finally
            {
                this.thisWorking = false;
            }
        }

        // ###########################################################################################
        // `target`'s capture, not yet accepted, into its baseline - or a new one, named as its signal name says and in the machine state
        // the box says - as ONE step of the draft, saved (ComponentPane.PutCaptureAsync). True when it is in, or there was none.
        // ###########################################################################################
        private async Task<bool> AcceptCaptureAsync(BaselineCaptureTarget target)
        {
            if (target.Waiting is not { } capture || this.thisPane is not { } pane)
            {
                return true;
            }

            // The baselines as they are now - a line with a capture waiting is never dropped by it.
            this.FollowRows();

            BaselineRunStep step = target.Step;
            ComponentPaneBaselineRow? row = target.Row;

            if (row is null && step.Name.Trim().Length == 0)
            {
                this.SayProblem(ComponentPaneWording.RunNeedsName);
                return false;
            }

            this.SayProblem(null);

            string? file;
            this.thisPutting = true;
            this.thisPuttingTarget = target;

            try
            {
                file = await pane.PutCaptureAsync(this, capture, row, step.PinText, step.Name.Trim(), this.MachineStateBox.Text?.Trim() ?? string.Empty, this.SayProblem);
            }
            finally
            {
                this.thisPutting = false;
                this.thisPuttingTarget = null;

                // Closed meanwhile by the pane (another component shown): what did not go in goes from the user's folder now.
                if (this.thisClosed)
                {
                    target.LetGoOfCapture();
                }

                // Its own close asked for meanwhile: closed now, once what asked for the put is done with it - "Accept and close"
                // going on to the next capture holds it again (OnClosing), and closes by itself at the end.
                if (this.thisCloseWanted && !this.thisClosed)
                {
                    Dispatcher.UIThread.Post(this.CloseAsWanted);
                }
            }

            if (file is not null)
            {
                target.Accepted(file);
            }

            // What changed meanwhile, and its own baseline, as one: accepted, the line holds the baseline it made.
            if (file is not null || this.thisFollowWanted)
            {
                this.thisFollowWanted = false;
                this.FollowRows();
            }

            return file is not null;
        }

        // ###########################################################################################
        // The line at `index` shown - its pin lit on the case, its baseline large with every field the popup shows of it, the
        // oscilloscope set as it was measured (`setScope`; a pin with no baseline leaves it as it is).
        // ###########################################################################################
        internal void GoTo(int index, bool setScope)
        {
            if (index < 0 || index >= this.thisTargets.Count)
            {
                return;
            }

            this.CommitName();
            this.thisCurrentIndex = index;
            this.ProblemText.IsVisible = false;

            this.thisShowingSelection = true;

            try
            {
                this.TargetsList.SelectedIndex = index;

                // The lines may have been put in just now (FollowRows) - laid out first, as a list's rows changed since its last layout
                // pass cannot be scrolled to (ListScrolling).
                ListScrolling.ToRow(this.TargetsList, index);
            }
            finally
            {
                this.thisShowingSelection = false;
            }

            this.ShowCurrent();

            if (setScope && this.thisTargets[index].Row is { } row)
            {
                this.SetScopeFor(row);
            }
        }

        private void ShowCurrent()
        {
            if (this.thisCurrentIndex < 0 || this.thisCurrentIndex >= this.thisTargets.Count)
            {
                this.ShowNoLine();
                return;
            }

            BaselineCaptureTarget target = this.thisTargets[this.thisCurrentIndex];
            BaselineRunStep step = target.Step;
            ComponentPaneBaselineRow? row = target.Row;
            string machineState = this.MachineStateBox.Text ?? string.Empty;

            BoldRuns.ShowMarked(this.CurrentText, ComponentPaneWording.RunCurrent(step));
            this.CurrentDetails.Text = ComponentPaneWording.RunStepDetails(step, row?.MachineState ?? machineState);

            // A new baseline's signal name may be typed - it needs one, until its capture is accepted. Written only when the box says
            // another name: shown again as it is typed in (a commit in the pane), a space typed last and the caret stay (code review,
            // 2026-11-21 - CommitName keeps it trimmed).
            this.NamePanel.IsVisible = row is null;

            if (!string.Equals(this.NameBox.Text?.Trim() ?? string.Empty, step.Name.Trim(), StringComparison.Ordinal))
            {
                this.NameBox.Text = step.Name;
            }

            this.ShowFacts(target.WaitingEntry(row, machineState, this.CapturedBaselineRegion) ?? row?.Committed, target.IsComparable ? target.BeforeEntry : null);
            this.ShowPictures(target);
            this.ShowCase();
            this.ShowNext();
            this.ShowButtons();
        }

        // ###########################################################################################
        // No line shown - none left, every pin and baseline taken away in the pane (code review, 2026-11-22: the window kept showing the
        // pin before, its fields, its picture and its name box): nothing of a pin shown, and nothing to capture (ShowButtons).
        // ###########################################################################################
        private void ShowNoLine()
        {
            BoldRuns.ShowMarked(this.CurrentText, string.Empty);
            this.CurrentDetails.Text = string.Empty;
            this.NamePanel.IsVisible = false;
            this.ShowFacts(null, null);
            this.SingleFrame.IsVisible = false;
            this.ComparePanel.IsVisible = false;
            this.ShowImage(this.CurrentImage, null);
            this.ShowImage(this.NowImage, null);
            this.OldImage.Source = null;
            this.ShowCase();
            this.ShowNext();
            this.ShowButtons();
        }

        // ---- the fields, lined up ---------------------------------------------------------------

        // Each field's name and value block, placed once - the grid's rows and columns are the markup's, never changed in code.
        private void CreateFactCells()
        {
            foreach ((string name, int row, int column) in FactPlaces)
            {
                var label = new TextBlock { Text = name };
                var value = new TextBlock();

                Grid.SetRow(label, row);
                Grid.SetColumn(label, column);
                Grid.SetRow(value, row);
                Grid.SetColumn(value, column + 2);

                this.FactsGrid.Children.Add(label);
                this.FactsGrid.Children.Add(value);
                this.thisFactCells[name] = (label, value);
            }
        }

        // ###########################################################################################
        // Every field of `baseline` - the one shown, or a capture not yet accepted as it would be - in its place, the value bold (owner
        // request, 2026-11-20: "do "bold" the data and not the headers ... align the values, so it looks nicer and easier to get an
        // overview") and what it was where it changed since `before`; a field the board does not have (what it applies to, with no
        // regions) not shown. No baseline, no fields.
        // ###########################################################################################
        private void ShowFacts(ComponentImageEntry? baseline, ComponentImageEntry? before)
        {
            this.FactsGrid.IsVisible = baseline is not null;

            if (baseline is null)
            {
                return;
            }

            var facts = ComponentPaneWording.RunFacts(baseline, this.thisPane?.BoardHasRegions == true, before)
                .ToDictionary(fact => fact.Name, fact => fact.Value, StringComparer.Ordinal);

            foreach ((string name, (TextBlock label, TextBlock value)) in this.thisFactCells)
            {
                bool shown = facts.TryGetValue(name, out string? marked);

                label.IsVisible = shown;
                value.IsVisible = shown;

                if (shown)
                {
                    BoldRuns.ShowMarked(value, marked!);
                }
            }
        }

        // ---- the oscilloscope -------------------------------------------------------------------

        // The oscilloscope set as `row` was measured - a capture waits for it.
        private void SetScopeFor(ComponentPaneBaselineRow row)
        {
            if (this.thisPane?.Scope is not { } scope || !this.thisConnected)
            {
                return;
            }

            ComponentImageEntry measured = row.Committed;
            var baseline = new ComponentImageEntry
            {
                BoardLabel = this.thisPane.ShownLabel,
                Region = measured.Region,
                Pin = measured.Pin,
                Name = measured.Name,
                File = measured.File,
                TimeDiv = measured.TimeDiv,
                VoltsDiv = measured.VoltsDiv,
                TriggerLevelVolts = measured.TriggerLevelVolts,
            };

            // ###########################################################################################
            // *** THE LAST PIN SHOWN, NEVER EVERY PIN PASSED (code review, 2026-11-22). *** Each pin went to the oscilloscope as it was
            // shown - Left held, or "1" then "2" typed, queued five to thirty set-ups, and a capture waited for them all while the
            // oscilloscope stepped through pins nobody looked at. One set-up is sent at a time; what waits is only the latest, and a
            // capture waits for that (thisScopeSet).
            // ###########################################################################################
            this.thisScopeWanted = baseline;

            if (this.thisScopeSet.IsCompleted)
            {
                this.thisScopeSet = this.SendScopeWantedAsync(scope);
            }
        }

        // The settings of the pin shown last, waiting for the set-up under way to end - null with none waiting.
        private ComponentImageEntry? thisScopeWanted;

        // The latest pin's settings sent, one set-up after another until none waits - on the UI thread, so the field needs no lock.
        private async Task SendScopeWantedAsync(IBaselineScope scope)
        {
            while (this.thisScopeWanted is { } baseline)
            {
                this.thisScopeWanted = null;
                await BaselineCaptureWindow.SetScopeAsync(scope, baseline);
            }
        }

        private static async Task SetScopeAsync(IBaselineScope scope, ComponentImageEntry baseline)
        {
            try
            {
                await scope.ApplyAsync(baseline, CancellationToken.None);
            }
            catch (Exception ex)
            {
                // Said in the "Oscilloscope" tab's output as every sync is; a capture goes ahead with the oscilloscope as it is.
                Logger.Warning($"Setting the oscilloscope for pin [{baseline.Pin}] of [{baseline.BoardLabel}] failed - [{ex.Message}]");
            }
        }

        // ---- capturing --------------------------------------------------------------------------

        private async void OnCaptureClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.CaptureCurrentAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Capturing a pin's oscilloscope baseline failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The pin shown captured: the oscilloscope's screen and settings, waiting in the user's own folder - NOT YET ACCEPTED, in no
        // draft - and the pin STAYS shown, to be checked: beside its old baseline when it had one. A capture before it, not accepted,
        // goes. A pin captured that was not picked is picked: it is one of the run now.
        // ###########################################################################################
        internal async Task CaptureCurrentAsync()
        {
            if (this.thisWorking || !this.thisConnected || this.thisPane is not { } pane ||
                this.thisCurrentIndex < 0 || this.thisCurrentIndex >= this.thisTargets.Count)
            {
                return;
            }

            this.thisWorking = true;

            try
            {
                this.CommitName();
                this.SayProblem(null);

                BaselineCaptureTarget target = this.thisTargets[this.thisCurrentIndex];
                BaselineRunStep step = target.Step;

                // The oscilloscope set for this pin first, and the old baseline's picture kept before a capture accepted can replace
                // it - under the window's "please wait", as every wait the user watches is (code review, 2026-11-21).
                Task scopeSet = this.thisScopeSet;

                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SettingOscilloscope, async () =>
                {
                    await scopeSet;

                    if (step.IsUpdate && target.BeforeBytes is null)
                    {
                        target.BeforeBytes = await this.ReadOldPictureBytesAsync(step.OriginalFile);
                    }
                });

                ComponentPane.ScopeCapture? capture = await pane.TakeCaptureAsync(this, this.SayProblem);

                if (capture is null)
                {
                    return;
                }

                // Closed while it was being taken - the window's own close, which no "please wait" covers.
                if (this.thisPane is null || this.thisLettingGo || this.thisClosed)
                {
                    ComponentPane.DeleteCapture(capture.Path);
                    return;
                }

                // Its line gone meanwhile - the pane's baselines changed under it (FollowRows): held by a line nobody sees, it was never
                // let go of (code review, 2026-11-21). Not kept, and said.
                if (!this.thisTargets.Contains(target))
                {
                    ComponentPane.DeleteCapture(capture.Path);
                    this.SayProblem(ComponentPaneWording.RunCaptureLost);
                    return;
                }

                target.Hold(capture);
                this.thisPictureReads.Add(this.ShowOnceReadAsync(target, target.WaitingPictureRead!));
                target.SetPicked(true);

                this.OnPickedChanged();
                this.ShowCurrent();
                this.Activate();
            }
            finally
            {
                this.thisWorking = false;
            }
        }

        private void SayProblem(string? problem)
        {
            this.ProblemText.Text = problem ?? string.Empty;
            this.ProblemText.IsVisible = !string.IsNullOrEmpty(problem);
        }
    }
}
