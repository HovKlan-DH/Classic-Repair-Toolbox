using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Tours;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // *** THE INTRODUCTION TOURS (owner request, 2026-11-11: "can we make the tool itself provide basic introduction per
    // tab, when the user see it the first time? ... It should not be everything it should introduce, but just the most
    // important topics"). *** Every tab but "Resources", "About" and "Maintainer" has one (TourCatalogue; Main.TourTabItems,
    // TabItemsWithoutTour), drawn by Main/TourOverlay over the whole window. The first introduction taken ends on a card
    // saying the other tabs have one too (OtherToursCard).
    //
    //   - A TAB'S TOUR IS OFFERED THE FIRST TIME IT IS SHOWN - to everybody, users updating from CRT 3.x included (owner
    //     decision: "Everyone should see it"): a moment after the window opens for the tab it opens on, else a moment
    //     after the tab is picked (TourAfterOpening, TourAfterTabShown), so the tab is laid out and its lists filled.
    //   - IT WAITS while the first board is still being read, a wait is on (BusyOverlay) or a tool is on (below) - the
    //     "Schematics" tour also while the board's KiCad data is read, for its panels' step - and is not offered over
    //     "CRT has to be updated", nor over another tour.
    //   - ONCE ENDED - "Done", "Skip tour" or Escape - it is remembered (UserSettings.IntroductionToursSeen, the version
    //     TourPlan says), and not offered again until CRT brings something new to the tab: then "Tour of what's new",
    //     with what EVERY version since brought (TourCatalogue's header). Closing CRT in the middle remembers nothing.
    //   - "Take the introduction tours again" on the "Configuration" tab forgets them all and shows the "Schematics" tab,
    //     whose tour starts as it is shown.
    //   - Another tab shown under a tour - by code, as the tour takes every click - puts it away, remembering nothing.
    //   - "--show-tours" shows each tab's tour the first time the tab is shown in the run, whatever was seen, and
    //     remembers nothing; "--show-tours=3.1.1" shows the tour of what came since that version (owner request: "I need
    //     a possibility to enforce this every time when in implementation mode"; SimulationOptions).
    //
    // A step is shown only when what it frames is on screen as the tour starts (TourCatalogue's header) - one with no
    // step left waits for the next time the tab is shown, remembering nothing - and a paragraph about a control not on
    // screen is left out of its step (TourPart.Shows; owner question, 2026-11-11: "that may not include the regions
    // buttons? Then what?"). WHAT A SHOWING LEAVES OUT, OR SAYS ONLY IN PART, IS REMEMBERED AS IT ENDS
    // (UserSettings.IntroductionTourStepsLeftOut) and comes back WHOLE the first time the tab is shown with all of it on
    // screen (TourPlan.LeftOutBefore; code review, 2026-11-12: "Export to PDF" and highlighting a trace never reached a
    // new user whose first look had no workbook, or no KiCad data).
    //
    // It also waits while a tool of the schematic image is on - a picture tool, a worklog's area being marked, "Draw
    // trace" (IsToolOnForTour) - and stops offering once the window closes.
    //
    // NEVER IN A TEST unless asked (LetToursStartForTests): only StartAsync, which no test runs, opens the gate.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // How long after the window opens the first tour comes - the window settled, the launch's notices said.
        internal static readonly TimeSpan TourAfterOpening = TimeSpan.FromSeconds(1);

        // How long after a tab is picked its tour comes - the tab laid out, its lists filled.
        internal static readonly TimeSpan TourAfterTabShown = TimeSpan.FromMilliseconds(400);

        // How many times a tour waiting for the first board, a wait or a picture tool looks again, a second apart - a
        // test's own, so a wait running out can be seen without two minutes of it.
        internal int TourWaitsAtMost { get; set; } = 120;

        private bool thisToursMayStart;

        // The tabs whose tour "--show-tours" has shown in this run.
        private readonly HashSet<TourTab> thisToursShownThisRun = [];

        private IDisposable? thisTourTimer;

        // The tour showing: whose, what it holds, whether "--show-tours" asked for it, and the steps of its plan it leaves
        // out or says only in part - remembered as it ends (TourPlan.LeftOutAfter).
        private (TourTab Tab, TourPlan Plan, bool Forced, IReadOnlyList<string> LeftOut)? thisTourShowing;

        // Whether tours are forced (--show-tours) - the command line's, or a test's.
        internal SimulationOptions TourOptions { get; set; } = SimulationOptions.Current;

        // The gate StartAsync opens, for a test of the tours.
        internal void LetToursStartForTests() => this.thisToursMayStart = true;

        // From the constructor.
        private void WireTours()
        {
            this.TourOverlay.Ended += this.OnTourEnded;
            this.TabConfiguration.TakeToursAgainRequested += this.TakeToursAgain;
            this.Closed += this.OnClosedStopOfferingTours;
            this.Activated += this.OnActivatedOfferTour;
            this.AddHandler(KeyDownEvent, this.OnKeyForTour, RoutingStrategies.Tunnel, handledEventsToo: true);
        }

        // ###########################################################################################
        // A TOUR WAITS FOR THE PERSON TO PAUSE TYPING (code review, 2026-11-12: one that had waited for the board's KiCad
        // data opened in the middle of "U1 6526" typed in the filter - the letters lost, Space pressing "Next", and an
        // Escape meant for the filter skipping the introduction for good). A key pressed in CRT's window - none the tour
        // takes - holds it back TourAfterTyping more. Keys only: a click is how a tab is picked, and its tour comes then.
        // ###########################################################################################
        internal static readonly TimeSpan TourAfterTyping = TimeSpan.FromSeconds(1.5);

        private DateTime thisLastKeyUtc = DateTime.MinValue;

        private void OnKeyForTour(object? sender, KeyEventArgs e)
        {
            if (!this.TourOverlay.IsShowing)
                this.thisLastKeyUtc = DateTime.UtcNow;
        }

        private bool IsTypingForTour => DateTime.UtcNow - this.thisLastKeyUtc < Main.TourAfterTyping;

        // For a test: the last key pressed so long ago.
        internal void LastKeyForTourWasAgoForTests(TimeSpan ago) => this.thisLastKeyUtc = DateTime.UtcNow - ago;

        // ###########################################################################################
        // CRT's window back in front - a dialog over it closed, or the user back from another program: a tour whose wait
        // ran out while the window was NOT in front is offered now, not only at the next tab picked (code review,
        // 2026-11-12). Only that one: offered at every activation, a tour of what an earlier one left out came up in the
        // middle of work as a component's window closed.
        // ###########################################################################################
        private bool thisTourWaitsForTheFront;

        private void OnActivatedOfferTour(object? sender, EventArgs e)
        {
            if (!this.thisTourWaitsForTheFront)
                return;

            this.thisTourWaitsForTheFront = false;
            this.OfferTourSoon(Main.TourAfterTabShown);
        }

        // Whether CRT's window is in front - neither behind a dialog or another program, nor minimized.
        private bool IsInFrontForTour => this.IsActive && this.WindowState != WindowState.Minimized;

        // For a test: the window coming back to the front, and a tour's wait having run out while it was not.
        internal void ComeBackToTheFrontForTests() => this.OnActivatedOfferTour(this, EventArgs.Empty);

        internal void TourWaitedOutBehindForTests() => this.thisTourWaitsForTheFront = true;

        // ###########################################################################################
        // A closed window offers nothing more: a tour waiting looked again every second for up to two minutes, on the
        // tests' shared dispatcher too, keeping the closed window alive (code review, 2026-11-12).
        // ###########################################################################################
        private void OnClosedStopOfferingTours(object? sender, EventArgs e)
        {
            this.Closed -= this.OnClosedStopOfferingTours;
            this.Activated -= this.OnActivatedOfferTour;
            this.thisToursMayStart = false;
            this.thisTourTimer?.Dispose();
            this.thisTourTimer = null;
        }

        // Whether the pending tour timer is running - a test of the window closing.
        internal bool IsTourTimerRunningForTests => this.thisTourTimer is not null;

        // ###########################################################################################
        // From StartAsync: tours may start from now on - the first a moment after the window opens.
        // ###########################################################################################
        private void StartOfferingTours()
        {
            this.thisToursMayStart = true;

            if (this.IsVisible)
            {
                this.OfferTourSoon(Main.TourAfterOpening);
            }
            else
            {
                this.Opened += this.OnOpenedOfferTour;
            }
        }

        private void OnOpenedOfferTour(object? sender, EventArgs e)
        {
            this.Opened -= this.OnOpenedOfferTour;
            this.OfferTourSoon(Main.TourAfterOpening);
        }

        // ###########################################################################################
        // The tab shown's tour in a moment - from the window opening, coming back to the front and every tab picked. A
        // newer ask replaces the one waiting; while it must wait (TourMustWait), it looks again a second later - but only
        // while the tab has a tour to offer at all: one seen, with nothing new, is not waited for (code review, 2026-11-12:
        // every tab picked looked for up to two minutes while KiCad data was read or a tool was on). A wait that runs out
        // shows NOTHING - the tour comes the next time the tab is shown, or, held back by the window not being in front,
        // as it comes back (code review, 2026-11-12: shown anyway, it came over KiCad data still being read and left out
        // its steps).
        // ###########################################################################################
        private void OfferTourSoon(TimeSpan after, int attempt = 0)
        {
            if (!this.thisToursMayStart)
                return;

            this.thisTourTimer?.Dispose();
            this.thisTourTimer = DispatcherTimer.RunOnce(() =>
            {
                this.thisTourTimer = null;

                if (this.TourToOffer() is not { } offer || this.OnlyLeftOutStepsNoneSayableWhole(offer.Plan))
                    return;

                if (this.TourMustWait())
                {
                    if (attempt < this.TourWaitsAtMost)
                    {
                        this.OfferTourSoon(TimeSpan.FromSeconds(1), attempt + 1);
                    }
                    else
                    {
                        this.thisTourWaitsForTheFront = !this.IsInFrontForTour;
                        Logger.Debug($"Introduction tour [{offer.Tab}] waited long enough - offered the next time its tab is shown{(this.thisTourWaitsForTheFront ? ", or CRT's window comes back to the front" : string.Empty)}");
                    }

                    return;
                }

                this.OfferTourNow();
            }, after);
        }

        // ###########################################################################################
        // Whether a tour about to be offered waits a second more: the first board still being read, a wait on, a tool of
        // the schematic image on, CRT's window not in front - a dialog over it, another program, minimized (code review,
        // 2026-11-12: a tour opened behind the worklog editor) - and on the "Schematics" tab the board's KiCad data still
        // being read, or the panels' step would find "Global settings" and "Important KiCad signals" not there yet (owner
        // request, 2026-11-11: "show the 3 panels"). Coming back to the front asks again (OnActivatedOfferTour).
        // ###########################################################################################
        internal bool TourMustWait() =>
            !this.FirstBoardShown.IsCompleted ||
            this.BusyOverlay.IsBusy ||
            this.IsToolOnForTour ||
            !this.IsInFrontForTour ||
            this.IsTypingForTour ||
            this.IsPopupOpenForTour() ||
            this.IsKiCadStillLoadingForTour;

        // For a test that expected a tour: what held it back, by name - null for nothing.
        internal string? TourWaitReasonForTests() =>
            !this.thisToursMayStart ? "gate shut" :
            this.TourOverlay.IsShowing ? "a tour showing" :
            !this.IsVisible ? "window hidden" :
            !this.FirstBoardShown.IsCompleted ? "first board" :
            this.BusyOverlay.IsBusy ? "busy" :
            this.IsToolOnForTour ? "tool" :
            !this.IsActive ? "not active" :
            this.WindowState == WindowState.Minimized ? "minimized" :
            this.IsTypingForTour ? "typing" :
            this.IsPopupOpenForTour() ? $"popup open: {string.Join(", ", Main.OpenPopups.Where(popup => ReferenceEquals(Main.WindowOf(popup), this)).Select(popup => popup.Name ?? popup.TemplatedParent?.GetType().Name ?? popup.PlacementTarget?.GetType().Name ?? "?"))}" :
            this.IsKiCadStillLoadingForTour ? "KiCad loading" :
            this.TourToOffer() is null ? "nothing to offer" :
            null;

        // On the "Schematics" tab the board's KiCad data still being read - what it shows is not all there yet.
        private bool IsKiCadStillLoadingForTour =>
            this.TourTabShown() == TourTab.Schematics && this.TabSchematicsControl.IsKiCadStillLoading;

        // ###########################################################################################
        // A plan of steps an earlier showing left out alone, none of which can be said whole now - nothing to wait for:
        // the steps are about what is not on screen, which no wait brings (code review, 2026-11-12: a board with KiCad
        // data and no important signals looked again every second for two minutes at every visit). Asked again the next
        // time the tab is shown. KiCad data still being read is the one wait that brings controls - it is waited for.
        // ###########################################################################################
        private bool OnlyLeftOutStepsNoneSayableWhole(TourPlan plan) =>
            plan.Kind == TourKind.LeftOut &&
            !this.IsKiCadStillLoadingForTour &&
            !plan.Steps.Any(this.CanSayTourStepWhole);

        // ###########################################################################################
        // A tool of the schematic image on, or asked for and waiting for its image or the KiCad data - a picture tool,
        // marking a worklog's area ("Add worklog", which shows the "Schematics" tab from any tab) or "Draw trace": a tour
        // over it would take the keys the tool is used by, and the tab's own key handler, which sees handled keys too,
        // would end the tool on the Escape that skips the tour (code review, 2026-11-12) - and one waiting would start
        // under the tour.
        // ###########################################################################################
        private bool IsToolOnForTour =>
            this.IsPictureToolOn ||
            this.TabSchematicsControl.IsPictureToolPending ||
            this.TabSchematicsControl.IsWorklogEntryModeActive ||
            this.TabSchematicsControl.IsDrawingTraces ||
            ColorDropperButton.IsAnyPicking;

        // ###########################################################################################
        // A drop-down list, a list of suggestions, a hover card, a flyout (a color picker's) or a context menu open in
        // CRT's window: it stands over the tour, and a click there would pick a board, a device or a color under a tour
        // that takes no clicks (code review, 2026-11-12). Each is a Popup: a drop-down's stands in the window's tree, a
        // flyout's and a context menu's beside it, placed at a control of the window - WindowOf asks both.
        //
        // KEPT AS THEY OPEN AND CLOSE (OpenPopups), never found by walking the window: a tour waiting asked once a second
        // for two minutes, and the walk went through thousands of controls on the "Contribute" tab (code review,
        // 2026-11-12).
        // ###########################################################################################
        private static readonly HashSet<Popup> OpenPopups = [];

        static Main()
        {
            Popup.IsOpenProperty.Changed.AddClassHandler<Popup>((popup, change) =>
            {
                if (change.GetNewValue<bool>())
                {
                    Main.OpenPopups.Add(popup);
                }
                else
                {
                    Main.OpenPopups.Remove(popup);
                }
            });
        }

        private bool IsPopupOpenForTour()
        {
            // One closed or let go of without telling (its window gone) is forgotten here.
            Main.OpenPopups.RemoveWhere(popup => !popup.IsOpen || Main.WindowOf(popup) is null);
            return Main.OpenPopups.Any(popup => ReferenceEquals(Main.WindowOf(popup), this));
        }

        // The window a popup stands in - by its place in the tree, else by the control it is placed at.
        private static TopLevel? WindowOf(Popup popup) =>
            TopLevel.GetTopLevel(popup) ?? (popup.PlacementTarget is Control target ? TopLevel.GetTopLevel(target) : null);

        // ###########################################################################################
        // Shows the tab shown's tour when it has one to show now - true when it did. It asks what the timer waits for,
        // the very same list (TourMustWait - code review, 2026-11-12: two lists meant to agree, and they did not).
        // ###########################################################################################
        internal bool OfferTourNow()
        {
            if (!this.thisToursMayStart || this.TourOverlay.IsShowing || !this.IsVisible || this.TourMustWait())
            {
                return false;
            }

            if (this.TourToOffer() is not { } offer)
                return false;

            (TourTab tab, TourOfTab tour, TourPlan plan, bool forced) = offer;

            // A step an earlier showing left out comes back only WHOLE - shown in part again, it would never be told whole.
            List<TourStep> shown = plan.Steps
                .Where(step => plan.LeftOutBefore.Contains(step) ? this.CanSayTourStepWhole(step) : this.IsTourStepOnScreen(step))
                .ToList();

            if (shown.Count == 0)
            {
                Logger.Debug($"Introduction tour [{tab}] waits: nothing it frames is on screen");
                return false;
            }

            List<TourCardStep> cards = shown
                .Select(step =>
                {
                    (Action? entering, Action? leaving) = this.TourStageOf(step);

                    // A paragraph about a control not on screen is left out - "PAL" and "NTSC" on a board without regions.
                    return new TourCardStep(
                        step.Headline,
                        step.Body.TextFor(this.IsTourTargetOnScreen),
                        plan.IsNew(step) ? TourWording.NewIn(step.Since) : null,
                        step.Targets.Count == 0 ? null : overlay => this.LocateTourTargets(step.Targets, overlay),
                        entering,
                        leaving);
                })
                .ToList();

            // The first introduction taken ends on the other tabs (owner request, 2026-11-11).
            if (plan.Kind == TourKind.Introduction && this.OtherToursCard(tab, forced) is { } otherTours)
            {
                cards.Add(otherTours);
            }

            // What the first card says follows what is shown: a tour of what changed whose new steps are all off screen
            // shows only what an earlier one left out, and says so.
            int leftOutShown = shown.Count(plan.LeftOutBefore.Contains);
            TourKind kindShown = plan.Kind == TourKind.WhatIsNew && leftOutShown == shown.Count ? TourKind.LeftOut : plan.Kind;

            var presentation = new TourPresentation(
                TourWording.Kicker(kindShown),
                kindShown == TourKind.WhatIsNew,
                TourWording.WelcomeHeadline(tour),
                TourWording.WelcomeBody(tour, kindShown, shown.Count - leftOutShown, plan.NewInOf(shown), leftOutShown),
                TourWording.TakeAgainFoot,
                cards);

            // What this showing leaves out, or says only in part, is remembered as it ends - to come back once it can be
            // said whole (code review, 2026-11-12: "Export to PDF" never reached a new user with no workbook yet).
            IReadOnlyList<string> leftOut = plan.LeftOutAfter(shown.Where(this.CanSayTourStepWhole));

            this.thisTourShowing = (tab, plan, forced, leftOut);

            if (forced)
                this.thisToursShownThisRun.Add(tab);

            Logger.Info($"Introduction tour [{tab}] shown - [{plan.Kind}], [{cards.Count}] of [{plan.Steps.Count}] steps, [{leftOut.Count}] left out{(forced ? $", forced by [{SimulationOptions.ShowToursArg}]" : string.Empty)}");
            this.TourOverlay.Show(presentation);
            return true;
        }

        // ###########################################################################################
        // The tab shown's tour and what of it to show - null when it has none to offer: no tour, "CRT has to be updated"
        // over it, shown already in this run by "--show-tours", or seen with nothing new and nothing left out. Pure
        // reading: what is on screen decides later, in OfferTourNow.
        // ###########################################################################################
        private (TourTab Tab, TourOfTab Tour, TourPlan Plan, bool Forced)? TourToOffer()
        {
            if (this.TourTabShown() is not TourTab tab)
                return null;

            // "CRT has to be updated" covers the tab: nothing on it can be used, so nothing is shown of it.
            if (tab == TourTab.Contribute && this.TabDrafts.IsUpdateRequiredShown)
                return null;

            SimulationOptions options = this.TourOptions;
            bool forced = options.ShowTours;

            if (forced && this.thisToursShownThisRun.Contains(tab))
                return null;

            TourOfTab tour = TourCatalogue.Of(tab);
            TourPlan plan = forced
                ? TourPlan.For(tour, options.ToursSeenAs)
                : TourPlan.For(tour, UserSettings.IntroductionTourSeen(tour.TabName), UserSettings.IntroductionTourStepsLeftOut(tour.TabName));

            return plan.Kind == TourKind.None ? null : (tab, tour, plan, forced);
        }

        // ###########################################################################################
        // Another tab shown under a tour - only code can do it, the tour takes every click - puts the tour away and
        // remembers nothing, so it is offered again the next time its tab is shown. Left up, it stood over the other tab.
        // ###########################################################################################
        private void PutTourAwayForAnotherTab()
        {
            if (this.thisTourShowing is not { } showing || this.TourTabShown() == showing.Tab)
                return;

            Logger.Info($"Introduction tour [{showing.Tab}] put away: another tab was shown under it");
            this.thisTourShowing = null;

            if (showing.Forced)
                this.thisToursShownThisRun.Remove(showing.Tab);

            this.TourOverlay.Skip();
        }

        // ###########################################################################################
        // A tour ended - "Done", "Skip tour" or Escape: remembered, unless "--show-tours" asked for it. What it left out
        // comes back only from a tour taken to its end: one SKIPPED was turned down, the steps it left out with it (code
        // review, 2026-11-12: skipped on its first card, it came back as "what an earlier tour left out").
        // ###########################################################################################
        private void OnTourEnded(bool completed)
        {
            if (this.thisTourShowing is not { } showing)
                return;

            this.thisTourShowing = null;
            Logger.Info($"Introduction tour [{showing.Tab}] {(completed ? "taken to its end" : "skipped")}");

            if (!showing.Forced)
            {
                UserSettings.SetIntroductionTourSeen(TourCatalogue.Of(showing.Tab).TabName, showing.Plan.VersionToRemember, completed ? showing.LeftOut : []);
            }
        }

        // ###########################################################################################
        // "Take the introduction tours again" (the "Configuration" tab): every tour as never seen, and the "Schematics" tab
        // shown - its tour comes as it is shown, the others as their tabs are. Pressed on the "Configuration" tab, the
        // only place it is, so showing the "Schematics" tab always changes the tab (code review, 2026-11-12: a branch for
        // the "Schematics" tab already shown could never run).
        // ###########################################################################################
        private void TakeToursAgain()
        {
            UserSettings.ForgetIntroductionToursSeen();
            this.thisToursShownThisRun.Clear();
            Logger.Info("Introduction tours to be taken again");

            this.MainTabControl.SelectedItem = this.SchematicsTabItem;
        }

        // ###########################################################################################
        // What a step makes the window show, and how it is put back as the step is left (TourStage) - never stored: the
        // "Schematics" tab's panels expanded for the step alone (TabSchematics.ExpandPanelsForTour).
        // ###########################################################################################
        private (Action? Entering, Action? Leaving) TourStageOf(TourStep step) => step.Stage switch
        {
            TourStage.SchematicsPanelsExpanded => (this.TabSchematicsControl.ExpandPanelsForTour, this.TabSchematicsControl.PutPanelsBackAfterTour),
            TourStage.ScrolledIntoView => this.ScrolledIntoView(step),
            _ => (null, null),
        };

        // Every tab's TabItem with a tour, by its tour - MainToursTests fails on a tab of the window in neither list.
        internal static readonly IReadOnlyDictionary<TourTab, string> TourTabItems = new Dictionary<TourTab, string>
        {
            [TourTab.Schematics] = "SchematicsTabItem",
            [TourTab.Overview] = "OverviewTabItem",
            [TourTab.Workbooks] = "WorkbooksTabItem",
            [TourTab.Oscilloscope] = "OscilloscopeTabItem",
            [TourTab.Contribute] = "ContributeTabItem",
            [TourTab.Configuration] = "ConfigurationTabItem",
            [TourTab.Feedback] = "FeedbackTabItem",
        };

        // The tabs with no tour, on purpose (owner request, 2026-11-12: "no need to have an introduction for the "About"
        // and "Resources" tabs", and the Maintainer tab "is also much more limited and specialized").
        internal static readonly IReadOnlyList<string> TabItemsWithoutTour = ["ResourcesTabItem", "MaintainerTabItem", "AboutTabItem"];

        // Which tab is shown, by its tour - null for none.
        private TourTab? TourTabShown()
        {
            object? selected = this.MainTabControl?.SelectedItem;

            foreach ((TourTab tab, string item) in Main.TourTabItems)
            {
                if (selected is not null && ReferenceEquals(selected, this.FindTourTarget(item)))
                    return tab;
            }

            return null;
        }

        // ###########################################################################################
        // The last card of the first introduction taken - "go to another tab to see its introduction" (TourWording) -
        // framing the other tabs that have one. Only while no other tab's tour has been seen: then it is known. "--show-tours" asks the
        // run instead, as it remembers nothing.
        // ###########################################################################################
        private TourCardStep? OtherToursCard(TourTab tab, bool forced)
        {
            bool anotherSeen = forced
                ? this.thisToursShownThisRun.Any(shown => shown != tab)
                : TourCatalogue.All.Any(tour => tour.Tab != tab && UserSettings.IntroductionTourSeen(tour.TabName) is not null);

            if (anotherSeen)
                return null;

            List<string> others = Main.TourTabItems.Where(item => item.Key != tab).Select(item => item.Value).ToList();

            return new TourCardStep(
                TourWording.OtherToursHeadline,
                TourWording.OtherToursBody,
                null,
                overlay => this.LocateTourTargets(others, overlay));
        }

        // ###########################################################################################
        // A page scrolled so a step's controls are in view, and scrolled back as it was once the step is left. The LAST is
        // brought into view - the page scrolled no further than it, so a group below the fold that fits shows whole (the
        // "Configuration" tab's tours, problems and folders - its first alone left the rest under the window's edge) - and
        // then the FIRST, which on a window too low for the group scrolls back to it: what the words begin with is seen
        // (code review, 2026-11-12: on a low window the first button stayed above the fold, out of the frame).
        // ###########################################################################################
        private (Action Entering, Action Leaving) ScrolledIntoView(TourStep step)
        {
            ScrollViewer? scrolled = null;
            Vector before = default;

            return (
                () =>
                {
                    if (step.Targets.Count == 0 || this.FindTourTarget(step.Targets[0]) is not Control first)
                        return;

                    scrolled = first.FindAncestorOfType<ScrollViewer>();
                    before = scrolled?.Offset ?? default;
                    (this.FindTourTarget(step.Targets[^1]) ?? first).BringIntoView();

                    // Laid out where the last put it, so the first is asked of where it now stands.
                    this.UpdateLayout();
                    first.BringIntoView();
                },
                () =>
                {
                    if (scrolled is not null)
                        scrolled.Offset = before;

                    scrolled = null;
                });
        }

        // A step is shown while what it frames is on screen - or always, one framing nothing. A step that scrolls its page
        // to what it frames (TourStage.ScrolledIntoView) counts it wherever the page stands: it is brought into view.
        private bool IsTourStepOnScreen(TourStep step) =>
            step.Targets.Count == 0 ||
            this.LocateTourTargets(step.Targets, this.TourOverlay, inViewOnly: step.Stage != TourStage.ScrolledIntoView) is not null;

        internal bool IsTourTargetOnScreen(string path) => this.LocateTourTargets([path], this.TourOverlay) is not null;

        // A step said WHOLE: shown, and every paragraph's control on screen too.
        private bool CanSayTourStepWhole(TourStep step) =>
            this.IsTourStepOnScreen(step) &&
            step.Body.Parts.All(part => part.Shows is null || this.IsTourTargetOnScreen(part.Shows));

        // ###########################################################################################
        // Where the controls a step names stand in `overlay`, all of them in one rectangle - those on screen only, null
        // when none is. ON SCREEN IS IN VIEW: each is cut to what the scrolling panels it stands in show of it, and one
        // scrolled out of them all counts for nothing (code review, 2026-11-12: "Add component" scrolled out of the
        // Contribute tab's left column was framed over the tab strip, and counted as said). `inViewOnly` false counts
        // one wherever its page stands - for a step that scrolls it into view as it starts.
        // ###########################################################################################
        private Rect? LocateTourTargets(IReadOnlyList<string> paths, Visual overlay, bool inViewOnly = true)
        {
            Rect? all = null;

            foreach (string path in paths)
            {
                if (this.FindTourTarget(path) is not Control control ||
                    !control.IsEffectivelyVisible ||
                    control.Bounds.Width <= 0 ||
                    control.Bounds.Height <= 0 ||
                    control.TranslatePoint(new Point(0, 0), overlay) is not Point at)
                {
                    continue;
                }

                var rect = new Rect(at, control.Bounds.Size);

                if (inViewOnly)
                {
                    if (Main.InViewOf(control, rect, overlay) is not Rect seen)
                        continue;

                    rect = seen;
                }

                all = all is Rect sofar ? sofar.Union(rect) : rect;
            }

            return all;
        }

        // What every scrolling panel `control` stands in shows of `rect` (in `overlay`) - null for nothing.
        private static Rect? InViewOf(Control control, Rect rect, Visual overlay)
        {
            Rect seen = rect;

            for (Visual? at = control.GetVisualParent(); at is not null; at = at.GetVisualParent())
            {
                if (at is not ScrollViewer viewer || viewer.TranslatePoint(new Point(0, 0), overlay) is not Point corner)
                    continue;

                seen = seen.Intersect(new Rect(corner, viewer.Bounds.Size));

                if (seen.Width <= 0 || seen.Height <= 0)
                    return null;
            }

            return seen;
        }

        // ###########################################################################################
        // A control a step names: "Name" in this window, "Tab/Name" in the tab of that name - TourCatalogue's paths.
        // ###########################################################################################
        internal Control? FindTourTarget(string path)
        {
            Control? at = this;

            foreach (string name in path.Split('/'))
            {
                at = at?.FindControl<Control>(name);
            }

            return at;
        }
    }
}
