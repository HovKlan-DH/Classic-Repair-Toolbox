using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.Geometry;
using Handlers.Theming;
using Handlers.Tours;
using System;
using System.Collections.Generic;

namespace CRT
{
    // One step as the overlay shows it: its words, the version badge of a tour of what changed (null otherwise), and
    // where what it is about stands in the overlay right now - null when it is not on screen; a step with no Locate
    // shows its card in the middle. Entering and Leaving set the window up for the step and put it back (the "Schematics"
    // tour's panels shown expanded) - Leaving runs however the step is left, the tour's end included.
    internal sealed record TourCardStep(
        string Headline,
        string Body,
        string? Badge,
        Func<Visual, Rect?>? Locate,
        Action? Entering = null,
        Action? Leaving = null);

    // A whole tour: the first card's words, then the steps.
    internal sealed record TourPresentation(string Kicker, bool IsWhatIsNew, string Headline, string Body, string Foot, IReadOnlyList<TourCardStep> Steps);

    // ###########################################################################################
    // *** AN INTRODUCTION TOUR OVER A WHOLE WINDOW (owner request, 2026-11-11: "showing a how-to use the most important
    // functionalities for this page, highlighting or framing the parts with an elaborative text"). ***
    //
    // A first card says what tour it is ("Introduction tour", or "Tour of what's new") and what the tab is for, with "Skip
    // tour" and "Start tour". Then each step: the window dimmed but for a framed hole round what it is about
    // (TourShade), and a card beside it (TourGeometry) with a dot per step, its words, "Skip tour", "Back" and "Next" -
    // "Done" on the last.
    //
    // *** A VIEW-DEMO ONLY (owner decision, 2026-11-11: "No - this is a view-demo only"). *** Nothing under it can be
    // clicked, the framed part included: the overlay takes every press and every wheel turn, and every key on the
    // window's TUNNEL route, as BusyOverlay does - the keys are its own: Escape skips, Right, Enter, Space and Page Down
    // go on, Left, Backspace and Page Up go back. The focus is its own too (TakeFocus), so no handler under it that
    // takes handled keys sees one.
    //
    // The frame follows what it frames after every layout pass (LayoutUpdated): a window resized, or thumbnails still
    // arriving, move the hole and the card with them.
    //
    // Main's alone, so it stands in Main/ (code review, 2026-11-12: in Controls/, which holds what more than one tab or
    // window uses). It touches nothing of CRT's orchestration all the same - Main.Tours.cs decides when a tour is shown,
    // what it holds, and what is remembered once it ends (Ended).
    // ###########################################################################################
    public partial class TourOverlay : UserControl
    {
        private const double WelcomeWidth = 420;
        private const double StepWidth = 340;

        private TourPresentation? thisPresentation;

        // -1: the first card; 0..n-1: a step.
        private int thisStep = -1;

        // The step whose Entering ran and whose Leaving has not yet.
        private TourCardStep? thisEntered;

        // The window's keys while a tour shows - disposed to give them back.
        private IDisposable? thisKeysTaken;

        public TourOverlay()
        {
            this.InitializeComponent();

            this.AddHandler(PointerPressedEvent, (_, e) => e.Handled = true);
            this.AddHandler(PointerReleasedEvent, (_, e) => e.Handled = true);
            this.AddHandler(PointerWheelChangedEvent, (_, e) => e.Handled = true);
            this.LayoutUpdated += (_, _) => this.Place();
        }

        // Raised once a tour ends: true when its last step was reached ("Done"), false when it was skipped.
        internal event Action<bool>? Ended;

        internal bool IsShowing => this.thisPresentation is not null;

        // -1 for the first card, else the step shown.
        internal int StepShown => this.thisStep;

        // The hole round what the step shown is about, in the overlay - null on the first card or a step in the middle.
        internal Rect? HoleShown => this.Shade.Hole;

        internal Rect CardShown => new(Canvas.GetLeft(this.Card), Canvas.GetTop(this.Card), this.Card.Bounds.Width, this.Card.Bounds.Height);

        // ###########################################################################################
        // Shows a tour from its first card. A tour already showing is replaced.
        // ###########################################################################################
        internal void Show(TourPresentation presentation)
        {
            ArgumentNullException.ThrowIfNull(presentation);

            this.LeaveStep();
            this.thisPresentation = presentation;
            this.thisStep = -1;
            this.IsVisible = true;
            this.TakeKeys();
            this.ShowCard();
            this.TakeFocus();
        }

        // ###########################################################################################
        // THE FOCUS IS THE TOUR'S WHILE IT SHOWS, on its own button (code review, 2026-11-12): marking a key handled stops
        // no handler that takes handled keys too, and those under it acted on the tour's keys - Down moved the caret in the
        // "Feedback" tab's text, Shift changed the "Schematics" tab's hover. With the focus here no key's route passes
        // through them. Given back to what had it as the tour ends, while that is still on screen.
        // ###########################################################################################
        private IInputElement? thisFocusBefore;

        private void TakeFocus()
        {
            if (TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() is IInputElement focused &&
                !ReferenceEquals(focused, this.NextButton) &&
                !ReferenceEquals(focused, this.BackButton) &&
                !ReferenceEquals(focused, this.SkipButton))
            {
                this.thisFocusBefore = focused;
            }

            // Hidden until now, the overlay has not been laid out the first time it shows - its button stands in no tree
            // yet and takes no focus. Then it takes it once laid out.
            if (!this.NextButton.Focus())
            {
                Dispatcher.UIThread.Post(() =>
                {
                    if (this.IsShowing)
                        this.NextButton.Focus();
                }, DispatcherPriority.Loaded);
            }
        }

        private void GiveFocusBack()
        {
            IInputElement? before = this.thisFocusBefore;
            this.thisFocusBefore = null;

            if (before is Control { IsEffectivelyVisible: true, IsEffectivelyEnabled: true } control && TopLevel.GetTopLevel(control) is not null)
                control.Focus();
        }

        internal void Next()
        {
            if (this.thisPresentation is not TourPresentation tour)
                return;

            if (this.thisStep + 1 >= tour.Steps.Count)
            {
                this.End(completed: true);
                return;
            }

            this.thisStep++;
            this.ShowCard();
        }

        // Back from a step to the one before - never from the first step to the first card, as "Back" there is greyed out
        // and the keys say what the buttons do (code review, 2026-11-12: Left on the first step showed the first card).
        internal void Back()
        {
            if (this.thisPresentation is null || this.thisStep <= 0)
                return;

            this.thisStep--;
            this.ShowCard();
        }

        internal void Skip()
        {
            if (this.thisPresentation is not null)
                this.End(completed: false);
        }

        private void End(bool completed)
        {
            this.LeaveStep();
            this.thisPresentation = null;
            this.thisStep = -1;
            this.IsVisible = false;
            this.GiveFocusBack();
            this.GiveKeysBackOnceInputIsTaken();
            this.Ended?.Invoke(completed);
        }

        // ###########################################################################################
        // The keys come back once the input queued by now has been taken (code review, 2026-11-12): a tour ended by a key -
        // Space on "Done" - has that key's typed character still to come, and given back at once, the space went into the
        // text box under the tour on Linux (X11 queues the character as input of its own; Windows drops it after a key
        // marked handled). Input runs above Background, so it comes first. A tour shown again meanwhile keeps them.
        // ###########################################################################################
        private void GiveKeysBackOnceInputIsTaken()
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (!this.IsShowing)
                    this.GiveKeysBack();
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // The card for where the tour stands: the first card, or a step.
        // ###########################################################################################
        private void ShowCard()
        {
            if (this.thisPresentation is not TourPresentation tour)
                return;

            bool welcome = this.thisStep < 0;
            TourCardStep? step = welcome ? null : tour.Steps[this.thisStep];

            if (!ReferenceEquals(step, this.thisEntered))
            {
                this.LeaveStep();
                this.thisEntered = step;
                step?.Entering?.Invoke();
            }

            this.Card.Width = welcome ? TourOverlay.WelcomeWidth : TourOverlay.StepWidth;
            this.Card.Classes.Set("Welcome", welcome);

            this.KickerRow.IsVisible = welcome;
            this.KickerIcon.Text = tour.IsWhatIsNew ? "" : "";
            this.KickerText.Text = tour.Kicker;

            this.ProgressRow.IsVisible = !welcome;
            this.ShowProgress(this.thisStep, tour.Steps.Count);

            this.NewBadge.IsVisible = step?.Badge is not null;
            this.NewBadgeText.Text = step?.Badge ?? string.Empty;

            this.HeadlineText.Text = welcome ? tour.Headline : step!.Headline;
            // A step's words may say their own bold pieces, "<b>Shift</b>" (owner's texts, 2026-11-11).
            BoldRuns.ShowMarked(this.BodyText, welcome ? tour.Body : step!.Body);
            BoldRuns.ShowMarked(this.FootText, tour.Foot);
            this.FootText.IsVisible = welcome && tour.Foot.Length > 0;

            this.SkipButton.Content = TourWording.SkipButton;
            this.BackButton.Content = TourWording.BackButton;
            this.BackButton.IsVisible = !welcome;
            this.BackButton.IsEnabled = this.thisStep > 0;
            this.NextButton.Content = welcome
                ? TourWording.StartButton
                : this.thisStep == tour.Steps.Count - 1 ? TourWording.DoneButton : TourWording.NextButton;

            this.Place();
        }

        // The step entered left: what its Entering set up put back.
        private void LeaveStep()
        {
            if (this.thisEntered is not TourCardStep entered)
                return;

            this.thisEntered = null;
            entered.Leaving?.Invoke();
        }

        private void ShowProgress(int step, int of)
        {
            this.ProgressDots.Children.Clear();

            for (int index = 0; index < of; index++)
            {
                var dot = new Ellipse();
                dot.Classes.Add("TourDot");

                if (index <= step)
                    dot.Classes.Add("Done");

                this.ProgressDots.Children.Add(dot);
            }

            this.ProgressText.Text = step < 0 ? string.Empty : TourWording.StepCounter(step + 1, of);
        }

        // ###########################################################################################
        // The hole round what the step is about, and the card beside it - again after every layout pass, so both follow
        // a window resized or a target that moved. Setting what is already set changes nothing, so no pass loops.
        // ###########################################################################################
        private void Place()
        {
            if (this.thisPresentation is not TourPresentation tour)
                return;

            Size area = this.Bounds.Size;

            if (area.Width <= 0 || area.Height <= 0)
                return;

            TourCardStep? step = this.thisStep < 0 ? null : tour.Steps[this.thisStep];
            Rect? hole = TourGeometry.HoleFor(step?.Locate?.Invoke(this), area);

            this.Card.Measure(Size.Infinity);
            (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, this.Card.DesiredSize, area);

            Canvas.SetLeft(this.Card, card.X);
            Canvas.SetTop(this.Card, card.Y);

            if (this.Shade.Width != area.Width)
                this.Shade.Width = area.Width;

            if (this.Shade.Height != area.Height)
                this.Shade.Height = area.Height;

            this.Shade.Show(hole, card, side);
        }

        // ###########################################################################################
        // Every key on the window is the tour's while it shows - taken on the tunnel route, before anything under it
        // sees one, typed characters too - by OverlayCover.TakeKeys, the one copy BusyOverlay and UpdateRequiredOverlay
        // use (code review, 2026-11-12), with the tour's own keys read beside it.
        // ###########################################################################################
        private void TakeKeys()
        {
            if (this.thisKeysTaken is not null || TopLevel.GetTopLevel(this) is not TopLevel window)
                return;

            window.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel, handledEventsToo: true);
            IDisposable taken = OverlayCover.TakeKeys(window);

            this.thisKeysTaken = new KeysGivenBack(() =>
            {
                taken.Dispose();
                window.RemoveHandler(KeyDownEvent, this.OnWindowKeyDown);
            });
        }

        private void GiveKeysBack()
        {
            IDisposable? taken = this.thisKeysTaken;
            this.thisKeysTaken = null;
            taken?.Dispose();
        }

        private sealed class KeysGivenBack(Action giveBack) : IDisposable
        {
            public void Dispose() => giveBack();
        }

        // The tour's keys. Marked handled here too: the Escape that ends the tour gives the keys back on its way down, and
        // must not go on to whatever has the focus under it.
        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            // While the color dropper picks, the keys are its own - its Escape must not skip the tour too (CLAUDE.md: a
            // handler taking a key whatever is handled asks IsAnyPicking; code review, 2026-11-12).
            if (ColorDropperButton.IsAnyPicking)
                return;

            e.Handled = true;

            switch (e.Key)
            {
                case Key.Escape:
                    this.Skip();
                    break;

                case Key.Right or Key.Enter or Key.Space or Key.PageDown:
                    this.Next();
                    break;

                case Key.Left or Key.Back or Key.PageUp:
                    this.Back();
                    break;
            }
        }

        private void OnSkipClick(object? sender, RoutedEventArgs e) => this.Skip();

        private void OnBackClick(object? sender, RoutedEventArgs e) => this.Back();

        private void OnNextClick(object? sender, RoutedEventArgs e) => this.Next();
    }
}
