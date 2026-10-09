using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Media;
using Avalonia.Threading;

namespace CRT
{
    // ###########################################################################################
    // "PLEASE WAIT" OVER A WHOLE WINDOW - the one way CRT shows that the user has
    // to wait (owner decisions: 2026-09-27 for pushing a board back from BETA, "please dim
    // everything or alike, so it is visible for the user he should wait until it finishes"; and
    // 2026-09-28, "I want this method everywhere in the entire project where there is a Wait").
    //
    // While work runs under it:
    //   - it takes every click AND every key from the first moment (a focused button would
    //     otherwise still take Enter), and shows the busy pointer;
    //   - once the wait has lasted RevealAfter it dims: the rest of the window FADES (its opacity
    //     drops - a dark tint alone all but vanished in the dark theme) and a card in the middle
    //     says what is happening, with a moving bar and how long it has been;
    //   - it ends the moment the work ends, however it ends - or when WaitLimit gives up after two
    //     minutes with nothing happening, and the caller then says what became of the work.
    //
    // *** WHY IT WAITS RevealAfter BEFORE DIMMING. *** Most waits are a fraction of a second, and a
    // window that dims and brightens again on every click reads as flicker, not as information.
    // Input is blocked from the start regardless, so nothing can be started in that moment.
    //
    // *** ONE PER WINDOW, FOUND FROM ANY CONTROL IN IT (For). *** A tab or a panel asks for its
    // window's overlay rather than owning one, so a wait anywhere in a window dims all of it. Work
    // started while one is already running NESTS: the newer sentence shows until it ends, and the
    // window is released only when the outermost wait ends.
    // ###########################################################################################
    public partial class BusyOverlay : UserControl
    {
        // How long a wait runs before the window dims.
        public static readonly TimeSpan RevealAfter = TimeSpan.FromMilliseconds(300);

        // What the rest of the window fades to once revealed.
        internal const double Fade = 0.35;

        private static readonly IBrush ScrimBrush = new SolidColorBrush(Color.FromArgb(0x33, 0, 0, 0));

        // The sentence of every wait running, innermost last.
        private readonly List<string> thisMessages = [];

        // What each faded control's opacity was, so ending the wait puts back exactly that.
        private readonly Dictionary<Control, double> thisFaded = [];

        // Updates the elapsed time on the card. It does NOT decide when to dim: polled at 250 ms, a
        // 300 ms RevealAfter dimmed at about 500 ms (code review, 2026-09-30) - thisPendingReveal does.
        private readonly DispatcherTimer thisTimer;
        private readonly Stopwatch thisElapsed = new();

        // The one-shot that dims the window RevealAfter into the outermost wait; disposed when that
        // wait ends, so it can never dim the next one early.
        private IDisposable? thisPendingReveal;

        // Bumped on every outermost start, so a report posted late by a wait that has ended can
        // never change the sentence of the next one.
        private int thisGeneration;
        private bool thisRevealed;
        private TopLevel? thisBlockedTopLevel;

        public BusyOverlay()
        {
            this.InitializeComponent();

            this.thisTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            this.thisTimer.Tick += (_, _) => this.OnTick();
        }

        // Whether a wait is running (input is blocked from this moment).
        public bool IsBusy => this.thisMessages.Count > 0;

        // Whether the window has dimmed and the card is showing.
        public bool IsRevealed => this.thisRevealed;

        // The sentence on the card.
        public string Message => this.MessageText.Text ?? string.Empty;

        // Replaces the two-minute clock for a test, which can then say exactly when it passes
        // (see WaitLimit.RunAsync's `limitPassed`).
        internal Func<CancellationToken, Task>? LimitOverrideForTests { get; set; }

        // Replaces the dispatcher's one-shot timer that dims the window, for a test: it is handed
        // the reveal and its delay, and runs the reveal when it chooses.
        internal Func<Action, TimeSpan, IDisposable>? ScheduleRevealOverrideForTests { get; set; }

        // ###########################################################################################
        // Runs `work` under this overlay. See WaitLimit for the two-minute rule and what TimedOut
        // means; see the class header for what is shown.
        // ###########################################################################################
        public async Task<WaitResult<T>> RunAsync<T>(string message, Func<WaitContext, Task<T>> work)
        {
            ArgumentNullException.ThrowIfNull(work);

            int generation = this.Begin(message);

            try
            {
                return await WaitLimit.RunAsync(
                    work,
                    this.LimitOverrideForTests,
                    report => Dispatcher.UIThread.Post(() => this.Apply(report, generation)));
            }
            finally
            {
                this.End();
            }
        }

        public async Task<WaitResult<bool>> RunAsync(string message, Func<WaitContext, Task> work)
        {
            ArgumentNullException.ThrowIfNull(work);

            return await this.RunAsync<bool>(message, async context =>
            {
                await work(context);
                return true;
            });
        }

        // ###########################################################################################
        // Holds the window across SEVERAL steps - a decision, then the lists read again - each of
        // which runs under its own RunAsync and so its own two-minute limit. Without it the window
        // brightened for a moment between two steps and dimmed again, which reads as done and then
        // not done. This adds no limit of its own: the steps inside have theirs.
        // ###########################################################################################
        public async Task HoldAsync(string message, Func<Task> steps)
        {
            ArgumentNullException.ThrowIfNull(steps);

            this.Begin(message);

            try
            {
                await steps();
            }
            finally
            {
                this.End();
            }
        }

        public static Task HoldAsync(StyledElement anchor, string message, Func<Task> steps) =>
            BusyOverlay.For(anchor) is BusyOverlay overlay
                ? overlay.HoldAsync(message, steps)
                : steps();

        // ###########################################################################################
        // The overlay of the window `element` is in - found by walking up to the window and down to
        // its one BusyOverlay, so a tab needs no reference to it. Null when the window has none.
        // ###########################################################################################
        public static BusyOverlay? For(StyledElement? element)
        {
            StyledElement? root = element;

            while (root?.Parent is StyledElement parent)
                root = parent;

            if (root is null)
                return null;

            return root as BusyOverlay ?? root.GetLogicalDescendants().OfType<BusyOverlay>().FirstOrDefault();
        }

        // ###########################################################################################
        // Runs `work` under the overlay of `anchor`'s window. A window without one still gets the
        // two-minute limit - it just shows nothing - so a missing host can never make work hang.
        // ###########################################################################################
        public static Task<WaitResult<T>> RunAsync<T>(StyledElement anchor, string message, Func<WaitContext, Task<T>> work) =>
            BusyOverlay.For(anchor) is BusyOverlay overlay
                ? overlay.RunAsync(message, work)
                : WaitLimit.RunAsync(work);

        public static Task<WaitResult<bool>> RunAsync(StyledElement anchor, string message, Func<WaitContext, Task> work) =>
            BusyOverlay.For(anchor) is BusyOverlay overlay
                ? overlay.RunAsync(message, work)
                : WaitLimit.RunAsync(work);

        // ###########################################################################################
        // Work on THIS machine that cannot be stopped halfway - a draft workbook being written, a
        // PDF being rendered. Under the overlay for up to the limit; if it is still going then, the
        // overlay lifts, `stillRunning` is told (to say so), and this carries on waiting for the
        // work ITSELF and returns its real result - "validating if it did finish" for local work,
        // which has no server to ask afterwards.
        //
        // The work is given no token: stopping a half-written workbook would be worse than waiting.
        // ###########################################################################################
        public static async Task<T> RunLocalAsync<T>(StyledElement anchor, string message, Func<Task<T>> work, Action? stillRunning = null)
        {
            ArgumentNullException.ThrowIfNull(work);

            WaitResult<T> waited = await BusyOverlay.RunAsync(anchor, message, _ => work());

            if (!waited.IsTimedOut)
                return waited.Value!;

            stillRunning?.Invoke();

            return await waited.Work;
        }

        public static Task RunLocalAsync(StyledElement anchor, string message, Func<Task> work, Action? stillRunning = null)
        {
            ArgumentNullException.ThrowIfNull(work);

            return BusyOverlay.RunLocalAsync<bool>(anchor, message, async () =>
            {
                await work();
                return true;
            }, stillRunning);
        }

        // Dims at once, as if RevealAfter had passed - for a test, which does not run the timer.
        internal void RevealNowForTests() => this.Reveal();

        // -----------------------------------------------------------------------------------

        private int Begin(string message)
        {
            this.thisMessages.Add(message);
            this.ShowMessage(message);

            if (this.thisMessages.Count > 1)
                return this.thisGeneration;

            this.thisGeneration++;
            this.thisRevealed = false;
            this.Scrim.Background = Brushes.Transparent;
            this.Card.IsVisible = false;
            this.ElapsedText.Text = string.Empty;
            this.IsVisible = true;

            this.BlockKeyboard();

            this.thisElapsed.Restart();
            this.thisTimer.Start();
            this.ScheduleReveal(this.thisGeneration);

            return this.thisGeneration;
        }

        private void End()
        {
            if (this.thisMessages.Count > 0)
                this.thisMessages.RemoveAt(this.thisMessages.Count - 1);

            // An inner wait ended: the outer one's sentence comes back.
            if (this.thisMessages.Count > 0)
            {
                this.ShowMessage(this.thisMessages[^1]);
                return;
            }

            this.thisTimer.Stop();
            this.CancelPendingReveal();
            this.thisElapsed.Reset();
            this.thisRevealed = false;
            this.Card.IsVisible = false;
            this.Scrim.Background = Brushes.Transparent;
            this.IsVisible = false;

            this.RestoreFaded();
            this.UnblockKeyboard();
        }

        private void ShowMessage(string message)
        {
            this.MessageText.Text = message;
            this.Bar.IsIndeterminate = true;
        }

        // A report from the work: a new sentence, and a position for the bar when it knows one.
        private void Apply(WaitReport report, int generation)
        {
            if (generation != this.thisGeneration || !this.IsBusy)
                return;

            if (!string.IsNullOrWhiteSpace(report.Message))
                this.MessageText.Text = report.Message;

            if (report.Fraction is double fraction)
            {
                this.Bar.IsIndeterminate = false;
                this.Bar.Value = Math.Clamp(fraction, 0, 1) * 100;
            }
            else
            {
                this.Bar.IsIndeterminate = true;
            }
        }

        private void OnTick()
        {
            if (this.IsBusy && this.thisRevealed)
                this.ElapsedText.Text = BusyOverlay.FormatElapsed(this.thisElapsed.Elapsed);
        }

        // Dims the window RevealAfter from now, unless the wait of `generation` has ended by then.
        // The generation check covers a timer the dispatcher had already queued when it was
        // cancelled.
        private void ScheduleReveal(int generation)
        {
            this.CancelPendingReveal();

            void RevealIfStillThisWait()
            {
                if (generation == this.thisGeneration)
                    this.Reveal();
            }

            this.thisPendingReveal = this.ScheduleRevealOverrideForTests is { } schedule
                ? schedule(RevealIfStillThisWait, BusyOverlay.RevealAfter)
                : DispatcherTimer.RunOnce(RevealIfStillThisWait, BusyOverlay.RevealAfter);
        }

        private void CancelPendingReveal()
        {
            this.thisPendingReveal?.Dispose();
            this.thisPendingReveal = null;
        }

        private void Reveal()
        {
            if (!this.IsBusy || this.thisRevealed)
                return;

            this.thisRevealed = true;
            this.Scrim.Background = BusyOverlay.ScrimBrush;
            this.Card.IsVisible = true;
            this.ElapsedText.Text = BusyOverlay.FormatElapsed(this.thisElapsed.Elapsed);

            // Everything else in the host's grid fades - the overlay itself stays at full strength.
            if (this.Parent is Panel host)
            {
                foreach (Control sibling in host.Children)
                {
                    if (ReferenceEquals(sibling, this) || this.thisFaded.ContainsKey(sibling))
                        continue;

                    this.thisFaded[sibling] = sibling.Opacity;
                    sibling.Opacity = BusyOverlay.Fade;
                }
            }
        }

        private void RestoreFaded()
        {
            foreach ((Control control, double opacity) in this.thisFaded)
                control.Opacity = opacity;

            this.thisFaded.Clear();
        }

        // In words - see WaitWording.Elapsed.
        internal static string FormatElapsed(TimeSpan elapsed) => WaitWording.Elapsed(elapsed);

        // ###########################################################################################
        // *** KEYS ARE TAKEN ON THE TUNNEL ROUTE, AT THE WINDOW. *** A click cannot get past the
        // overlay, but a key goes to whatever has focus - so the button that started the work, still
        // focused, would take Enter and start it again. Handled on the way DOWN, before any control
        // sees it.
        // ###########################################################################################
        private void BlockKeyboard()
        {
            if (TopLevel.GetTopLevel(this) is not TopLevel topLevel)
                return;

            this.thisBlockedTopLevel = topLevel;
            topLevel.AddHandler(InputElement.KeyDownEvent, BusyOverlay.Swallow, RoutingStrategies.Tunnel, handledEventsToo: true);
            topLevel.AddHandler(InputElement.KeyUpEvent, BusyOverlay.Swallow, RoutingStrategies.Tunnel, handledEventsToo: true);
            topLevel.AddHandler(InputElement.TextInputEvent, BusyOverlay.SwallowText, RoutingStrategies.Tunnel, handledEventsToo: true);
        }

        private void UnblockKeyboard()
        {
            if (this.thisBlockedTopLevel is not TopLevel topLevel)
                return;

            topLevel.RemoveHandler(InputElement.KeyDownEvent, BusyOverlay.Swallow);
            topLevel.RemoveHandler(InputElement.KeyUpEvent, BusyOverlay.Swallow);
            topLevel.RemoveHandler(InputElement.TextInputEvent, BusyOverlay.SwallowText);
            this.thisBlockedTopLevel = null;
        }

        private static readonly EventHandler<KeyEventArgs> Swallow = (_, e) => e.Handled = true;

        private static readonly EventHandler<TextInputEventArgs> SwallowText = (_, e) => e.Handled = true;
    }
}
