using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.Geometry;
using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // The NOTICES: the cards floating over the tabs' bottom-right corner (NoticeStack in
    // Main.axaml) - an update, data that needs a newer CRT, a data check, launch sync turned off,
    // a draft's board moving under it, a source to switch back from, the label editor's lock. Each
    // card's own part shows and hides it (Main.Updates.cs, Main.DataSyncStatus.cs, ...); this part
    // owns what all of them share. See Main.axaml.cs for the file map. "--demo-notices" plays them
    // all (Main.NoticeDemo.cs).
    //
    // *** THEY FLOAT, THEY DO NOT TAKE ROOM (owner report, 2026-10-11). *** They were rows above
    // the tabs, and every one that came or went shifted the whole window - the launch data check
    // alone did it twice - and on the Schematics tab the image is fitted to its container, so it
    // was resized as well. Never give a notice a row of its own again; MainWindowTests holds the
    // tabs still while each card comes and goes.
    //
    // *** A CARD THAT HAS JUST APPEARED LETS CLICKS THROUGH. *** It appears on its own, at a moment
    // nobody chose, possibly right under a pointer about to click a thumbnail - and that click
    // would land on the card's Install or close button instead. The tooltips taught the same
    // lesson (Handlers/Theme/ToolTipPlacement). For NoticeClickGuard after it appears, a card is
    // not hit-testable, so the press reaches what the user aimed at.
    //
    // *** A CARD WITH NOTHING TO DO GOES BY ITSELF *** after NoticeGoesAfter (HideNoticeLater) -
    // today the data check's "all up to date". Anything new put on the card first cancels that.
    // ###########################################################################################
    public partial class Main
    {
        // How long a card that has just appeared lets clicks through to what is under it.
        internal static readonly TimeSpan NoticeClickGuard = TimeSpan.FromMilliseconds(500);

        // How long a card with nothing to do stays before it goes by itself.
        internal static readonly TimeSpan NoticeGoesAfter = TimeSpan.FromSeconds(8);

        // The pending timer per card, with a token: a card hidden and shown again starts its guard
        // over, and a timer the dispatcher had already queued for the earlier showing must not act
        // on the new one. One dictionary per kind of timer.
        private readonly Dictionary<Control, (object Token, IDisposable Timer)> thisNoticeGuards = new();
        private readonly Dictionary<Control, (object Token, IDisposable Timer)> thisNoticeHides = new();

        // Runs a notice's timer instead of a DispatcherTimer - a headless test's timers never tick.
        // Both kinds come through here; the TimeSpan says which.
        internal Func<Action, TimeSpan, IDisposable>? ScheduleNoticeTimerOverrideForTests { get; set; }

        // ###########################################################################################
        // Sizes the stack to the tabs (NoticeStackGeometry: at most half their width) and guards
        // every card from the moment it is shown. Called once from the constructor, after the
        // markup is loaded - the cards start hidden, so none is showing yet.
        // ###########################################################################################
        private void WireNotices()
        {
            this.NoticeStack.Margin = new Thickness(0, 0, NoticeStackGeometry.EdgeMargin, NoticeStackGeometry.EdgeMargin);
            this.NoticeStack.MaxWidth = NoticeStackGeometry.Widest;
            this.MainTabControl.SizeChanged += (_, e) =>
                this.NoticeStack.MaxWidth = NoticeStackGeometry.MaxWidth(e.NewSize.Width);

            foreach (Control card in this.NoticeStack.Children)
            {
                card.PropertyChanged += (_, e) =>
                {
                    if (e.Property != Visual.IsVisibleProperty)
                    {
                        return;
                    }

                    if (e.NewValue is true)
                    {
                        this.GuardNoticeAgainstStrayClick(card);
                    }
                    else
                    {
                        // Hidden by any route - its close button, a click on it, the flow that owns
                        // it: a hide still pending would take down its NEXT showing early.
                        this.CancelNoticeHide(card);
                    }
                };
            }
        }

        private void GuardNoticeAgainstStrayClick(Control card)
        {
            card.IsHitTestVisible = false;
            this.StartNoticeTimer(this.thisNoticeGuards, card, Main.NoticeClickGuard, () => card.IsHitTestVisible = true);
        }

        // ###########################################################################################
        // Hides `card` NoticeGoesAfter from now - for a notice with nothing to do. Call it after
        // putting the words on the card; anything put on the card later calls CancelNoticeHide.
        // ###########################################################################################
        private void HideNoticeLater(Control card)
        {
            this.StartNoticeTimer(this.thisNoticeHides, card, Main.NoticeGoesAfter, () => card.IsVisible = false);
        }

        private void CancelNoticeHide(Control card)
        {
            if (this.thisNoticeHides.Remove(card, out var pending))
            {
                pending.Timer.Dispose();
            }
        }

        // Starts `after` for `card` in `timers`, replacing one already running there.
        private void StartNoticeTimer(
            Dictionary<Control, (object Token, IDisposable Timer)> timers,
            Control card,
            TimeSpan delay,
            Action after)
        {
            if (timers.Remove(card, out var pending))
            {
                pending.Timer.Dispose();
            }

            object token = new();

            void Elapsed()
            {
                if (timers.TryGetValue(card, out var current) && ReferenceEquals(current.Token, token))
                {
                    timers.Remove(card);
                    after();
                }
            }

            IDisposable timer = this.ScheduleNoticeTimerOverrideForTests is { } schedule
                ? schedule(Elapsed, delay)
                : DispatcherTimer.RunOnce(Elapsed, delay);

            timers[card] = (token, timer);
        }
    }
}
