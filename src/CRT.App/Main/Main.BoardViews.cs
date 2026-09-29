using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.OnlineHandling;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // BOARD VIEWS (owner request, 2026-09-27): counting a view each time a published board has been
    // on screen for ten seconds, and sending the views home - see CRT.Data's BoardViewContract.
    //
    // This part is only the three timers. When a board counts is BoardViewTracker's rule, and keeping
    // and sending views is BoardViewReporter's; both are tested. See Main.axaml.cs for the file map.
    //
    //   - LoadSelectedBoardAsync calls NoteBoardOnScreen after every board load, with the board now
    //     on screen or null. Only a PUBLISHED board is counted: a contributor's draft-only system
    //     (and a legacy "_UserContribution" one) never leaves this machine.
    //   - The count timer fires when the board on screen has had its ten seconds.
    //   - The send timer fires BoardViewSendDelay after a view was counted, so several boards in a
    //     row go as one report. It is not restarted by the next view, so a view waits at most that.
    //   - The retry timer (2026-09-27) fires every BoardViewRetryInterval for as long as the window
    //     is open, and sends when anything still waits - a send that found no network, at launch
    //     or later, is tried again without waiting for the next board change or the next start.
    // ###########################################################################################
    public partial class Main
    {
        private readonly BoardViewTracker thisBoardViewTracker = new();
        private DispatcherTimer? thisBoardViewCountTimer;
        private DispatcherTimer? thisBoardViewSendTimer;
        private DispatcherTimer? thisBoardViewRetryTimer;

        // Called once, when the window first opens.
        private void StartBoardViewRetries()
        {
            if (this.thisBoardViewRetryTimer is not null)
                return;

            this.thisBoardViewRetryTimer = new DispatcherTimer { Interval = AppConfig.BoardViewRetryInterval };
            this.thisBoardViewRetryTimer.Tick += this.OnBoardViewRetryTimerTick;
            this.thisBoardViewRetryTimer.Start();
        }

        private void OnBoardViewRetryTimerTick(object? sender, EventArgs e)
        {
            // Nothing waiting is the usual case, and costs nothing: no request is made.
            if (BoardViewReporter.HasWaiting)
                _ = BoardViewReporter.SendWaitingAsync();
        }

        private void NoteBoardOnScreen(HardwareBoardEntry? entry)
        {
            string? systemId = entry is { IsPublished: true }
                ? SystemDescriptorRules.SystemIdFromExcelDataFile(entry.ExcelDataFile)
                : null;

            DateTimeOffset now = DateTimeOffset.UtcNow;

            // The same board re-read: its view, counted or waiting, carries on.
            if (!this.thisBoardViewTracker.Show(systemId, now))
                return;

            this.thisBoardViewCountTimer?.Stop();

            if (this.thisBoardViewTracker.Remaining(now) is TimeSpan left)
                this.StartBoardViewCountTimer(left);
        }

        private void StartBoardViewCountTimer(TimeSpan after)
        {
            if (this.thisBoardViewCountTimer is null)
            {
                this.thisBoardViewCountTimer = new DispatcherTimer();
                this.thisBoardViewCountTimer.Tick += this.OnBoardViewCountTimerTick;
            }

            this.thisBoardViewCountTimer.Stop();

            // A timer cannot be set to nothing; a moment is as good.
            this.thisBoardViewCountTimer.Interval = after > TimeSpan.FromMilliseconds(50) ? after : TimeSpan.FromMilliseconds(50);
            this.thisBoardViewCountTimer.Start();
        }

        private void OnBoardViewCountTimerTick(object? sender, EventArgs e)
        {
            this.thisBoardViewCountTimer?.Stop();

            DateTimeOffset now = DateTimeOffset.UtcNow;

            if (this.thisBoardViewTracker.TakeDue(now) is string systemId)
            {
                BoardViewReporter.Record(systemId, UserSettings.DownloadDataFromTestSource, now);
                this.ScheduleBoardViewSend();
                return;
            }

            // A timer can fire a hair early: wait out the rest.
            if (this.thisBoardViewTracker.Remaining(now) is TimeSpan left)
                this.StartBoardViewCountTimer(left);
        }

        private void ScheduleBoardViewSend()
        {
            if (this.thisBoardViewSendTimer is null)
            {
                this.thisBoardViewSendTimer = new DispatcherTimer { Interval = AppConfig.BoardViewSendDelay };
                this.thisBoardViewSendTimer.Tick += this.OnBoardViewSendTimerTick;
            }

            if (!this.thisBoardViewSendTimer.IsEnabled)
                this.thisBoardViewSendTimer.Start();
        }

        private void OnBoardViewSendTimerTick(object? sender, EventArgs e)
        {
            this.thisBoardViewSendTimer?.Stop();

            // Fire-and-forget: SendWaitingAsync never throws, and nothing here waits on the network.
            _ = BoardViewReporter.SendWaitingAsync();
        }
    }
}
