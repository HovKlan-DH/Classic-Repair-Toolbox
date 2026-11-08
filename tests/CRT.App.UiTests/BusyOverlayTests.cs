using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// BusyOverlay - the "please wait" CRT shows over a whole window (owner decisions,
// 2026-09-27 and 2026-09-28: "I want this method everywhere in the entire project where there is a
// Wait").
//
// What must hold: input is blocked from the first moment; the window dims only once the wait has
// lasted long enough to notice; it ALWAYS lifts again - on success, on an exception and when the
// two-minute limit passes - and puts back exactly the opacity everything had.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BusyOverlayTests
{
    // A window laid out the way every host lays one out: the content, then the overlay last,
    // spanning the grid.
    private static (Window Window, Grid Root, Border Content, Button Button, BusyOverlay Overlay) BuildHost()
    {
        var button = new Button { Content = "Start" };
        var content = new Border { Opacity = 0.8, Child = new StackPanel { Children = { button } } };
        var overlay = new BusyOverlay();
        var root = new Grid { Children = { content, overlay } };
        var window = new Window { Content = root, Width = 600, Height = 400 };

        return (window, root, content, button, overlay);
    }

    // Lets the dispatcher run the continuations a finished task queued, until `done` holds - never
    // a fixed sleep. Fails rather than hanging if it never does.
    private static async Task SettleAsync(Func<bool> done)
    {
        for (int attempt = 0; attempt < 500 && !done(); attempt++)
        {
            await Task.Delay(2);
            Dispatcher.UIThread.RunJobs();
        }

        Assert.True(done());
    }

    [Fact]
    public void An_idle_overlay_is_hidden_and_takes_nothing()
    {
        UiTest.Run(() =>
        {
            BusyOverlay overlay = BuildHost().Overlay;

            Assert.False(overlay.IsVisible);
            Assert.False(overlay.IsBusy);
            Assert.False(overlay.IsRevealed);
        });
    }

    // ###########################################################################################
    // *** INPUT IS BLOCKED AT ONCE, THE WINDOW DIMS ONLY ONCE THE WAIT IS NOTICEABLE. *** Most waits
    // are a fraction of a second; dimming for each would flicker. Revealed, the rest of the window
    // fades and the card names the work - and afterwards every opacity is exactly what it was
    // (the content here was 0.8, not 1).
    // ###########################################################################################
    [Fact]
    public async Task While_work_runs_input_is_blocked_then_the_window_dims_and_comes_back_exactly()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, content, _, overlay) = BuildHost();
            var finish = new TaskCompletionSource<int>();

            Task<WaitResult<int>> waiting = overlay.RunAsync("Publishing Commodore/C64/250407 to BETA.", _ => finish.Task);

            Assert.True(overlay.IsVisible);
            Assert.True(overlay.IsBusy);
            Assert.False(overlay.IsRevealed);
            Assert.Equal(0.8, content.Opacity);
            Assert.Equal("Publishing Commodore/C64/250407 to BETA.", overlay.Message);

            overlay.RevealNowForTests();

            Assert.True(overlay.IsRevealed);
            Assert.Equal(BusyOverlay.Fade, content.Opacity);
            Assert.True(overlay.FindControl<Border>("Card")!.IsVisible);

            finish.SetResult(7);
            WaitResult<int> result = await waiting;

            Assert.Equal(7, result.Value);
            Assert.False(overlay.IsVisible);
            Assert.False(overlay.IsBusy);
            Assert.Equal(0.8, content.Opacity);
        });
    }

    // Stands in for the dispatcher's one-shot timer: records each reveal asked for, with its delay,
    // and runs one only when the test says so.
    private sealed class RecordingSchedule
    {
        public List<(Action Reveal, TimeSpan Delay, Cancellation Handle)> Requests { get; } = [];

        public IDisposable Schedule(Action reveal, TimeSpan delay)
        {
            var handle = new Cancellation();
            this.Requests.Add((reveal, delay, handle));
            return handle;
        }

        public sealed class Cancellation : IDisposable
        {
            public bool IsDisposed { get; private set; }

            public void Dispose() => this.IsDisposed = true;
        }
    }

    // ###########################################################################################
    // *** THE WINDOW DIMS AT RevealAfter, NOT AT THE NEXT POLL AFTER IT (code review, 2026-09-30).
    // *** The reveal was checked on the 250 ms elapsed-time tick, so a 300 ms RevealAfter dimmed at
    // about 500 ms - a wait of 300-500 ms never dimmed at all, and every longer one dimmed 200 ms
    // late. It is now scheduled once, for exactly RevealAfter, when the outermost wait begins; a
    // wait nested inside it asks for no second one.
    // ###########################################################################################
    [Fact]
    public async Task The_window_dims_exactly_RevealAfter_into_a_wait_and_only_once_for_nested_waits()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, content, _, overlay) = BuildHost();
            var schedule = new RecordingSchedule();
            overlay.ScheduleRevealOverrideForTests = schedule.Schedule;
            var finishOuter = new TaskCompletionSource<int>();
            var finishInner = new TaskCompletionSource<int>();

            Task<WaitResult<int>> outer = overlay.RunAsync("Outer", _ => finishOuter.Task);
            Task<WaitResult<int>> inner = overlay.RunAsync("Inner", _ => finishInner.Task);

            var request = Assert.Single(schedule.Requests);
            Assert.Equal(BusyOverlay.RevealAfter, request.Delay);
            Assert.False(overlay.IsRevealed);

            request.Reveal();

            Assert.True(overlay.IsRevealed);
            Assert.Equal(BusyOverlay.Fade, content.Opacity);

            finishInner.SetResult(1);
            await inner;
            finishOuter.SetResult(2);
            await outer;

            Assert.True(request.Handle.IsDisposed);
            Assert.Equal(0.8, content.Opacity);
        });
    }

    // ###########################################################################################
    // *** A REVEAL LEFT OVER FROM AN ENDED WAIT NEVER DIMS THE NEXT ONE EARLY. *** A 100 ms wait
    // followed at once by another would otherwise have the first one's timer dim the second only
    // 200 ms in - the flicker RevealAfter exists to prevent.
    // ###########################################################################################
    [Fact]
    public async Task A_reveal_scheduled_for_a_wait_that_has_ended_leaves_the_next_wait_undimmed()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, content, _, overlay) = BuildHost();
            var schedule = new RecordingSchedule();
            overlay.ScheduleRevealOverrideForTests = schedule.Schedule;

            await overlay.RunAsync("First", _ => Task.FromResult(1));

            var finishSecond = new TaskCompletionSource<int>();
            Task<WaitResult<int>> second = overlay.RunAsync("Second", _ => finishSecond.Task);

            Assert.Equal(2, schedule.Requests.Count);
            Assert.True(schedule.Requests[0].Handle.IsDisposed);

            // The first wait's timer fires anyway (a dispatcher that had already queued it).
            schedule.Requests[0].Reveal();

            Assert.False(overlay.IsRevealed);
            Assert.Equal(0.8, content.Opacity);

            schedule.Requests[1].Reveal();

            Assert.True(overlay.IsRevealed);

            finishSecond.SetResult(2);
            await second;
        });
    }

    // ###########################################################################################
    // *** IT ALWAYS LIFTS. *** A refusal, a lost connection or an exception must not leave the user
    // looking at a window that never comes back.
    // ###########################################################################################
    [Fact]
    public async Task An_exception_lifts_the_overlay_and_reaches_the_caller()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, content, _, overlay) = BuildHost();

            await Assert.ThrowsAsync<InvalidOperationException>(() => overlay.RunAsync("Working", async _ =>
            {
                overlay.RevealNowForTests();
                await Task.Yield();
                throw new InvalidOperationException("lost");
            }));

            Assert.False(overlay.IsVisible);
            Assert.Equal(0.8, content.Opacity);
        });
    }

    // ###########################################################################################
    // *** AFTER TWO MINUTES IT LIFTS BY ITSELF (owner decision, 2026-09-28). *** The caller is told
    // it timed out - not that it failed - so it can go and check what became of the work.
    // ###########################################################################################
    [Fact]
    public async Task When_the_limit_passes_the_overlay_lifts_and_the_caller_is_told_it_timed_out()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, _, overlay) = BuildHost();
            var limit = new TaskCompletionSource();
            overlay.LimitOverrideForTests = _ => limit.Task;

            Task<WaitResult<int>> waiting = overlay.RunAsync("Waiting", _ => new TaskCompletionSource<int>().Task);

            Assert.True(overlay.IsBusy);

            limit.SetResult();
            WaitResult<int> result = await waiting;

            Assert.True(result.IsTimedOut);
            Assert.False(overlay.IsVisible);
        });
    }

    // A wait started inside another shows its own sentence, and the window is released only when
    // the outermost one ends.
    [Fact]
    public async Task A_wait_inside_another_shows_its_sentence_and_the_outer_one_keeps_the_window()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, _, overlay) = BuildHost();
            var innerDone = new TaskCompletionSource();
            var outerDone = new TaskCompletionSource();

            Task outer = overlay.RunAsync("Publishing to production.", async _ =>
            {
                await overlay.RunAsync("Reading the lists again.", _ => innerDone.Task);
                await outerDone.Task;
            });

            Assert.Equal("Reading the lists again.", overlay.Message);

            innerDone.SetResult();
            await SettleAsync(() => overlay.Message == "Publishing to production.");

            Assert.True(overlay.IsBusy);
            Assert.Equal("Publishing to production.", overlay.Message);

            outerDone.SetResult();
            await outer;

            Assert.False(overlay.IsBusy);
            Assert.False(overlay.IsVisible);
        });
    }

    // What the work reports reaches the card - and a position makes the bar stop wandering.
    [Fact]
    public async Task A_report_changes_the_sentence_and_a_position_fills_the_bar()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, _, overlay) = BuildHost();
            var finish = new TaskCompletionSource();
            WaitContext? context = null;

            Task waiting = overlay.RunAsync("Sending feedback.", given =>
            {
                context = given;
                return finish.Task;
            });

            context!.Report("Sending to server... 40%", 0.4);
            Dispatcher.UIThread.RunJobs();

            ProgressBar bar = overlay.FindControl<ProgressBar>("Bar")!;
            Assert.Equal("Sending to server... 40%", overlay.Message);
            Assert.False(bar.IsIndeterminate);
            Assert.Equal(40, bar.Value, 3);

            finish.SetResult();
            await waiting;
        });
    }

    // A report posted by a wait that has already ended must not rewrite the NEXT wait's sentence.
    [Fact]
    public async Task A_late_report_from_an_ended_wait_leaves_the_next_one_alone()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, _, overlay) = BuildHost();
            WaitContext? first = null;

            await overlay.RunAsync("First.", given =>
            {
                first = given;
                return Task.CompletedTask;
            });

            var finish = new TaskCompletionSource();
            Task second = overlay.RunAsync("Second.", _ => finish.Task);

            first!.Report("Late from the first.");
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Second.", overlay.Message);

            finish.SetResult();
            await second;
        });
    }

    // ###########################################################################################
    // *** KEYS ARE TAKEN TOO. *** The overlay stops every click, but a key goes to whatever has
    // focus - so the button that started the work would take Enter and start it again. A real key
    // press through the headless input stack, watching the button's own Click.
    // ###########################################################################################
    [Fact]
    public async Task While_busy_a_focused_button_does_not_take_Enter_and_afterwards_it_does_again()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, _, _, button, overlay) = BuildHost();
            window.Show();
            button.Focus();

            int clicks = 0;
            button.Click += (_, _) => clicks++;

            var finish = new TaskCompletionSource();
            Task waiting = overlay.RunAsync("Working.", _ => finish.Task);

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);
            Assert.Equal(0, clicks);

            finish.SetResult();
            await waiting;

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);
            Assert.Equal(1, clicks);

            window.Close();
        });
    }

    // Any control in a window finds that window's one overlay; a window without one finds none.
    [Fact]
    public void Any_control_in_a_window_finds_its_overlay()
    {
        UiTest.Run(() =>
        {
            var (_, _, _, button, overlay) = BuildHost();

            Assert.Same(overlay, BusyOverlay.For(button));

            var lonely = new Button();
            _ = new Window { Content = new StackPanel { Children = { lonely } } };

            Assert.Null(BusyOverlay.For(lonely));
        });
    }

    // A window with no overlay still gets the limit - the work runs and its value comes back.
    [Fact]
    public async Task Work_in_a_window_without_an_overlay_still_runs()
    {
        await UiTest.RunAsync(async () =>
        {
            var lonely = new Button();
            _ = new Window { Content = lonely };

            WaitResult<int> result = await BusyOverlay.RunAsync(lonely, "Working.", _ => Task.FromResult(3));

            Assert.Equal(3, result.Value);
        });
    }

    // The overlay shows WaitWording's words, not a clock of its own (owner request, 2026-09-28).
    [Fact]
    public void The_overlay_shows_the_elapsed_time_in_words()
    {
        Assert.Equal("1 minute and 42 seconds elapsed", BusyOverlay.FormatElapsed(TimeSpan.FromSeconds(102)));
    }

    // The line is EMPTY for the first second (WaitWording.Elapsed), and must still take its
    // height then - or the card grows by a line, and jumps, when "1 second elapsed" appears.
    [Fact]
    public void The_elapsed_line_is_as_tall_empty_as_with_words_so_the_card_does_not_jump()
    {
        UiTest.Run(() =>
        {
            TextBlock line = new BusyOverlay().ElapsedText;

            line.Text = string.Empty;
            line.Measure(Avalonia.Size.Infinity);
            double empty = line.DesiredSize.Height;

            line.Text = WaitWording.Elapsed(TimeSpan.FromSeconds(1));
            line.Measure(Avalonia.Size.Infinity);

            Assert.True(empty > 0);
            Assert.Equal(line.DesiredSize.Height, empty);
        });
    }

    // ###########################################################################################
    // *** A HOLD KEEPS THE WINDOW ACROSS SEVERAL STEPS. *** A decision and the lists read after it
    // are two waits; between them the window used to brighten for a moment and dim again. Held, it
    // stays busy - and revealed - from the first step to the last, and lets go after.
    // ###########################################################################################
    [Fact]
    public async Task A_hold_keeps_the_window_busy_between_its_steps()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, _, overlay) = BuildHost();
            bool busyBetween = false;
            bool revealedBetween = false;

            await overlay.HoldAsync("Approving.", async () =>
            {
                overlay.RevealNowForTests();

                await overlay.RunAsync("Sending the decision.", _ => Task.CompletedTask);

                busyBetween = overlay.IsBusy;
                revealedBetween = overlay.IsRevealed;

                await overlay.RunAsync("Reading the lists.", _ => Task.CompletedTask);
            });

            Assert.True(busyBetween);
            Assert.True(revealedBetween);
            Assert.False(overlay.IsBusy);
            Assert.False(overlay.IsVisible);
        });
    }

    // ###########################################################################################
    // *** LOCAL WORK PAST THE LIMIT CARRIES ON, AND ITS REAL RESULT STILL ARRIVES. *** A workbook
    // being written cannot be stopped halfway. The overlay lifts at the limit and says so; the caller
    // still gets the work's own answer when it ends - "validating if it did finish" with no server.
    // ###########################################################################################
    [Fact]
    public async Task Local_work_past_the_limit_lifts_the_overlay_says_so_and_still_returns_its_result()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, button, overlay) = BuildHost();
            var limit = new TaskCompletionSource();
            var slow = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            overlay.LimitOverrideForTests = _ => limit.Task;

            bool told = false;
            Task<string> saving = BusyOverlay.RunLocalAsync(button, "Saving to your draft...", () => slow.Task, () => told = true);

            Assert.True(overlay.IsBusy);

            limit.SetResult();
            await SettleAsync(() => told);

            Assert.False(overlay.IsBusy);
            Assert.False(saving.IsCompleted);

            slow.SetResult("saved");
            Assert.Equal("saved", await saving);
        });
    }

    [Fact]
    public async Task Local_work_inside_the_limit_returns_its_result_and_says_nothing_extra()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, _, _, button, overlay) = BuildHost();
            bool told = false;

            string result = await BusyOverlay.RunLocalAsync(button, "Saving...", () => Task.FromResult("saved"), () => told = true);

            Assert.Equal("saved", result);
            Assert.False(told);
            Assert.False(overlay.IsBusy);
        });
    }
}
