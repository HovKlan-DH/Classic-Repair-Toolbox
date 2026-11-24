using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Oscilloscope;
using Handlers.Theming;
using static ClassicRepairToolbox.Tests.Ui.BaselineCaptureTesting;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "CAPTURE FROM OSCILLOSCOPE PIN-BY-PIN" AND "CAPTURE AGAIN" FROM THE COMPONENT PANE (ComponentPane.BaselineRun.cs; owner requests,
// 2026-11-18: "how I easily can take an entire oscilloscope baseline in one go - e.g. for updating it", and 2026-11-20: "open this
// in a new window", "Capture again" opening it on its pin): the buttons and the notice on the "Oscilloscope baseline" page, the
// window closed by the pane (another component shown, the component gone), the pane's own changes reaching it, and each capture
// taken into the user's folder and put into the draft only while its baseline is still the pane's. The window itself is
// BaselineCaptureWindowTests'; what both share is BaselineCaptureTesting. The oscilloscope is a fake (FakeBaselineScope).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentPaneBaselineRunTests : IDisposable
{
    private readonly Draft thisDraft = new();

    public void Dispose() => this.thisDraft.Dispose();

    // The draft - pin 1's earlier baseline the picture `earlier`, when given.
    private DraftSession OpenDraft(byte[]? earlier = null) => this.thisDraft.Open(earlier);

    // ###########################################################################################
    // The oscilloscope gone, the window stays and says so - a capture not yet accepted kept, nothing captured or sent - and goes on
    // once it is back; a capture taken again lets go of the one before it. Another component shown closes it WITHOUT asking: a capture
    // not yet accepted goes, in no draft. One with no pins has nothing to go round, said in the pane.
    // ###########################################################################################
    [Fact]
    public async Task The_window_waits_for_the_oscilloscope_and_closes_as_another_component_is_shown()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            bool closed = false;
            window.Closed += (_, _) => closed = true;

            var written = new List<string>();
            pane.CaptureWrittenForTests = written.Add;

            Assert.True(In<StackPanel>(window, "ScopeConnectedPanel").IsVisible);
            Assert.False(In<Border>(window, "NoScopeNotice").IsVisible);

            await window.CaptureCurrentAsync();

            pane.ShowScopeConnected(false);

            Assert.Same(window, pane.CaptureWindowForTests);
            Assert.True(In<Border>(window, "NoScopeNotice").IsVisible);
            Assert.Equal(ComponentPaneWording.ScopeNotConnected, In<TextBlock>(window, "NoScopeText").Text);
            Assert.False(In<StackPanel>(window, "ScopeConnectedPanel").IsVisible);
            Assert.False(In<Button>(window, "CaptureButton").IsEnabled);

            await window.CaptureCurrentAsync();
            window.KeyPress(Key.Up, RawInputModifiers.None, PhysicalKey.ArrowUp, keySymbol: null);
            Assert.Equal(1, scope.Captures);
            Assert.Empty(scope.TriggerSteps);
            Assert.True(window.TargetsForTests[0].IsWaiting);
            Assert.Equal("Scope baseline/U27_1.png", Baselines(session)[0].File);

            // Back: captured again, the capture before it let go of.
            pane.ShowScopeConnected(true);
            Assert.False(In<Border>(window, "NoScopeNotice").IsVisible);
            Assert.True(In<Button>(window, "CaptureButton").IsEnabled);
            await window.CaptureCurrentAsync();
            Assert.Equal(2, scope.Captures);
            await Settled(pane, window);
            Assert.False(File.Exists(written[0]));
            Assert.True(File.Exists(written[1]));

            // Another component shown: closed without asking, the capture let go of, the draft as it was.
            pane.Show(session, "U26", "PAL", Published());
            Assert.True(closed);
            Assert.False(window.IsAskingAboutWaitingForTests);
            Assert.Null(pane.CaptureWindowForTests);
            Assert.False(File.Exists(written[1]));
            Assert.Equal("Scope baseline/U27_1.png", Baselines(session)[0].File);

            pane.Show(session, "C1", "PAL", Published());
            Assert.Null(pane.OpenBaselineCaptureWindow());
            Assert.Equal(ComponentPaneWording.RunNoPins, Get<TextBlock>(pane, "PaneMessageText").Text);

            await pane.ThumbnailsLoadedForTests;
            host.Close();
        });
    }

    // ###########################################################################################
    // The buttons and the panels they open stand OVER the baselines (owner request, 2026-11-18: "should be shown on top, as
    // they will be hidden if it has any capture"), "Capture from my oscilloscope" is gone (owner request, 2026-11-20), and while
    // no oscilloscope is connected a notice says what one offers - with the "Oscilloscope" tab turned off, where it is turned on
    // (owner request, 2026-11-18: "I at first thought it was not implemented").
    // ###########################################################################################
    [Fact]
    public void The_buttons_stand_over_the_baselines_and_a_notice_says_what_an_oscilloscope_offers()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, new FakeBaselineScope());

            StackPanel section = Get<StackPanel>(pane, "BaselinesSection");
            int list = section.Children.IndexOf(Get<ItemsControl>(pane, "BaselinesItemsControl"));

            Assert.True(section.Children.IndexOf(Get<StackPanel>(pane, "AddBaselineButtons")) < list);
            Assert.True(section.Children.IndexOf(Get<Border>(pane, "NewBaselinePanel")) < list);

            Assert.DoesNotContain(
                Get<StackPanel>(pane, "AddBaselineButtons").Children.OfType<Button>(),
                added => Equals(added.Content, "Capture from my oscilloscope"));

            // Connected: no notice.
            Assert.False(Get<Border>(pane, "ScopeHintNotice").IsVisible);

            pane.ShowScopeConnected(false);
            Assert.True(Get<Border>(pane, "ScopeHintNotice").IsVisible);
            Assert.Equal(ComponentPaneWording.ScopeHint(tabTurnedOn: true), Get<TextBlock>(pane, "ScopeHintText").Text);

            pane.ShowScopeConnected(false, tabTurnedOn: false);
            Assert.Equal(ComponentPaneWording.ScopeHint(tabTurnedOn: false), Get<TextBlock>(pane, "ScopeHintText").Text);

            pane.ShowScopeConnected(true);
            Assert.False(Get<Border>(pane, "ScopeHintNotice").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Capture again" on a pin far down the list, its line put in just then, shows that line IN VIEW (code review, 2026-11-24): a
    // list's rows put in since its last layout pass cannot be scrolled to without a layout first (ListScrolling).
    // ###########################################################################################
    [Fact]
    public async Task Capture_again_on_a_pin_far_down_shows_its_line_in_the_window()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());

            // The window open already on the first pin, and pin 14's baseline coming from another editor: its line put in as
            // "Capture again" asks for it (BaselineCaptureWindow.FollowRows).
            BaselineCaptureWindow window = Open(pane);

            AddBaseline(session, new ComponentImageEntry { BoardLabel = "U27", Pin = "14", Name = "VCC", File = "Scope baseline/U27_3.png", TimeDiv = "1us", VoltsDiv = "1V" });

            await pane.CaptureAgainInWindowAsync(pane.BaselinesForTests.Single(row => row.Committed.Pin == "14"));

            Assert.Same(window, pane.CaptureWindowForTests);
            ListBox lines = In<ListBox>(window, "TargetsList");
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            Dispatcher.UIThread.RunJobs();

            ScrollViewer scroller = lines.GetVisualDescendants().OfType<ScrollViewer>().First();
            int index = window.CurrentIndexForTests;

            Assert.Equal("14", window.TargetsForTests[index].Step.PinText);
            Assert.True(scroller.Extent.Height > scroller.Viewport.Height, "The list is longer than its view.");
            Assert.True(
                lines.ContainerFromIndex(index) is Control line &&
                line.IsEffectivelyVisible &&
                line.TranslatePoint(default, scroller) is { } at &&
                at.Y >= -1 &&
                at.Y + line.Bounds.Height <= scroller.Viewport.Height + 1,
                "Pin 14's line is in view.");

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // "CAPTURE AGAIN" on a baseline opens the window on its pin and captures there (owner request, 2026-11-20: "this button should
    // open the "Capture from oscilloscope pin-by-pin" window and go to that pin and capture the image from scope and then show it in
    // this view instead. Then its data can be controlled by the user before it maybe will be accepted"): the oscilloscope set as the
    // baseline was measured first, the capture not yet accepted - the draft as it was - until "Accept and continue to next pin". With
    // the window open already, it goes to that pin there.
    // ###########################################################################################
    [Fact]
    public async Task Capture_again_opens_the_window_on_its_pin_and_captures_there_to_be_accepted()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.002, null, null) };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);

            ComponentPaneBaselineRow row = pane.BaselinesForTests[1];
            Assert.True(row.CaptureOffered);

            await pane.CaptureAgainInWindowAsync(row);

            BaselineCaptureWindow window = pane.CaptureWindowForTests!;
            Assert.NotNull(window);
            Assert.True(window.IsVisible);
            Assert.Equal(2, window.CurrentIndexForTests);
            Assert.Equal(("3", "1us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));
            Assert.Equal(1, scope.Captures);
            Assert.True(window.TargetsForTests[2].IsWaiting);
            Assert.Equal("New capture - not yet accepted   2ms, 1V, 1.5V", BoldRuns.TextOf(In<TextBlock>(window, "NewHeaderText")));
            Assert.Equal("Scope baseline/U27_3.png", Baselines(session)[1].File);

            await window.AcceptAndContinueAsync();

            ComponentImageEntry pin3 = Baselines(session)[1];
            Assert.NotEqual("Scope baseline/U27_3.png", pin3.File);
            Assert.Equal(("1Y", "2ms", "1V", "1.5V"), (pin3.Name, pin3.TimeDiv, pin3.VoltsDiv, pin3.TriggerLevelVolts));
            Assert.Equal("Scope baseline/U27_1.png", Baselines(session)[0].File);
            Assert.Equal(DraftStepWording.ComponentBaselines("U27"), session.UndoDescription);

            // With the window open: the same window, on the pin of the baseline asked for.
            await pane.CaptureAgainInWindowAsync(pane.BaselinesForTests[0]);
            Assert.Same(window, pane.CaptureWindowForTests);
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal(2, scope.Captures);
            Assert.True(window.TargetsForTests[0].IsWaiting);

            await Settled(pane, window);
            window.LetGo();

            pane.ShowScopeConnected(false);
            Assert.False(pane.BaselinesForTests[0].CaptureOffered);

            host.Close();
        });
    }

    // ###########################################################################################
    // A BASELINE WITH NO IMAGE YET - a table edit took it, and Submit refuses it until it has one - is a line of its own, and its
    // "Capture again" captures into it (code review, 2026-11-21: the window opened and did nothing, and its capture went in as a
    // second baseline beside it).
    // ###########################################################################################
    [Fact]
    public async Task Capture_again_on_a_baseline_with_no_image_captures_into_that_very_baseline()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.0005, null, null) };

        await UiTest.RunAsync(async () =>
        {
            AddBaseline(session, new ComponentImageEntry { BoardLabel = "U27", Pin = "2", Name = "1B", TimeDiv = "1ms" });

            (ComponentPane pane, Window host) = Shown(session, scope);
            ComponentPaneBaselineRow noImage = pane.BaselinesForTests.Single(row => row.Committed.Pin == "2");

            await pane.CaptureAgainInWindowAsync(noImage);
            BaselineCaptureWindow window = pane.CaptureWindowForTests!;

            BaselineCaptureTarget pin2 = window.TargetsForTests[1];
            Assert.Same(pin2, window.CurrentForTests);
            Assert.Same(noImage, pin2.Row);
            Assert.True(pin2.IsWaiting);
            Assert.True(In<Border>(window, "SingleFrame").IsVisible);
            Assert.False(In<StackPanel>(window, "NamePanel").IsVisible);

            await window.AcceptAndContinueAsync();

            Assert.Equal(3, Baselines(session).Count);
            ComponentImageEntry filled = Baselines(session).Single(row => row.Pin == "2");
            Assert.StartsWith("Scope baseline/U27_2_1B", filled.File, StringComparison.Ordinal);
            Assert.Equal("500us", filled.TimeDiv);
            Assert.True(pin2.IsAccepted);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // A CAPTURE GOES ONLY INTO A BASELINE THE PANE STILL HAS (code review, 2026-11-21): the window's "please wait" covers the window
    // alone, and the baselines read again while the capture was copied in - an "Undo", another editor's change - had it written into
    // a row the pane no longer held, the window saying "Accepted". It waits on, the reason said, and goes in when accepted again.
    // ###########################################################################################
    [Fact]
    public async Task A_capture_is_not_put_into_a_baseline_the_pane_let_go_of_while_it_went_in()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.0001, null, null) };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            await window.CaptureCurrentAsync();

            Task accepting = window.AcceptAndContinueAsync();

            // While it is copied in, another editor's change: the pane reads its baselines again.
            session.Apply("Another editor", content => content with { Board = content.Board.WithRevisionDate("2026-11-21") });
            await accepting;

            Assert.Equal("Scope baseline/U27_1.png", Baselines(session)[0].File);
            Assert.True(window.TargetsForTests[0].IsWaiting);
            Assert.Equal(ComponentPaneWording.RunCaptureNotPut, In<TextBlock>(window, "ProblemText").Text);
            Assert.Equal(0, window.CurrentIndexForTests);

            await window.AcceptAndContinueAsync();
            Assert.StartsWith("Scope baseline/U27_1_1A", Baselines(session)[0].File, StringComparison.Ordinal);
            Assert.True(window.TargetsForTests[0].IsAccepted);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // The pane's own changes reach the window too (code review, 2026-11-21): a baseline given another pin is that pin's line, its old
    // pin a new line again - a capture of pin 3 never goes into the baseline now pin 4's - and a capture whose line went while it was
    // being taken is not kept, and said.
    // ###########################################################################################
    [Fact]
    public async Task A_baseline_given_another_pin_in_the_pane_is_that_pins_line_and_a_capture_whose_line_went_is_not_kept()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            pane.BaselinesForTests[1].Pin = "4";
            pane.CommitHeld();

            Assert.Equal(14, window.TargetsForTests.Count);
            Assert.Equal(("3  1Y", "No baseline yet"), (window.TargetsForTests[2].Label, window.TargetsForTests[2].StateText));
            Assert.Null(window.TargetsForTests[2].Row);
            Assert.Equal("4  1Y", window.TargetsForTests[3].Label);
            Assert.Same(pane.BaselinesForTests[1], window.TargetsForTests[3].Row);

            // Pin 5 being captured as another editor gives it a baseline: its line goes, and the capture with it.
            var written = new List<string>();
            pane.CaptureWrittenForTests = written.Add;
            Assert.True(window.GoToPin("5", pickIt: true));
            scope.CaptureGate = new TaskCompletionSource<byte[]?>();
            Task capturing = window.CaptureCurrentAsync();
            await Until(() => scope.Captures == 1, "the capture asked of the oscilloscope");

            AddBaseline(session, new ComponentImageEntry { BoardLabel = "U27", Pin = "5", Name = "2A", File = "Scope baseline/U27_1.png" });
            Dispatcher.UIThread.RunJobs();

            scope.CaptureGate.SetResult(FakeBaselineScope.Png);
            await capturing;

            Assert.Equal(ComponentPaneWording.RunCaptureLost, In<TextBlock>(window, "ProblemText").Text);
            Assert.False(File.Exists(Assert.Single(written)));
            Assert.DoesNotContain(window.TargetsForTests, target => target.IsWaiting);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // THE WINDOW GOES WITH ITS COMPONENT (code review, 2026-11-21): the component deleted - here by another editor - and the window
    // stayed open, and accepting a capture wrote baselines for a component the board no longer had. Its captures go with it.
    // ###########################################################################################
    [Fact]
    public async Task The_window_closes_as_its_component_goes_and_its_captures_with_it()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;
            await window.CaptureCurrentAsync();
            await Settled(pane, window);

            session.Apply("Another editor", content => content with
            {
                Board = content.Board.WithComponentSections(components: content.Board.Components.Where(row => row.BoardLabel != "U27").ToList()),
            });
            Dispatcher.UIThread.RunJobs();

            Assert.Null(pane.CaptureWindowForTests);
            Assert.False(window.IsVisible);
            Assert.False(File.Exists(waiting));

            host.Close();
        });
    }

    // ###########################################################################################
    // A capture waits in AppConfig.ScopeCaptureWaitFolder until it is accepted - the user's OWN app data folder (code review,
    // 2026-10-26: under the shared /tmp another account could swap the screenshot before it was added to the draft and sent;
    // AppConfigTests holds the folder to that), a folder of the test process's own here (TestScopeCaptures). Accepted into the draft,
    // it is gone.
    // ###########################################################################################
    [Fact]
    public async Task A_capture_waits_in_the_capture_folder_and_goes_once_accepted()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;

            await pane.CaptureAgainInWindowAsync(pane.BaselinesForTests[0]);
            BaselineCaptureWindow window = pane.CaptureWindowForTests!;

            Assert.NotNull(waiting);
            Assert.StartsWith(AppConfig.ScopeCaptureWaitFolder + Path.DirectorySeparatorChar, waiting, StringComparison.Ordinal);
            Assert.True(File.Exists(waiting));

            await Settled(pane, window);
            await window.AcceptAndContinueAsync();
            Assert.False(File.Exists(waiting));

            await Settled(pane, window);
            window.Close();
            host.Close();
        });
    }

    // ###########################################################################################
    // CAPTURES NOT YET ACCEPTED HOLD THE PANE ON THEIR COMPONENT (code review, 2026-11-22: another component, a part of the board,
    // "Table" or a region picked in CRT's window closed the capture window, every screen deleted unasked): the host asks the pane
    // (KeepsItselfOpen) and is refused while any waits, the window brought forward asking what its own close asks. Another region
    // is no reason while it is the one they were taken in. Answered, the pane is free again.
    // ###########################################################################################
    [Fact]
    public async Task Captures_not_yet_accepted_hold_the_pane_on_their_component_and_the_window_asks()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);

            // Nothing captured: nothing holds the pane.
            Assert.False(pane.KeepsItselfOpen());

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;
            await window.CaptureCurrentAsync();
            await Settled(pane, window);

            Assert.True(pane.KeepsItselfOpen());
            Assert.True(window.IsAskingAboutWaitingForTests);
            Assert.Same(window, pane.CaptureWindowForTests);
            Assert.True(File.Exists(waiting));

            // Another region than theirs holds it too; their own does not.
            Assert.True(pane.KeepsWaitingCaptures("NTSC"));
            Assert.False(pane.KeepsWaitingCaptures("PAL"));

            // "Keep capturing": still held, asked again at the next pick.
            Click(In<Button>(window, "KeepCapturingButton"));
            Assert.False(window.IsAskingAboutWaitingForTests);
            Assert.True(pane.KeepsItselfOpen());
            Assert.True(window.IsAskingAboutWaitingForTests);

            // "Discard and close": the pane is free.
            Click(In<Button>(window, "DiscardAndCloseButton"));
            Assert.Null(pane.CaptureWindowForTests);
            Assert.False(pane.KeepsItselfOpen());
            Assert.False(File.Exists(waiting));

            host.Close();
        });
    }

    // ###########################################################################################
    // THE REGION CHANGED UNDER CAPTURES NOT YET ACCEPTED LETS NONE GO (second code review, 2026-11-22): a region picked is held while
    // any waits, so it changes only by an edit - "Board: Regions", or its "Undo" - and the window closed with them. It follows the
    // region the pane shows, where they would go in, and asks; with none waiting it closes as before.
    // ###########################################################################################
    [Fact]
    public async Task The_region_changed_under_captures_not_yet_accepted_keeps_the_window_and_asks()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;
            await window.CaptureCurrentAsync();
            await Settled(pane, window);

            pane.UseRegion("NTSC");

            Assert.Same(window, pane.CaptureWindowForTests);
            Assert.Equal("NTSC", window.RunRegion);
            Assert.True(window.IsAskingAboutWaitingForTests);
            Assert.True(window.TargetsForTests[0].IsWaiting);
            Assert.True(File.Exists(waiting));

            // Accepted, it goes in as the pane shows the component now.
            await window.AcceptAndContinueAsync();
            Assert.StartsWith("Scope baseline/U27_1_1A", Baselines(session)[0].File, StringComparison.Ordinal);

            // Nothing waiting: another region closes the window, as before.
            await Settled(pane, window);
            pane.UseRegion("PAL");
            Assert.Null(pane.CaptureWindowForTests);

            host.Close();
        });
    }

    // ###########################################################################################
    // ONE FILE GOES INTO THE DRAFT AT A TIME, EACH UNDER A NAME OF ITS OWN (code review, 2026-11-22): a capture accepted in the window
    // goes in under the window's "please wait" alone, so the pane could add a file at the same moment - a baseline of the same pin and
    // signal was given the very name the capture was, one copy replacing the other. The pane's waits for the capture, under a name of
    // its own.
    // ###########################################################################################
    [Fact]
    public async Task A_baseline_added_in_the_pane_while_a_capture_goes_in_waits_and_takes_a_name_of_its_own()
    {
        DraftSession session = this.OpenDraft();
        string screenshot = this.thisDraft.Pickable("IMG_0042.png", Png);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);
            await window.CaptureCurrentAsync();

            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(screenshot);
            pane.OpenNewBaseline();
            Get<TextBox>(pane, "NewBaselinePinBox").Text = "1";
            Get<TextBox>(pane, "NewBaselineNameBox").Text = "1A";
            await pane.ChooseNewBaselineScreenshotAsync();

            // Pin 1's capture goes in - and while it is copied in, a baseline of pin 1 is added in the pane.
            Task accepting = window.AcceptAndContinueAsync();
            Task adding = pane.AddNewBaselineAsync();
            await Task.WhenAll(accepting, adding);

            List<string> files = Baselines(session).Where(baseline => baseline.Pin == "1").Select(baseline => baseline.File).ToList();
            Assert.Equal(2, files.Count);
            Assert.Equal(2, files.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.All(files, file => Assert.True(session.ListedFiles.ContainsKey(file), file));
            Assert.True(window.TargetsForTests[0].IsAccepted);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // "CAPTURE AGAIN" WHILE THE WINDOW IS BUSY SAYS WHY NOTHING HAPPENS (code review, 2026-11-22: the window came to the front, and
    // nothing was captured or said).
    // ###########################################################################################
    [Fact]
    public async Task Capture_again_while_the_window_is_busy_says_why()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            scope.CaptureGate = new TaskCompletionSource<byte[]?>();
            Task capturing = window.CaptureCurrentAsync();
            await Until(() => window.IsBusy && scope.Captures == 1, "the capture asked of the oscilloscope");

            await pane.CaptureAgainInWindowAsync(pane.BaselinesForTests[1]);
            Assert.Equal(ComponentPaneWording.RunWindowBusy, Get<TextBlock>(pane, "PaneMessageText").Text);
            Assert.Equal(1, scope.Captures);

            scope.CaptureGate.SetResult(Png);
            await capturing;

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // THE WINDOW SAYS THE COMPONENT AS THE PANE NOW NAMES IT (code review, 2026-11-22: renamed U27A in the pane, the window still said
    // U27 while its captures went in under U27A) - and a baseline's pin written otherwise in the pane is the line's pin text.
    // ###########################################################################################
    [Fact]
    public void The_window_follows_a_rename_in_the_pane_and_a_baselines_pin_text()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);
            Assert.Equal(ComponentPaneWording.CaptureWindowTitle("U27"), window.Title);

            Get<TextBox>(pane, "BoardLabelBox").Text = "U27A";
            pane.CommitHeld();

            Assert.Same(window, pane.CaptureWindowForTests);
            Assert.Equal(ComponentPaneWording.CaptureWindowTitle("U27A"), window.Title);
            Assert.Equal(pane.ShownTitle, In<TextBlock>(window, "ComponentText").Text);
            Assert.Contains("U27A", In<TextBlock>(window, "ComponentText").Text, StringComparison.Ordinal);

            // Pin 3's baseline written "3 [running]" in the pane: its line says so.
            pane.BaselinesForTests[1].Pin = "3 [running]";
            pane.CommitHeld();

            Assert.Equal("3 [running]", window.TargetsForTests.Single(target => ReferenceEquals(target.Row, pane.BaselinesForTests[1])).Step.PinText);

            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // A COMMIT IN THE PANE THAT CHANGES NOTHING LEAVES THE WINDOW ALONE (code review, 2026-11-21): every box left and typing pause
    // commits the pins and the baselines, and each built the window's lines again - twice - and wrote its "Signal name" box back
    // trimmed, the space just typed gone with the caret. Only a change reaches it, and the box keeps what is typed in it still.
    // ###########################################################################################
    [Fact]
    public void A_commit_that_changes_nothing_leaves_the_window_and_its_name_box_alone()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);

            // Pin 5, with no baseline: its signal name being typed, a space last.
            Assert.True(window.GoToPin("5", pickIt: true));
            TextBox name = In<TextBox>(window, "NameBox");
            name.Text = "Clock ";
            name.CaretIndex = name.Text.Length;
            int followed = window.RowsFollowedForTests;

            pane.CommitHeld();

            Assert.Equal(followed, window.RowsFollowedForTests);
            Assert.Equal(("Clock ", 6), (name.Text, name.CaretIndex));

            // Pin 3's baseline given pin 4 in the pane: the window follows, once.
            pane.BaselinesForTests[1].Pin = "4";
            pane.CommitHeld();

            Assert.Equal(followed + 1, window.RowsFollowedForTests);
            Assert.Equal("Clock ", name.Text);

            window.LetGo();
            host.Close();
        });
    }
}
