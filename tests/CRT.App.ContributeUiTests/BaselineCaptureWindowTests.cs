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
// BaselineCaptureWindow - "CAPTURE FROM OSCILLOSCOPE PIN-BY-PIN" (owner requests, 2026-11-18: "how I easily can take an entire
// oscilloscope baseline in one go", the settings read back what the baseline stores; and 2026-11-20: "open this in a new window
// (remember size and position) ... support the same keyboard options as available in the component info window", then "it must
// stay on the pin, so I can check that it is good ... "Accept and continue to next pin"", "this must be accepted before it will be
// used", the old baseline beside the new one, every picture clicked to see it large, and the buttons standing still). Opened by the
// pane (ComponentPaneBaselineRunTests has the pane's side); what both share is BaselineCaptureTesting. The oscilloscope is a fake
// (FakeBaselineScope).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BaselineCaptureWindowTests : IDisposable
{
    private readonly Draft thisDraft = new();

    public void Dispose() => this.thisDraft.Dispose();

    // The draft - pin 1's earlier baseline the picture `earlier`, when given.
    private DraftSession OpenDraft(byte[]? earlier = null) => this.thisDraft.Open(earlier);

    // ###########################################################################################
    // Offered while the oscilloscope is connected, it opens ONE window: every pin a line, the pins with a baseline picked and
    // framed on the case, the first shown and the oscilloscope set as it was measured, what comes next said - and a quick pick
    // (going to the first pin it picks - owner request, 2026-11-20: "it should automatically jump to the first entry in that scope"),
    // a checkbox, a pin clicked on the case or a line clicked picks and shows again.
    // ###########################################################################################
    [Fact]
    public void The_window_opens_on_the_pins_with_a_baseline_and_picks_by_quick_pick_checkbox_or_the_case()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);

            Button button = Get<Button>(pane, "BaselineRunButton");
            Assert.Equal("Capture from oscilloscope pin-by-pin", button.Content);
            Assert.True(button.IsVisible);
            pane.ShowScopeConnected(false);
            Assert.False(button.IsVisible);
            pane.ShowScopeConnected(true);

            BaselineCaptureWindow window = Open(pane);

            Assert.True(window.IsVisible);
            Assert.Same(window, pane.CaptureWindowForTests);
            Assert.Same(window, pane.OpenBaselineCaptureWindow());
            Assert.Equal(ComponentPaneWording.CaptureWindowTitle("U27"), window.Title);
            Assert.Equal(ComponentPaneWording.RunKeys.Count, In<ItemsControl>(window, "KeysList").ItemCount);

            Assert.Equal(14, window.TargetsForTests.Count);
            Assert.Equal(["1", "3"], Picked(window));
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal("Pin 1 - 1A", Said(window, "CurrentText"));
            Assert.Equal(("1", "50us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));

            PinoutDrawing drawing = In<PinoutDrawing>(window, "CasePinout");
            Assert.Equal([1, 3], drawing.ChosenPinsForTests.Order());
            Assert.Equal(1, drawing.MarkedPinForTests);
            Assert.Equal("0 of 2 picked baselines accepted", Said(window, "ProgressText"));
            Assert.Equal("Next: pin 3 - 1Y", Said(window, "NextText"));
            Assert.True(window.TargetsForTests[2].IsNext);
            Assert.Equal(("Has a baseline", "No baseline yet"), (window.TargetsForTests[0].StateText, window.TargetsForTests[1].StateText));

            // The gaps: every pin with no baseline - the ground and supply pins too (owner report, 2026-11-18) - pin 2 the first shown.
            Click(In<Button>(window, "PickWithoutBaselineButton"));
            Assert.Equal(12, Picked(window).Count);
            Assert.Contains("7", Picked(window));
            Assert.Contains("14", Picked(window));
            Assert.DoesNotContain("1", Picked(window));
            Assert.Equal(1, window.CurrentIndexForTests);
            Assert.Equal(2, drawing.MarkedPinForTests);

            // Back to the pins with a baseline: pin 1 shown, the oscilloscope set for it.
            Click(In<Button>(window, "PickWithBaselineButton"));
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal("1", scope.Applied[^1].Pin);

            Click(In<Button>(window, "PickEveryPinButton"));
            Assert.Equal(14, Picked(window).Count);
            Assert.Equal(0, window.CurrentIndexForTests);

            // None picked: nothing to go to - the pin shown stays.
            window.GoTo(5, setScope: false);
            Click(In<Button>(window, "PickNoPinButton"));
            Assert.Empty(Picked(window));
            Assert.Equal(5, window.CurrentIndexForTests);
            Assert.Empty(drawing.ChosenPinsForTests);
            Assert.Equal(ComponentPaneWording.RunNothingPicked, Said(window, "ProgressText"));
            Assert.False(In<Border>(window, "NextBorder").IsVisible);

            // A checkbox in the list picks its pin.
            window.TargetsForTests[6].IsPicked = true;
            Assert.Equal([7], drawing.ChosenPinsForTests);
            Assert.Equal("0 of 1 picked baseline accepted", Said(window, "ProgressText"));
            Assert.True(In<Border>(window, "NextBorder").IsVisible);

            // A pin clicked on the case is shown and picked - with no baseline, the oscilloscope is left as it is.
            int applied = scope.Applied.Count;
            drawing.RaiseEvent(new PinPickedEventArgs(PinoutDrawing.PinPickedEvent, 5));
            Assert.Equal(4, window.CurrentIndexForTests);
            Assert.Equal(["5", "7"], Picked(window));
            Assert.Equal(5, drawing.MarkedPinForTests);
            Assert.Equal(applied, scope.Applied.Count);

            // A line clicked is shown, the oscilloscope set as it was measured.
            In<ListBox>(window, "TargetsList").SelectedIndex = 2;
            Assert.Equal(2, window.CurrentIndexForTests);
            Assert.Equal(("3", "1us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));

            window.Close();
            Assert.Null(pane.CaptureWindowForTests);

            host.Close();
        });
    }

    // ###########################################################################################
    // CAPTURE, CHECK, ACCEPT - AND ONLY THEN THE DRAFT (owner requests, 2026-11-20: "it must stay on the pin, so I can check that it
    // is good and its data looks good ... "Accept and continue to next pin". The "Previous pin" should not "unmark" it as accepted",
    // then "New capture - not yet accepted", "so it is clear this must be accepted before it will be used"): a capture - the image and
    // what the oscilloscope says it is set to - waits on its pin, in no draft; accepted, it goes into the baseline as ONE step, saved,
    // and the next picked pin not accepted is shown, the oscilloscope set for it; going back leaves it accepted, and a capture taken
    // again waits again.
    // ###########################################################################################
    [Fact]
    public async Task A_capture_waits_on_its_pin_and_is_in_the_draft_only_once_accepted()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.0001, 2, 1.0) };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            BaselineCaptureTarget pin1 = window.TargetsForTests[0];

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;

            await window.CaptureCurrentAsync();

            // Waiting: the draft as it was, the screen kept in the user's folder.
            ComponentImageEntry before = Baselines(session)[0];
            Assert.Equal(("Scope baseline/U27_1.png", "50us", "1V", "1.5V"), (before.File, before.TimeDiv, before.VoltsDiv, before.TriggerLevelVolts));
            Assert.True(File.Exists(waiting));

            // Still on pin 1, not yet accepted, the oscilloscope left as it is.
            int applied = scope.Applied.Count;
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.True(pin1.IsWaiting);
            Assert.False(pin1.IsAccepted);
            Assert.Equal("Captured - not yet accepted", pin1.StateText);
            Assert.Equal("100us, 2V, 1V", pin1.Settings);
            Assert.Equal("0 of 2 picked baselines accepted", Said(window, "ProgressText"));
            Assert.Empty(In<PinoutDrawing>(window, "CasePinout").DonePinsForTests);
            Assert.StartsWith("Captured - not yet accepted, so not in your draft.", In<TextBlock>(window, "CurrentDetails").Text, StringComparison.Ordinal);

            // Its every field as it would be accepted, the VALUE bold (owner request, 2026-11-20: "I want the value to be the
            // important one here") - and what changed since the baseline it is to replace.
            Assert.Equal(
                ["1", "1A", "(none)", "100us  (was 50us)", "2V  (was 1V)", "1V  (was 1.5V)", "At the READY prompt", "(none)"],
                new[]
                {
                    ComponentPaneWording.FactPin, ComponentPaneWording.FactSignalName, ComponentPaneWording.FactExpectedReading,
                    ComponentPaneWording.FactTimeDiv, ComponentPaneWording.FactVoltsDiv, ComponentPaneWording.FactTriggerLevel,
                    ComponentPaneWording.FactMachineState, ComponentPaneWording.FactNote,
                }.Select(window.FactShownForTests));
            Assert.Equal(["100us"], BoldRuns.BoldPiecesOf(window.FactValueBlockForTests(ComponentPaneWording.FactTimeDiv)));
            Assert.Empty(BoldRuns.BoldPiecesOf(window.FactValueBlockForTests(ComponentPaneWording.FactNote)));

            // No regions on this board: what it applies to is not shown.
            Assert.Null(window.FactShownForTests(ComponentPaneWording.FactAppliesTo));

            // "Accept and continue to next pin": into the draft as ONE step, saved, the waiting screen gone - and on to pin 3 with the
            // oscilloscope set for it.
            Click(In<Button>(window, "AcceptButton"));
            await Until(() => pin1.IsAccepted, "the capture accepted");

            ComponentImageEntry captured = Baselines(session)[0];
            Assert.StartsWith("Scope baseline/U27_1_1A", captured.File, StringComparison.Ordinal);
            Assert.Equal(("1A", "100us", "2V", "1V", "At the READY prompt"), (captured.Name, captured.TimeDiv, captured.VoltsDiv, captured.TriggerLevelVolts, captured.MachineState));
            Assert.Equal(DraftStepWording.ComponentBaselines("U27"), session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.False(File.Exists(waiting));

            Assert.Equal("Accept and continue to next pin", In<Button>(window, "AcceptButton").Content);
            Assert.False(pin1.IsWaiting);
            Assert.Equal(2, window.CurrentIndexForTests);
            Assert.Equal(("3", "1us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));
            Assert.Equal(applied + 1, scope.Applied.Count);
            Assert.Equal([1], In<PinoutDrawing>(window, "CasePinout").DonePinsForTests);
            Assert.Equal("1 of 2 picked baselines accepted", Said(window, "ProgressText"));
            Assert.Equal(ComponentPaneWording.RunShownIsLast, Said(window, "NextText"));

            // Going back leaves it accepted.
            Click(In<Button>(window, "PreviousButton"));
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.True(pin1.IsAccepted);
            Assert.Equal("Accepted. Enter captures it again. Machine state: At the READY prompt", In<TextBlock>(window, "CurrentDetails").Text);
            Assert.Equal("Next: pin 3 - 1Y", Said(window, "NextText"));

            // Pin 3, an answer the oscilloscope did not give leaving the baseline's as it was; accepted, nothing is next.
            window.GoBy(1);
            scope.Settings = new ScopeSettingsReading(null, null, 0.8);
            await window.CaptureCurrentAsync();
            Assert.Equal("1.5V", Baselines(session)[1].TriggerLevelVolts);

            await window.AcceptAndContinueAsync();

            ComponentImageEntry pin3 = Baselines(session)[1];
            Assert.Equal(("1us", "1V", "0.8V"), (pin3.TimeDiv, pin3.VoltsDiv, pin3.TriggerLevelVolts));
            Assert.Equal(2, window.CurrentIndexForTests);
            Assert.Equal("Every picked pin is accepted.", Said(window, "NextText"));
            Assert.Equal("2 of 2 picked baselines accepted", Said(window, "ProgressText"));
            Assert.DoesNotContain(window.TargetsForTests, target => target.IsNext);

            // Captured again, it waits again - the accepted baseline kept in the draft until this one is accepted.
            await window.CaptureCurrentAsync();
            Assert.True(window.TargetsForTests[2].IsWaiting);
            Assert.Equal("1 of 2 picked baselines accepted", Said(window, "ProgressText"));
            Assert.Equal(pin3.File, Baselines(session)[1].File);

            // One "Undo" takes back one accepted capture - and the window stays, finding its baselines again.
            Assert.True(session.Undo());
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("1.5V", Baselines(session)[1].TriggerLevelVolts);
            Assert.Same(window, pane.CaptureWindowForTests);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // A pin with no baseline gets a new one once its capture is accepted - its signal name the pinout's, changed until it is accepted
    // ("Then its data can be controlled by the user before it maybe will be accepted"), in the window's machine state - and one with
    // no name is not accepted until one is typed; captured again, it is the same new baseline. With nothing captured, "Accept and
    // continue to next pin" only goes on.
    // ###########################################################################################
    [Fact]
    public async Task A_pin_with_no_baseline_gets_a_new_one_once_accepted_and_one_with_no_name_waits_for_it()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            // The machine state of the baselines measured last, to start with.
            TextBox machineState = In<TextBox>(window, "MachineStateBox");
            Assert.Equal("At the READY prompt", machineState.Text);

            // The quick pick goes to the first pin it picks.
            int applied = scope.Applied.Count;
            window.PickPins(pins => pins.Where(pin => pin.Pin is "2" or "5").ToList());
            Assert.Equal(1, window.CurrentIndexForTests);
            machineState.Text = "Running a game";
            Assert.True(window.GoToPin("2", pickIt: false));

            Assert.Equal("No baseline yet - Enter captures one, to be accepted. Machine state: Running a game", In<TextBlock>(window, "CurrentDetails").Text);
            Assert.True(In<TextBlock>(window, "NoImageText").IsVisible);
            Assert.Equal("1B", In<TextBox>(window, "NameBox").Text);

            // Nothing to set the oscilloscope by: it stays as it is.
            Assert.Equal(applied, scope.Applied.Count);

            scope.Settings = new ScopeSettingsReading(5e-07, 0.5, 1.4);
            await window.CaptureCurrentAsync();

            // Waiting: no new baseline in the draft yet - its picture and its data shown, its name still to be typed over.
            Assert.Equal(2, Baselines(session).Count);
            Assert.True(In<Border>(window, "SingleFrame").IsVisible);
            Assert.False(In<Grid>(window, "ComparePanel").IsVisible);
            Assert.Equal("New capture - not yet accepted   500ns, 500mV, 1.4V", BoldRuns.TextOf(In<TextBlock>(window, "SingleHeaderText")));
            Assert.Equal("500ns", window.FactShownForTests(ComponentPaneWording.FactTimeDiv));
            Assert.Equal("1B", window.FactShownForTests(ComponentPaneWording.FactSignalName));
            Assert.True(In<StackPanel>(window, "NamePanel").IsVisible);

            // The name typed before accepting is the one it is accepted under.
            In<TextBox>(window, "NameBox").Text = "B1";
            await window.AcceptAndContinueAsync();

            ComponentImageEntry added = Baselines(session)[^1];
            Assert.Equal(
                ("2", "B1", "Scope baseline/U27_2_B1.png", "500ns", "500mV", "1.4V", "Running a game"),
                (added.Pin, added.Name, added.File, added.TimeDiv, added.VoltsDiv, added.TriggerLevelVolts, added.MachineState));

            // On to pin 5, named as the 74LS08's datasheet names it; emptied, its capture is not accepted until one is typed.
            Assert.Equal(4, window.CurrentIndexForTests);
            TextBox name = In<TextBox>(window, "NameBox");
            Assert.False(string.IsNullOrEmpty(name.Text));
            name.Text = "";
            await window.CaptureCurrentAsync();
            Assert.Equal(2, scope.Captures);
            Assert.True(window.TargetsForTests[4].IsWaiting);

            await window.AcceptAndContinueAsync();
            Assert.True(In<TextBlock>(window, "ProblemText").IsVisible);
            Assert.Equal(ComponentPaneWording.RunNeedsName, In<TextBlock>(window, "ProblemText").Text);
            Assert.Equal(4, window.CurrentIndexForTests);
            Assert.False(window.TargetsForTests[4].IsAccepted);
            Assert.Equal(3, Baselines(session).Count);

            name.Text = "2A";
            await window.AcceptAndContinueAsync();

            Assert.Equal(("5", "2A"), (Baselines(session)[^1].Pin, Baselines(session)[^1].Name));
            Assert.Equal("5  2A", window.TargetsForTests[4].Label);
            Assert.Equal(4, Baselines(session).Count);

            // Pin 2 accepted already: no other picked pin to go on to.
            Assert.Equal(4, window.CurrentIndexForTests);

            // Gone back to, captured and accepted again, the new baseline is replaced - not added a second time.
            Assert.True(window.GoToPin("2", pickIt: false));
            await window.CaptureCurrentAsync();
            await window.AcceptAndContinueAsync();
            Assert.Equal(4, Baselines(session).Count);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // THE OLD BASELINE FRAMED BESIDE THE NEW CAPTURE (owner requests, 2026-11-20: "I want also to see the old baseline image",
    // "have it more clear/visible what is the old and new image", "a red'ish background color"): the old on the left, red, the capture
    // on the right, a strip over each naming it with its settings - the capture's amber while not yet accepted, green once accepted -
    // and, captured again, still the baseline the run started from beside it.
    // ###########################################################################################
    [Fact]
    public async Task A_pin_captured_again_shows_its_old_baseline_framed_beside_the_new_capture()
    {
        DraftSession session = this.OpenDraft(earlier: Screen(4, 3, SkiaSharp.SKColors.Black));
        var scope = new FakeBaselineScope { Screen = Screen(5, 3, SkiaSharp.SKColors.White), Settings = new ScopeSettingsReading(0.0001, null, null) };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            TextBlock Header(string name) => In<TextBlock>(window, name);

            // Before a capture: the baseline alone, named so.
            Assert.True(In<Border>(window, "SingleFrame").IsVisible);
            Assert.False(In<Grid>(window, "ComparePanel").IsVisible);
            Assert.Equal("Baseline   50us, 1V, 1.5V", BoldRuns.TextOf(Header("SingleHeaderText")));
            Assert.Equal(["Baseline"], BoldRuns.BoldPiecesOf(Header("SingleHeaderText")));

            await window.CaptureCurrentAsync();
            await Until(() => window.CurrentForTests!.OldPicture is not null, "the old baseline's picture read");

            Assert.True(In<Grid>(window, "ComparePanel").IsVisible);
            Assert.False(In<Border>(window, "SingleFrame").IsVisible);
            Assert.Equal("Old baseline   50us, 1V, 1.5V", BoldRuns.TextOf(Header("OldHeaderText")));
            Assert.Equal("New capture - not yet accepted   100us, 1V, 1.5V", BoldRuns.TextOf(Header("NewHeaderText")));
            Assert.Equal(["Old baseline"], BoldRuns.BoldPiecesOf(Header("OldHeaderText")));

            // The old baseline's own picture on the left - its bytes kept before a capture accepted can replace it - the capture on
            // the right. (The headless platform decodes every picture as 1 x 1: the bytes say which.)
            byte[] old = Screen(4, 3, SkiaSharp.SKColors.Black);
            Assert.Equal(old, window.CurrentForTests!.BeforeBytes);
            Assert.Equal("Scope baseline/U27_1.png", Baselines(session)[0].File);
            await Until(() => ReferenceEquals(window.CurrentForTests!.OldPicture, In<Image>(window, "OldImage").Source), "the old baseline's picture shown");
            await Until(() => In<Image>(window, "NowImage").Source is not null, "the new capture's picture shown");
            Assert.Same(window.CurrentForTests!.WaitingPicture, In<Image>(window, "NowImage").Source);

            // The capture amber while not yet accepted, the old one red.
            Assert.Contains("ToCheck", In<Border>(window, "NewFrame").Classes);
            Assert.DoesNotContain("Accepted", In<Border>(window, "NewFrame").Classes);
            Assert.DoesNotContain("ToCheck", In<Border>(window, "OldFrame").Classes);

            // The old baseline red (owner request, 2026-11-20: "a red'ish background color"), from the theme - and each strip rounded
            // as the inside of its frame, so the frame's curve shows whole at the top corners ("missing a pixel in its rounding").
            Border oldHeader = In<Border>(window, "OldHeader");
            Assert.Contains("Old", oldHeader.Classes);
            Assert.Contains("Old", In<Border>(window, "OldFrame").Classes);
            Assert.True(Application.Current!.TryGetResource("CaptureOld_Bg", oldHeader.ActualThemeVariant, out object? red));
            Assert.Equal(((Avalonia.Media.ISolidColorBrush)red!).Color, ((Avalonia.Media.ISolidColorBrush)oldHeader.Background!).Color);

            foreach ((string frame, string header) in new[] { ("OldFrame", "OldHeader"), ("NewFrame", "NewHeader") })
            {
                Border outer = In<Border>(window, frame);
                double inner = outer.CornerRadius.TopLeft - outer.BorderThickness.Top;

                Assert.Equal(new CornerRadius(inner, inner, 0, 0), In<Border>(window, header).CornerRadius);
            }

            // Accepted, and gone back to: green - the old baseline still the one the run started from, though the draft has it no more.
            await window.AcceptAndContinueAsync();
            window.GoBy(-1);
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal("New capture - accepted   100us, 1V, 1.5V", BoldRuns.TextOf(Header("NewHeaderText")));
            Assert.Contains("Accepted", In<Border>(window, "NewFrame").Classes);
            Assert.DoesNotContain("ToCheck", In<Border>(window, "NewFrame").Classes);
            Assert.Null(pane.ResolveFileOfBoard("Scope baseline/U27_1.png"));
            Assert.Same(window.CurrentForTests!.OldPicture, In<Image>(window, "OldImage").Source);

            // Captured again: still the baseline the run started from beside it - and not yet accepted again.
            scope.Settings = new ScopeSettingsReading(0.0002, null, null);
            await window.CaptureCurrentAsync();

            Assert.Equal("Old baseline   50us, 1V, 1.5V", BoldRuns.TextOf(Header("OldHeaderText")));
            Assert.Equal(old, window.CurrentForTests!.BeforeBytes);
            Assert.Equal("New capture - not yet accepted   200us, 1V, 1.5V", BoldRuns.TextOf(Header("NewHeaderText")));
            Assert.Contains("ToCheck", In<Border>(window, "NewFrame").Classes);
            Assert.Equal("200us  (was 50us)", window.FactShownForTests(ComponentPaneWording.FactTimeDiv));

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // EVERY PICTURE IS CLICKED TO SEE IT LARGE (owner request, 2026-11-20: "click any image or thumbnail I see in this window and
    // then it will do the same image view as e.g. when working in worklogs"): the pin's picture, a thumbnail in the list - whose
    // press shows no other pin and sets no oscilloscope - and, captured again, the old baseline and the new capture; under each its
    // settings named, the VALUES bold (owner request, the same day: "again do remember to have the values "bold"").
    // ###########################################################################################
    [Fact]
    public async Task Every_picture_opens_large_its_values_bold_and_a_thumbnails_press_picks_no_other_pin()
    {
        DraftSession session = this.OpenDraft(earlier: Screen(4, 3, SkiaSharp.SKColors.Black));
        var scope = new FakeBaselineScope { Screen = Screen(4, 3, SkiaSharp.SKColors.White) };
        var viewers = new List<WorklogPhotoViewerWindow>();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane, viewers);
            await Settled(pane, window);

            // The pin shown's picture, pressed.
            Press(window, In<Panel>(window, "SinglePicture"));
            WorklogPhotoViewerWindow shown = Assert.Single(viewers);
            Assert.Equal("Pin 1 - 1A", shown.ShownTitleForTests);
            await Until(() => shown.ShownForTests is not null, "the picture read for the viewer");

            // Pin 3's thumbnail, pressed: its picture large - and still pin 1 shown, the oscilloscope as it was.
            int applied = scope.Applied.Count;
            Control line = In<ListBox>(window, "TargetsList").ContainerFromIndex(2)!;
            Panel thumbnail = line.GetVisualDescendants().OfType<Panel>().Single(panel => panel.Classes.Contains("Thumbnail"));
            Press(window, thumbnail);

            Assert.Equal(2, viewers.Count);
            Assert.Equal("Pin 3 - 1Y", viewers[1].ShownTitleForTests);
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal(applied, scope.Applied.Count);

            // Its settings each named, the values bold - and gone at a click outside it.
            Assert.Equal("T/DIV: 1us   |   V/DIV: 1V   |   Trigger level: 1.5V", RichNote.PlainOf(viewers[1].CommentForTests));
            Assert.Equal(["1us", "1V", "1.5V"], BoldOfNote(viewers[1].CommentForTests));
            Assert.All(viewers, viewer => Assert.True(viewer.ClosesWhenLeft));

            // Captured again: the old baseline and the new capture, not yet accepted, each pressed.
            await window.CaptureCurrentAsync();
            await Until(() => window.CurrentForTests!.OldPicture is not null && window.CurrentForTests!.WaitingPicture is not null, "the pictures read");

            Press(window, In<Panel>(window, "OldPicture"));
            Assert.Equal("Pin 1 - 1A - Old baseline", viewers[^1].ShownTitleForTests);
            Assert.Same(window.CurrentForTests!.OldPicture, viewers[^1].ShownForTests);
            Assert.Equal("T/DIV: 50us   |   V/DIV: 1V   |   Trigger level: 1.5V", RichNote.PlainOf(viewers[^1].CommentForTests));
            Assert.Equal(["50us", "1V", "1.5V"], BoldOfNote(viewers[^1].CommentForTests));

            Press(window, In<Panel>(window, "NewPicture"));
            Assert.Equal("Pin 1 - 1A - New capture", viewers[^1].ShownTitleForTests);
            Assert.Same(window.CurrentForTests!.WaitingPicture, viewers[^1].ShownForTests);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // "Click to see it larger" FOLLOWS THE POINTER over a picture (owner request, 2026-11-20: "have the hover/helper text follow the
    // mouse, as it now shows below image") - beside it, drawn in the window and taking no click, never a tooltip (whose own window
    // under the pointer takes the next click); off a picture, or over one with nothing to show, it goes.
    // ###########################################################################################
    [Fact]
    public async Task The_hint_follows_the_pointer_over_a_picture_and_goes_off_it()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);
            await Settled(pane, window);

            Panel picture = In<Panel>(window, "SinglePicture");
            Border hint = In<Border>(window, "PictureHint");

            Point Move(Control control, double x, double y)
            {
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                Dispatcher.UIThread.RunJobs();

                Point at = control.TranslatePoint(new Point(x, y), window)!.Value;
                window.MouseMove(at);
                Dispatcher.UIThread.RunJobs();
                return at;
            }

            Point first = Move(picture, 20, 30);
            Assert.True(hint.IsVisible);
            Assert.Equal(ComponentPaneWording.PictureHint, In<TextBlock>(window, "PictureHintText").Text);
            Assert.Equal(first.X + PointerHintGeometry.Right, Canvas.GetLeft(hint), 3);
            Assert.Equal(first.Y + PointerHintGeometry.Below, Canvas.GetTop(hint), 3);

            Point second = Move(picture, 120, 60);
            Assert.Equal(second.X + PointerHintGeometry.Right, Canvas.GetLeft(hint), 3);
            Assert.Equal(second.Y + PointerHintGeometry.Below, Canvas.GetTop(hint), 3);

            // Taking no click: what is under it is pressed.
            Assert.False(In<Canvas>(window, "HintLayer").IsHitTestVisible);
            Assert.Null(hint.GetValue(ToolTip.TipProperty));
            Assert.Null(picture.GetValue(ToolTip.TipProperty));

            // Off the picture, it goes.
            Move(In<Button>(window, "CaptureButton"), 5, 5);
            Assert.False(hint.IsVisible);

            // Over a pin with no picture, none.
            window.GoTo(1, setScope: false);
            Move(picture, 20, 30);
            Assert.False(hint.IsVisible);

            window.Close();
            host.Close();
        });
    }

    // ###########################################################################################
    // A picture seen large goes at a click outside it, as at Escape and its "Close" (owner request, 2026-11-20: "when I click with
    // mouse outside it, then it should also close the window") - shown beside the window, never as a dialog, which would refuse
    // that click; and "Close capturing" closes the window (owner request, the same day).
    // ###########################################################################################
    [Fact]
    public async Task A_picture_seen_large_goes_as_the_window_behind_it_is_clicked_and_close_capturing_closes()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);
            await Settled(pane, window);

            // The real viewer, not a stand-in.
            window.ShowPictureOverrideForTests = null;
            window.OpenPicture(window.CurrentForTests!);

            WorklogPhotoViewerWindow viewer = Assert.Single(window.OwnedWindows.OfType<WorklogPhotoViewerWindow>());
            Dispatcher.UIThread.RunJobs();
            Assert.True(viewer.IsVisible);
            await Until(() => viewer.ShownForTests is not null, "the picture read for the viewer");

            // Clicked outside it: the window system tells it it is no longer the active window - which the headless platform, making
            // no other window inactive as one is activated, is asked to say here - and the viewer goes.
            // (Its hook is Avalonia's internal - read by reflection.)
            object platform = viewer.PlatformImpl!;
            var deactivated = (Action?)platform.GetType()
                .GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic)
                .First(property => property.Name.EndsWith("Deactivated", StringComparison.Ordinal) && property.PropertyType == typeof(Action))
                .GetValue(platform);
            Assert.NotNull(deactivated);
            deactivated();
            Dispatcher.UIThread.RunJobs();
            Assert.False(viewer.IsVisible);
            Assert.Empty(window.OwnedWindows.OfType<WorklogPhotoViewerWindow>());

            // Nothing waiting to be accepted: "Close capturing" closes at once.
            Click(In<Button>(window, "CloseCaptureButton"));
            Assert.False(window.IsVisible);
            Assert.Null(pane.CaptureWindowForTests);

            host.Close();
        });
    }

    // ###########################################################################################
    // CLOSING WITH A CAPTURE NOT YET ACCEPTED ASKS FIRST - it is in no draft, so a slip of the hand would lose it: "Close capturing",
    // Escape or the window's own close shows a panel over the buttons. "Keep capturing" (or Escape) stays, "Discard and close" lets it
    // go - its screen gone, the draft as it was - and "Accept and close" puts it into the draft first.
    // ###########################################################################################
    [Fact]
    public async Task Closing_with_a_capture_not_yet_accepted_asks_to_accept_or_discard_it()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.0002, null, null) };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;
            await window.CaptureCurrentAsync();

            Click(In<Button>(window, "CloseCaptureButton"));
            Assert.True(window.IsVisible);
            Assert.True(window.IsAskingAboutWaitingForTests);
            Assert.Equal("1 capture is not yet accepted, so it is not in your draft. Accept it before closing, or discard it?", Said(window, "NotAcceptedText"));

            // Escape, or "Keep capturing": still capturing.
            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);
            Assert.False(window.IsAskingAboutWaitingForTests);
            Assert.True(window.IsVisible);

            window.Close();
            Assert.True(window.IsAskingAboutWaitingForTests);
            Click(In<Button>(window, "KeepCapturingButton"));
            Assert.False(window.IsAskingAboutWaitingForTests);
            Assert.True(window.IsVisible);

            // "Discard and close": gone, its screen with it, the draft as it was.
            await Settled(pane, window);
            window.Close();
            Click(In<Button>(window, "DiscardAndCloseButton"));
            Assert.False(window.IsVisible);
            Assert.Null(pane.CaptureWindowForTests);
            Assert.False(File.Exists(waiting));
            Assert.Equal("Scope baseline/U27_1.png", Baselines(session)[0].File);

            // "Accept and close": into the draft first.
            window = Open(pane);
            await window.CaptureCurrentAsync();
            window.Close();
            Click(In<Button>(window, "AcceptAndCloseButton"));
            await Until(() => !window.IsVisible, "the window closed once its capture is in");

            Assert.StartsWith("Scope baseline/U27_1_1A", Baselines(session)[0].File, StringComparison.Ordinal);
            Assert.Equal("200us", Baselines(session)[0].TimeDiv);
            Assert.False(File.Exists(waiting));

            await pane.ThumbnailsLoadedForTests;
            host.Close();
        });
    }

    // ###########################################################################################
    // THE BUTTONS NEVER MOVE (owner request, 2026-11-20: "The layout needs to be a little more static for the 3 buttons ... as it
    // can jump up and down, depending if it has image or not"): a pin with a picture, a new one with its name box and no picture,
    // and one captured again beside its earlier baseline all show them in the same place.
    // ###########################################################################################
    [Fact]
    public async Task The_buttons_stand_in_the_same_place_whatever_the_pin_shown()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            List<Point> Places()
            {
                window.UpdateLayout();
                return new[] { "CaptureButton", "PreviousButton", "AcceptButton" }
                    .Select(name => In<Button>(window, name).TranslatePoint(new Point(0, 0), window)!.Value)
                    .ToList();
            }

            List<Point> withPicture = Places();

            window.GoTo(1, setScope: true);
            Assert.True(In<StackPanel>(window, "NamePanel").IsVisible);
            Assert.Equal(withPicture, Places());

            window.GoTo(0, setScope: true);
            await window.CaptureCurrentAsync();
            await Until(() => window.CurrentForTests!.OldPicture is not null, "the old baseline's picture read");
            Assert.True(In<Grid>(window, "ComparePanel").IsVisible);
            Assert.Equal(withPicture, Places());

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // The component info window's keys (owner request, 2026-11-20: "support the same keyboard options as available in the
    // component info window, when having a scope connected"): Left the picked pin before, Right "Accept and continue to next pin",
    // the number keys to a pin, Space to the first; Up and Down the trigger level, numpad + and - the time per division, 1 and 2
    // the volts, "." "*" "/" the model's commands; Enter captures and stays, Escape closes - and in a text box the arrows and
    // digits are the box's.
    // ###########################################################################################
    [Fact]
    public async Task The_component_info_windows_keys_go_round_the_pins_set_the_oscilloscope_capture_and_close()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            In<Button>(window, "CaptureButton").Focus();

            void Hit(Key key, PhysicalKey physical) => window.KeyPress(key, RawInputModifiers.None, physical, keySymbol: null);

            // Nothing captured, Right only goes on - round the picked pins - and Left goes back.
            Hit(Key.Right, PhysicalKey.ArrowRight);
            await Until(() => window.CurrentIndexForTests == 2, "Right going on");
            Hit(Key.Right, PhysicalKey.ArrowRight);
            await Until(() => window.CurrentIndexForTests == 0, "Right going round");
            Hit(Key.Left, PhysicalKey.ArrowLeft);
            Assert.Equal(2, window.CurrentIndexForTests);
            Assert.DoesNotContain(window.TargetsForTests, target => target.IsAccepted);

            Hit(Key.Up, PhysicalKey.ArrowUp);
            Hit(Key.Down, PhysicalKey.ArrowDown);
            Hit(Key.Add, PhysicalKey.NumPadAdd);
            Hit(Key.Subtract, PhysicalKey.NumPadSubtract);
            Hit(Key.NumPad1, PhysicalKey.NumPad1);
            Hit(Key.NumPad2, PhysicalKey.NumPad2);
            Hit(Key.Decimal, PhysicalKey.NumPadDecimal);
            Hit(Key.Multiply, PhysicalKey.NumPadMultiply);
            Hit(Key.Divide, PhysicalKey.NumPadDivide);

            Assert.Equal([1, -1], scope.TriggerSteps);
            Assert.Equal([-1, 1], scope.TimeDivSteps);
            Assert.Equal([1.0, 2.0], scope.VoltsDivSet);
            Assert.Equal([ScopeCommandPalette.ClearStatistics, ScopeCommandPalette.Single, ScopeCommandPalette.Run], scope.Commands);
            Assert.Equal(2, window.CurrentIndexForTests);

            // The number keys: the pin typed so far, shown at once - "1" then "14" - and forgotten once typing pauses.
            Hit(Key.D1, PhysicalKey.Digit1);
            Assert.Equal(0, window.CurrentIndexForTests);
            Hit(Key.D4, PhysicalKey.Digit4);
            Assert.Equal(13, window.CurrentIndexForTests);
            Assert.Equal("Pin 14", In<TextBlock>(window, "TypedPinText").Text);
            await Until(() => !In<Border>(window, "TypedPinBorder").IsVisible, "the typed pin forgotten as typing pauses");

            Hit(Key.D9, PhysicalKey.Digit9);
            Hit(Key.D9, PhysicalKey.Digit9);
            Assert.Equal(8, window.CurrentIndexForTests);
            Assert.Equal("No pin 99", In<TextBlock>(window, "TypedPinText").Text);
            await Until(() => !In<Border>(window, "TypedPinBorder").IsVisible, "the typed pin forgotten again");

            Hit(Key.Space, PhysicalKey.Space);
            Assert.Equal(0, window.CurrentIndexForTests);

            // In a text box the arrows, digits and the oscilloscope's keys are the box's.
            In<TextBox>(window, "MachineStateBox").Focus();
            Hit(Key.Right, PhysicalKey.ArrowRight);
            Hit(Key.Up, PhysicalKey.ArrowUp);
            Hit(Key.D3, PhysicalKey.Digit3);
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal(2, scope.TriggerSteps.Count);

            // Enter captures, from the box too, and stays on the pin, not yet accepted; Right then accepts it and goes on.
            Hit(Key.Enter, PhysicalKey.Enter);
            await Until(() => window.TargetsForTests[0].IsWaiting, "the Enter key's capture");
            Assert.Equal(1, scope.Captures);
            Assert.Equal(0, window.CurrentIndexForTests);

            In<Button>(window, "CaptureButton").Focus();
            Hit(Key.Right, PhysicalKey.ArrowRight);
            await Until(() => window.TargetsForTests[0].IsAccepted, "the Right key's accepting");
            Assert.Equal(2, window.CurrentIndexForTests);

            await Settled(pane, window);

            // Nothing waiting: Escape closes at once.
            Hit(Key.Escape, PhysicalKey.Escape);
            Assert.Null(pane.CaptureWindowForTests);
            Assert.False(window.IsVisible);

            host.Close();
        });
    }

    // ###########################################################################################
    // The window opens where it was left, at the size it had (owner request, 2026-11-20: "remember size and position") -
    // maximized again when it was, its size before kept - over CRT's window when that place is on no screen any more, and made to fit
    // a screen smaller than the size it had.
    // ###########################################################################################
    [Fact]
    public void The_window_opens_where_it_was_left_and_is_remembered_as_it_closes()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());

            WindowPlace? remembered = null;
            pane.CaptureWindowPlace = () => new WindowPlace("Normal", 1200, 760, 40, 50);
            pane.RememberCaptureWindowPlace = place => remembered = place;

            BaselineCaptureWindow window = Open(pane);
            Assert.Equal((1200d, 760d), (window.Width, window.Height));
            Assert.Equal(new PixelPoint(40, 50), window.Position);

            window.Width = 1250;
            window.Position = new PixelPoint(60, 70);
            Dispatcher.UIThread.RunJobs();
            window.Close();

            Assert.Equal(new WindowPlace("Normal", 1250, 760, 60, 70), remembered);

            // Maximized: opened maximized, and its size and place before kept for when it is restored.
            pane.CaptureWindowPlace = () => remembered! with { State = "Maximized" };
            window = Open(pane);
            Assert.Equal(WindowState.Maximized, window.WindowState);
            window.Close();
            Assert.Equal(new WindowPlace("Maximized", 1250, 760, 60, 70), remembered);

            // A place on no screen: over CRT's window, at the size it had.
            pane.CaptureWindowPlace = () => new WindowPlace("Normal", 1200, 760, 1_000_000, 1_000_000);
            window = Open(pane);
            Assert.Equal(WindowStartupLocation.CenterOwner, window.WindowStartupLocation);
            Assert.NotEqual(new PixelPoint(1_000_000, 1_000_000), window.Position);
            Assert.Equal(1200d, window.Width);
            window.Close();

            // Left larger than the screen it opens on: made to fit it, its foot - the button bar - on the screen (code review,
            // 2026-11-21: left on a large monitor, it opened off the bottom of a laptop's).
            pane.CaptureWindowPlace = () => new WindowPlace("Normal", 1200, 100_000, 40, 50);
            window = Open(pane);
            Avalonia.Platform.Screen screen = window.Screens.ScreenFromPoint(new PixelPoint(40, 50))!;
            double foot = window.Position.Y + ((window.Height + DialogPlacement.FrameAllowance.Height) * screen.Scaling);
            Assert.True(window.Height < 100_000);
            Assert.True(foot <= screen.WorkingArea.Bottom, $"its foot at {foot}, the screen's at {screen.WorkingArea.Bottom}");
            window.Close();

            host.Close();
        });
    }

    // ###########################################################################################
    // THE LINES FOLLOW THE DRAFT (code review, 2026-11-21): an "Undo" of an accepted capture takes its line back - not accepted,
    // not counted, not framed on the case, no baseline - and a baseline another editor adds meanwhile gets a line of its own, in its
    // pin's place, picked as that pin was; "Capture again" on it captures into it.
    // ###########################################################################################
    [Fact]
    public async Task An_undo_of_an_accepted_capture_takes_its_line_back_and_a_baseline_added_meanwhile_gets_its_own()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            PinoutDrawing drawing = In<PinoutDrawing>(window, "CasePinout");

            window.PickPins(pins => pins.Where(pin => pin.Pin == "5").ToList());
            BaselineCaptureTarget pin5 = window.TargetsForTests[4];
            Assert.Same(pin5, window.CurrentForTests);

            await window.CaptureCurrentAsync();
            await window.AcceptAndContinueAsync();
            Assert.True(pin5.IsAccepted);
            Assert.Equal(3, Baselines(session).Count);
            Assert.Equal([5], drawing.DonePinsForTests);

            Assert.True(session.Undo());
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(2, Baselines(session).Count);
            Assert.Same(pin5, window.TargetsForTests[4]);
            Assert.False(pin5.IsAccepted);
            Assert.Null(pin5.Row);
            Assert.Equal("No baseline yet", pin5.StateText);
            Assert.Equal("0 of 1 picked baseline accepted", Said(window, "ProgressText"));
            Assert.Empty(drawing.DonePinsForTests);

            // A "Redo" brings it back: accepted again (code review, 2026-11-21) - and taken back once more.
            Assert.True(session.Redo());
            Dispatcher.UIThread.RunJobs();
            Assert.True(pin5.IsAccepted);
            Assert.NotNull(pin5.Row);
            Assert.Equal("1 of 1 picked baseline accepted", Said(window, "ProgressText"));
            Assert.Equal([5], drawing.DonePinsForTests);

            Assert.True(session.Undo());
            Dispatcher.UIThread.RunJobs();
            Assert.False(pin5.IsAccepted);

            // Another editor gives pin 5 a baseline: its own line, in the place of pin 5's new one, picked as that was.
            AddBaseline(session, new ComponentImageEntry { BoardLabel = "U27", Pin = "5", Name = "2A", File = "Scope baseline/U27_3.png", TimeDiv = "2us" });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(14, window.TargetsForTests.Count);
            BaselineCaptureTarget added = window.TargetsForTests[4];
            Assert.NotSame(pin5, added);
            Assert.Equal(("5  2A", "Has a baseline", true), (added.Label, added.StateText, added.IsPicked));
            Assert.Same(pane.BaselinesForTests[2], added.Row);

            // "Capture again" on it: its line, the oscilloscope set as it was measured, the capture into it once accepted.
            await pane.CaptureAgainInWindowAsync(pane.BaselinesForTests[2]);
            Assert.Same(added, window.CurrentForTests);
            Assert.Equal(("5", "2us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));
            Assert.True(added.IsWaiting);

            await window.AcceptAndContinueAsync();
            Assert.Equal(3, Baselines(session).Count);
            Assert.StartsWith("Scope baseline/U27_5_2A", Baselines(session)[2].File, StringComparison.Ordinal);
            Assert.True(added.IsAccepted);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // A pin clicked on the case picks EVERY line of it - a pin measured twice has two (code review, 2026-11-21: only the first was
    // picked, and the second never visited) - and the quick picks ask whether a pin has a baseline as it is NOW: one given its first
    // in this run is one "with", no longer one "without".
    // ###########################################################################################
    [Fact]
    public async Task A_pin_clicked_on_the_case_picks_every_line_of_it_and_the_quick_picks_go_by_the_baselines_now()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            AddBaseline(session, new ComponentImageEntry { BoardLabel = "U27", Pin = "3 [running]", Name = "1Y", File = "Scope baseline/U27_3.png" });

            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);
            Assert.Equal(15, window.TargetsForTests.Count);

            Click(In<Button>(window, "PickNoPinButton"));
            In<PinoutDrawing>(window, "CasePinout").RaiseEvent(new PinPickedEventArgs(PinoutDrawing.PinPickedEvent, 3));
            Assert.Equal(["3", "3 [running]"], Picked(window));

            // Pin 2 given its first baseline in this run.
            window.PickPins(pins => pins.Where(pin => pin.Pin == "2").ToList());
            await window.CaptureCurrentAsync();
            await window.AcceptAndContinueAsync();
            Assert.True(window.TargetsForTests[1].IsAccepted);

            Click(In<Button>(window, "PickWithBaselineButton"));
            Assert.Contains("2", Picked(window));

            Click(In<Button>(window, "PickWithoutBaselineButton"));
            Assert.DoesNotContain("2", Picked(window));

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // A CAPTURE STILL BEING TAKEN AS THE WINDOW CLOSES IS NOT KEPT (code review, 2026-11-21): the window's own close is covered by no
    // "please wait", and a screen coming in after it was held by a closed window and left in the user's folder for good.
    // ###########################################################################################
    [Fact]
    public async Task A_capture_that_comes_in_after_the_window_closed_is_not_kept()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { CaptureGate = new TaskCompletionSource<byte[]?>() };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            var written = new List<string>();
            pane.CaptureWrittenForTests = written.Add;

            Task capturing = window.CaptureCurrentAsync();
            await Until(() => scope.Captures == 1, "the capture asked of the oscilloscope");

            // Nothing waiting yet: it closes at once.
            window.Close();
            Assert.False(window.IsVisible);

            scope.CaptureGate.SetResult(FakeBaselineScope.Png);
            await capturing;

            string screen = Assert.Single(written);
            Assert.False(File.Exists(screen));
            Assert.DoesNotContain(window.TargetsForTests, target => target.IsWaiting);

            await window.ImagesReadForTests;
            await pane.ThumbnailsLoadedForTests;
            host.Close();
        });
    }

    // ###########################################################################################
    // QUITTING CRT WITH CAPTURES NOT YET ACCEPTED ASKS FIRST, AND IS CALLED OFF (code review, 2026-11-22: they were deleted without a
    // word, though the window's own close asks): CRT asks this window before its own Closing, and the window holds the quitting and
    // asks what its own close asks. The captures stay to be accepted (code review, 2026-11-21: a quitting called off had let them
    // go). Discarded, CRT quits.
    // ###########################################################################################
    [Fact]
    public async Task Quitting_with_captures_not_yet_accepted_asks_first_and_keeps_them()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.0002, null, null) };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;
            await window.CaptureCurrentAsync();

            bool crtAsked = false;
            host.Closing += (_, _) => crtAsked = true;
            host.Close();

            Assert.True(host.IsVisible);
            Assert.True(window.IsVisible);
            Assert.True(window.IsAskingAboutWaitingForTests);
            Assert.False(crtAsked);
            Assert.True(File.Exists(waiting));

            // Still there to be accepted.
            await window.AcceptAndContinueAsync();
            Assert.StartsWith("Scope baseline/U27_1_1A", Baselines(session)[0].File, StringComparison.Ordinal);
            Assert.Equal("200us", Baselines(session)[0].TimeDiv);

            // Another waiting: asked again; discarded, CRT quits.
            await window.CaptureCurrentAsync();
            string second = waiting!;
            await Settled(pane, window);
            host.Close();

            Assert.True(host.IsVisible);
            Assert.True(File.Exists(second));

            Click(In<Button>(window, "DiscardAndCloseButton"));
            Assert.False(window.IsVisible);
            Assert.False(File.Exists(second));

            host.Close();
            Assert.False(host.IsVisible);
            Assert.True(crtAsked);
        });
    }

    // ###########################################################################################
    // A CAPTURE WAITING OPENS AS ITSELF, READ OR NOT (code review, 2026-11-21): one whose picture was not read yet - or could not be -
    // opened the OLD baseline under the title "New capture"; it opens from where it waits, with the capture's own settings.
    // ###########################################################################################
    [Fact]
    public async Task A_capture_not_read_opens_as_itself_never_as_the_old_baseline()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope { Settings = new ScopeSettingsReading(0.0001, null, null) };
        var viewers = new List<WorklogPhotoViewerWindow>();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane, viewers);
            await Settled(pane, window);

            // Its picture cannot be read: gone from where it waits as soon as it is written.
            pane.CaptureWrittenForTests = path => File.Delete(path);
            await window.CaptureCurrentAsync();

            BaselineCaptureTarget pin1 = window.CurrentForTests!;
            await Until(() => pin1.WaitingPictureRead!.IsCompleted, "the capture's picture tried");
            Assert.Null(pin1.WaitingPicture);

            window.OpenPicture(pin1);

            WorklogPhotoViewerWindow viewer = Assert.Single(viewers);
            Assert.Equal("Pin 1 - 1A - New capture", viewer.ShownTitleForTests);
            Assert.Equal("T/DIV: 100us   |   V/DIV: 1V   |   Trigger level: 1.5V", RichNote.PlainOf(viewer.CommentForTests));
            Assert.Null(viewer.ShownForTests);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // NO KEY WHILE THE WINDOW WAITS (code review, 2026-11-21): a capture waits for the oscilloscope to be set for its pin, with no
    // "please wait" up yet - and Left then set it for another pin before the screen and the settings were read, so the capture was
    // stored with that pin's settings. And Enter on a focused button presses that button - never a capture - Space on a line's
    // checkbox checks it, and with the close question showing Enter captures nothing.
    // ###########################################################################################
    [Fact]
    public async Task No_key_works_while_the_window_waits_and_enter_and_space_are_a_focused_controls()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            In<Button>(window, "CaptureButton").Focus();

            void Hit(Key key, PhysicalKey physical) => window.KeyPress(key, RawInputModifiers.None, physical, keySymbol: null);

            // The oscilloscope slow to be set for pin 1: the capture waits for it.
            var applying = new TaskCompletionSource();
            scope.ApplyGate = applying;
            window.GoTo(0, setScope: true);
            Task capturing = window.CaptureCurrentAsync();

            int applied = scope.Applied.Count;
            Hit(Key.Left, PhysicalKey.ArrowLeft);
            Hit(Key.Up, PhysicalKey.ArrowUp);
            Hit(Key.D3, PhysicalKey.Digit3);
            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal(applied, scope.Applied.Count);
            Assert.Empty(scope.TriggerSteps);

            scope.ApplyGate = null;
            applying.SetResult();
            await capturing;
            Assert.True(window.TargetsForTests[0].IsWaiting);

            // Enter is the footswitch's capture wherever the keyboard is - on a button a click left it on too, never that button's.
            window.GoTo(2, setScope: false);
            In<Button>(window, "PreviousButton").Focus();
            Hit(Key.Enter, PhysicalKey.Enter);
            await Until(() => scope.Captures == 2 && !window.IsBusy && window.TargetsForTests[2].IsWaiting, "Enter's capture taken");
            Assert.Equal(2, window.CurrentIndexForTests);

            // Space on a line's checkbox checks it - the pin shown stays.
            Control line = In<ListBox>(window, "TargetsList").ContainerFromIndex(4)!;
            CheckBox check = line.GetVisualDescendants().OfType<CheckBox>().First();
            check.Focus();
            Hit(Key.Space, PhysicalKey.Space);
            window.KeyRelease(Key.Space, RawInputModifiers.None, PhysicalKey.Space, keySymbol: null);
            Dispatcher.UIThread.RunJobs();
            Assert.True(window.TargetsForTests[4].IsPicked);
            Assert.Equal(2, window.CurrentIndexForTests);

            // With the close question showing, it alone is answered - the keyboard put on "Keep capturing": Enter and Right from the
            // machine state, Enter and Space on "Capture" and "Accept and continue to next pin" (where a click left the keyboard) do
            // nothing, and Enter on "Keep capturing" presses it.
            window.Close();
            Assert.True(window.IsAskingAboutWaitingForTests);
            Assert.True(In<Button>(window, "KeepCapturingButton").IsFocused);
            In<TextBox>(window, "MachineStateBox").Focus();
            Hit(Key.Enter, PhysicalKey.Enter);
            In<ListBox>(window, "TargetsList").ContainerFromIndex(2)!.Focus();
            Hit(Key.Right, PhysicalKey.ArrowRight);

            foreach (string name in new[] { "CaptureButton", "AcceptButton" })
            {
                In<Button>(window, name).Focus();
                Hit(Key.Enter, PhysicalKey.Enter);
                Hit(Key.Space, PhysicalKey.Space);
                window.KeyRelease(Key.Space, RawInputModifiers.None, PhysicalKey.Space, keySymbol: null);
            }

            Dispatcher.UIThread.RunJobs();
            Assert.Equal(2, scope.Captures);
            Assert.False(window.TargetsForTests[2].IsAccepted);
            Assert.True(window.IsAskingAboutWaitingForTests);

            In<Button>(window, "KeepCapturingButton").Focus();
            Hit(Key.Enter, PhysicalKey.Enter);
            Dispatcher.UIThread.RunJobs();
            Assert.False(window.IsAskingAboutWaitingForTests);
            Assert.Equal(2, scope.Captures);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // While a capture goes into the draft the window's own close waits for it (code review, 2026-11-21: Escape and the close question
    // could act on a capture that was going in anyway) - and then CLOSES, asked once (second code review, the same day: the close was
    // dropped without a word, to be asked again), the oscilloscope not set for a next pin nobody will see. And a baseline the pane
    // holds back, its signal name emptied, takes no capture: said, the capture waiting on, the draft as it was (it read "Accepted" and
    // the screen was gone).
    // ###########################################################################################
    [Fact]
    public async Task Closing_waits_for_a_capture_going_in_then_closes_and_a_baseline_held_back_takes_none()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            await window.CaptureCurrentAsync();
            int applied = scope.Applied.Count;
            Task accepting = window.AcceptAndContinueAsync();
            window.Close();
            Assert.True(window.IsVisible);
            Assert.False(window.IsAskingAboutWaitingForTests);
            await accepting;
            Assert.True(window.TargetsForTests[0].IsAccepted);

            Dispatcher.UIThread.RunJobs();
            Assert.False(window.IsVisible);
            Assert.Null(pane.CaptureWindowForTests);
            Assert.Equal(applied, scope.Applied.Count);

            // Pin 3's signal name emptied in the pane: held back there.
            window = Open(pane);
            pane.BaselinesForTests[1].Name = "";
            pane.CommitHeld();

            string? waiting = null;
            pane.CaptureWrittenForTests = path => waiting = path;
            Assert.True(window.GoToPin("3", pickIt: false));
            await window.CaptureCurrentAsync();
            await window.AcceptAndContinueAsync();

            Assert.Equal(
                ComponentPaneWording.BaselineHeldBack("An oscilloscope baseline needs its signal name."),
                In<TextBlock>(window, "ProblemText").Text);
            Assert.True(window.TargetsForTests[2].IsWaiting);
            Assert.True(File.Exists(waiting));
            Assert.Equal("Scope baseline/U27_3.png", Baselines(session)[1].File);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // THE OSCILLOSCOPE BACK IS SET FOR THE PIN SHOWN (code review, 2026-11-21): a pin moved to while it was gone was captured with
    // the oscilloscope as it was set for the pin before.
    // ###########################################################################################
    [Fact]
    public void The_oscilloscope_back_is_set_for_the_pin_shown()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            Assert.Equal("1", scope.Applied[^1].Pin);

            pane.ShowScopeConnected(false);
            window.GoTo(2, setScope: true);
            int applied = scope.Applied.Count;

            pane.ShowScopeConnected(true);
            Assert.Equal(applied + 1, scope.Applied.Count);
            Assert.Equal(("3", "1us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));

            // Still connected, told again: nothing sent again.
            pane.ShowScopeConnected(true);
            Assert.Equal(applied + 1, scope.Applied.Count);

            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // "ACCEPT AND CLOSE" STOPPING ON A CAPTURE IT CANNOT ACCEPT SAYS WHY (code review, 2026-11-21): going to that pin hid the reason,
    // and the button seemed to do nothing. A new pin's capture with no signal name: the window stays, on it, saying so.
    // ###########################################################################################
    [Fact]
    public async Task Accept_and_close_stopping_on_a_capture_says_why_on_its_pin()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);

            window.PickPins(pins => pins.Where(pin => pin.Pin == "5").ToList());
            In<TextBox>(window, "NameBox").Text = "";
            await window.CaptureCurrentAsync();
            window.GoTo(0, setScope: false);

            window.Close();
            Assert.True(window.IsAskingAboutWaitingForTests);
            await window.AcceptWaitingAndCloseAsync();

            Assert.True(window.IsVisible);
            Assert.Equal(4, window.CurrentIndexForTests);
            Assert.True(In<TextBlock>(window, "ProblemText").IsVisible);
            Assert.Equal(ComponentPaneWording.RunNeedsName, In<TextBlock>(window, "ProblemText").Text);
            Assert.True(window.TargetsForTests[4].IsWaiting);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // ON A BOARD WITH REGIONS A NEW PIN'S CAPTURE IS SHOWN IN THE REGION IT GOES INTO (code review, 2026-11-21: "Applies to: All
    // regions" before accepting, PAL in the draft after) - one rule for both (BaselineSettingWords.Captured).
    // ###########################################################################################
    [Fact]
    public async Task On_a_board_with_regions_a_new_pins_capture_is_shown_in_the_region_it_goes_into()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            session.Apply("Regions", content => content with { Board = content.Board.WithRegions(["PAL", "NTSC"]) });

            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);

            window.PickPins(pins => pins.Where(pin => pin.Pin == "5").ToList());
            await window.CaptureCurrentAsync();
            Assert.Equal("PAL", window.FactShownForTests(ComponentPaneWording.FactAppliesTo));

            await window.AcceptAndContinueAsync();
            Assert.Equal(("5", "PAL"), (Baselines(session)[^1].Pin, Baselines(session)[^1].Region));

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // NOR DOES THE MOUSE MOVE TO ANOTHER PIN WHILE A CAPTURE IS UNDER WAY (code review, 2026-11-21: "Previous pin" clicked while the
    // oscilloscope was still being set for the pin captured set it for another, and the screen was stored with that pin's settings):
    // "Previous pin", a line clicked, a quick pick and a pin clicked on the case all wait. And the case is drawn again as the pane's
    // package changes - a pin it adds is one to click.
    // ###########################################################################################
    [Fact]
    public async Task No_click_moves_to_another_pin_while_a_capture_is_under_way_and_the_case_follows_the_package()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            PinoutDrawing drawing = In<PinoutDrawing>(window, "CasePinout");

            var applying = new TaskCompletionSource();
            scope.ApplyGate = applying;
            window.GoTo(0, setScope: true);
            Task capturing = window.CaptureCurrentAsync();

            // Waiting for the oscilloscope is a wait the contributor watches - said, under the window's "please wait".
            BusyOverlay overlay = BusyOverlay.For(window)!;
            Assert.True(overlay.IsBusy);
            Assert.Equal(CrtWaitWording.SettingOscilloscope, overlay.Message);

            int applied = scope.Applied.Count;
            Click(In<Button>(window, "PreviousButton"));
            In<ListBox>(window, "TargetsList").SelectedIndex = 2;
            Click(In<Button>(window, "PickEveryPinButton"));
            drawing.RaiseEvent(new PinPickedEventArgs(PinoutDrawing.PinPickedEvent, 7));

            Assert.Equal(0, window.CurrentIndexForTests);
            Assert.Equal(0, In<ListBox>(window, "TargetsList").SelectedIndex);
            Assert.Equal(["1", "3"], Picked(window));
            Assert.Equal(applied, scope.Applied.Count);

            scope.ApplyGate = null;
            applying.SetResult();
            await capturing;
            Assert.True(window.TargetsForTests[0].IsWaiting);

            // A DIP-16 now: the case has its pins 15 and 16 to click.
            Assert.Equal(14, drawing.LayoutForTests!.Pins.Count);
            Get<AutoCompleteBox>(pane, "PackageBox").Text = "DIP-16";
            pane.CommitHeld();
            Assert.Equal(16, window.TargetsForTests.Count(target => target.Step.Pin.Number is not null));
            Assert.Equal(16, drawing.LayoutForTests!.Pins.Count);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // THE WINDOW FITS THE SCREEN OF CRT'S WINDOW THE FIRST TIME TOO (code review, 2026-11-22: at its 1280 x 860, a laptop of 768 had
    // its button bar and keys under the screen's edge) - and a place kept on no screen any more is fitted to that screen.
    // ###########################################################################################
    [Fact]
    public void The_window_fits_the_screen_it_opens_on_with_no_place_kept_or_one_on_no_screen()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());

            // A laptop's 1366 x 728 working area: 728 high less the frame's 40, as wide as it was.
            pane.CaptureWindowScreenForTests = (new PixelRect(0, 0, 1366, 728), 1.0);
            BaselineCaptureWindow window = Open(pane);
            Assert.Equal((1280d, 688d), (window.Width, window.Height));
            window.Close();

            pane.CaptureWindowPlace = () => new WindowPlace("Normal", 1200, 1000, 1_000_000, 1_000_000);
            window = Open(pane);
            Assert.Equal((1200d, 688d), (window.Width, window.Height));
            window.Close();

            host.Close();
        });
    }

    // ###########################################################################################
    // ONLY THE LAST PIN SHOWN GOES TO THE OSCILLOSCOPE (code review, 2026-11-22: each pin passed - Left held, "1" then "2" typed - queued
    // a set-up of its own, and a capture waited for them all while the oscilloscope stepped through pins nobody looked at): one set-up
    // at a time, the latest waiting, and a capture takes the pin shown last.
    // ###########################################################################################
    [Fact]
    public async Task Only_the_last_pin_shown_is_sent_to_the_oscilloscope()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);
            int before = scope.Applied.Count;

            // Pins 1 and 3 - the two with a baseline - shown in turn while the first set-up is still going out.
            var applying = new TaskCompletionSource();
            scope.ApplyGate = applying;
            window.GoTo(2, setScope: true);
            window.GoTo(0, setScope: true);
            window.GoTo(2, setScope: true);
            window.GoTo(0, setScope: true);
            window.GoTo(2, setScope: true);

            Assert.Equal(before + 1, scope.Applied.Count);

            scope.ApplyGate = null;
            applying.SetResult();
            await window.CaptureCurrentAsync();

            Assert.Equal(before + 2, scope.Applied.Count);
            Assert.Equal(("3", "1us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));
            Assert.True(window.TargetsForTests[2].IsWaiting);

            await Settled(pane, window);
            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // ANOTHER LINE SHOWN IN PLACE OF ONE GONE IS SET FOR (code review, 2026-11-22): the baseline shown taken away - deleted, or given
    // another pin - the window showed the pin's other baseline, but the oscilloscope stayed set for the one gone, and Enter captured it
    // with those settings.
    // ###########################################################################################
    [Fact]
    public void Another_line_shown_in_place_of_one_gone_sets_the_oscilloscope_for_it()
    {
        DraftSession session = this.OpenDraft();
        var scope = new FakeBaselineScope();

        UiTest.Run(() =>
        {
            // Pin 3 with a second baseline, at 2us.
            AddBaseline(session, new ComponentImageEntry { BoardLabel = "U27", Pin = "3", Name = "1Y", File = "Scope baseline/U27_1.png", TimeDiv = "2us", VoltsDiv = "1V", TriggerLevelVolts = "1.5V" });

            (ComponentPane pane, Window host) = Shown(session, scope);
            BaselineCaptureWindow window = Open(pane);

            int first = window.TargetsForTests.ToList().FindIndex(target => target.Row?.Committed.TimeDiv == "1us");
            window.GoTo(first, setScope: true);
            Assert.Equal("1us", scope.Applied[^1].TimeDiv);

            // The 1us baseline taken away by another editor: the 2us one is shown, and the oscilloscope set for it.
            session.Apply("Another editor", content => content with
            {
                Board = content.Board.WithComponentSections(
                    componentImages: content.Board.ComponentImages.Where(image => !(image.Pin == "3" && image.TimeDiv == "1us")).ToList()),
            });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("2us", window.CurrentForTests!.Row!.Committed.TimeDiv);
            Assert.Equal(("3", "2us"), (scope.Applied[^1].Pin, scope.Applied[^1].TimeDiv));

            window.LetGo();
            host.Close();
        });
    }

    // ###########################################################################################
    // WITH NO LINE LEFT, NO PIN IS SHOWN (code review, 2026-11-22: every pin and baseline taken away in the pane, the window kept the
    // pin before - its heading, fields, picture and name box), and there is nothing to capture.
    // ###########################################################################################
    [Fact]
    public void With_no_line_left_no_pin_is_shown_and_nothing_can_be_captured()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window host) = Shown(session, new FakeBaselineScope());
            BaselineCaptureWindow window = Open(pane);
            Assert.NotEqual(string.Empty, Said(window, "CurrentText"));

            // Another editor takes its case, its type number, its pins and its baselines away: nothing to go round.
            session.Apply("Another editor", content =>
            {
                BoardData board = ComponentPaneEdits.WithField(content.Board, "U27", "PAL", ComponentField.Package, string.Empty);
                board = ComponentPaneEdits.WithField(board, "U27", "PAL", ComponentField.TechnicalNameOrValue, string.Empty);

                return content with
                {
                    Board = board.WithComponentSections(
                        componentPins: board.ComponentPins.Where(pin => pin.BoardLabel != "U27").ToList(),
                        componentImages: board.ComponentImages.Where(image => image.BoardLabel != "U27").ToList()),
                };
            });
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(window.TargetsForTests);
            Assert.Equal(string.Empty, Said(window, "CurrentText"));
            Assert.Equal(string.Empty, In<TextBlock>(window, "CurrentDetails").Text);
            Assert.False(In<StackPanel>(window, "NamePanel").IsVisible);
            Assert.Null(window.FactShownForTests(ComponentPaneWording.FactPin));
            Assert.Null(In<Image>(window, "CurrentImage").Source);
            Assert.False(In<Button>(window, "CaptureButton").IsEnabled);

            window.LetGo();
            host.Close();
        });
    }
}
