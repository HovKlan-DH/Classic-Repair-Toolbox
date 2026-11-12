using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;
using Handlers.Theming;
using Handlers.Tours;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// TourOverlay - an introduction tour over a whole window (owner request, 2026-11-11). What must hold: the first card says
// what tour it is; each step frames what it is about with its card beside it, not over it; "Back", "Next", "Done",
// "Skip tour" and the keys move through it and end it; and it is a VIEW-DEMO ONLY (owner decision: "No - this is a
// view-demo only") - nothing under it takes a click or a key, the framed part included.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TourOverlayTests
{
    // A window as a host lays one out: a button and a text box, then the overlay over them.
    private sealed class Host
    {
        public Host()
        {
            this.Button.Click += (_, _) => this.Clicks++;

            var content = new Canvas { Width = 900, Height = 600 };
            Canvas.SetLeft(this.Button, 40);
            Canvas.SetTop(this.Button, 40);
            Canvas.SetLeft(this.TextBox, 40);
            Canvas.SetTop(this.TextBox, 120);
            content.Children.Add(this.Button);
            content.Children.Add(this.TextBox);

            this.Window = new Window { Content = new Grid { Children = { content, this.Overlay } }, Width = 900, Height = 600 };
            this.Overlay.Ended += completed => this.Endings.Add(completed);
        }

        public Window Window { get; }

        public Button Button { get; } = new() { Content = "Under the tour", Width = 160, Height = 40 };

        public TextBox TextBox { get; } = new() { Width = 200 };

        public TourOverlay Overlay { get; } = new();

        public int Clicks { get; set; }

        public List<bool> Endings { get; } = [];

        // The first step frames the button, the second stands in the middle.
        public TourPresentation Tour(bool whatIsNew = false) => new(
            whatIsNew ? TourWording.Kicker(TourKind.WhatIsNew) : TourWording.Kicker(TourKind.Introduction),
            whatIsNew,
            "The \"Schematics\" tab",
            "What the tab is for.",
            TourWording.TakeAgainFoot,
            [
                new TourCardStep("The button", "It does something.", whatIsNew ? "New in 4.1.0" : null, overlay => this.Button.TranslatePoint(new Point(0, 0), overlay) is Point at ? new Rect(at, this.Button.Bounds.Size) : null),
                new TourCardStep("In the middle", "About nothing in particular.", null, null),
            ]);

        public void Show(TourPresentation tour)
        {
            this.Window.Show();
            Settle();
            this.Overlay.Show(tour);
            Settle();
        }

        public Rect ButtonRect => new(this.Button.TranslatePoint(new Point(0, 0), this.Window)!.Value, this.Button.Bounds.Size);

        // A press and a release on what is drawn now - a press hits the last frame drawn.
        public void Click(Point point)
        {
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            Dispatcher.UIThread.RunJobs();
            this.Window.MouseDown(point, MouseButton.Left);
            this.Window.MouseUp(point, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
        }

        public void Key(Key key, PhysicalKey physical)
        {
            this.Window.KeyPress(key, RawInputModifiers.None, physical, keySymbol: null);
            this.Window.KeyRelease(key, RawInputModifiers.None, physical, keySymbol: null);
            Dispatcher.UIThread.RunJobs();
        }
    }

    private static void Settle()
    {
        for (int pass = 0; pass < 3; pass++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
    }

    private static string? Text(TourOverlay overlay, string name) => overlay.GetControl<TextBlock>(name).Text;

    private static Button ButtonOf(TourOverlay overlay, string name) => overlay.GetControl<Button>(name);

    [Fact]
    public void Hidden_until_a_tour_is_shown()
    {
        UiTest.Run(() =>
        {
            var host = new Host();

            Assert.False(host.Overlay.IsVisible);
            Assert.False(host.Overlay.IsShowing);
        });
    }

    // *** THE FIRST CARD SAYS WHAT TOUR IT IS (owner request, 2026-11-11: "it should clearly show this is an
    // "Introduction tour""). *** Nothing framed yet, "Start tour" and "Skip tour".
    [Fact]
    public void The_first_card_says_it_is_an_introduction_tour_and_frames_nothing()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());

            Assert.True(host.Overlay.IsVisible);
            Assert.Equal(-1, host.Overlay.StepShown);
            Assert.Equal("Introduction tour", Text(host.Overlay, "KickerText"));
            Assert.Equal("The \"Schematics\" tab", Text(host.Overlay, "HeadlineText"));
            Assert.Equal(BoldRuns.Plain(TourWording.TakeAgainFoot), BoldRuns.TextOf(host.Overlay.GetControl<TextBlock>("FootText")));
            Assert.Equal(["Configuration"], BoldRuns.BoldPiecesOf(host.Overlay.GetControl<TextBlock>("FootText")));
            Assert.Null(host.Overlay.HoleShown);
            Assert.Equal("Start tour", ButtonOf(host.Overlay, "NextButton").Content);
            Assert.Equal("Skip tour", ButtonOf(host.Overlay, "SkipButton").Content);
            Assert.False(ButtonOf(host.Overlay, "BackButton").IsVisible);
            Assert.Contains("Welcome", host.Overlay.GetControl<Border>("Card").Classes);
        });
    }

    [Fact]
    public void A_tour_of_what_changed_says_so_and_badges_each_step_with_its_version()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour(whatIsNew: true));

            Assert.Equal("Tour of what's new", Text(host.Overlay, "KickerText"));

            host.Overlay.Next();
            Settle();

            Assert.True(host.Overlay.GetControl<Border>("NewBadge").IsVisible);
            Assert.Equal("New in 4.1.0", Text(host.Overlay, "NewBadgeText"));
        });
    }

    // Each step frames what it is about, its card beside it - never over it - and says where the tour stands.
    [Fact]
    public void A_step_frames_what_it_is_about_with_its_card_beside_it()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());

            host.Overlay.Next();
            Settle();

            Assert.Equal(0, host.Overlay.StepShown);
            Rect hole = Assert.NotNull(host.Overlay.HoleShown);
            Assert.True(hole.Contains(host.ButtonRect), $"the hole {hole} does not hold the button {host.ButtonRect}");
            Assert.False(host.Overlay.CardShown.Intersects(hole), $"the card {host.Overlay.CardShown} covers the hole {hole}");
            Assert.Equal("Step 1 of 2", Text(host.Overlay, "ProgressText"));
            Assert.Equal("The button", Text(host.Overlay, "HeadlineText"));
            Assert.Equal("Next", ButtonOf(host.Overlay, "NextButton").Content);
            Assert.True(ButtonOf(host.Overlay, "BackButton").IsVisible);
            Assert.False(ButtonOf(host.Overlay, "BackButton").IsEnabled);
            Assert.False(host.Overlay.GetControl<Border>("NewBadge").IsVisible);
        });
    }

    // A step's words say their own bold pieces (owner's texts, 2026-11-11: "hold down <b>Shift</b>") - drawn bold,
    // the tags never on the card; and the next card's plain words are plain again.
    [Fact]
    public void A_step_s_marked_words_are_bold_and_its_tags_are_never_shown()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(new TourPresentation(
                TourWording.Kicker(TourKind.Introduction),
                false,
                "The \"Schematics\" tab",
                "What the tab is for.",
                TourWording.TakeAgainFoot,
                [
                    new TourCardStep("Keys", "Hold down <b>Shift</b> and click.", null, null),
                    new TourCardStep("Plain", "Nothing bold here.", null, null),
                ]));

            host.Overlay.Next();
            Settle();

            TextBlock body = host.Overlay.GetControl<TextBlock>("BodyText");
            Assert.Equal("Hold down Shift and click.", BoldRuns.TextOf(body));
            Assert.Equal(["Shift"], BoldRuns.BoldPiecesOf(body));

            host.Overlay.Next();
            Settle();

            Assert.Equal("Nothing bold here.", BoldRuns.TextOf(body));
            Assert.Empty(BoldRuns.BoldPiecesOf(body));
        });
    }

    // A step that sets the window up for itself (the "Schematics" tour's panels shown expanded) puts it back however it
    // is left - "Next", "Back", or the tour skipped - and the first card sets up nothing.
    [Fact]
    public void A_step_s_set_up_is_put_back_however_the_step_is_left()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            var said = new List<string>();

            host.Show(new TourPresentation(
                TourWording.Kicker(TourKind.Introduction),
                false,
                "The \"Schematics\" tab",
                "What the tab is for.",
                TourWording.TakeAgainFoot,
                [
                    new TourCardStep("First", "One.", null, null, () => said.Add("enter 1"), () => said.Add("leave 1")),
                    new TourCardStep("Second", "Two.", null, null, () => said.Add("enter 2"), () => said.Add("leave 2")),
                ]));

            Assert.Empty(said);

            host.Overlay.Next();
            host.Overlay.Next();
            host.Overlay.Back();
            host.Overlay.Skip();

            Assert.Equal(["enter 1", "leave 1", "enter 2", "leave 2", "enter 1", "leave 1"], said);
        });
    }

    [Fact]
    public void The_last_step_says_Done_and_ends_the_tour_as_taken()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());

            host.Overlay.Next();
            host.Overlay.Next();
            Settle();

            Assert.Equal(1, host.Overlay.StepShown);
            Assert.Null(host.Overlay.HoleShown);
            Assert.Equal("Done", ButtonOf(host.Overlay, "NextButton").Content);
            Assert.True(ButtonOf(host.Overlay, "BackButton").IsEnabled);

            host.Overlay.Back();
            Assert.Equal(0, host.Overlay.StepShown);

            host.Overlay.Next();
            host.Overlay.Next();

            Assert.False(host.Overlay.IsShowing);
            Assert.False(host.Overlay.IsVisible);
            Assert.Equal([true], host.Endings);
        });
    }

    [Fact]
    public void Skip_tour_ends_it_as_skipped_from_any_card()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());
            host.Overlay.Skip();

            host.Show(host.Tour());
            host.Overlay.Next();
            host.Overlay.Skip();

            Assert.False(host.Overlay.IsShowing);
            Assert.Equal([false, false], host.Endings);
        });
    }

    // The card's own buttons, pressed for real.
    [Fact]
    public void The_cards_buttons_move_through_the_tour()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());

            Button next = ButtonOf(host.Overlay, "NextButton");
            host.Click(next.TranslatePoint(new Point(next.Bounds.Width / 2, next.Bounds.Height / 2), host.Window)!.Value);
            Settle();

            Assert.Equal(0, host.Overlay.StepShown);

            Button skip = ButtonOf(host.Overlay, "SkipButton");
            host.Click(skip.TranslatePoint(new Point(skip.Bounds.Width / 2, skip.Bounds.Height / 2), host.Window)!.Value);

            Assert.Equal([false], host.Endings);
        });
    }

    [Fact]
    public void The_keys_are_the_tours()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());

            host.Key(Key.Right, PhysicalKey.ArrowRight);
            Assert.Equal(0, host.Overlay.StepShown);

            host.Key(Key.Enter, PhysicalKey.Enter);
            Assert.Equal(1, host.Overlay.StepShown);

            host.Key(Key.Left, PhysicalKey.ArrowLeft);
            Assert.Equal(0, host.Overlay.StepShown);

            // On the first step "Back" is greyed out, and its keys do nothing either - the first card is not shown again
            // (code review, 2026-11-12).
            Assert.False(ButtonOf(host.Overlay, "BackButton").IsEnabled);
            host.Key(Key.Left, PhysicalKey.ArrowLeft);
            host.Key(Key.Back, PhysicalKey.Backspace);
            host.Key(Key.PageUp, PhysicalKey.PageUp);
            Assert.Equal(0, host.Overlay.StepShown);

            host.Key(Key.Escape, PhysicalKey.Escape);
            Assert.False(host.Overlay.IsShowing);
            Assert.Equal([false], host.Endings);
        });
    }

    // ###########################################################################################
    // The focus is the tour's while it shows, on its own button - a handler under it that takes handled keys too saw the
    // tour's keys (code review, 2026-11-12: Down moved the caret in the "Feedback" tab's text) - and goes back to what had
    // it as the tour ends.
    // ###########################################################################################
    [Fact]
    public void The_focus_is_the_tours_while_it_shows_and_goes_back_as_it_ends()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Window.Show();
            Settle();
            host.TextBox.Focus();
            Assert.True(host.TextBox.IsFocused);

            host.Overlay.Show(host.Tour());
            Settle();

            Assert.False(host.TextBox.IsFocused);
            Assert.True(ButtonOf(host.Overlay, "NextButton").IsFocused);

            host.Overlay.Next();
            host.Overlay.Skip();
            Settle();

            Assert.True(host.TextBox.IsFocused);
        });
    }

    // ###########################################################################################
    // A tour ended by a key keeps the keys until the input queued with it has been taken (code review, 2026-11-12): on
    // Linux the space typed with the Space that pressed "Done" went into the text box under the tour. Once taken, typing
    // reaches the text box again.
    // ###########################################################################################
    [Fact]
    public void The_character_typed_with_the_key_that_ends_a_tour_reaches_nothing()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Window.Show();
            Settle();
            host.TextBox.Focus();

            host.Overlay.Show(host.Tour());
            host.Overlay.Next();
            host.Overlay.Next();
            Settle();
            Assert.Equal("Done", ButtonOf(host.Overlay, "NextButton").Content);

            // Raised as the platform hands them over - the key, then its character, no dispatcher job in between. The
            // headless KeyPress runs every job queued before it returns, so it cannot show the order.
            host.TextBox.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Space, Source = host.TextBox });
            host.TextBox.RaiseEvent(new TextInputEventArgs { RoutedEvent = InputElement.TextInputEvent, Text = " ", Source = host.TextBox });

            Assert.False(host.Overlay.IsShowing);
            Assert.Equal([true], host.Endings);
            Assert.True(string.IsNullOrEmpty(host.TextBox.Text));

            Settle();
            host.TextBox.RaiseEvent(new TextInputEventArgs { RoutedEvent = InputElement.TextInputEvent, Text = "a", Source = host.TextBox });

            Assert.Equal("a", host.TextBox.Text);
        });
    }

    // *** A VIEW-DEMO ONLY (owner decision, 2026-11-11). *** The button stands in the framed hole of the first step: a
    // click on it reaches nothing - nor does typing reach the text box that had the keyboard. Once the tour ends, both
    // work again.
    [Fact]
    public void Nothing_under_the_tour_takes_a_click_or_a_key_the_framed_part_included()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Window.Show();
            Settle();
            host.TextBox.Focus();

            host.Overlay.Show(host.Tour());
            host.Overlay.Next();
            Settle();

            Assert.True(host.Overlay.HoleShown!.Value.Contains(host.ButtonRect));

            host.Click(host.ButtonRect.Center);
            host.Window.KeyTextInput("abc");
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(0, host.Clicks);
            Assert.True(string.IsNullOrEmpty(host.TextBox.Text));
            Assert.True(host.Overlay.IsShowing);

            host.Overlay.Skip();
            Settle();
            host.Click(host.ButtonRect.Center);

            Assert.Equal(1, host.Clicks);
        });
    }

    // The frame follows what it frames - a window resized, or the control moving as its panel fills.
    [Fact]
    public void The_frame_follows_what_it_frames()
    {
        UiTest.Run(() =>
        {
            var host = new Host();
            host.Show(host.Tour());
            host.Overlay.Next();
            Settle();

            Canvas.SetLeft(host.Button, 300);
            Settle();

            Assert.True(host.Overlay.HoleShown!.Value.Contains(host.ButtonRect), $"the hole {host.Overlay.HoleShown} stayed where the button was");
        });
    }
}
