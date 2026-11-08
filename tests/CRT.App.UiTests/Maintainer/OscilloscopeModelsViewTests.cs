using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > "Oscilloscope models", OscilloscopeModelsView (CRT 4.0 step 5): the models sent in from
// CRT, each a card the administrator publishes or turns down. Drawn without a server - the list and
// the decisions answered by the test, which also sees what was sent.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class OscilloscopeModelsViewTests
{
    private static OscilloscopeModelQueueEntry Entry(long id = 41) => new(
        id,
        new DateTimeOffset(2026, 10, 29, 12, 0, 0, TimeSpan.Zero),
        "Voltcraft",
        "DOV702",
        new Dictionary<string, string> { ["identify"] = "*IDN?", ["dumpImage"] = ":DISP:DATA?" },
        "Anna",
        "anna@example.com",
        "Firmware 1.2 - everything but Single works",
        true,
        "CRT 4.0.0",
        null);

    private static Task<ReviewApiResult<OscilloscopeModelQueueAnswer>> List(params OscilloscopeModelQueueEntry[] entries) =>
        Task.FromResult(ReviewApiResult<OscilloscopeModelQueueAnswer>.Ok(new OscilloscopeModelQueueAnswer(entries)));

    [Fact]
    public async Task With_nothing_waiting_it_says_so()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new OscilloscopeModelsView { ReadOverrideForTests = () => List() };

            await view.LoadAsync();

            Assert.Equal(OscilloscopeModelsDisplay.Empty, view.NoneWaitingLineForTests);
        });
    }

    // Everything the administrator decides on, and the two buttons in their colours.
    [Fact]
    public async Task A_model_waiting_is_a_card_with_everything_to_decide_on()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new OscilloscopeModelsView { ReadOverrideForTests = () => List(Entry()) };

            await view.LoadAsync();

            IReadOnlyList<string> card = Assert.Single(view.CardsForTests());
            Assert.Equal("Voltcraft DOV702", card[0]);
            Assert.Equal(OscilloscopeModelsDisplay.NewModel, card[1]);
            Assert.StartsWith("Sent by Anna (anna@example.com)", card[2]);
            Assert.Equal(OscilloscopeModelsDisplay.SentTypedIn, card[3]);
            Assert.Equal(OscilloscopeModelsDisplay.SuitePassed, card[4]);
            Assert.Equal("Firmware 1.2 - everything but Single works", card[5]);
            Assert.Contains("Identify | *IDN?", card[6]);
            Assert.Contains("Screen picture | :DISP:DATA?", card[6]);
            Assert.Equal("[reason]", card[7]);
            Assert.Equal("[Publish:Accept] [Reject:Danger]", card[8]);
        });
    }

    // ###########################################################################################
    // *** NOTHING A STRANGER CAN TYPE IS DRAWN AS GOOD NEWS (code review, 2026-11-02 - Security). *** A
    // card reads who sent it from an account, or says the name and address were typed in, in the
    // warning color. The suite's line was green, on the sender's own word alone, and is plain now.
    // ###########################################################################################
    [Theory]
    [InlineData(false, OscilloscopeModelsDisplay.SentTypedIn, true)]
    [InlineData(true, OscilloscopeModelsDisplay.SentWithAccount, false)]
    public async Task A_card_says_who_the_sender_is_and_draws_no_word_of_the_senders_as_good_news(bool signedIn, string line, bool warned)
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new OscilloscopeModelsView { ReadOverrideForTests = () => List(Entry() with { SignedIn = signedIn }) };

            await view.LoadAsync();

            IReadOnlyList<string> card = Assert.Single(view.CardsForTests());
            Assert.Equal(line, card[3]);

            // Which of the card's two looks each line has: dimmed, or at full strength in the warning color.
            var body = (Avalonia.Controls.StackPanel)((Avalonia.Controls.Border)view.GetControl<Avalonia.Controls.StackPanel>("ModelsPanel").Children[0]).Child!;

            Assert.Equal(warned ? 1 : 0.8, body.Children[3].Opacity);
            Assert.Equal(0.8, body.Children[4].Opacity);
        });
    }

    // ###########################################################################################
    // The list follows the count the queue answered, quietly: nothing asked while it agrees, and a
    // quiet read that fails says nothing over what is on show.
    // ###########################################################################################
    [Fact]
    public async Task Following_the_count_reads_only_when_it_differs_and_fails_without_a_word()
    {
        await UiTest.RunAsync(async () =>
        {
            int reads = 0;
            bool fails = false;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () =>
                {
                    reads++;
                    return fails
                        ? Task.FromResult(ReviewApiResult<OscilloscopeModelQueueAnswer>.Failed(ReviewApiFailure.Unreachable, "Could not reach the server."))
                        : List(Entry());
                },
            };

            await view.LoadAsync();
            Assert.Equal(1, reads);

            await view.FollowWaitingCountAsync(1, onScreen: false);
            Assert.Equal(1, reads);

            await view.FollowWaitingCountAsync(2, onScreen: false);
            Assert.Equal(2, reads);

            fails = true;
            await view.FollowWaitingCountAsync(2, onScreen: false);

            Assert.Equal(3, reads);
            Assert.Null(view.MessageForTests);
            Assert.Single(view.CardsForTests());
        });
    }

    // ###########################################################################################
    // *** ON SCREEN THE LIST IS READ AT EVERY FULL CHECK, NOT ONLY WHEN THE COUNT MOVES (code review,
    // 2026-11-02). *** A contributor re-sent a corrected model of the same name: the server marks the
    // pending copy replaced and adds the new one, so the count stayed 1. The card kept the old commands
    // and the old id, and Publish on it answered "decided meanwhile" until Refresh was pressed. Off screen
    // the count is still all that is asked: nothing shows the list there, and showing it reads it.
    // ###########################################################################################
    [Fact]
    public async Task On_screen_a_model_replaced_by_a_newer_one_shows_though_the_count_stays()
    {
        await UiTest.RunAsync(async () =>
        {
            OscilloscopeModelQueueEntry waiting = Entry();
            int reads = 0;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () =>
                {
                    reads++;
                    return List(waiting);
                },
            };

            await view.LoadAsync();

            waiting = Entry(42) with { Message = "Corrected" };

            await view.FollowWaitingCountAsync(1, onScreen: false);
            Assert.Equal(1, reads);
            Assert.Equal("Firmware 1.2 - everything but Single works", Assert.Single(view.CardsForTests())[5]);

            await view.FollowWaitingCountAsync(1, onScreen: true);
            Assert.Equal(2, reads);
            Assert.Equal("Corrected", Assert.Single(view.CardsForTests())[5]);
        });
    }

    // A read showing nothing new makes no card again: a card made again throws away the press going down
    // on its buttons and moves the focus, and this happens every minute while the screen is open.
    [Fact]
    public async Task On_screen_a_read_with_nothing_new_makes_no_card_again()
    {
        await UiTest.RunAsync(async () =>
        {
            int reads = 0;
            int? counted = null;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () =>
                {
                    reads++;
                    return List(Entry());
                },
            };

            await view.LoadAsync();

            view.ModelsWaitingRead = count => counted = count;

            Avalonia.Controls.Control card = view.GetControl<Avalonia.Controls.StackPanel>("ModelsPanel").Children[0];

            // The queue's answer said 2: what was read is what the badges count, card or no card.
            await view.FollowWaitingCountAsync(2, onScreen: true);

            Assert.Equal(2, reads);
            Assert.Same(card, view.GetControl<Avalonia.Controls.StackPanel>("ModelsPanel").Children[0]);
            Assert.Equal(1, counted);
        });
    }

    // Made again, the card would take the box from under the caret - the words survive that, the
    // typing does not. The next check asks again, once the focus has left it.
    [Fact]
    public async Task Following_the_count_waits_while_a_reason_is_being_typed()
    {
        await UiTest.RunAsync(async () =>
        {
            int reads = 0;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () =>
                {
                    reads++;
                    return List(Entry());
                },
            };

            var window = new Avalonia.Controls.Window { Content = view, Width = 900, Height = 700 };
            window.Show();

            await view.LoadAsync();
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Avalonia.Controls.TextBox reason = Avalonia.VisualTree.VisualExtensions
                .GetVisualDescendants(view.GetControl<Avalonia.Controls.StackPanel>("ModelsPanel"))
                .OfType<Avalonia.Controls.TextBox>()
                .Single();

            reason.Focus();
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            Assert.True(reason.IsKeyboardFocusWithin);

            await view.FollowWaitingCountAsync(2, onScreen: true);
            Assert.Equal(1, reads);

            Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(view)
                .OfType<Avalonia.Controls.Button>()
                .First(button => button.Content is "Publish")
                .Focus();
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            Assert.False(reason.IsKeyboardFocusWithin);

            await view.FollowWaitingCountAsync(2, onScreen: true);
            Assert.Equal(2, reads);

            window.Close();
        });
    }

    // Signed out while it was read: nothing of the answer goes on show (the panel's half of the tab's
    // own rule, TabMaintainerSessionTests).
    [Fact]
    public async Task A_quiet_read_landing_after_signing_out_shows_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            var gate = new TaskCompletionSource<ReviewApiResult<OscilloscopeModelQueueAnswer>>();
            var view = new OscilloscopeModelsView { ReadOverrideForTests = () => gate.Task };
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "a@example.com", "A");

            view.Initialize(null, session);

            Task reading = view.FollowWaitingCountAsync(1, onScreen: false);

            view.Initialize(null, null);
            view.Clear();

            gate.SetResult(ReviewApiResult<OscilloscopeModelQueueAnswer>.Ok(new OscilloscopeModelQueueAnswer([Entry()])));
            await reading;

            Assert.Empty(view.CardsForTests());
        });
    }

    // Publish sends the id with no reason; the server's sentence stays, and the list is read again.
    [Fact]
    public async Task Publish_sends_the_model_says_what_the_server_did_and_reads_the_list_again()
    {
        await UiTest.RunAsync(async () =>
        {
            var decided = new List<(long Id, string? Reason)>();
            int reads = 0;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () => ++reads == 1 ? List(Entry()) : List(),
                DecideOverrideForTests = (id, reason) =>
                {
                    decided.Add((id, reason));
                    return Task.FromResult(ReviewApiResult<OscilloscopeModelDecisionAnswer>.Ok(
                        new OscilloscopeModelDecisionAnswer(true, true, "Voltcraft DOV702 is published.")));
                },
            };

            await view.LoadAsync();
            await view.DecideAsync(41, reason: null);

            Assert.Equal([(41L, (string?)null)], decided);
            Assert.Equal("Voltcraft DOV702 is published.", view.MessageForTests);
            Assert.Equal(OscilloscopeModelsDisplay.Empty, view.NoneWaitingLineForTests);
        });
    }

    // A reject with no reason asks for one, and sends nothing; with one, sends it.
    [Fact]
    public async Task Reject_asks_for_a_reason_first()
    {
        await UiTest.RunAsync(async () =>
        {
            var decided = new List<(long Id, string? Reason)>();

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () => List(Entry()),
                DecideOverrideForTests = (id, reason) =>
                {
                    decided.Add((id, reason));
                    return Task.FromResult(ReviewApiResult<OscilloscopeModelDecisionAnswer>.Ok(
                        new OscilloscopeModelDecisionAnswer(false, null, "Voltcraft DOV702 will not go in.")));
                },
            };

            await view.LoadAsync();
            await view.DecideAsync(41, "   ");

            Assert.Empty(decided);
            Assert.Equal(OscilloscopeModelsDisplay.NoReason, view.MessageForTests);

            await view.DecideAsync(41, "Single does not work - please check it.");

            Assert.Equal([(41L, (string?)"Single does not work - please check it.")], decided);
        });
    }

    // A model a publish put into BETA only says so, and offers only Publish - it cannot be rejected
    // any more (code review, 2026-10-29; the server refuses it too).
    [Fact]
    public async Task A_model_in_BETA_offers_only_Publish_and_says_why()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new OscilloscopeModelsView { ReadOverrideForTests = () => List(Entry() with { InBeta = true }) };

            await view.LoadAsync();

            IReadOnlyList<string> card = Assert.Single(view.CardsForTests());
            Assert.Equal(OscilloscopeModelsDisplay.InBeta, card[1]);
            Assert.DoesNotContain("[reason]", card);
            Assert.Equal("[Publish:Accept]", card[^1]);
        });
    }

    // ###########################################################################################
    // *** A REASON TYPED ON ONE CARD OUTLIVES A DECISION ON ANOTHER (code review, 2026-10-29). *** The
    // list read again after every decision made every card anew, and a long reason half typed on card
    // B was gone once card A was published.
    // ###########################################################################################
    [Fact]
    public async Task A_reason_typed_on_one_card_is_kept_through_a_decision_on_another()
    {
        await UiTest.RunAsync(async () =>
        {
            int reads = 0;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () => ++reads == 1 ? List(Entry(41), Entry(42)) : List(Entry(42)),
                DecideOverrideForTests = (_, _) => Task.FromResult(ReviewApiResult<OscilloscopeModelDecisionAnswer>.Ok(
                    new OscilloscopeModelDecisionAnswer(true, true, "Voltcraft DOV702 is published."))),
            };

            await view.LoadAsync();
            view.TypeReasonForTests(42, "Single does not work on mine either -");
            await view.DecideAsync(41, reason: null);

            Assert.Single(view.CardsForTests());
            Assert.Equal("Single does not work on mine either -", view.ReasonForTests(42));
        });
    }

    // A publish that wrote BETA but not all after it says both.
    [Fact]
    public async Task What_did_not_happen_after_BETA_is_said_with_the_publish()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () => List(),
                DecideOverrideForTests = (_, _) => Task.FromResult(ReviewApiResult<OscilloscopeModelDecisionAnswer>.Ok(
                    new OscilloscopeModelDecisionAnswer(true, false, "Voltcraft DOV702 is published.", "The stable source has no main data file."))),
            };

            await view.DecideAsync(41, reason: null);

            Assert.Equal("Voltcraft DOV702 is published. The stable source has no main data file.", view.MessageForTests);
        });
    }

    // ###########################################################################################
    // *** A DECISION WITH NO ANSWER IS CHECKED, NOT CALLED A FAILURE (code review, 2026-11-02). *** It
    // said "The server did not answer within 2 minutes. Try again in a moment." in red, the sentence
    // for a READ - even when the server had finished, the list read again then showing the card gone
    // with no word that it was published. Pressed again, a model still listed got "decided meanwhile".
    // The list read again now says which it is: gone, it finished; there, it has not (yet).
    // ###########################################################################################
    private static OscilloscopeModelsView TimedOut(params OscilloscopeModelQueueEntry[] afterwards)
    {
        int reads = 0;

        return new OscilloscopeModelsView
        {
            ReadOverrideForTests = () => ++reads == 1 ? List(Entry()) : List(afterwards),
            DecideOverrideForTests = (_, _) => Task.FromResult(
                ReviewApiResult<OscilloscopeModelDecisionAnswer>.Failed(ReviewApiFailure.TimedOut, WaitWording.NoAnswer)),
        };
    }

    [Fact]
    public async Task A_publish_with_no_answer_whose_model_left_the_list_is_said_to_have_finished()
    {
        await UiTest.RunAsync(async () =>
        {
            OscilloscopeModelsView view = TimedOut();

            await view.LoadAsync();
            await view.DecideAsync(41, reason: null);

            Assert.Equal(
                "The server did not answer within 2 minutes, but it did finish: Voltcraft DOV702 has left the list - it is published.",
                view.MessageForTests);
        });
    }

    [Fact]
    public async Task A_reject_with_no_answer_whose_model_left_the_list_is_said_to_have_finished()
    {
        await UiTest.RunAsync(async () =>
        {
            OscilloscopeModelsView view = TimedOut();

            await view.LoadAsync();
            await view.DecideAsync(41, "Single does not work.");

            Assert.Equal(
                "The server did not answer within 2 minutes, but it did finish: Voltcraft DOV702 has left the list - it will not go in, and the contributor is told why.",
                view.MessageForTests);
        });
    }

    [Fact]
    public async Task A_decision_with_no_answer_whose_model_is_still_listed_says_so_and_not_to_press_at_once()
    {
        await UiTest.RunAsync(async () =>
        {
            OscilloscopeModelsView view = TimedOut(Entry());

            await view.LoadAsync();
            await view.DecideAsync(41, reason: null);

            Assert.Equal(
                "The server did not answer within 2 minutes, and Voltcraft DOV702 is still in the list. " +
                "It may still be working on it - wait a minute and look again before trying again.",
                view.MessageForTests);
        });
    }

    [Fact]
    public async Task A_decision_with_no_answer_and_no_list_says_it_could_not_be_checked()
    {
        await UiTest.RunAsync(async () =>
        {
            int reads = 0;

            var view = new OscilloscopeModelsView
            {
                ReadOverrideForTests = () => ++reads == 1
                    ? List(Entry())
                    : Task.FromResult(ReviewApiResult<OscilloscopeModelQueueAnswer>.Failed(ReviewApiFailure.Unreachable, "Could not reach the server.")),
                DecideOverrideForTests = (_, _) => Task.FromResult(
                    ReviewApiResult<OscilloscopeModelDecisionAnswer>.Failed(ReviewApiFailure.TimedOut, WaitWording.NoAnswer)),
            };

            await view.LoadAsync();
            await view.DecideAsync(41, reason: null);

            Assert.Equal(
                "The server did not answer within 2 minutes, and whether it finished could not be checked. " +
                "Wait a minute and look again before trying again.",
                view.MessageForTests);
        });
    }
}
