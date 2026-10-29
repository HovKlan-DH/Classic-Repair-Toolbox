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
            Assert.Equal(OscilloscopeModelsDisplay.SuitePassed, card[3]);
            Assert.Equal("Firmware 1.2 - everything but Single works", card[4]);
            Assert.Contains("Identify | *IDN?", card[5]);
            Assert.Contains("Screen picture | :DISP:DATA?", card[5]);
            Assert.Equal("[reason]", card[6]);
            Assert.Equal("[Publish:Accept] [Reject:Danger]", card[7]);
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
}
