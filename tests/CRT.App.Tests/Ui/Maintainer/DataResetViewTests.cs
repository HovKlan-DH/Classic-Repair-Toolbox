using System.Net;
using System.Text;
using System.Text.Json;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > "Reset contribution data" (owner request, 2026-10-04: a clean start at go-live), with the
// server answering from AnsweringHttpHandler - no network.
//
// What matters: the counts are shown; the button stays off until RESET is typed, and stays off
// whatever is typed while the server's switch is off; a reset sends back EXACTLY the counts shown
// (their fingerprint); and every outcome reads the counts again, so the screen shows what is left.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DataResetViewTests
{
    private static readonly ReviewSession Session =
        new("token", DateTimeOffset.UtcNow.AddDays(30), 1, "admin@example.com", "Admin");

    private static readonly string Shown = new('f', 64);

    private static string Json(object answer) => JsonSerializer.Serialize(answer, answer.GetType(), ReviewApiContract.WireSettings);

    private static DataResetPlanAnswer Plan(bool enabled, bool empty = false) =>
        empty
            ? new DataResetPlanAnswer(enabled, new string('0', 64), 0, 0, 1, 0, 0, 0, 1, 0, 0)
            : new DataResetPlanAnswer(enabled, DataResetViewTests.Shown, 14, 6, 1, 4, 2, 7, 310, 1200, 80,
                enabled ? null : "Resetting is switched off on this server.");

    private sealed record Server(DataResetView View, List<string> Calls, List<string> ResetBodies, List<int> AfterResets);

    // ###########################################################################################
    // The view against a server answering the counts with `enabled`, and the reset with `resetStatus`
    // (200 with what went, or a refusal). After a reset that went, the counts read zero.
    // ###########################################################################################
    private static Server Open(bool enabled, HttpStatusCode resetStatus = HttpStatusCode.OK)
    {
        var calls = new List<string>();
        var bodies = new List<string>();
        var afterResets = new List<int>();
        bool reset = false;

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(request =>
        {
            calls.Add($"{request.Method} {request.RequestUri!.AbsolutePath}");

            if (request.RequestUri.AbsolutePath != "/api/admin/reset")
                return new HttpResponseMessage(HttpStatusCode.NotFound);

            if (request.Method == HttpMethod.Get)
                return AnsweringHttpHandler.Json(DataResetViewTests.Json(DataResetViewTests.Plan(enabled, empty: reset)));

            bodies.Add(request.Content!.ReadAsStringAsync().Result);

            if (resetStatus != HttpStatusCode.OK)
            {
                return new HttpResponseMessage(resetStatus)
                {
                    Content = new StringContent("{\"error\":\"refused by the server\"}", Encoding.UTF8, "application/json")
                };
            }

            reset = true;
            return AnsweringHttpHandler.Json(DataResetViewTests.Json(new DataResetAnswer(14, 6, 4, 2, 7, 310, 1200, 80, 23)));
        })));

        var view = new DataResetView();
        view.Initialize(client, DataResetViewTests.Session);
        view.AfterReset = () =>
        {
            afterResets.Add(1);
            return Task.CompletedTask;
        };

        return new Server(view, calls, bodies, afterResets);
    }

    [Fact]
    public async Task The_counts_are_shown_and_the_button_waits_for_the_word()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = DataResetViewTests.Open(enabled: true);

            await server.View.LoadAsync();

            Assert.Equal(9, server.View.CountLinesForTests().Count);
            Assert.StartsWith("[14] submissions", server.View.CountLinesForTests()[0], StringComparison.Ordinal);
            Assert.Null(server.View.SwitchedOffForTests);
            Assert.Equal(DataResetWording.ResetButton, server.View.ResetButtonForTests.Content);

            Assert.True(server.View.ConfirmBoxForTests.IsEnabled);
            Assert.False(server.View.ResetButtonForTests.IsEnabled);

            server.View.ConfirmBoxForTests.Text = "reset";
            Assert.False(server.View.ResetButtonForTests.IsEnabled);

            server.View.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;
            Assert.True(server.View.ResetButtonForTests.IsEnabled);
        });
    }

    // ###########################################################################################
    // *** THE SERVER'S SWITCH. *** Off, the counts still show - with the server's words for how to
    // switch it on - and nothing can be typed or pressed.
    // ###########################################################################################
    [Fact]
    public async Task While_the_server_has_it_switched_off_the_counts_show_and_nothing_can_be_pressed()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = DataResetViewTests.Open(enabled: false);

            await server.View.LoadAsync();

            Assert.Equal(9, server.View.CountLinesForTests().Count);
            Assert.Equal("Resetting is switched off on this server.", server.View.SwitchedOffForTests);
            Assert.False(server.View.ConfirmBoxForTests.IsEnabled);

            server.View.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;
            Assert.False(server.View.ResetButtonForTests.IsEnabled);

            await server.View.ResetAsync();
            Assert.Empty(server.ResetBodies);
        });
    }

    [Fact]
    public async Task A_reset_sends_the_counts_shown_says_what_went_and_reads_the_counts_again()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = DataResetViewTests.Open(enabled: true);

            await server.View.LoadAsync();
            server.View.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;

            await server.View.ResetAsync();

            string body = Assert.Single(server.ResetBodies);
            Assert.Equal(DataResetViewTests.Shown, JsonSerializer.Deserialize<DataResetRequest>(body, ReviewApiContract.WireSettings)!.Fingerprint);

            Assert.StartsWith("The contribution data was reset: [14] submissions", server.View.DoneForTests, StringComparison.Ordinal);
            Assert.Null(server.View.MessageForTests);
            Assert.Single(server.AfterResets);

            // Read again: what is left is shown, and the word must be typed again for another.
            Assert.Equal(["GET /api/admin/reset", "POST /api/admin/reset", "GET /api/admin/reset"], server.Calls);
            Assert.StartsWith("[0] submissions", server.View.CountLinesForTests()[0], StringComparison.Ordinal);
            Assert.Equal(string.Empty, server.View.ConfirmBoxForTests.Text);
            Assert.False(server.View.ResetButtonForTests.IsEnabled);
        });
    }

    // Something arrived since the counts were shown: nothing went, the counts are read again, and
    // the screen says to look at them before confirming again.
    [Fact]
    public async Task A_reset_refused_as_changed_reads_the_counts_again_and_says_so()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = DataResetViewTests.Open(enabled: true, resetStatus: HttpStatusCode.Conflict);

            await server.View.LoadAsync();
            server.View.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;

            await server.View.ResetAsync();

            Assert.Equal(DataResetWording.Changed, server.View.MessageForTests);
            Assert.Null(server.View.DoneForTests);
            Assert.Empty(server.AfterResets);
            Assert.Equal(3, server.Calls.Count);
        });
    }

    // ###########################################################################################
    // *** A TIMED-OUT RESET IS JUDGED FROM THE COUNTS READ AGAIN (code review, 2026-11-02). *** It said
    // "if they are all zero, it did" - which can never be: the reset's own history entry is one, and
    // API usage and board views grow again by themselves. And the queue and the boards stayed as they
    // were until the next minute check, as AfterReset was skipped. The submissions, the accounts and
    // the board records at zero say it: `landed` is whether the reset the server never answered went.
    // ###########################################################################################
    [Fact]
    public async Task A_reset_with_no_answer_that_landed_is_said_and_the_tab_reads_its_lists_again()
    {
        await UiTest.RunAsync(async () =>
        {
            (DataResetView view, List<int> afterResets) = await DataResetViewTests.TimedOutAsync(landed: true);

            Assert.Equal(
                "The server did not answer within 2 minutes, but it did finish: the contribution data is reset.",
                view.MessageForTests);
            Assert.Single(afterResets);
            Assert.StartsWith("[0] submissions", view.CountLinesForTests()[0], StringComparison.Ordinal);
        });
    }

    [Fact]
    public async Task A_reset_with_no_answer_that_did_not_land_says_what_is_still_there()
    {
        await UiTest.RunAsync(async () =>
        {
            (DataResetView view, List<int> afterResets) = await DataResetViewTests.TimedOutAsync(landed: false);

            Assert.Equal(
                "The server did not answer within 2 minutes, and the contribution data is still there. " +
                "It may still be working on it - wait a minute and look again before trying again.",
                view.MessageForTests);
            Assert.Empty(afterResets);
        });
    }

    // The counts could not be read either: nothing is known, and it says so.
    [Fact]
    public async Task A_reset_with_no_answer_and_no_counts_says_it_could_not_be_checked()
    {
        await UiTest.RunAsync(async () =>
        {
            (DataResetView view, List<int> afterResets) = await DataResetViewTests.TimedOutAsync(landed: false, countsRead: false);

            Assert.Equal(
                "The server did not answer within 2 minutes, and whether it finished could not be checked. " +
                "Wait a minute and look again before trying again.",
                view.MessageForTests);
            Assert.Empty(afterResets);
        });
    }

    // The reset's request never answers and its wait runs out at once (the test's clock, not two real
    // minutes), the view hosted with an overlay as CRT's window hosts it.
    private static async Task<(DataResetView View, List<int> AfterResets)> TimedOutAsync(bool landed, bool countsRead = true)
    {
        var afterResets = new List<int>();
        bool posted = false;

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler((request, token) =>
        {
            if (request.Method != HttpMethod.Get)
            {
                posted = true;
                return AnsweringHttpHandler.NeverAsync(token);
            }

            if (posted && !countsRead)
                return Task.FromResult(AnsweringHttpHandler.Refused());

            return Task.FromResult(AnsweringHttpHandler.Json(
                DataResetViewTests.Json(DataResetViewTests.Plan(enabled: true, empty: posted && landed))));
        })));

        var view = new DataResetView();
        var overlay = new BusyOverlay();
        _ = new Avalonia.Controls.Grid { Children = { view, overlay } };

        view.Initialize(client, DataResetViewTests.Session);
        view.AfterReset = () =>
        {
            afterResets.Add(1);
            return Task.CompletedTask;
        };

        await view.LoadAsync();

        int limits = 0;
        overlay.LimitOverrideForTests = token => ++limits == 1 ? Task.CompletedTask : Task.Delay(Timeout.Infinite, token);

        view.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;
        await view.ResetAsync();

        return (view, afterResets);
    }

    // ###########################################################################################
    // *** THE WORD TYPED DOES NOT WAIT FOR THE NEXT ONE (code review, 2026-11-02). *** Signing out and
    // leaving the entry kept RESET in the box, so the next administrator on this computer - or the
    // same one coming back - got new counts with the button already on and nothing typed.
    // ###########################################################################################
    [Fact]
    public async Task Signing_out_empties_the_confirmation_word()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = DataResetViewTests.Open(enabled: true);

            await server.View.LoadAsync();
            server.View.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;

            server.View.Clear();

            Assert.True(string.IsNullOrEmpty(server.View.ConfirmBoxForTests.Text));

            await server.View.LoadAsync();

            Assert.False(server.View.ResetButtonForTests.IsEnabled);
        });
    }

    [Fact]
    public async Task Coming_back_to_the_entry_shows_new_counts_with_the_word_to_type_again()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = DataResetViewTests.Open(enabled: true);

            await server.View.LoadAsync();
            server.View.ConfirmBoxForTests.Text = DataResetWording.ConfirmWord;
            Assert.True(server.View.ResetButtonForTests.IsEnabled);

            await server.View.LoadAsync();

            Assert.True(string.IsNullOrEmpty(server.View.ConfirmBoxForTests.Text));
            Assert.False(server.View.ResetButtonForTests.IsEnabled);
        });
    }
}
