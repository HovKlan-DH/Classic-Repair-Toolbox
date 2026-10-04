using System.Net;
using System.Text.Json;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > "API usage" (owner request, 2026-10-04), with the server answering from
// AnsweringHttpHandler - no network. What matters: it asks for the days chosen (90 at first, and
// again when another is chosen), and shows the launches and then the routes in ApiUsageDisplay's
// groups and order.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ApiUsageViewTests
{
    private static readonly ReviewSession Session =
        new("token", DateTimeOffset.UtcNow.AddDays(30), 1, "admin@example.com", "Admin");

    private static readonly DateTimeOffset Last = new(2026, 10, 4, 10, 0, 0, TimeSpan.Zero);

    private static (ApiUsageView View, List<string> Asked) Open()
    {
        var asked = new List<string>();

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(request =>
        {
            asked.Add(request.RequestUri!.PathAndQuery);

            if (request.RequestUri.AbsolutePath != "/api/admin/api-usage")
                return new HttpResponseMessage(HttpStatusCode.NotFound);

            var answer = new ApiUsageAnswer(
                90,
                [
                    new ApiUsageRoute("POST", "/api/usage/check-in", "Forever", true, 900, ApiUsageViewTests.Last,
                        [new ApiUsageVersion("2.5.0", 900, ApiUsageViewTests.Last)]),
                    new ApiUsageRoute("POST", "/api/review/systems/edit", "Maintainer", false, 0, null, []),
                ],
                [new ApiUsageInstallations("3.0.0", 42, 305)]);

            return AnsweringHttpHandler.Json(JsonSerializer.Serialize(answer, ReviewApiContract.WireSettings));
        })));

        var view = new ApiUsageView();
        view.Initialize(client, ApiUsageViewTests.Session);

        return (view, asked);
    }

    [Fact]
    public async Task It_shows_the_launches_then_each_group_of_routes()
    {
        await UiTest.RunAsync(async () =>
        {
            (ApiUsageView view, List<string> asked) = ApiUsageViewTests.Open();

            await view.LoadAsync();

            Assert.Equal(["/api/admin/api-usage?days=90"], asked);

            IReadOnlyList<string> lines = view.LinesForTests();

            Assert.Equal(
                [
                    ApiUsageDisplay.InstallationsHeading,
                    "3.0.0: [42] installations, [305] launches",
                    ApiUsageDisplay.MaintainerHeading,
                    "POST /api/review/systems/edit",
                    "No calls in the last [90] days",
                    ApiUsageDisplay.ForeverHeading,
                    "POST /api/usage/check-in",
                    $"[900] calls, the last on {SubmissionReceiptPresenter.FormatDate(ApiUsageViewTests.Last)}",
                    $"2.5.0: [900] calls, the last on {SubmissionReceiptPresenter.FormatDate(ApiUsageViewTests.Last)}",
                ],
                lines);
            Assert.Null(view.MessageForTests);
        });
    }

    // Choosing another window asks again, for those days.
    [Fact]
    public async Task Choosing_other_days_asks_for_them()
    {
        await UiTest.RunAsync(async () =>
        {
            (ApiUsageView view, List<string> asked) = ApiUsageViewTests.Open();

            Assert.Equal(ApiUsageDisplay.DefaultDays, view.DaysChosen);

            view.DaysBoxForTests.SelectedIndex = 0;

            // The box's SelectionChanged runs the read; give it its turn.
            for (int turn = 0; turn < 20 && asked.Count == 0; turn++)
                await Task.Delay(10);

            Assert.Equal(30, view.DaysChosen);
            Assert.Equal(["/api/admin/api-usage?days=30"], asked);
        });
    }
}
