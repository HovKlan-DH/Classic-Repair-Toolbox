using System.Net;
using System.Text;
using System.Text.Json;
using Avalonia.Controls;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > "Delete a system" (owner request, 2026-10-03: "list all systems and have a delete button
// on each (with a confirmation box)"), with the server answering from AnsweringHttpHandler - no
// network - and the confirmation answered through ConfirmOverrideForTests, since ShowDialog blocks
// headlessly (the window itself is DeleteSystemWindowTests').
//
// What matters: every system has its button; a system the server will not let go says why and is
// never confirmed or sent; a confirmed delete sends back EXACTLY the plan that was shown (its
// fingerprint) with the reason; a cancelled one sends nothing.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemDeletionViewTests
{
    private const string Deleted = "Commodore/C64/999999";
    private const string Kept = "Commodore/C64/250407";

    private static readonly ReviewSession Session =
        new("token", DateTimeOffset.UtcNow.AddDays(30), 1, "admin@example.com", "Admin");

    private static SystemOverviewEntry System(string id, bool inBeta = true)
    {
        string[] parts = id.Split('/');
        return new SystemOverviewEntry(id, parts[0], parts[1], parts[2], inBeta, inBeta, false, true, null, null, null, 1);
    }

    private static string Json(object answer) => JsonSerializer.Serialize(answer, answer.GetType(), ReviewApiContract.WireSettings);

    private static SystemDeletePlanAnswer Plan(string? blockedBecause = null) =>
        new(SystemDeletionViewTests.Deleted, "Commodore", "C64", "999999", new string('f', 64),
            12, 11, true, true, true, 5, 2, 0,
            [new SystemDeleteOpenSubmission(14, "pending", "anna@example.com", "Corrected U8.", DateTimeOffset.UtcNow)],
            blockedBecause);

    private sealed record Server(SystemDeletionView View, List<string> Paths, List<string> DeleteBodies, List<SystemDeletePlanAnswer> Confirmed);

    // ###########################################################################################
    // The view against a server that lists both systems until one is deleted, plans with `plan`,
    // and deletes; confirmation answered with `confirm`.
    // ###########################################################################################
    private static Server Open(SystemDeletePlanAnswer plan, (bool Confirmed, string Reason) confirm)
    {
        var paths = new List<string>();
        var bodies = new List<string>();
        var confirmed = new List<SystemDeletePlanAnswer>();
        bool gone = false;

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(request =>
        {
            string path = request.RequestUri!.AbsolutePath;
            paths.Add(path);

            switch (path)
            {
                case "/api/review/systems":
                    IReadOnlyList<SystemOverviewEntry> systems = gone
                        ? [SystemDeletionViewTests.System(SystemDeletionViewTests.Kept)]
                        : [SystemDeletionViewTests.System(SystemDeletionViewTests.Kept), SystemDeletionViewTests.System(SystemDeletionViewTests.Deleted, inBeta: false)];
                    return AnsweringHttpHandler.Json(SystemDeletionViewTests.Json(new SystemOverviewAnswer(systems)));

                case "/api/admin/systems/delete/plan":
                    return AnsweringHttpHandler.Json(SystemDeletionViewTests.Json(plan));

                case "/api/admin/systems/delete":
                    bodies.Add(request.Content!.ReadAsStringAsync().Result);
                    gone = true;
                    return AnsweringHttpHandler.Json(SystemDeletionViewTests.Json(
                        new SystemDeleteAnswer(SystemDeletionViewTests.Deleted, 12, 11, 5, 1)));

                default:
                    return new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent("{\"error\":\"no such route\"}", Encoding.UTF8, "application/json")
                    };
            }
        })));

        var view = new SystemDeletionView();
        view.Initialize(client, SystemDeletionViewTests.Session);
        view.ConfirmOverrideForTests = shown =>
        {
            confirmed.Add(shown);
            return Task.FromResult(confirm);
        };

        return new Server(view, paths, bodies, confirmed);
    }

    // Every system, by name, each with its own red Delete button - none labelled with "...".
    [Fact]
    public async Task Every_system_is_listed_with_a_Delete_button_of_its_own()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = SystemDeletionViewTests.Open(SystemDeletionViewTests.Plan(), (false, string.Empty));

            await server.View.LoadAsync();

            Assert.Equal(["Commodore / C64 / 250407", "Commodore / C64 / 999999"], server.View.NamesForTests());

            IReadOnlyList<Button> buttons = server.View.DeleteButtons();
            Assert.Equal(2, buttons.Count);
            Assert.All(buttons, button => Assert.Equal(SystemDeletionWording.DeleteButton, button.Content));
            Assert.Equal(
                [SystemDeletionViewTests.Kept, SystemDeletionViewTests.Deleted],
                buttons.Select(button => ((SystemOverviewEntry)button.Tag!).SystemId));
        });
    }

    // ###########################################################################################
    // *** CONFIRMED: THE PLAN THAT WAS SHOWN IS THE ONE SENT BACK. *** Its fingerprint and the
    // reason reach the server; afterwards the list is read again, without it, and what went is said.
    // ###########################################################################################
    [Fact]
    public async Task A_confirmed_delete_sends_the_shown_plans_fingerprint_and_the_reason()
    {
        await UiTest.RunAsync(async () =>
        {
            SystemDeletePlanAnswer plan = SystemDeletionViewTests.Plan();
            Server server = SystemDeletionViewTests.Open(plan, (true, "It was a test system."));
            int afterDelete = 0;
            server.View.AfterDelete = () => { afterDelete++; return Task.CompletedTask; };

            await server.View.LoadAsync();
            await server.View.DeleteAsync(SystemDeletionViewTests.System(SystemDeletionViewTests.Deleted));

            Assert.Equal(SystemDeletionViewTests.Deleted, Assert.Single(server.Confirmed).SystemId);

            SystemDeleteRequest sent = JsonSerializer.Deserialize<SystemDeleteRequest>(
                Assert.Single(server.DeleteBodies), ReviewApiContract.WireSettings)!;

            Assert.Equal(SystemDeletionViewTests.Deleted, sent.SystemId);
            Assert.Equal(plan.Fingerprint, sent.Fingerprint);
            Assert.Equal("It was a test system.", sent.Reason);

            Assert.Equal(["Commodore / C64 / 250407"], server.View.NamesForTests());
            Assert.Equal(SystemDeletionWording.Done(new SystemDeleteAnswer(SystemDeletionViewTests.Deleted, 12, 11, 5, 1)), server.View.MessageForTests);
            Assert.Equal(1, afterDelete);
        });
    }

    // Cancelled: nothing is sent, the list stays as it was.
    [Fact]
    public async Task A_cancelled_confirmation_deletes_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = SystemDeletionViewTests.Open(SystemDeletionViewTests.Plan(), (false, string.Empty));

            await server.View.LoadAsync();
            await server.View.DeleteAsync(SystemDeletionViewTests.System(SystemDeletionViewTests.Deleted));

            Assert.Single(server.Confirmed);
            Assert.DoesNotContain("/api/admin/systems/delete", server.Paths);
            Assert.Equal(2, server.View.NamesForTests().Count);
            Assert.All(server.View.DeleteButtons(), button => Assert.True(button.IsEnabled));
        });
    }

    // ###########################################################################################
    // *** A SYSTEM THAT CANNOT GO SAYS WHY, AND NOTHING IS CONFIRMED OR SENT. *** The server's own
    // sentence (an older main Excel data file lists it, another board uses one of its files) - not
    // a confirmation box that could only be refused afterwards.
    // ###########################################################################################
    [Fact]
    public async Task A_blocked_system_shows_the_servers_reason_and_opens_no_confirmation()
    {
        await UiTest.RunAsync(async () =>
        {
            const string why = "[Classic-Repair-Toolbox.xlsx] in the BETA data lists this system.";
            Server server = SystemDeletionViewTests.Open(SystemDeletionViewTests.Plan(blockedBecause: why), (true, "Test."));

            await server.View.LoadAsync();
            await server.View.DeleteAsync(SystemDeletionViewTests.System(SystemDeletionViewTests.Deleted));

            Assert.Empty(server.Confirmed);
            Assert.DoesNotContain("/api/admin/systems/delete", server.Paths);
            Assert.Equal(why, server.View.MessageForTests);
        });
    }

    // Signed out: the previous account's list and message go.
    [Fact]
    public async Task Clearing_leaves_nothing_of_the_list_or_the_message()
    {
        await UiTest.RunAsync(async () =>
        {
            Server server = SystemDeletionViewTests.Open(SystemDeletionViewTests.Plan(blockedBecause: "No."), (false, string.Empty));

            await server.View.LoadAsync();
            await server.View.DeleteAsync(SystemDeletionViewTests.System(SystemDeletionViewTests.Deleted));

            server.View.Clear();

            Assert.Empty(server.View.NamesForTests());
            Assert.Null(server.View.MessageForTests);
        });
    }
}
