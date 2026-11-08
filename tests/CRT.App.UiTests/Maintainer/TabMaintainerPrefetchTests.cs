using System.Reflection;
using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The entry "Contributor Submissions" opens on, read ahead while the tab is away (owner request,
// 2026-10-02: "When I launched CRT app, then it was idle for some minutes. Then I did go into the
// "Maintainer" tab ... and then I saw that "Wait" screen") - TabMaintainer.Prefetch.cs.
//
// A fake server (AnsweringHttpHandler, no network) records every request and whether CRT's
// "please wait" was up while it was asked. What is pinned: the read happens with the tab away and
// only when the entry changed; opening that entry then asks nothing BEFORE it is on screen, with no
// overlay; and the one quiet re-read after it brings a newer table in.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerPrefetchTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(long id = 42, bool? awaitsYou = true) =>
        new(id, "Commodore/C64/250407", "pending", "Corrected U8.", "c@example.com",
            new DateTimeOffset(2026, 10, 2, 9, 0, 0, TimeSpan.Zero), false, false, awaitsYou);

    private static string TableJson(int version, string friendlyName) =>
        JsonSerializer.Serialize(
            new ReviewTableData(
                version,
                new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA" }] },
                new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = friendlyName }] }),
            JsonSerializerOptions.Web);

    // The server: each request recorded with whether the overlay was up, the table at `tableVersion`.
    private sealed class Server
    {
        public readonly List<(string Path, bool OverlayUp)> Requests = [];
        public int TableVersion = 1;
        public BusyOverlay? Overlay;

        public AnsweringHttpHandler Handler => new(request =>
        {
            string path = request.RequestUri!.AbsolutePath;
            this.Requests.Add((path, this.Overlay is { IsVisible: true, IsBusy: true }));

            if (path.EndsWith("/table", StringComparison.Ordinal))
                return AnsweringHttpHandler.Json(TableJson(this.TableVersion, $"CIA 6526 v{this.TableVersion}"));

            if (path.Contains("/submissions/", StringComparison.Ordinal))
            {
                long id = long.Parse(path[(path.LastIndexOf('/') + 1)..]);

                return AnsweringHttpHandler.Json(
                    $$"""{"canPublish":true,"submission":{"id":{{id}},"boardId":"Commodore/C64/250407","state":"pending"},"findings":[]}""");
            }

            return AnsweringHttpHandler.Refused();
        });
    }

    // The tab signed in against the fake server, with `rows` in its queue - not on screen.
    private static (TabMaintainer Tab, Server Server) SignedIn(params ReviewQueueRow[] rows)
    {
        var tab = new TabMaintainer();
        var server = new Server();
        server.Overlay = MaintainerTabHost.AddOverlay(tab);

        typeof(TabMaintainer).GetField("thisClient", Any)!
            .SetValue(tab, new ReviewApiClient("https://review.invalid", new HttpClient(server.Handler)));
        typeof(TabMaintainer).GetField("thisSession", Any)!
            .SetValue(tab, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

        tab.ApplyQueueResponse(new ReviewQueueResponse(true, rows, false));

        return (tab, server);
    }

    private static int? ShownTableVersion(TabMaintainer tab) =>
        (typeof(TabMaintainer).GetField("thisTable", Any)!.GetValue(tab) as ReviewTableData)?.Version;

    [Fact]
    public async Task With_the_tab_away_the_entry_it_opens_on_is_read_and_held()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, Server server) = SignedIn(Row(42), Row(43));

            await tab.PrefetchEntryToOpenForTests();

            Assert.Equal(Row(42), tab.PrefetchedRowForTests);
            Assert.Equal(2, server.Requests.Count);
            Assert.DoesNotContain(server.Requests, request => request.OverlayUp);

            // Nothing on screen moved: no submission chosen, no table opened.
            Assert.Null(tab.SelectedQueueRowForTests);
            Assert.False(tab.IsTableOpen);
        });
    }

    // The badge's check runs every minute; the server is asked again only when the entry changed.
    [Fact]
    public async Task The_entry_is_read_again_only_when_the_queue_says_it_changed()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, Server server) = SignedIn(Row(42));

            await tab.PrefetchEntryToOpenForTests();
            await tab.PrefetchEntryToOpenForTests();

            Assert.Equal(2, server.Requests.Count);

            // Another approver acted: the same submission no longer waits for this account.
            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(42, awaitsYou: false)], false), background: true);
            await tab.PrefetchEntryToOpenForTests();

            Assert.Equal(4, server.Requests.Count);
            Assert.Equal(Row(42, awaitsYou: false), tab.PrefetchedRowForTests);
        });
    }

    // ###########################################################################################
    // *** THE REPORTED CASE. *** The tab is shown; it opens on the entry read ahead - the table and
    // the decision bar are there at once, and the overlay never went up. The two requests after it
    // are the quiet re-read, also with no overlay.
    // ###########################################################################################
    [Fact]
    public async Task Opening_the_entry_read_ahead_shows_it_with_no_wait()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, Server server) = SignedIn(Row(42));
            await tab.PrefetchEntryToOpenForTests();
            server.Requests.Clear();

            tab.SelectOnEntryForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(42, tab.SelectedQueueRowForTests?.Id);
            Assert.True(tab.IsTableOpen);
            Assert.True(tab.TableEditorForTests.HasTable);
            Assert.True(tab.FindControl<StackPanel>("DecisionPanel")!.IsVisible);

            Assert.DoesNotContain(server.Requests, request => request.OverlayUp);
            Assert.Null(tab.PrefetchedRowForTests);
        });
    }

    // What the queue row cannot show - another maintainer saving the table - is brought in by the
    // re-read: a newer table replaces the one read ahead.
    [Fact]
    public async Task A_table_saved_by_somebody_else_since_the_read_replaces_the_one_shown()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, Server server) = SignedIn(Row(42));
            await tab.PrefetchEntryToOpenForTests();

            server.TableVersion = 2;

            tab.SelectOnEntryForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(2, ShownTableVersion(tab));
        });
    }

    // A queue row that changed after the read is not opened from it: the ordinary open, under the
    // overlay, as before.
    [Fact]
    public async Task An_entry_whose_row_changed_since_the_read_is_opened_the_ordinary_way()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, Server server) = SignedIn(Row(42));
            await tab.PrefetchEntryToOpenForTests();

            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(42, awaitsYou: false)], false), background: true);
            server.Requests.Clear();

            tab.SelectOnEntryForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.IsTableOpen);
            Assert.All(server.Requests, request => Assert.True(request.OverlayUp));
            Assert.NotEmpty(server.Requests);
        });
    }

    // On screen there is nothing to read ahead: the entry is simply opened there.
    [Fact]
    public async Task Nothing_is_read_ahead_while_the_tab_is_on_screen()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            var server = new Server();

            typeof(TabMaintainer).GetField("thisClient", Any)!
                .SetValue(tab, new ReviewApiClient("https://review.invalid", new HttpClient(server.Handler)));
            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(tab, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

            var window = new Window { Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            // The queue arrives with the tab already on screen.
            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(42)], false), background: true);
            server.Requests.Clear();

            await tab.PrefetchEntryToOpenForTests();

            Assert.Empty(server.Requests);
            Assert.Null(tab.PrefetchedRowForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // The tab with Row(42) read ahead and opened, the quiet re-read's detail held until `release` is
    // set - and the approval answered as `approve` does.
    // ###########################################################################################
    private static async Task<(TabMaintainer Tab, TaskCompletionSource Release)> OpenedWithTheReReadHeldAsync(
        Func<Task<HttpResponseMessage>> approve,
        int tableVersionOnReRead = 1)
    {
        var release = new TaskCompletionSource();
        bool hold = false;
        int tableVersion = 1;

        var server = new AnsweringHttpHandler(async (request, _) =>
        {
            string path = request.RequestUri!.AbsolutePath;

            if (path.EndsWith("/approve", StringComparison.Ordinal))
                return await approve();

            if (path.EndsWith("/table", StringComparison.Ordinal))
                return AnsweringHttpHandler.Json(TableJson(tableVersion, "CIA 6526"));

            if (path.EndsWith("/submissions/42", StringComparison.Ordinal))
            {
                if (hold)
                {
                    hold = false;
                    await release.Task;
                }

                return AnsweringHttpHandler.Json(
                    """{"canPublish":true,"submission":{"id":42,"boardId":"Commodore/C64/250407","state":"pending"},"findings":[]}""");
            }

            return AnsweringHttpHandler.Refused();
        });

        var tab = new TabMaintainer();
        MaintainerTabHost.AddOverlay(tab);

        typeof(TabMaintainer).GetField("thisClient", Any)!
            .SetValue(tab, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
        typeof(TabMaintainer).GetField("thisSession", Any)!
            .SetValue(tab, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

        tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(42)], false));

        await tab.PrefetchEntryToOpenForTests();

        hold = true;
        tableVersion = tableVersionOnReRead;

        tab.SelectOnEntryForTests();
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(42, tab.SelectedQueueRowForTests?.Id);
        Assert.True(tab.FindControl<StackPanel>("DecisionPanel")!.IsVisible);
        Assert.NotNull(tab.LastPrefetchReReadForTests);

        return (tab, release);
    }

    // ###########################################################################################
    // *** NOR OVER A CELL BEING TYPED IN (code review, 2026-11-02). *** The re-read asked whether the
    // table held unsaved changes, which a cell still in its editor does not count as yet: a newer table
    // arriving then built the table again and threw away what was being typed. The same gap as the
    // Boards screen's quiet read (BoardDetailView.Table.cs).
    // ###########################################################################################
    [Fact]
    public async Task The_re_read_after_opening_leaves_a_cell_being_typed_in_alone()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, TaskCompletionSource release) = await OpenedWithTheReReadHeldAsync(
                () => Task.FromResult(AnsweringHttpHandler.Refused()),
                tableVersionOnReRead: 2);

            var window = new Window { Content = tab.Parent, Width = 1400, Height = 900 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            BoardTableEditor editor = tab.TableEditorForTests;
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");
            BoardTableSheet components = editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;

            editor.SelectSheet(components);
            Dispatcher.UIThread.RunJobs();

            editor.SelectCell(components.Rows.Single(), BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName));
            grid.Focus();
            Assert.True(grid.BeginEdit());
            Dispatcher.UIThread.RunJobs();

            release.SetResult();
            await tab.LastPrefetchReReadForTests!;

            Assert.True(editor.IsEditingCell);
            Assert.Equal(1, ShownTableVersion(tab));

            window.Close();
        });
    }

    // The first of two approvals: the submission stays selected, with what became of it beside it.
    private static HttpResponseMessage FirstOfTwoApprovals()
    {
        var settings = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        ReviewApiContract.ApplyWireSettings(settings);

        return AnsweringHttpHandler.Json(JsonSerializer.Serialize(
            new ReviewDecisionAnswer("approved", WaitingFor: [ApproverRole.Administrator]), settings));
    }

    private static Task Approve(TabMaintainer tab) =>
        (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(tab, [ReviewDecisionKind.Approve])!;

    // ###########################################################################################
    // *** THE RE-READ NEVER LANDS ON A DECISION (code review, 2026-11-02). *** Approve pressed at once
    // on the entry read ahead, the older answer arriving after it with the submission still selected
    // (a two-person approval): showing it cleared what the decision said and put the old approval
    // state back. ReapplyApprovalGate has always kept this rule; this path did not.
    // ###########################################################################################
    [Fact]
    public async Task The_re_read_after_opening_leaves_a_decision_made_meanwhile_alone()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, TaskCompletionSource release) =
                await OpenedWithTheReReadHeldAsync(() => Task.FromResult(FirstOfTwoApprovals()));

            await Approve(tab);

            TextBlock message = tab.FindControl<TextBlock>("DecisionMessageText")!;
            string said = message.Text!;

            Assert.True(message.IsVisible);
            Assert.NotEqual(string.Empty, said);

            release.SetResult();
            await tab.LastPrefetchReReadForTests!;
            Dispatcher.UIThread.RunJobs();

            Assert.True(message.IsVisible);
            Assert.Equal(said, message.Text);
        });
    }

    // And while the decision is on its way: its buttons stay off, so a second click cannot send it twice.
    [Fact]
    public async Task The_re_read_after_opening_does_not_turn_the_buttons_back_on_while_a_decision_is_sent()
    {
        await UiTest.RunAsync(async () =>
        {
            var answer = new TaskCompletionSource<HttpResponseMessage>();

            (TabMaintainer tab, TaskCompletionSource release) = await OpenedWithTheReReadHeldAsync(() => answer.Task);

            Task deciding = Approve(tab);
            Dispatcher.UIThread.RunJobs();

            Button reject = tab.FindControl<Button>("RejectButton")!;
            Assert.False(reject.IsEnabled);

            release.SetResult();
            await tab.LastPrefetchReReadForTests!;
            Dispatcher.UIThread.RunJobs();

            Assert.False(reject.IsEnabled);

            answer.SetResult(FirstOfTwoApprovals());
            await deciding;

            Assert.True(tab.FindControl<TextBlock>("DecisionMessageText")!.IsVisible);
        });
    }

    // With no decision anywhere near it, the re-read is what it always was: what the server says now.
    [Fact]
    public async Task The_re_read_after_opening_still_shows_what_the_server_says_now_with_no_decision_made()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, TaskCompletionSource release) =
                await OpenedWithTheReReadHeldAsync(() => Task.FromResult(FirstOfTwoApprovals()));

            object? shownBefore = typeof(TabMaintainer).GetField("thisShownDetail", Any)!.GetValue(tab);

            release.SetResult();
            await tab.LastPrefetchReReadForTests!;
            Dispatcher.UIThread.RunJobs();

            Assert.NotSame(shownBefore, typeof(TabMaintainer).GetField("thisShownDetail", Any)!.GetValue(tab));
        });
    }

    // Signed out - or another account - what was read is gone with the session it was read for.
    [Fact]
    public async Task Signing_out_forgets_what_was_read_ahead()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, _) = SignedIn(Row(42));
            await tab.PrefetchEntryToOpenForTests();

            tab.UseSessionForTests(null);

            Assert.Null(tab.PrefetchedRowForTests);
        });
    }
}
