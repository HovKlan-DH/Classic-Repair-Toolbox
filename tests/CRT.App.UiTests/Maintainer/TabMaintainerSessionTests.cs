using System.Net;
using System.Reflection;
using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The signed-in session on the Maintainer tab (TabMaintainer.Session.cs, .QueueRefresh.cs) - the
// code review of 2026-10-13: the remembered sign-in follows the server's sliding expiry, a 401
// under unsaved table changes stops the checks and says what is left to do, and a system where a
// sign-in cannot be remembered says so under "Sign in". Built, never shown, and never the real
// session file: the tab is handed its session and client (UseSessionForTests, UseClientForTests).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerSessionTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(long id) =>
        new(id, "Commodore/C64/250407", "pending", "Corrected U8.", "c@example.com", null);

    private static ReviewTableData Table() =>
        new(
            Version: 1,
            Published: new SubmissionRows
            {
                ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Name = "Pinout", File = "Commodore/C64/250407/old.png" }]
            },
            Submitted: new SubmissionRows
            {
                ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Name = "Pinout", File = "Commodore/C64/250407/new.png" }]
            });

    private static ReviewSession Session(DateTimeOffset expires, string token = "token") =>
        new(token, expires, 7, "dh@example.com", "Dennis");

    // ###########################################################################################
    // *** THE SESSION IN USE SLIDES AS THE SERVER SLIDES IT (code review, 2026-10-13). *** A
    // request the server accepted, sent with less than half the lifetime left, moves the expiry a
    // full lifetime on - and the rest of CRT (SignedInChanged) hears of it. An answer for a token no
    // longer in use changes nothing.
    // ###########################################################################################
    [Fact]
    public void A_request_the_server_accepted_slides_the_session_in_use()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            DateTimeOffset now = DateTimeOffset.UtcNow;
            tab.UseSessionForTests(TabMaintainerSessionTests.Session(now.AddDays(10)));

            int changes = 0;
            tab.SignedInChanged += () => changes++;

            tab.SlideSession("someone-else", now);
            Assert.Equal(now.AddDays(10), tab.SignedIn!.ExpiresUtc);

            tab.SlideSession("token", now);
            Assert.Equal(now.AddDays(30), tab.SignedIn!.ExpiresUtc);
            Assert.Equal(1, changes);

            // Again at once: more than half left now, so nothing moves and nobody is told.
            tab.SlideSession("token", now.AddMinutes(1));
            Assert.Equal(now.AddDays(30), tab.SignedIn!.ExpiresUtc);
            Assert.Equal(1, changes);
        });
    }

    // End to end: the tab's client answering a request is what slides it - posted to the UI thread.
    [Fact]
    public async Task The_tabs_own_client_slides_the_session_on_an_accepted_request()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            ReviewSession session = TabMaintainerSessionTests.Session(DateTimeOffset.UtcNow.AddDays(5));
            tab.UseSessionForTests(session);

            var server = new AnsweringHttpHandler(_ => AnsweringHttpHandler.Json("""{"isAdministrator":false,"submissions":[]}"""));
            var client = new ReviewApiClient("https://review.invalid", new HttpClient(server));
            tab.UseClientForTests(client);

            Assert.True((await client.GetQueueAsync(session)).IsOk);
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.SignedIn!.ExpiresUtc > DateTimeOffset.UtcNow.AddDays(29));
        });
    }

    // ###########################################################################################
    // *** A 401 UNDER UNSAVED TABLE CHANGES STOPS THE CHECKS AND SAYS WHAT IS LEFT TO DO (code
    // review, 2026-10-13). *** The panel was kept for the changes, but the minute check went on
    // asking, and "Sign in to continue." came back every minute with no way out said. Now it is
    // the expired session's case: the checks stop, the table stays, one sentence says it.
    // ###########################################################################################
    [Fact]
    public async Task A_401_under_unsaved_table_changes_stops_the_checks_and_says_what_is_left_to_do()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            var server = new AnsweringHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Unauthorized));

            tab.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            tab.UseSessionForTests(TabMaintainerSessionTests.Session(DateTimeOffset.UtcNow.AddDays(10)));
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(tab, null);

            tab.OpenTableForTests(TabMaintainerSessionTests.Row(42), TabMaintainerSessionTests.Table());
            BoardTableDocument document = tab.TableEditorForTests.CommitAndGetDocument()!;
            document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(row => !row.IsDeleted)
                .Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";

            DispatcherTimer timer = (DispatcherTimer)typeof(TabMaintainer).GetField("thisQueueTimer", Any)!.GetValue(tab)!;
            Assert.True(timer.IsEnabled);

            await (Task)typeof(TabMaintainer).GetMethod("RefreshQueueAsync", Any)!.Invoke(tab, [true])!;

            Assert.False(timer.IsEnabled);
            Assert.Equal(ReviewTableWording.SessionEndedUnderUnsavedEdits(expired: false), tab.FindControl<TextBlock>("QueueMessageText")!.Text);
            Assert.True(tab.IsTableOpen);
            Assert.True(tab.HasUnsavedTableEdits);
            Assert.False(tab.FindControl<StackPanel>("SignInPanel")!.IsVisible);
        });
    }

    // Without unsaved changes a 401 is the sign-in screen, as always.
    [Fact]
    public async Task A_401_with_nothing_unsaved_goes_to_the_sign_in_screen()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            var server = new AnsweringHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Unauthorized));

            tab.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            tab.UseSessionForTests(TabMaintainerSessionTests.Session(DateTimeOffset.UtcNow.AddDays(10)));
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(tab, null);

            await (Task)typeof(TabMaintainer).GetMethod("RefreshQueueAsync", Any)!.Invoke(tab, [true])!;

            Assert.True(tab.FindControl<StackPanel>("SignInPanel")!.IsVisible);
            Assert.Null(tab.SignedIn);
        });
    }

    // ###########################################################################################
    // *** TAKEN OFF THE LAST BOARD, NOTHING IS WAITING (code review, 2026-11-02). *** The queue and BETA
    // checks answer 403 from then on, every minute; that read as "could not ask" and left the lists
    // alone, so the tab's badge and the buttons went on saying "2" until signing out, and "hide the tab
    // while no work is waiting" never hid it. A 403 is not "could not ask": the server has said this
    // account reviews nothing. The lists go, the server's sentence stays on the line.
    // ###########################################################################################
    private static (TabMaintainer Tab, string Words) TakenOffTheLastBoard()
    {
        const string words = "This account does not maintain any board.";
        var tab = new TabMaintainer();

        var server = new AnsweringHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Forbidden)
        {
            Content = new StringContent($"{{\"error\":\"{words}\"}}", System.Text.Encoding.UTF8, "application/json")
        });

        tab.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(server)));
        tab.UseSessionForTests(TabMaintainerSessionTests.Session(DateTimeOffset.UtcNow.AddDays(10)));
        typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(tab, null);

        tab.ApplyQueueResponse(new ReviewQueueResponse(
            CanPublish: true,
            Submissions: [TabMaintainerSessionTests.Row(41) with { BoardId = "Commodore/C64/250407" }, TabMaintainerSessionTests.Row(42) with { BoardId = "Commodore/C128/310378" }],
            IsAdministrator: false));

        tab.ApplyBetaListAsync(
            new ProductionListResponse(true, [new ProductionBoardRow("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-25", "hash-1", "2026-May-14", null, true)]),
            background: true).GetAwaiter().GetResult();

        return (tab, words);
    }

    [Fact]
    public async Task A_maintainer_taken_off_the_last_board_has_nothing_waiting()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, string words) = TabMaintainerSessionTests.TakenOffTheLastBoard();
            Assert.Equal(3, tab.TabAttention);

            await (Task)typeof(TabMaintainer).GetMethod("RefreshQueueAsync", Any)!.Invoke(tab, [true])!;

            Assert.Equal(0, tab.TabAttention);
            Assert.Null(tab.ModeBadgeForTests(MaintainerMode.Review));
            Assert.Null(tab.ModeBadgeForTests(MaintainerMode.Beta));
            Assert.Empty(tab.QueueTextsForTests());
            Assert.Empty(tab.BetaTextsForTests());

            // A real answer, so "hide the tab while no work is waiting" can hide it - and said why.
            Assert.True(tab.BadgeKnown);
            Assert.Equal(words, tab.FindControl<TextBlock>("QueueMessageText")!.Text);
        });
    }

    // The same refusal under a table with unsaved changes takes nothing away: the lists stay beside the
    // table, the sentence says it, and the next check asks again.
    [Fact]
    public async Task A_maintainer_taken_off_the_last_board_keeps_the_lists_under_unsaved_table_changes()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer tab, string words) = TabMaintainerSessionTests.TakenOffTheLastBoard();

            tab.OpenTableForTests(TabMaintainerSessionTests.Row(41) with { BoardId = "Commodore/C64/250407" }, TabMaintainerSessionTests.Table());
            BoardTableDocument document = tab.TableEditorForTests.CommitAndGetDocument()!;
            document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!.Rows.First(row => !row.IsDeleted)
                .Cells[BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName)].Text = "Changed";

            await (Task)typeof(TabMaintainer).GetMethod("RefreshQueueAsync", Any)!.Invoke(tab, [true])!;

            Assert.Equal(3, tab.TabAttention);
            Assert.NotEmpty(tab.QueueTextsForTests());
            Assert.True(tab.IsTableOpen);
            Assert.True(tab.HasUnsavedTableEdits);
            Assert.Equal(words, tab.FindControl<TextBlock>("QueueMessageText")!.Text);
        });
    }

    // ###########################################################################################
    // *** AN ANSWER THAT LANDS AFTER SIGNING OUT BRINGS NOTHING BACK (code review, 2026-11-02). *** The
    // minute check's reads never looked at the session after their await. Landing after the sign-in
    // screen they refilled the hidden lists, set "known" and the administrator's flag, and redrew the
    // tab's badge - and the timer being stopped, the signed-out tab kept the count until the next
    // sign-in, whose account started with the old queue and the administrator's entries.
    // ###########################################################################################
    private static async Task LandsAfterSigningOutAsync(string method, object[] arguments, string answer, Action<TabMaintainer> assertNothingBack)
    {
        var tab = new TabMaintainer();
        var gate = new TaskCompletionSource<string>();

        tab.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(
            async (_, _) => AnsweringHttpHandler.Json(await gate.Task)))));
        tab.UseSessionForTests(TabMaintainerSessionTests.Session(DateTimeOffset.UtcNow.AddDays(10)));
        typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(tab, null);

        Task asked = (Task)typeof(TabMaintainer).GetMethod(method, Any)!.Invoke(tab, arguments)!;
        Assert.False(asked.IsCompleted);

        typeof(TabMaintainer).GetMethod("ShowSignInPanel", Any)!.Invoke(tab, null);

        gate.SetResult(answer);
        await asked;

        assertNothingBack(tab);
        Assert.Equal(0, tab.TabAttention);
        Assert.Null(tab.ModeBadgeForTests(MaintainerMode.Review));
        Assert.Null(tab.ModeBadgeForTests(MaintainerMode.Beta));
        Assert.Null(tab.ModeBadgeForTests(MaintainerMode.Boards));
    }

    [Fact]
    public async Task A_queue_answer_landing_after_signing_out_brings_nothing_back()
    {
        await UiTest.RunAsync(() => TabMaintainerSessionTests.LandsAfterSigningOutAsync(
            "RefreshQueueAsync",
            [true],
            """{"isAdministrator":true,"oscilloscopeModelsWaiting":3,"submissions":[{"id":41,"boardId":"Commodore/C64/250407","state":"pending","summary":"Corrected U8.","contactEmail":"c@example.com","awaitsYou":true}]}""",
            tab =>
            {
                Assert.Empty(tab.QueueTextsForTests());
                Assert.Equal(["MyAccountItem", "ServerVersionItem"], tab.AccountEntriesShownForTests);
                Assert.False((bool)typeof(TabMaintainer).GetField("thisQueueKnown", Any)!.GetValue(tab)!);
            }));
    }

    [Fact]
    public async Task A_BETA_list_landing_after_signing_out_brings_nothing_back()
    {
        await UiTest.RunAsync(() => TabMaintainerSessionTests.LandsAfterSigningOutAsync(
            "RefreshBetaAsync",
            [true, false],
            """{"configured":true,"boards":[{"boardId":"Commodore/C64/250407","manufacturer":"Commodore","hardware":"C64","board":"250407","betaRevision":"2026-September-25","betaContentHash":"hash-1","productionRevision":"2026-May-14"}]}""",
            tab =>
            {
                Assert.Empty(tab.BetaTextsForTests());
                Assert.False((bool)typeof(TabMaintainer).GetField("thisBetaKnown", Any)!.GetValue(tab)!);
            }));
    }

    [Fact]
    public async Task A_boards_list_landing_after_signing_out_brings_nothing_back()
    {
        await UiTest.RunAsync(() => TabMaintainerSessionTests.LandsAfterSigningOutAsync(
            "RefreshBoardsAsync",
            [true],
            System.Text.Json.JsonSerializer.Serialize(
                new BoardOverviewAnswer([new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1)]),
                ReviewApiContract.WireSettings),
            tab =>
            {
                Assert.Empty(tab.BoardsTextsForTests());
                Assert.False((bool)typeof(TabMaintainer).GetField("thisBoardsKnown", Any)!.GetValue(tab)!);
            }));
    }

    // The drop-down listing the minute check reads away from the Boards screen is the same read: its
    // answer turned the Boards button's badge into an attention badge on the signed-out tab.
    [Fact]
    public async Task A_listing_landing_after_signing_out_brings_nothing_back()
    {
        await UiTest.RunAsync(() => TabMaintainerSessionTests.LandsAfterSigningOutAsync(
            "RefreshListingAsync",
            [],
            System.Text.Json.JsonSerializer.Serialize(
                new BoardListingAnswer(
                    true,
                    [],
                    [new UnlistedBoardEntry("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, true, null, new BoardPlacement("Commodore 128", "310378 Open128", string.Empty, null))]),
                ReviewApiContract.WireSettings),
            tab => Assert.Null(typeof(TabMaintainer).GetField("thisListing", Any)!.GetValue(tab))));
    }

    // ###########################################################################################
    // *** WHERE A SIGN-IN CANNOT BE REMEMBERED, THE SIGN-IN SCREEN SAYS SO (code review,
    // 2026-10-13). *** Linux and macOS store no token (no DPAPI), and nothing said so - the password
    // asked at every launch read as CRT forgetting. One grey line under "Sign in", off Windows only.
    // ###########################################################################################
    [Fact]
    public void Under_Sign_in_a_system_that_cannot_remember_the_sign_in_says_so()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            TextBlock note = tab.FindControl<TextBlock>("NotRememberedText")!;

            // As this system is.
            Assert.Equal(!ReviewSessionStore.IsSupported, note.IsVisible);

            tab.ShowNotRememberedNote(isSupported: false);
            Assert.True(note.IsVisible);
            Assert.Equal(ReviewSessionStore.NotRememberedLine, note.Text);

            tab.ShowNotRememberedNote(isSupported: true);
            Assert.False(note.IsVisible);
        });
    }

    // ###########################################################################################
    // *** A PRE-RELEASE SIGNS IN TO THE TEST INSTANCE AND SAYS SO (2026-10-25). *** Its accounts are
    // its own; a maintainer typing the live password would otherwise read the refusal as a broken
    // account. Shown as CRT is built, and both ways through the seam.
    // ###########################################################################################
    [Fact]
    public void Above_Sign_in_a_CRT_talking_to_the_test_instance_says_so()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            TextBlock note = tab.FindControl<TextBlock>("TestServerSignInText")!;

            Assert.Equal(AppConfig.UsesTestServer, note.IsVisible);

            tab.ShowTestServerNote(usesTestServer: true);
            Assert.True(note.IsVisible);
            Assert.Equal(Handlers.Online.ContributionServer.SignInNote, note.Text);

            tab.ShowTestServerNote(usesTestServer: false);
            Assert.False(note.IsVisible);
        });
    }
}
