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
