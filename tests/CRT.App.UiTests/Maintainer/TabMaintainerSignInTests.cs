using System.Net;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// THE MAINTAINER TAB'S SIGN-IN SINCE 2026-11-25 (accounts for everybody): the sign-in form, "I forgot my password" and "I have
// an invitation" are CRT's account window's (AccountWindowTests); the tab keeps what it is, "Sign in" - which asks for that
// window - and takes over what the window signed in (UseSignInAsync), the session staying the tab's. An account that
// maintains no board is said so instead of lists the server would refuse.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerSignInTests
{
    private static readonly ReviewSession Session = new("bearer-1", DateTimeOffset.UtcNow.AddDays(20), 7, "anna@example.com", "Anna");

    private static AccountAnswer Account(bool maintainer) =>
        new(7, "anna@example.com", "Anna", true, IsAdministrator: false, MaintainerOf: maintainer ? ["Commodore/C64/250407"] : [], CreatedUtc: DateTimeOffset.UtcNow);

    // A client whose /me answers `account` (read again on each request, so a test can change it), every other request refused;
    // every path asked kept in `asked`.
    private static ReviewApiClient Client(Func<AccountAnswer> account, List<string> asked) =>
        new("https://crt.example", new HttpClient(new AnsweringHttpHandler(request =>
        {
            asked.Add(request.RequestUri!.AbsolutePath);

            return request.RequestUri.AbsolutePath == "/api/accounts/me"
                ? AccountWindowTests.Answer(HttpStatusCode.OK, account())
                : AnsweringHttpHandler.Refused();
        })));

    [Fact]
    public void Sign_in_on_the_tab_asks_for_the_account_window()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            int asked = 0;
            tab.SignInWanted += () => asked++;

            tab.GetControl<Button>("SignInButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Equal(1, asked);
        });
    }

    // ###########################################################################################
    // *** A SIGN-IN MADE IN THE ACCOUNT WINDOW IS THE TAB'S SESSION. *** Taken over with the client it was got with, shared
    // with the rest of CRT (SignedInChanged), and - the tab not on screen - its lists left for when it is first shown: only
    // the account is read now.
    // ###########################################################################################
    [Fact]
    public async Task A_sign_in_from_the_account_window_becomes_the_tabs_session()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            MaintainerTabHost.AddOverlay(tab);

            var asked = new List<string>();
            ReviewApiClient client = TabMaintainerSignInTests.Client(() => TabMaintainerSignInTests.Account(maintainer: true), asked);
            int changes = 0;
            tab.SignedInChanged += () => changes++;

            await tab.UseSignInAsync(TabMaintainerSignInTests.Session, client, readLists: false);

            Assert.Same(TabMaintainerSignInTests.Session, tab.SignedIn);
            Assert.Same(client, tab.ClientInUse);
            Assert.True(changes >= 1);
            Assert.True(tab.GetControl<Grid>("QueuePanel").IsVisible);
            Assert.False(tab.SaysAccountMaintainsNothingForTests);
            Assert.Equal(["/api/accounts/me"], asked);

            // The queue's minute check, started with the queue, must not outlive the test on the shared dispatcher.
            typeof(TabMaintainer).GetMethod("StopQueueChecks", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(tab, null);
        });
    }

    // ###########################################################################################
    // *** AN ACCOUNT THAT MAINTAINS NO BOARD IS SAID SO (2026-11-25). *** Anybody may have an account now; signed in with
    // one that is in no board's pool and no administrator, the tab says so - who is signed in, and why there is nothing here -
    // instead of the queue, and reads none of the lists. Once the server says the account maintains a board, the queue comes
    // back; signed out, the panel offers "Sign in" again.
    // ###########################################################################################
    [Fact]
    public async Task An_account_that_maintains_no_board_is_told_so_instead_of_the_queue()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            MaintainerTabHost.AddOverlay(tab);

            var asked = new List<string>();
            ReviewApiClient client = TabMaintainerSignInTests.Client(() => TabMaintainerSignInTests.Account(maintainer: false), asked);

            await tab.UseSignInAsync(TabMaintainerSignInTests.Session, client, readLists: true);

            Assert.True(tab.SaysAccountMaintainsNothingForTests);
            Assert.False(tab.GetControl<Grid>("QueuePanel").IsVisible);
            Assert.True(tab.GetControl<StackPanel>("SignInPanel").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("SignedOutPart").IsVisible);
            Assert.Equal(
                "Signed in as Anna (anna@example.com). This account is not a maintainer of any board, so there is nothing for it here. " +
                "The administrator makes an account a board's maintainer.",
                TabMaintainer.TextOf(tab.GetControl<TextBlock>("NotMaintainerText")));
            Assert.Equal(["/api/accounts/me"], asked);

            // Made a maintainer since: the queue is back.
            tab.ApplyAccount(TabMaintainerSignInTests.Account(maintainer: true));
            Dispatcher.UIThread.RunJobs();

            Assert.False(tab.SaysAccountMaintainsNothingForTests);
            Assert.True(tab.GetControl<Grid>("QueuePanel").IsVisible);

            // Signed out: "Sign in" again, nothing said of the account.
            await tab.SignOutAsync();

            Assert.True(tab.GetControl<StackPanel>("SignedOutPart").IsVisible);
            Assert.False(tab.SaysAccountMaintainsNothingForTests);
            Assert.Null(tab.SignedIn);
        });
    }
}
