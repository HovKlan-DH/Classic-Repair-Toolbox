using System.Net;
using Avalonia.Controls;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE ACCOUNT WINDOW OPENED FROM THE "MAINTAINER" TAB (Main.Account.cs): its "Sign in" opens it on the view for where the
// account stands, and what the window signs in is the Maintainer tab's session - the one the whole of CRT uses (the Feedback
// tab's address). The account icon in CRT's top right corner is GONE (owner decision, 2026-11-28: no accounts for everybody).
//
// Main is BUILT, never shown, with UserSettings and the workbook folder pointed at temp files - MainMaintainerTabTests' setup.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainAccountTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private static readonly ReviewSession Session = new("bearer-1", DateTimeOffset.UtcNow.AddDays(20), 7, "anna@example.com", "Anna");

    public MainAccountTests()
    {
        this.RedirectToTemp();
    }

    public void Dispose()
    {
        this.RedirectToTemp();
        this.thisWorkspace.Dispose();
    }

    private void RedirectToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        // An EMPTY file, not a missing one: LoadFrom keeps the settings already in memory when the file does not exist.
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

    // Owner decision, 2026-11-28: "Entirely remove the new "Account" possibility" - nothing in CRT's window offers an account to
    // everybody; the Maintainer tab's "Sign in" is the way in.
    [Fact]
    public void CRT_s_window_has_no_account_icon()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            Assert.Null(window.FindControl<Button>("AccountButton"));
        });
    }

    // "Sign in" on the Maintainer tab opens the window on the sign-in, or on "My account" while signed in.
    [Fact]
    public async Task The_tabs_Sign_in_opens_the_account_window_where_the_account_stands()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();
            var opened = new List<AccountWindow.AccountView>();
            window.ShowAccountWindowOverrideForTests = shown =>
            {
                opened.Add(shown.ViewShown);
                return Task.CompletedTask;
            };

            window.TabMaintainer.GetControl<Button>("SignInButton").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            window.TabMaintainer.UseSessionForTests(MainAccountTests.Session);
            window.TabMaintainer.UseClientForTests(new ReviewApiClient("https://crt.example", new HttpClient(new AnsweringHttpHandler(_ => AnsweringHttpHandler.Refused()))));
            await window.OpenAccountWindowAsync();

            Assert.Equal([AccountWindow.AccountView.SignIn, AccountWindow.AccountView.SignedIn], opened);
        });
    }

    // ###########################################################################################
    // *** WHAT THE ACCOUNT WINDOW SIGNS IN IS THE SIGN-IN OF THE WHOLE OF CRT. *** Through Main's own host: the Maintainer tab
    // holds the session, and the Feedback tab takes the account's address - with "Enable Maintainer tab" off too.
    // ###########################################################################################
    [Fact]
    public async Task A_sign_in_made_in_the_account_window_reaches_the_tab_and_the_Feedback_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            UserSettings.ContactEmail = "typed@example.com";
            var window = new CRT.Main();

            var client = new ReviewApiClient("https://crt.example", new HttpClient(new AnsweringHttpHandler(request =>
                request.RequestUri!.AbsolutePath == "/api/accounts/me"
                    ? AccountWindowTests.Answer(HttpStatusCode.OK, new AccountAnswer(7, "anna@example.com", "Anna", true, false, [], DateTimeOffset.UtcNow))
                    : AnsweringHttpHandler.Refused())));

            await window.AccountHostForWindow().UseSignIn(MainAccountTests.Session, client);

            Assert.Same(MainAccountTests.Session, window.TabMaintainer.SignedIn);
            Assert.Equal("anna@example.com", window.TabFeedback.GetControl<TextBox>("EmailTextBox").Text);
        });
    }
}
