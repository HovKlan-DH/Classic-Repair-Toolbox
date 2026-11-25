using System.Net;
using System.Text;
using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// CRT'S ACCOUNT WINDOW (2026-11-25, accounts for everybody - AccountWindow): signing in, "Create an account", "I forgot my
// password", "I have an invitation" and, signed in, "My account". The sign-in and the invitation were the "Maintainer" tab's
// until that day - their tests came with them (TabMaintainerInvitationTests, and the two notes from TabMaintainerSessionTests).
// The server is a function (AnsweringHttpHandler); the host - Main's, in CRT - is a test's, seeing what the window hands it.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class AccountWindowTests
{
    private const string Password = "correct horse battery staple";

    private static readonly ReviewSession Signed = new("bearer-1", DateTimeOffset.UtcNow.AddDays(20), 7, "dennis@example.com", "Dennis");

    // What a test's host saw the window hand over.
    internal sealed class Seen
    {
        public List<ReviewSession> SignIns { get; } = [];

        public List<AccountAnswer> AccountsChanged { get; } = [];

        public int SignOuts { get; set; }

        // The server's "not signed in" handed on - the host then holds no session, as the tab's does once it ended it.
        public int SessionsEnded { get; set; }

        public List<HttpRequestMessage> Requests { get; } = [];

        public List<string> Bodies { get; } = [];
    }

    // A host for the window - the server answering through `answer`, every request and its body kept in `seen`.
    internal static AccountHost Host(Func<HttpRequestMessage, HttpResponseMessage> answer, Seen seen, ReviewSession? signedIn = null)
    {
        ReviewApiClient NewClient() =>
            new("https://crt.example", new HttpClient(new AnsweringHttpHandler(request =>
            {
                seen.Requests.Add(request);
                seen.Bodies.Add(request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty);
                return answer(request);
            })));

        ReviewApiClient? inUse = signedIn is null ? null : NewClient();
        ReviewSession? current = signedIn;

        return new AccountHost(
            SignedIn: () => current,
            ClientInUse: () => inUse,
            NewClient: NewClient,
            UseSignIn: (session, _) =>
            {
                seen.SignIns.Add(session);
                return Task.CompletedTask;
            },
            AccountChanged: seen.AccountsChanged.Add,
            SignOut: () =>
            {
                seen.SignOuts++;
                return Task.CompletedTask;
            },
            SessionEnded: () =>
            {
                seen.SessionsEnded++;
                current = null;
            },
            SignInIsRemembered: true,
            UsesTestServer: false);
    }

    internal static HttpResponseMessage Answer(HttpStatusCode status, object body) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(body, ReviewApiContract.WireSettings), Encoding.UTF8, "application/json") };

    private static HttpResponseMessage Unexpected(HttpRequestMessage request) =>
        throw new InvalidOperationException($"No request was expected to {request.RequestUri}.");

    private static AccountWindow Shown(AccountHost host)
    {
        var window = new AccountWindow();
        window.Initialize(host);
        window.Show();
        Dispatcher.UIThread.RunJobs();
        return window;
    }

    private static void Press(Window window, string button)
    {
        window.GetControl<Button>(button).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        Dispatcher.UIThread.RunJobs();
    }

    private static string Text(Window window, string name) => window.GetControl<TextBlock>(name).Text ?? string.Empty;

    // A wait that gives up says so (CLAUDE.md, "Headless UI tests").
    private static async Task Until(Func<bool> condition, string what)
    {
        for (int attempt = 0; attempt < 300 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), what);
    }

    private static bool Shows(Window window, string name) => window.GetControl<Control>(name).IsVisible;

    // -----------------------------------------------------------------------------------
    // The views
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** ONE VIEW AT A TIME (the invitation's rule since 2026-10-04: "it should show only the information relevant to that").
    // *** Each way from the sign-in shows its view and nothing of the sign-in form; Cancel comes back to it.
    // ###########################################################################################
    [Theory]
    [InlineData("CreateAccountButton", "RegisterPanel", "CancelRegisterButton")]
    [InlineData("ForgotPasswordButton", "ForgotPanel", "CancelForgotButton")]
    [InlineData("HaveInvitationButton", "InvitationPanel", "CancelInvitationButton")]
    public void Each_way_from_the_sign_in_shows_its_view_alone_and_Cancel_comes_back(string way, string view, string cancel)
    {
        UiTest.Run(() =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, seen));

            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.True(AccountWindowTests.Shows(window, "SignInFormPanel"));

            AccountWindowTests.Press(window, way);

            foreach (string panel in new[] { "SignInFormPanel", "RegisterPanel", "ForgotPanel", "InvitationPanel", "SignedInPanel" })
                Assert.Equal(panel == view, AccountWindowTests.Shows(window, panel));

            AccountWindowTests.Press(window, cancel);

            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.True(AccountWindowTests.Shows(window, "SignInFormPanel"));
            Assert.False(AccountWindowTests.Shows(window, view));
            Assert.Empty(seen.Requests);

            window.Close();
        });
    }

    // The words a mail or an account web page names a button by are the button's own (CRT.Data's AccountWording).
    [Fact]
    public void The_buttons_a_mail_or_a_web_page_names_carry_those_very_words()
    {
        UiTest.Run(() =>
        {
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));

            Assert.Equal(AccountWording.SignInButton, window.GetControl<Button>("SignInButton").Content);
            Assert.Equal(AccountWording.CreateAccountButton, window.GetControl<Button>("CreateAccountButton").Content);
            Assert.Equal(AccountWording.ForgotPasswordButton, window.GetControl<Button>("ForgotPasswordButton").Content);
            Assert.Equal(AccountWording.HaveInvitationButton, window.GetControl<Button>("HaveInvitationButton").Content);
            Assert.Equal(AccountWording.SendMailAgainButton, window.GetControl<Button>("SendMailAgainButton").Content);

            window.Close();
        });
    }

    // ###########################################################################################
    // The two notes the "Maintainer" tab's sign-in had (code review, 2026-10-13; 2026-10-25): where a sign-in cannot be
    // remembered (no DPAPI) it says so under "Sign in"; a CRT talking to the test instance says so above it.
    // ###########################################################################################
    [Fact]
    public void Under_and_over_Sign_in_the_notes_say_what_is_so_on_this_system_and_this_server()
    {
        UiTest.Run(() =>
        {
            AccountWindow plain = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));
            Assert.False(AccountWindowTests.Shows(plain, "NotRememberedText"));
            Assert.False(AccountWindowTests.Shows(plain, "TestServerSignInText"));
            plain.Close();

            AccountHost noted = AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()) with { SignInIsRemembered = false, UsesTestServer = true };
            AccountWindow window = AccountWindowTests.Shown(noted);

            Assert.True(AccountWindowTests.Shows(window, "NotRememberedText"));
            Assert.Equal(ReviewSessionStore.NotRememberedLine, AccountWindowTests.Text(window, "NotRememberedText"));
            Assert.True(AccountWindowTests.Shows(window, "TestServerSignInText"));
            Assert.Equal(Handlers.Online.ContributionServer.SignInNote, AccountWindowTests.Text(window, "TestServerSignInText"));

            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------
    // Signing in
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** A SIGN-IN IS HANDED OVER, AND THE WINDOW CLOSES. *** The session goes to the host - Main gives it to the Maintainer
    // tab, which owns it - and the password box is emptied: a filled one is a credential in plain sight.
    // ###########################################################################################
    [Fact]
    public async Task Signing_in_hands_the_session_over_closes_the_window_and_empties_the_password()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath == "/api/accounts/login"
                    ? AccountWindowTests.Answer(HttpStatusCode.OK, new SessionAnswer("bearer-9", DateTimeOffset.UtcNow.AddDays(30), new SessionAccountAnswer(9, "anna@example.com", "Anna", true)))
                    : AccountWindowTests.Unexpected(request),
                seen));

            window.GetControl<TextBox>("EmailTextBox").Text = "  anna@example.com ";
            window.GetControl<TextBox>("PasswordTextBox").Text = AccountWindowTests.Password;

            await window.SignInAsync();

            ReviewSession session = Assert.Single(seen.SignIns);
            Assert.Equal("bearer-9", session.BearerToken);
            Assert.Equal("Anna", session.DisplayName);
            Assert.Contains("\"email\":\"anna@example.com\"", seen.Bodies.Single(), StringComparison.Ordinal);
            Assert.Equal(string.Empty, window.GetControl<TextBox>("PasswordTextBox").Text);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public async Task A_refused_sign_in_says_so_hands_nothing_over_and_stays_open()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(_ => new HttpResponseMessage(HttpStatusCode.Unauthorized), seen));

            window.GetControl<TextBox>("EmailTextBox").Text = "anna@example.com";
            window.GetControl<TextBox>("PasswordTextBox").Text = "wrong";

            await window.SignInAsync();

            Assert.Empty(seen.SignIns);
            Assert.True(window.IsVisible);
            Assert.True(AccountWindowTests.Shows(window, "SignInMessageText"));
            Assert.NotEqual(string.Empty, AccountWindowTests.Text(window, "SignInMessageText"));

            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------
    // I forgot my password
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE ACCOUNT WEB PAGE'S TWO STEPS, IN ONE VIEW. *** Asking for a code says the server's neutral sentence - the same
    // for every address - and setting the password comes back to the sign-in with the address and the new password filled
    // in, one press from signed in: resetting is not signing in, the server issues no session for it.
    // ###########################################################################################
    [Fact]
    public async Task I_forgot_my_password_asks_for_a_code_then_sets_the_password_and_comes_back_filled_in()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath switch
                {
                    "/api/accounts/forgot-password" => AccountWindowTests.Answer(HttpStatusCode.Accepted, new { message = "If that address has an account, a reset link is on its way to it." }),
                    "/api/accounts/reset-password" => AccountWindowTests.Answer(HttpStatusCode.OK, new { message = "Your password is set. You can sign in with it now." }),
                    _ => AccountWindowTests.Unexpected(request)
                },
                seen));

            // The address typed for signing in is carried into step 1 - asking nothing yet.
            window.GetControl<TextBox>("EmailTextBox").Text = "anna@example.com";
            AccountWindowTests.Press(window, "ForgotPasswordButton");

            Assert.Equal("anna@example.com", window.GetControl<TextBox>("ForgotEmailTextBox").Text);
            Assert.Empty(seen.Requests);

            await window.SendResetCodeAsync();

            Assert.Equal("If that address has an account, a reset link is on its way to it.", AccountWindowTests.Text(window, "ForgotMessageText"));

            window.GetControl<TextBox>("ResetCodeTextBox").Text = "  the-code  ";
            window.GetControl<TextBox>("NewPasswordTextBox").Text = "a brand new long password";

            await window.SetPasswordAsync();

            Assert.Contains("\"token\":\"the-code\"", seen.Bodies[1], StringComparison.Ordinal);
            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.Equal("anna@example.com", window.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Equal("a brand new long password", window.GetControl<TextBox>("PasswordTextBox").Text);
            Assert.Equal("Your password is set. You can sign in with it now.", AccountWindowTests.Text(window, "SignInMessageText"));
            Assert.Equal(string.Empty, window.GetControl<TextBox>("ResetCodeTextBox").Text);
            Assert.Empty(seen.SignIns);

            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------
    // Create an account
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE DISPLAY NAME IS HELD TO ITS RULE AS IT IS TYPED (owner request, 2026-11-25: "only normal A-Z + 0-9 and then some
    // common special charatcers like: +-/@.[]() ... max length 30"). *** The server's own sentence under the box, gone once
    // the name keeps the rule - and a name breaking it is never sent.
    // ###########################################################################################
    [Fact]
    public async Task The_display_name_is_held_to_its_rule_as_it_is_typed_and_never_sent_breaking_it()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, seen));
            AccountWindowTests.Press(window, "CreateAccountButton");

            TextBox name = window.GetControl<TextBox>("RegisterNameTextBox");
            Assert.Equal(DisplayNames.MaximumLength, name.MaxLength);

            // TextChanged is posted (Avalonia), so the line follows once the jobs have run.
            name.Text = "Dennis Hansen";
            Dispatcher.UIThread.RunJobs();
            Assert.True(AccountWindowTests.Shows(window, "RegisterNameProblemText"));
            Assert.Equal(DisplayNames.Problem("Dennis Hansen", sayRule: false), AccountWindowTests.Text(window, "RegisterNameProblemText"));

            window.GetControl<TextBox>("RegisterEmailTextBox").Text = "dennis@example.com";
            window.GetControl<TextBox>("RegisterPasswordTextBox").Text = AccountWindowTests.Password;

            await window.RegisterAsync();

            Assert.Empty(seen.Requests);
            Assert.Equal(DisplayNames.Problem("Dennis Hansen", sayRule: false), AccountWindowTests.Text(window, "RegisterMessageText"));

            name.Text = "Dennis@CRT";
            Dispatcher.UIThread.RunJobs();
            Assert.False(AccountWindowTests.Shows(window, "RegisterNameProblemText"));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE ANSWER IS THE SERVER'S NEUTRAL ONE, AND MAKING AN ACCOUNT IS NOT SIGNING IN. *** The name, the address and the
    // password go to POST /api/accounts/register; the sign-in comes back with the address filled in and that sentence, and
    // the password is not kept - the address may have had an account already, and signing in waits for the mail.
    // ###########################################################################################
    [Fact]
    public async Task Creating_an_account_sends_name_address_and_password_and_comes_back_with_the_neutral_answer()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath == "/api/accounts/register"
                    ? AccountWindowTests.Answer(HttpStatusCode.Accepted, new { message = "If that address can receive mail, a message is on its way to it." })
                    : AccountWindowTests.Unexpected(request),
                seen));

            AccountWindowTests.Press(window, "CreateAccountButton");
            window.GetControl<TextBox>("RegisterNameTextBox").Text = " Dennis@CRT ";
            window.GetControl<TextBox>("RegisterEmailTextBox").Text = "dennis@example.com";
            window.GetControl<TextBox>("RegisterPasswordTextBox").Text = AccountWindowTests.Password;

            await window.RegisterAsync();

            RegisterAccountRequest sent = JsonSerializer.Deserialize<RegisterAccountRequest>(seen.Bodies.Single(), ReviewApiContract.WireSettings)!;
            Assert.Equal(new RegisterAccountRequest("dennis@example.com", AccountWindowTests.Password, "Dennis@CRT"), sent);

            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.Equal("dennis@example.com", window.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Equal(string.Empty, window.GetControl<TextBox>("PasswordTextBox").Text ?? string.Empty);
            Assert.Equal(string.Empty, window.GetControl<TextBox>("RegisterPasswordTextBox").Text);
            Assert.Equal(
                AccountWindow.AfterRegistering("If that address can receive mail, a message is on its way to it."),
                AccountWindowTests.Text(window, "SignInMessageText"));
            Assert.Empty(seen.SignIns);

            window.Close();
        });
    }

    // Every rule the server says was broken - a name taken among them - shown as it says them, the view kept.
    [Fact]
    public async Task A_refused_account_shows_every_rule_broken_and_keeps_the_view()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                _ => AccountWindowTests.Answer(HttpStatusCode.BadRequest, new { errors = new[] { DisplayNames.Taken, "The password must be at least 12 characters." } }),
                seen));

            AccountWindowTests.Press(window, "CreateAccountButton");
            window.GetControl<TextBox>("RegisterNameTextBox").Text = "Dennis";
            window.GetControl<TextBox>("RegisterEmailTextBox").Text = "dennis@example.com";
            window.GetControl<TextBox>("RegisterPasswordTextBox").Text = "short";

            await window.RegisterAsync();

            Assert.Equal(AccountWindow.AccountView.Register, window.ViewShown);
            Assert.Equal($"{DisplayNames.Taken} The password must be at least 12 characters.", AccountWindowTests.Text(window, "RegisterMessageText"));

            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------
    // I have an invitation (the "Maintainer" tab's until 2026-11-25)
    // -----------------------------------------------------------------------------------

    // Cancel empties the invitation whole - a password typed there must not sit in a hidden box, and the next "I have an
    // invitation" starts blank - and leaves the sign-in as it was.
    [Fact]
    public async Task Cancel_leaves_the_sign_in_as_it_was_and_empties_the_invitation()
    {
        await UiTest.RunAsync(async () =>
        {
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));
            int asked = 0;
            window.AcceptInvitationOverrideForTests = _ =>
            {
                asked++;
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Failed(ReviewApiFailure.Refused, "No."));
            };

            window.GetControl<TextBox>("EmailTextBox").Text = "anna@example.com";
            AccountWindowTests.Press(window, "HaveInvitationButton");
            window.GetControl<TextBox>("InvitationNameTextBox").Text = "Anna";
            window.GetControl<TextBox>("InvitationPasswordTextBox").Text = AccountWindowTests.Password;

            // An empty code is answered on the spot, in the invitation's own line.
            await window.AcceptInvitationAsync();
            Assert.Equal(0, asked);
            Assert.Equal("Paste the code from the invitation email first.", AccountWindowTests.Text(window, "InvitationMessageText"));

            AccountWindowTests.Press(window, "CancelInvitationButton");

            Assert.Equal("anna@example.com", window.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Equal(string.Empty, window.GetControl<TextBox>("InvitationNameTextBox").Text);
            Assert.Equal(string.Empty, window.GetControl<TextBox>("InvitationPasswordTextBox").Text);

            AccountWindowTests.Press(window, "HaveInvitationButton");
            Assert.False(AccountWindowTests.Shows(window, "InvitationMessageText"));

            window.Close();
        });
    }

    // Accepted: the sign-in comes back filled with the address invited and the password chosen, saying the account is ready.
    [Fact]
    public async Task Accepting_an_invitation_sends_code_name_and_password_and_fills_the_sign_in()
    {
        await UiTest.RunAsync(async () =>
        {
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));
            var sent = new List<AcceptInvitationRequest>();
            window.AcceptInvitationOverrideForTests = request =>
            {
                sent.Add(request);
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Ok(
                    new AcceptInvitationAnswer("anna@example.com", ["Commodore/C64/250407"], "Your account is ready.")));
            };

            AccountWindowTests.Press(window, "HaveInvitationButton");
            window.GetControl<TextBox>("InvitationCodeTextBox").Text = "  the-code  ";
            window.GetControl<TextBox>("InvitationNameTextBox").Text = "Anna";
            window.GetControl<TextBox>("InvitationPasswordTextBox").Text = AccountWindowTests.Password;

            await window.AcceptInvitationAsync();

            Assert.Equal([new AcceptInvitationRequest("the-code", "Anna", AccountWindowTests.Password)], sent);
            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.Equal("anna@example.com", window.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Equal(AccountWindowTests.Password, window.GetControl<TextBox>("PasswordTextBox").Text);
            Assert.Equal("Your account is ready.", AccountWindowTests.Text(window, "SignInMessageText"));
            Assert.Equal(string.Empty, window.GetControl<TextBox>("InvitationCodeTextBox").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** NO ANSWER IN TWO MINUTES GOES BACK TO THE SIGN-IN (code review, 2026-10-04). *** The account may exist by now, so
    // signing in is the next step: the sign-in comes back with the password chosen and the message under it.
    // ###########################################################################################
    [Fact]
    public async Task No_answer_to_an_invitation_goes_back_to_the_sign_in_with_the_password_filled_in()
    {
        await UiTest.RunAsync(async () =>
        {
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));
            window.AcceptInvitationOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<AcceptInvitationAnswer>.Failed(ReviewApiFailure.TimedOut, "No answer."));

            AccountWindowTests.Press(window, "HaveInvitationButton");
            window.GetControl<TextBox>("InvitationCodeTextBox").Text = "the-code";
            window.GetControl<TextBox>("InvitationNameTextBox").Text = "Anna";
            window.GetControl<TextBox>("InvitationPasswordTextBox").Text = AccountWindowTests.Password;

            await window.AcceptInvitationAsync();

            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.Equal(AccountWindowTests.Password, window.GetControl<TextBox>("PasswordTextBox").Text);
            Assert.Equal(MaintainerWaitWording.InvitationNoAnswer, AccountWindowTests.Text(window, "SignInMessageText"));
            Assert.Equal(string.Empty, window.GetControl<TextBox>("InvitationPasswordTextBox").Text);

            window.Close();
        });
    }

    // A refusal keeps everything typed and shows the server's words in the invitation's own line; a name the rule refuses
    // is said without asking.
    [Fact]
    public async Task A_refused_invitation_keeps_the_view_and_says_why()
    {
        await UiTest.RunAsync(async () =>
        {
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));
            int asked = 0;
            window.AcceptInvitationOverrideForTests = _ =>
            {
                asked++;
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Failed(
                    ReviewApiFailure.Refused, "That invitation has expired. Ask the administrator for a new one."));
            };

            AccountWindowTests.Press(window, "HaveInvitationButton");
            window.GetControl<TextBox>("InvitationCodeTextBox").Text = "old-code";
            window.GetControl<TextBox>("InvitationNameTextBox").Text = "Anna Hansen";
            window.GetControl<TextBox>("InvitationPasswordTextBox").Text = AccountWindowTests.Password;

            await window.AcceptInvitationAsync();
            Assert.Equal(0, asked);
            Assert.Equal(DisplayNames.Problem("Anna Hansen", sayRule: false), AccountWindowTests.Text(window, "InvitationMessageText"));

            window.GetControl<TextBox>("InvitationNameTextBox").Text = "Anna";
            await window.AcceptInvitationAsync();

            Assert.Equal(1, asked);
            Assert.Equal(AccountWindow.AccountView.Invitation, window.ViewShown);
            Assert.Equal("That invitation has expired. Ask the administrator for a new one.", AccountWindowTests.Text(window, "InvitationMessageText"));
            Assert.Equal("old-code", window.GetControl<TextBox>("InvitationCodeTextBox").Text);

            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------
    // Signed in: "My account"
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** SIGNED IN, THE WINDOW IS "MY ACCOUNT", AND SAYS WHEN THE ADDRESS IS NOT CONFIRMED. *** The account is read as the
    // window opens; a new link is one press away, and the notice goes once the server says the address is confirmed. Every
    // account the server hands back goes to the host (the Maintainer tab, which owns the session).
    // ###########################################################################################
    [Fact]
    public async Task Signed_in_the_window_is_My_account_and_says_when_the_address_is_not_confirmed()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            bool confirmed = false;

            AccountAnswer Account() =>
                new(7, "dennis@example.com", "Dennis", confirmed, IsAdministrator: false, MaintainerOf: [], CreatedUtc: DateTimeOffset.UtcNow);

            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath switch
                {
                    "/api/accounts/me" => AccountWindowTests.Answer(HttpStatusCode.OK, Account()),
                    "/api/accounts/me/verify-again" => AccountWindowTests.Answer(HttpStatusCode.OK, new AccountChangeAnswer("A new link is on its way.", Account())),
                    _ => AccountWindowTests.Unexpected(request)
                },
                seen,
                AccountWindowTests.Signed));

            Assert.Equal(AccountWindow.AccountView.SignedIn, window.ViewShown);
            await AccountWindowTests.Until(() => AccountWindowTests.Shows(window, "NotConfirmedPanel"), "the account read as the window opened says the address is not confirmed");
            Assert.Contains(seen.AccountsChanged, account => !account.IsVerified);

            await window.SendMailAgainAsync();

            Assert.Equal("/api/accounts/me/verify-again", seen.Requests[^1].RequestUri!.AbsolutePath);
            Assert.Equal("Bearer bearer-1", seen.Requests[^1].Headers.Authorization!.ToString());
            Assert.Equal("A new link is on its way.", AccountWindowTests.Text(window, "AccountMessageText"));
            Assert.True(AccountWindowTests.Shows(window, "NotConfirmedPanel"));

            // Confirmed meanwhile: the server says so, and the notice goes.
            confirmed = true;
            await window.SendMailAgainAsync();
            Assert.False(AccountWindowTests.Shows(window, "NotConfirmedPanel"));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A SESSION THE SERVER HAS ENDED IS ENDED HERE TOO *** (code review, 2026-11-25). With "Enable Maintainer tab" off nothing
    // else asks the server, so a sign-in ended elsewhere - a password reset signs out every computer - showed as signed in, the
    // window opened on "My account", and every change there answered "Sign in to continue." with no sign-in offered. The 401 to
    // the read as the window opens goes to the host (the tab's own end of a session), and the window is the sign-in, saying why.
    // ###########################################################################################
    [Fact]
    public async Task A_session_the_server_has_ended_takes_the_window_back_to_the_sign_in()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath == "/api/accounts/me"
                    ? new HttpResponseMessage(HttpStatusCode.Unauthorized)
                    : AccountWindowTests.Unexpected(request),
                seen,
                AccountWindowTests.Signed));

            await AccountWindowTests.Until(() => window.ViewShown == AccountWindow.AccountView.SignIn, "the 401 to the account read takes the window to the sign-in");

            Assert.Equal(1, seen.SessionsEnded);
            Assert.Equal(TabMaintainer.SessionEndedMessage, AccountWindowTests.Text(window, "SignInMessageText"));
            Assert.False(AccountWindowTests.Shows(window, "SignedInPanel"));

            window.Close();
        });
    }

    // The same 401 on a change made under "My account" - the name, here.
    [Fact]
    public async Task A_change_refused_as_not_signed_in_takes_the_window_back_to_the_sign_in()
    {
        await UiTest.RunAsync(async () =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath switch
                {
                    "/api/accounts/me" => AccountWindowTests.Answer(HttpStatusCode.OK, new AccountAnswer(7, "dennis@example.com", "Dennis", true, false, [], DateTimeOffset.UtcNow)),
                    "/api/accounts/me/name" => new HttpResponseMessage(HttpStatusCode.Unauthorized),
                    _ => AccountWindowTests.Unexpected(request)
                },
                seen,
                AccountWindowTests.Signed));

            await AccountWindowTests.Until(() => seen.AccountsChanged.Count == 1, "the account is read as the window opens");
            Assert.Equal(AccountWindow.AccountView.SignedIn, window.ViewShown);

            MyAccountView account = window.GetControl<MyAccountView>("MyAccountPanel");
            account.GetControl<TextBox>("NameTextBox").Text = "Dennis2";
            await account.SaveNameAsync();

            Assert.Equal(1, seen.SessionsEnded);
            Assert.Equal(AccountWindow.AccountView.SignIn, window.ViewShown);
            Assert.Equal(TabMaintainer.SessionEndedMessage, AccountWindowTests.Text(window, "SignInMessageText"));

            // Everything typed for the session gone with it.
            Assert.Equal(string.Empty, account.GetControl<TextBox>("NameTextBox").Text);

            window.Close();
        });
    }

    // Opened, the first box of the view takes the keyboard - the address typed first goes into it (code review, 2026-11-25:
    // the focus was asked for before there was a window to focus in).
    [Fact]
    public void Opened_the_sign_in_has_the_keyboard_in_its_address_box()
    {
        UiTest.Run(() =>
        {
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));

            Assert.Same(window.GetControl<TextBox>("EmailTextBox"), window.FocusManager?.GetFocusedElement());

            window.Close();
        });
    }

    // "Sign out" under "My account": the window closes and the sign-out is the host's - the tab's own, which asks about a
    // table's unsaved changes over CRT's window.
    [Fact]
    public void Signing_out_closes_the_window_and_hands_the_sign_out_over()
    {
        UiTest.Run(() =>
        {
            var seen = new Seen();
            AccountWindow window = AccountWindowTests.Shown(AccountWindowTests.Host(
                request => request.RequestUri!.AbsolutePath == "/api/accounts/me"
                    ? AccountWindowTests.Answer(HttpStatusCode.OK, new AccountAnswer(7, "dennis@example.com", "Dennis", true, false, [], DateTimeOffset.UtcNow))
                    : AccountWindowTests.Unexpected(request),
                seen,
                AccountWindowTests.Signed));

            window.GetControl<MyAccountView>("MyAccountPanel").GetControl<Button>("SignOutButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(1, seen.SignOuts);
            Assert.False(window.IsVisible);
        });
    }

    // ###########################################################################################
    // THE OWNER'S REPORT OF 2026-11-25 ("I cannot see this input field, as it is outside viewport and there is no possibility
    // for me to scroll down") was about the reset code under the "Maintainer" tab's sign-in. Every view of the window that
    // holds it now, its message lines filled, scrolls within a window as low as CRT's smallest - at twice the width, as the
    // headless font is twice as wide.
    // ###########################################################################################
    [Fact]
    public void Every_view_is_in_reach_in_a_window_as_low_as_CRTs_smallest()
    {
        UiTest.Run(() =>
        {
            var window = new AccountWindow();
            window.Initialize(AccountWindowTests.Host(AccountWindowTests.Unexpected, new Seen()));
            window.SizeToContent = SizeToContent.Manual;
            window.Width = 960;
            window.Height = 500;
            window.Show();

            try
            {
                foreach ((AccountWindow.AccountView view, string message) in new[]
                {
                    (AccountWindow.AccountView.SignIn, "SignInMessageText"),
                    (AccountWindow.AccountView.Register, "RegisterMessageText"),
                    (AccountWindow.AccountView.Forgot, "ResetMessageText"),
                    (AccountWindow.AccountView.Invitation, "InvitationMessageText")
                })
                {
                    window.ShowView(view);

                    TextBlock line = window.GetControl<TextBlock>(message);
                    line.Text = "A sentence the server might say, long enough to take a line of its own under the buttons.";
                    line.IsVisible = true;

                    ControlsInReach.Settle(window);
                    Assert.Empty(ControlsInReach.OutOfReach(window, sideways: false));
                }

                // Not vacuous: "I forgot my password" is taller than the window.
                window.ShowView(AccountWindow.AccountView.Forgot);
                ControlsInReach.Settle(window);

                ScrollViewer scroller = window.GetControl<ScrollViewer>("ViewScroller");
                Assert.True(scroller.Extent.Height > scroller.Viewport.Height + 1, "the view fits the window, so nothing here needed to scroll");
            }
            finally
            {
                window.Close();
            }
        });
    }
}
