using System.Net;
using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Media;
using Avalonia.Threading;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// "My account" (owner requests, 2026-10-03: the maintainer "can edit his/her own email address and
// name"; and 2026-10-04: the "Your account" window became the first entry under the Account screen,
// with Sign out), shown, with the server answering from AnsweringHttpHandler - no network.
//
// What matters is what reaches the rest of CRT: every account the server hands back is passed on at
// once (the tab, the remembered sign-in, the Feedback tab and the Submit dialog follow it), while a
// new address that is only waiting for its code passes on NOTHING - the address has not changed.
// These were AccountWindowTests until the window became this view.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MyAccountViewTests
{
    private static readonly ReviewSession Session =
        new("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");

    private static string AccountJson(string email, string name) =>
        $$"""{"id":7,"email":"{{email}}","displayName":"{{name}}","isVerified":true,"isAdministrator":false,"maintainerOf":[],"createdUtc":"2026-01-01T00:00:00+00:00"}""";

    private static HttpResponseMessage Changed(string message, string email, string name) =>
        AnsweringHttpHandler.Json($$"""{"message":"{{message}}","account":{{MyAccountViewTests.AccountJson(email, name)}}}""");

    private sealed record Opened(MyAccountView View, Window Window, List<AccountAnswer> PassedOn, List<string> Paths);

    private static Opened Open(Func<HttpRequestMessage, HttpResponseMessage> answer)
    {
        var paths = new List<string>();
        var passedOn = new List<AccountAnswer>();

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(request =>
        {
            paths.Add(request.RequestUri!.AbsolutePath);
            return answer(request);
        })));

        var view = new MyAccountView();
        view.Initialize(client, MyAccountViewTests.Session);
        view.AccountChanged = passedOn.Add;

        // Shown, so typing raises TextChanged - which is what turns the buttons on.
        var window = new Window { Content = view, Width = 800, Height = 900 };
        window.Show();

        return new Opened(view, window, passedOn, paths);
    }

    private static void Type(MyAccountView view, string name, string text)
    {
        view.FindControl<TextBox>(name)!.Text = text;
        Dispatcher.UIThread.RunJobs();
    }

    private static string? Typed(MyAccountView view, string name) => view.FindControl<TextBox>(name)!.Text;

    private static bool IsOn(MyAccountView view, string name) => view.FindControl<Button>(name)!.IsEnabled;

    private static bool Shown(MyAccountView view, string name) => view.FindControl<Control>(name)!.IsVisible;

    private static string? Said(MyAccountView view, string name) => view.FindControl<TextBlock>(name)!.Text;

    private static string LoggedIn(MyAccountView view) => TabMaintainer.TextOf(view.FindControl<TextBlock>("LoggedInAsText")!);

    private static Color ColourOf(MyAccountView view, string name) =>
        ((ISolidColorBrush)view.FindControl<TextBlock>(name)!.Foreground!).Color;

    private static void Click(MyAccountView view, string name) =>
        view.FindControl<Button>(name)!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    [Fact]
    public void It_opens_on_the_account_with_nothing_to_send_yet()
    {
        UiTest.Run(() =>
        {
            Opened opened = MyAccountViewTests.Open(_ => AnsweringHttpHandler.Refused());
            MyAccountView view = opened.View;

            Assert.Equal("Dennis", MyAccountViewTests.Typed(view, "NameTextBox"));
            Assert.Equal("Logged in as Dennis (dh@example.com)", MyAccountViewTests.LoggedIn(view));

            // The name bold, as under the lists (owner request, 2026-10-04).
            Run name = view.FindControl<TextBlock>("LoggedInAsText")!.Inlines!.OfType<Run>().Single(run => run.FontWeight == FontWeight.Bold);
            Assert.Equal("Dennis", name.Text);

            Assert.False(MyAccountViewTests.IsOn(view, "SaveNameButton"));
            Assert.False(MyAccountViewTests.IsOn(view, "SendCodeButton"));
            Assert.False(MyAccountViewTests.IsOn(view, "ChangePasswordButton"));
            Assert.False(MyAccountViewTests.Shown(view, "CodePanel"));
            Assert.Empty(opened.Paths);

            opened.Window.Close();
        });
    }

    [Fact]
    public async Task A_saved_name_is_passed_on_at_once()
    {
        await UiTest.RunAsync(async () =>
        {
            Opened opened = MyAccountViewTests.Open(_ => MyAccountViewTests.Changed("Your name is now Dennis H.", "dh@example.com", "Dennis H"));
            MyAccountView view = opened.View;

            MyAccountViewTests.Type(view, "NameTextBox", "Dennis H");
            Assert.True(MyAccountViewTests.IsOn(view, "SaveNameButton"));

            await view.SaveNameAsync();

            Assert.Equal(["/api/accounts/me/name"], opened.Paths);
            Assert.Equal("Dennis H", Assert.Single(opened.PassedOn).DisplayName);
            Assert.Equal("Dennis H", view.Session!.DisplayName);
            Assert.Equal("token", view.Session.BearerToken);
            Assert.Equal("Logged in as Dennis H (dh@example.com)", MyAccountViewTests.LoggedIn(view));

            Assert.Equal("Your name is now Dennis H.", MyAccountViewTests.Said(view, "NameMessageText"));
            Assert.Equal(Colors.SeaGreen, MyAccountViewTests.ColourOf(view, "NameMessageText"));

            // Nothing left to save.
            Assert.False(MyAccountViewTests.IsOn(view, "SaveNameButton"));

            opened.Window.Close();
        });
    }

    // ###########################################################################################
    // *** ASKING FOR A CODE CHANGES NOTHING. *** No current password is asked for (owner decision,
    // 2026-10-03), and no account is passed on - the address is still the old one until the code
    // comes back. The code is typed in after "I have received a code" (owner wording), so the box
    // stays put away until that is chosen; the code coming back passes the new address on and puts
    // the box away again.
    // ###########################################################################################
    [Fact]
    public async Task A_new_address_is_passed_on_only_once_its_code_comes_back()
    {
        await UiTest.RunAsync(async () =>
        {
            Opened opened = MyAccountViewTests.Open(request => request.RequestUri!.AbsolutePath.EndsWith("/confirm", StringComparison.Ordinal)
                ? MyAccountViewTests.Changed("Your email address is now bench@example.com.", "bench@example.com", "Dennis")
                : AnsweringHttpHandler.Json("""{"message":"A code is on its way to bench@example.com.","codeSent":true}"""));
            MyAccountView view = opened.View;

            Assert.False(MyAccountViewTests.IsOn(view, "SendCodeButton"));

            MyAccountViewTests.Type(view, "NewEmailTextBox", "bench@example.com");
            Assert.True(MyAccountViewTests.IsOn(view, "SendCodeButton"));

            await view.SendCodeAsync();

            Assert.False(MyAccountViewTests.Shown(view, "CodePanel"));
            Assert.Equal("A code is on its way to bench@example.com.", MyAccountViewTests.Said(view, "EmailMessageText"));
            Assert.Empty(opened.PassedOn);
            Assert.Equal("dh@example.com", view.Session!.Email);

            MyAccountViewTests.Click(view, "HaveCodeButton");
            Assert.True(MyAccountViewTests.Shown(view, "CodePanel"));

            MyAccountViewTests.Type(view, "CodeTextBox", "iAXr2z0PPffPwpmzHR-bOFo5ZPCPfHEK1hZbFAKAoYQ");
            await view.ConfirmCodeAsync();

            Assert.Equal(["/api/accounts/me/email", "/api/accounts/me/email/confirm"], opened.Paths);
            Assert.Equal("bench@example.com", Assert.Single(opened.PassedOn).Email);
            Assert.Equal("bench@example.com", view.Session!.Email);
            Assert.False(MyAccountViewTests.Shown(view, "CodePanel"));
            Assert.Equal(string.Empty, MyAccountViewTests.Typed(view, "NewEmailTextBox"));
            Assert.Equal("Logged in as Dennis (bench@example.com)", MyAccountViewTests.LoggedIn(view));

            opened.Window.Close();
        });
    }

    // ###########################################################################################
    // The owner's words (2026-10-03): "Name or handle", "New email address", "I have received a
    // code" - and no current-password box anywhere. The email section says how the code is typed
    // in. A code asked for earlier, with another screen shown or CRT closed in between, goes in the
    // same way. The view is headed "My account", the entry's own name.
    // ###########################################################################################
    [Fact]
    public void The_view_uses_the_owners_words_and_asks_for_no_current_password()
    {
        UiTest.Run(() =>
        {
            Opened opened = MyAccountViewTests.Open(_ => AnsweringHttpHandler.Refused());
            MyAccountView view = opened.View;

            List<string?> texts = view.GetLogicalDescendants().OfType<TextBlock>().Select(text => text.Text).ToList();

            Assert.Contains(MaintainerScreenWording.MyAccount, texts);
            Assert.Contains("Name or handle", texts);
            Assert.Contains("New email address", texts);
            Assert.DoesNotContain("Current password", texts);
            Assert.Contains(texts, text => text?.Contains("Type that code in with \"I have received a code\"", StringComparison.Ordinal) == true);
            Assert.Null(view.FindControl<TextBox>("CurrentPasswordTextBox"));
            Assert.Null(view.FindControl<TextBox>("EmailPasswordTextBox"));

            Button haveCode = view.FindControl<Button>("HaveCodeButton")!;
            Assert.Equal("I have received a code", haveCode.Content);

            haveCode.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Assert.True(MyAccountViewTests.Shown(view, "CodePanel"));

            opened.Window.Close();
        });
    }

    // ###########################################################################################
    // A refusal is the server's own sentence, in red, under the section it is about. The new
    // password is kept, so a rule broken can be fixed without typing both again. Nothing is
    // passed on.
    // ###########################################################################################
    [Fact]
    public async Task A_refusal_is_the_servers_sentence_in_red_where_the_change_was_made()
    {
        await UiTest.RunAsync(async () =>
        {
            Opened opened = MyAccountViewTests.Open(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent("""{"errors":["At least 12 characters."]}""", Encoding.UTF8, "application/json")
            });
            MyAccountView view = opened.View;

            MyAccountViewTests.Type(view, "NewPasswordTextBox", "too short");
            MyAccountViewTests.Type(view, "RepeatPasswordTextBox", "too short");
            Assert.True(MyAccountViewTests.IsOn(view, "ChangePasswordButton"));

            await view.ChangePasswordAsync();

            Assert.Equal(["/api/accounts/me/password"], opened.Paths);
            Assert.Equal("At least 12 characters.", MyAccountViewTests.Said(view, "PasswordMessageText"));
            Assert.Equal(Colors.IndianRed, MyAccountViewTests.ColourOf(view, "PasswordMessageText"));
            Assert.Equal("too short", MyAccountViewTests.Typed(view, "NewPasswordTextBox"));
            Assert.Empty(opened.PassedOn);

            opened.Window.Close();
        });
    }

    [Fact]
    public void Two_new_passwords_that_differ_are_said_and_send_nothing()
    {
        UiTest.Run(() =>
        {
            Opened opened = MyAccountViewTests.Open(_ => AnsweringHttpHandler.Refused());
            MyAccountView view = opened.View;

            MyAccountViewTests.Type(view, "NewPasswordTextBox", "a new one here");
            MyAccountViewTests.Type(view, "RepeatPasswordTextBox", "a new one her");

            Assert.Equal(MyAccountRules.PasswordsDiffer, MyAccountViewTests.Said(view, "PasswordMessageText"));
            Assert.False(MyAccountViewTests.IsOn(view, "ChangePasswordButton"));

            MyAccountViewTests.Type(view, "RepeatPasswordTextBox", "a new one here");

            Assert.False(MyAccountViewTests.Shown(view, "PasswordMessageText"));
            Assert.True(MyAccountViewTests.IsOn(view, "ChangePasswordButton"));

            opened.Window.Close();
        });
    }

    // ###########################################################################################
    // *** SIGN OUT IS HERE NOW, AND IT IS THE TAB'S TO DO (owner request, 2026-10-04: "removing the
    // "Account" and "Sign out" buttons from bottom-left area"). *** The button asks the tab -
    // which asks about unsaved table changes first - and sends nothing to the server itself.
    // ###########################################################################################
    [Fact]
    public void Sign_out_is_the_last_section_and_hands_the_sign_out_to_the_tab()
    {
        UiTest.Run(() =>
        {
            Opened opened = MyAccountViewTests.Open(_ => AnsweringHttpHandler.Refused());
            MyAccountView view = opened.View;
            int asked = 0;

            view.SignOutRequested = () =>
            {
                asked++;
                return Task.CompletedTask;
            };

            Button signOut = view.FindControl<Button>("SignOutButton")!;
            Assert.Equal("Sign out", signOut.Content);

            // The last of the sections, after the password.
            Button changePassword = view.FindControl<Button>("ChangePasswordButton")!;
            Assert.True(
                signOut.TranslatePoint(default, opened.Window)!.Value.Y > changePassword.TranslatePoint(default, opened.Window)!.Value.Y,
                "Sign out is not below the password section.");

            MyAccountViewTests.Click(view, "SignOutButton");
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(1, asked);
            Assert.Empty(opened.Paths);

            opened.Window.Close();
        });
    }

    // ###########################################################################################
    // *** SIGNED OUT, NOTHING TYPED FOR THE LAST ACCOUNT STAYS. *** A view that lives on in the tab
    // (it was a window, closed and gone, until 2026-10-04) would otherwise keep a half-typed new
    // password, a code and a new address for whoever signs in next on this computer.
    // ###########################################################################################
    [Fact]
    public void Signing_out_empties_every_box_and_turns_every_button_off()
    {
        UiTest.Run(() =>
        {
            Opened opened = MyAccountViewTests.Open(_ => AnsweringHttpHandler.Refused());
            MyAccountView view = opened.View;

            MyAccountViewTests.Type(view, "NewEmailTextBox", "bench@example.com");
            MyAccountViewTests.Click(view, "HaveCodeButton");
            MyAccountViewTests.Type(view, "CodeTextBox", "a code");
            MyAccountViewTests.Type(view, "NewPasswordTextBox", "a new one here");
            MyAccountViewTests.Type(view, "RepeatPasswordTextBox", "a new one her");

            view.Initialize(null, null);
            Dispatcher.UIThread.RunJobs();

            foreach (string box in new[] { "NameTextBox", "NewEmailTextBox", "CodeTextBox", "NewPasswordTextBox", "RepeatPasswordTextBox" })
                Assert.Equal(string.Empty, MyAccountViewTests.Typed(view, box));

            foreach (string button in new[] { "SaveNameButton", "SendCodeButton", "ConfirmEmailButton", "ChangePasswordButton" })
                Assert.False(MyAccountViewTests.IsOn(view, button), button);

            Assert.False(MyAccountViewTests.Shown(view, "CodePanel"));
            Assert.False(MyAccountViewTests.Shown(view, "PasswordMessageText"));
            Assert.Equal(string.Empty, MyAccountViewTests.LoggedIn(view));
            Assert.Null(view.Session);

            opened.Window.Close();
        });
    }

    // ###########################################################################################
    // A change made ELSEWHERE - the name and address read again at launch - reaches the view: its
    // "Logged in as" line, and the name box unless something else is being typed into it.
    // ###########################################################################################
    [Fact]
    public void A_change_made_elsewhere_is_followed_but_a_name_being_typed_is_kept()
    {
        UiTest.Run(() =>
        {
            Opened opened = MyAccountViewTests.Open(_ => AnsweringHttpHandler.Refused());
            MyAccountView view = opened.View;

            view.UseAccount(MyAccountViewTests.Session with { DisplayName = "Dennis H", Email = "bench@example.com" });

            Assert.Equal("Dennis H", MyAccountViewTests.Typed(view, "NameTextBox"));
            Assert.Equal("Logged in as Dennis H (bench@example.com)", MyAccountViewTests.LoggedIn(view));

            MyAccountViewTests.Type(view, "NameTextBox", "Somebody else");
            view.UseAccount(MyAccountViewTests.Session with { DisplayName = "Dennis HH", Email = "bench@example.com" });

            Assert.Equal("Somebody else", MyAccountViewTests.Typed(view, "NameTextBox"));
            Assert.Equal("Logged in as Dennis HH (bench@example.com)", MyAccountViewTests.LoggedIn(view));
            Assert.True(MyAccountViewTests.IsOn(view, "SaveNameButton"));

            opened.Window.Close();
        });
    }
}
