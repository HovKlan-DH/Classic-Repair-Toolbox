using Avalonia.Controls;
using Avalonia.LogicalTree;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The Maintainer tab as CRT's window handles it (2026-09-29: the separate CRT Maintainer
// application became a tab - Main.Maintainer.cs): hidden unless "Enable Maintainer tab" is on,
// placed between Drafts and Configuration, and - while selected - the sidebar and the worklog bar
// collapsed so the four screens get the window's width, then put back exactly as they were.
//
// Main is BUILT, never shown, with UserSettings and the workbook folder pointed at temp files -
// MainWindowTests' own setup, and in the "HeadlessUi" collection for the same reason it is (see
// its note on UserSettings).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainMaintainerTabTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MainMaintainerTabTests()
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
        // An EMPTY file, not a missing one: LoadFrom keeps the settings already in memory when the
        // file does not exist, so a missing one would carry the previous test's choices over.
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

    // Almost nobody running CRT has a maintainer account, so the tab is not there until asked for.
    [Fact]
    public void The_tab_is_hidden_until_it_is_turned_on()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            Assert.False(window.MaintainerTabItem.IsVisible);

            UserSettings.EnableMaintainerTab = true;
            window.ApplyMaintainerTabVisibility();

            Assert.True(window.MaintainerTabItem.IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE SIGN-IN REACHES THE TABS THAT ASK FOR AN ADDRESS (owner request, 2026-10-01: "When I
    // am a maintainer, and I have logged in, then I want to use that email address everywhere in
    // the CRT app - e.g. for the Feedback tab or in the Draft tab"). *** Through the real window:
    // signing in on the Maintainer tab puts the account's address in the Feedback tab, read only,
    // and hands the session to the Drafts tab for its Submit dialog; signing out brings the typed
    // address back. Fails if Main stops passing it on.
    // ###########################################################################################
    [Fact]
    public void Signing_in_puts_the_accounts_address_in_the_Feedback_and_Drafts_tabs()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.ContactEmail = "typed@example.com";

            var window = new CRT.Main();
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");
            var feedbackBox = window.TabFeedback.GetControl<TextBox>("EmailTextBox");

            Assert.Equal("typed@example.com", feedbackBox.Text);

            window.TabMaintainer.UseSessionForTests(session);

            Assert.Equal("dh@example.com", feedbackBox.Text);
            Assert.True(feedbackBox.IsReadOnly);
            Assert.True(window.TabFeedback.GetControl<TextBlock>("EmailAccountNoteText").IsVisible);
            Assert.Same(session, window.TabDrafts.MaintainerAccountForTests);

            window.TabMaintainer.UseSessionForTests(null);

            Assert.Equal("typed@example.com", feedbackBox.Text);
            Assert.False(feedbackBox.IsReadOnly);
            Assert.False(window.TabFeedback.GetControl<TextBlock>("EmailAccountNoteText").IsVisible);
            Assert.Null(window.TabDrafts.MaintainerAccountForTests);
        });
    }

    // ###########################################################################################
    // *** SIGNED IN IS SIGNED IN, WHATEVER THE CONFIGURATION TAB SAYS (owner request, 2026-10-02:
    // "As long as the maintainer is logged in, then use email from that, no matter what is checked
    // in "Configuration" tab. The maintainer will need to logoff to be forgotten"). *** Turning the
    // tab off took the sign-in back until then - this test asserted the opposite. Now only signing
    // out brings the typed address back.
    // ###########################################################################################
    [Fact]
    public void Turning_the_tab_off_keeps_its_sign_in_until_signing_out()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.ContactEmail = "typed@example.com";

            var window = new CRT.Main();
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");
            window.TabMaintainer.UseSessionForTests(session);

            UserSettings.EnableMaintainerTab = false;
            window.ApplyMaintainerTabVisibility();

            Assert.False(window.MaintainerTabItem.IsVisible);
            Assert.Equal("dh@example.com", window.TabFeedback.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Same(session, window.TabDrafts.MaintainerAccountForTests);

            window.TabMaintainer.UseSessionForTests(null);

            Assert.Equal("typed@example.com", window.TabFeedback.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Null(window.TabDrafts.MaintainerAccountForTests);
        });
    }

    // ###########################################################################################
    // And at LAUNCH with the tab off: the remembered sign-in is put in place - with no request at
    // all, since nothing of the tab runs - so the Feedback tab and the Submit dialog use the
    // account. Turning the tab on later reads the badge's lists then. The session and the server are
    // the test's own (no user file, no network).
    // ###########################################################################################
    [Fact]
    public void At_launch_with_the_tab_off_the_remembered_sign_in_is_used_without_asking_the_server()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = false;
            UserSettings.ContactEmail = "typed@example.com";

            var window = new CRT.Main();
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");
            var requests = new List<string>();

            window.TabMaintainer.RecallSessionOverrideForTests = _ => session;
            window.TabMaintainer.ClientOverrideForTests = () => new ReviewApiClient(
                "https://review.invalid",
                new HttpClient(new ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler(request =>
                {
                    requests.Add(request.RequestUri!.AbsolutePath);
                    return ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler.Refused();
                })));

            // What StartAsync does once the window is up.
            typeof(CRT.Main).GetMethod("StartMaintainerBadge", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .Invoke(window, null);

            Assert.Same(session, window.TabMaintainer.SignedIn);
            Assert.Equal("dh@example.com", window.TabFeedback.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Same(session, window.TabDrafts.MaintainerAccountForTests);
            Assert.Empty(requests);

            // Ticked later: the badge's lists are read now.
            UserSettings.EnableMaintainerTab = true;
            window.ApplyMaintainerTabVisibility();
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Assert.Contains(requests, path => path.EndsWith("/queue", StringComparison.Ordinal));
        });
    }

    // ###########################################################################################
    // *** HIDDEN FOR WANT OF WORK KEEPS THE SIGN-IN (owner report, 2026-10-02). *** It took the
    // sign-in back from 2026-10-01 (a code review), and the owner's own submission - sent just
    // after rejecting the last one in the queue, the tab hidden - arrived "Sent without an account"
    // while signed in. "Hide the Maintainer tab while no work is waiting" is tidiness, not signing
    // out: the Feedback tab and the Submit dialog keep the account throughout. This test asserted
    // the opposite until then.
    // ###########################################################################################
    [Fact]
    public async Task A_tab_hidden_for_want_of_work_keeps_its_sign_in_for_the_rest_of_CRT()
    {
        await UiTest.RunAsync(async () =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.ShowMaintainerTabOnlyWhenWorkWaiting = true;
            UserSettings.ContactEmail = "typed@example.com";

            var window = new CRT.Main();
            TabMaintainer tab = window.TabMaintainer;
            var feedbackBox = window.TabFeedback.GetControl<TextBox>("EmailTextBox");
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");

            tab.UseSessionForTests(session);

            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [], false));
            await tab.ApplyBetaListAsync(new ProductionListResponse(true, []), background: true);

            Assert.False(window.MaintainerTabItem.IsVisible);
            Assert.Equal("dh@example.com", feedbackBox.Text);
            Assert.Same(session, window.TabDrafts.MaintainerAccountForTests);

            tab.ApplyQueueResponse(new ReviewQueueResponse(true,
            [
                new ReviewQueueRow(1, "Commodore/C64/250407", "pending", "One", "c@example.com", null, AwaitsYou: true)
            ], false));

            Assert.True(window.MaintainerTabItem.IsVisible);
            Assert.Equal("dh@example.com", feedbackBox.Text);
            Assert.Same(session, window.TabDrafts.MaintainerAccountForTests);
        });
    }

    // Beside the other contribution work - Contribute, Drafts, Maintainer - and before Configuration.
    [Fact]
    public void The_tab_sits_after_Drafts_and_before_Configuration()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();
            var items = window.MainTabControl.Items.Cast<object>().ToList();

            int maintainer = items.IndexOf(window.MaintainerTabItem);

            Assert.Equal(items.IndexOf(window.DraftsTabItem) + 1, maintainer);
            Assert.Equal(items.IndexOf(window.ConfigurationTabItem) - 1, maintainer);

            // The header is a title and a badge since 2026-09-30, not a plain string.
            Assert.Equal("Maintainer", window.MaintainerTabHeaderText.Text);
        });
    }

    // ###########################################################################################
    // *** THE TAB'S BADGE: WHAT WAITS FOR THIS ACCOUNT, IN BOTH QUEUES (owner request, 2026-09-30:
    // "it should show as a badge in the "Maintainer" tab ... until it is fully processed, including
    // if it is awaiting in "BETA to PROD" queue"). *** Drawn by Main from what the tab counts - the
    // two attention badges added up - and gone when nothing waits.
    // ###########################################################################################
    [Fact]
    public async Task The_tabs_badge_adds_up_the_submissions_and_BETA_waiting_for_you()
    {
        await UiTest.RunAsync(async () =>
        {
            UserSettings.EnableMaintainerTab = true;

            var window = new CRT.Main();
            TabMaintainer tab = window.TabMaintainer;

            Assert.Null(window.MaintainerTabBadgeForTests);

            tab.ApplyQueueResponse(new ReviewQueueResponse(true,
            [
                new ReviewQueueRow(1, "Commodore/C64/250407", "pending", "One", "c@example.com", null, AwaitsYou: true),
                new ReviewQueueRow(2, "Commodore/C128/310378", "pending", "Two", "c@example.com", null, AwaitsYou: true),
                new ReviewQueueRow(3, "Commodore/VIC-20/250403", "pending", "With the other approver", "c@example.com", null, AwaitsYou: false)
            ], false));

            Assert.Equal("2", window.MaintainerTabBadgeForTests);

            await tab.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                new ProductionSystemRow("Commodore/C64/250407", "Commodore", "C64", "250407", null, "hash", null, null, AwaitsYou: true)
            ]), background: true);

            Assert.Equal("3", window.MaintainerTabBadgeForTests);

            // No tooltip on it - the owner wants none on these (2026-09-30).
            Assert.Null(ToolTip.GetTip(window.MaintainerTabItem));

            // All processed: no badge.
            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [], false));
            await tab.ApplyBetaListAsync(new ProductionListResponse(true, []), background: true);

            Assert.Null(window.MaintainerTabBadgeForTests);
        });
    }

    // ###########################################################################################
    // The minute check runs off screen only while the badge can be SEEN - the tab turned on and
    // CRT's window not minimised. A minimised window, or the tab turned off, asks the server
    // nothing a minute (and keeps no session alive on its behalf).
    // ###########################################################################################
    [Fact]
    public void The_badge_can_be_seen_only_with_the_tab_on_and_the_window_not_minimised()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            UserSettings.EnableMaintainerTab = false;
            Assert.False(window.MaintainerTabBadgeCanBeSeen);

            UserSettings.EnableMaintainerTab = true;
            Assert.True(window.MaintainerTabBadgeCanBeSeen);

            window.WindowState = WindowState.Minimized;
            Assert.False(window.MaintainerTabBadgeCanBeSeen);
        });
    }

    // ###########################################################################################
    // *** THE SIDEBAR AND THE WORKLOG BAR GO WHILE THE TAB IS SELECTED, AND COME BACK AS THEY
    // WERE. *** Neither has anything to do with reviewing, and the screens need the width. The
    // width put back is the one there before - not the saved setting, not a default - and the
    // saved setting is never overwritten with the collapsed 0.
    // ###########################################################################################
    [Fact]
    public void Selecting_the_tab_collapses_the_sidebar_and_worklog_bar_and_leaving_it_restores_them()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.EnableWorklog = true;
            UserSettings.LeftPanelWidth = 260;

            var window = new CRT.Main();
            window.RootGrid.ColumnDefinitions[0].Width = new GridLength(237);

            window.MainTabControl.SelectedItem = window.MaintainerTabItem;

            Assert.True(window.IsMaintainerLayoutActive);
            Assert.Equal(0, window.RootGrid.ColumnDefinitions[0].Width.Value);
            Assert.Equal(0, window.RootGrid.ColumnDefinitions[0].MinWidth);
            Assert.Equal(0, window.RootGrid.ColumnDefinitions[1].Width.Value);
            Assert.False(window.LeftPanel.IsVisible);
            Assert.False(window.MainSplitter.IsVisible);
            Assert.False(window.WorklogBar.IsVisible);

            window.MainTabControl.SelectedItem = window.SchematicsTabItem;

            Assert.False(window.IsMaintainerLayoutActive);
            Assert.Equal(new GridLength(237), window.RootGrid.ColumnDefinitions[0].Width);
            Assert.Equal(80, window.RootGrid.ColumnDefinitions[0].MinWidth);
            Assert.Equal(4, window.RootGrid.ColumnDefinitions[1].Width.Value);
            Assert.True(window.LeftPanel.IsVisible);
            Assert.True(window.MainSplitter.IsVisible);
            Assert.True(window.WorklogBar.IsVisible);

            Assert.Equal(260, UserSettings.LeftPanelWidth);
        });
    }

    // The worklog bar comes back only if it was on to begin with.
    [Fact]
    public void Leaving_the_tab_keeps_the_worklog_bar_hidden_when_the_worklog_is_off()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.EnableWorklog = false;

            var window = new CRT.Main();

            window.MainTabControl.SelectedItem = window.MaintainerTabItem;
            window.MainTabControl.SelectedItem = window.SchematicsTabItem;

            Assert.False(window.WorklogBar.IsVisible);
            Assert.True(window.LeftPanel.IsVisible);
        });
    }

    // Unticked while the tab is selected: the selection moves off it, and that alone restores the
    // layout - the window is never left with no sidebar and no Maintainer tab either.
    [Fact]
    public void Turning_the_tab_off_while_it_is_selected_moves_off_it_and_restores_the_layout()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;

            var window = new CRT.Main();
            window.MainTabControl.SelectedItem = window.MaintainerTabItem;

            UserSettings.EnableMaintainerTab = false;
            window.ApplyMaintainerTabVisibility();

            Assert.False(window.MaintainerTabItem.IsVisible);
            Assert.NotSame(window.MaintainerTabItem, window.MainTabControl.SelectedItem);
            Assert.False(window.IsMaintainerLayoutActive);
            Assert.True(window.LeftPanel.IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE COMPONENT FILTER NEVER TAKES FOCUS BACK OVER THE MAINTAINER TAB. *** The window pulls
    // focus to the component filter after every click (ShouldReturnFocusToComponentSearch). The tab
    // has a password box, a decision comment box and an invitation address, and the filter is
    // hidden with the sidebar - so it backs off whenever the tab is shown, not only with a table
    // open (the Drafts tab's rule).
    // ###########################################################################################
    [Fact]
    public void The_component_filter_does_not_take_focus_over_the_maintainer_tab()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;

            var window = new CRT.Main();

            window.MainTabControl.SelectedItem = window.SchematicsTabItem;
            Assert.True(window.ShouldReturnFocusToComponentSearch());

            window.MainTabControl.SelectedItem = window.MaintainerTabItem;
            Assert.False(window.ShouldReturnFocusToComponentSearch());
        });
    }

    // ###########################################################################################
    // *** THE TABLE'S FILTER IS REMEMBERED IN CRT'S SETTINGS (2026-09-29, as "Show changes only";
    // the colour-key pills since 2026-10-02). *** The separate application kept it in a file of its
    // own; Main now hands the tab UserSettings' MaintainerTableFilter, and every later PICK is
    // written back.
    // ###########################################################################################
    [Fact]
    public void The_tables_filter_comes_from_and_goes_back_to_CRTs_settings()
    {
        UiTest.Run(() =>
        {
            UserSettings.MaintainerTableFilter = BoardTableRowKinds.Errors | BoardTableRowKinds.Modified;

            var window = new CRT.Main();
            CRT.BoardTableEditor editor = window.TabMaintainer.TableEditorForTests;

            Assert.Equal(BoardTableRowKinds.Errors | BoardTableRowKinds.Modified, editor.FilterWanted);

            editor.TogglePill(BoardTableRowKinds.Errors);

            Assert.Equal(BoardTableRowKinds.Modified, UserSettings.MaintainerTableFilter);
        });
    }

    // ###########################################################################################
    // *** A CHANGED NAME OR ADDRESS REACHES THE REST OF CRT AT ONCE (owner request, 2026-10-03). ***
    // "My account" (the "Your account" window until 2026-10-04) hands every account the server
    // returns to the tab, which passes the new address to the Feedback tab and the Submit dialog -
    // the same path a sign-in takes - says it under the lists, and shows it in "My account" too.
    // The token is untouched: the next request still goes through.
    // ###########################################################################################
    [Fact]
    public void A_changed_account_reaches_the_feedback_tab_and_the_submit_dialog_at_once()
    {
        UiTest.Run(() =>
        {
            UserSettings.ContactEmail = "typed@example.com";

            var window = new CRT.Main();
            TabMaintainer tab = window.TabMaintainer;
            tab.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis"));

            // Signed in on the tab, which hands "My account" the session.
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(tab, null);

            tab.ApplyAccount(new AccountAnswer(7, "bench@example.com", "Dennis H", true, false, [], DateTimeOffset.UnixEpoch));

            Assert.Equal("bench@example.com", window.TabFeedback.GetControl<TextBox>("EmailTextBox").Text);
            Assert.Equal("bench@example.com", window.TabDrafts.MaintainerAccountForTests!.Email);
            Assert.Equal(new ReviewSession(tab.SignedIn!.BearerToken, tab.SignedIn.ExpiresUtc, 7, "bench@example.com", "Dennis H"), tab.SignedIn);
            Assert.Equal("token", tab.SignedIn.BearerToken);
            Assert.Equal("Dennis H (bench@example.com)", TabMaintainer.TextOf(tab.GetControl<TextBlock>("SignedInAsText")));
            Assert.Equal("Dennis H", tab.GetControl<MyAccountView>("MyAccountPanel").Session?.DisplayName);
        });
    }

    // ###########################################################################################
    // *** UNDER THE LISTS: WHO IS LOGGED IN, AND NOTHING ELSE (owner request, 2026-10-04: "only show
    // "Logged in as:<br /><b>Dennis</b> (dennis@...dk)" in the bottom-left corner"). *** The words on
    // a line of their own, then the name in bold and the address - and no "Account" or "Sign out"
    // button any more: both are under Account > "My account" (MyAccountViewTests).
    // ###########################################################################################
    [Fact]
    public void Under_the_lists_it_says_who_is_logged_in_and_carries_no_buttons()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            tab.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis"));
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(tab, null);

            StackPanel panel = tab.GetControl<StackPanel>("SignedInAsPanel");
            TextBlock[] lines = panel.Children.OfType<TextBlock>().ToArray();

            Assert.Equal(2, lines.Length);
            Assert.Equal("Logged in as:", lines[0].Text);
            Assert.Same(tab.GetControl<TextBlock>("SignedInAsText"), lines[1]);
            Assert.Equal("Dennis (dh@example.com)", TabMaintainer.TextOf(lines[1]));

            Avalonia.Controls.Documents.Run[] runs = lines[1].Inlines!.OfType<Avalonia.Controls.Documents.Run>().ToArray();
            Assert.Equal(["Dennis"], runs.Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold).Select(run => run.Text));

            Assert.Empty(panel.GetLogicalDescendants().OfType<Button>());
            Assert.Null(tab.FindControl<Button>("AccountButton"));
            Assert.Null(tab.FindControl<Button>("SignOutButton"));
        });
    }

    // ###########################################################################################
    // *** THE REMEMBERED NAME AND ADDRESS ARE READ AGAIN AT LAUNCH (2026-10-03). *** The stored
    // session keeps them as they were at sign-in; changed on another computer, this one would
    // otherwise keep sending the old address from the Feedback tab. Read once, quietly, with the
    // badge's lists - the session and the server are the test's own.
    // ###########################################################################################
    [Fact]
    public async Task At_launch_the_remembered_name_and_address_are_read_again_from_the_server()
    {
        await UiTest.RunAsync(async () =>
        {
            UserSettings.EnableMaintainerTab = true;

            var window = new CRT.Main();
            var session = new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis");
            var requests = new List<string>();

            window.TabMaintainer.RecallSessionOverrideForTests = _ => session;
            window.TabMaintainer.ClientOverrideForTests = () => new ReviewApiClient(
                "https://review.invalid",
                new HttpClient(new ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler(request =>
                {
                    requests.Add(request.RequestUri!.AbsolutePath);

                    return request.RequestUri.AbsolutePath == "/api/accounts/me"
                        ? ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler.Json(
                            """{"id":7,"email":"bench@example.com","displayName":"Dennis H","isVerified":true,"isAdministrator":false,"maintainerOf":[],"createdUtc":"2026-01-01T00:00:00+00:00"}""")
                        : ClassicRepairToolbox.Tests.Maintainer.AnsweringHttpHandler.Refused();
                })));

            await window.TabMaintainer.RestoreInBackgroundAsync();

            Assert.Single(requests, path => path == "/api/accounts/me");
            Assert.Equal("bench@example.com", window.TabMaintainer.SignedIn!.Email);
            Assert.Equal("Dennis H", window.TabMaintainer.SignedIn.DisplayName);
            Assert.Equal("bench@example.com", window.TabFeedback.GetControl<TextBox>("EmailTextBox").Text);

            // Once a launch: asked again, it is not read again.
            await window.TabMaintainer.RestoreInBackgroundAsync();
            Assert.Single(requests, path => path == "/api/accounts/me");
        });
    }
}
