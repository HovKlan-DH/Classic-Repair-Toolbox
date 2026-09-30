using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Styling;
using Avalonia.Threading;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// THE FOUR SCREENS (owner request, 2026-09-27): Review, BETA, Systems and Admin, chosen by the
// buttons at the top left - each its own list on the left and its own panel on the right, where
// three windows used to open over the queue. TabMaintainer.Modes.cs and its three siblings.
//
// The window is BUILT, never shown: its OnOpened restores the real signed-in session and fetches
// the live queue. With no client every refresh asks nobody, so the lists are filled through the
// same Apply methods a server answer goes through. Where layout matters, its CONTENT is moved
// into a plain window and that is shown.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerModesTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(long id, string system = "Commodore/C64/250407", bool? awaitsYou = true) =>
        new(id, system, "pending", $"Submission {id}.", "c@example.com", DateTimeOffset.UtcNow.AddDays(-1), false, false, awaitsYou);

    private static ProductionSystemRow Beta(string system = "Commodore/C64/250407", string hash = "hash-1", bool? awaitsYou = true) =>
        new(system, "Commodore", system.Split('/')[1], system.Split('/')[2], "2026-September-25", hash, "2026-May-14", null, awaitsYou);

    private static SystemOverviewEntry System(string system = "Commodore/C64/250407", int maintainers = 1) =>
        new(system, "Commodore", system.Split('/')[1], system.Split('/')[2], true, true, false, true, null, null, null, maintainers);

    private static void Queue(TabMaintainer main, bool isAdministrator, params ReviewQueueRow[] rows) =>
        main.ApplyQueueResponse(new ReviewQueueResponse(CanPublish: true, Submissions: rows, IsAdministrator: isAdministrator));

    private static void Mode(TabMaintainer main, MaintainerMode mode) =>
        main.ShowModeAsync(mode).GetAwaiter().GetResult();

    private static void BetaList(TabMaintainer main, bool background, params ProductionSystemRow[] rows) =>
        main.ApplyBetaListAsync(new ProductionListResponse(true, rows), background).GetAwaiter().GetResult();

    private static bool Shown(TabMaintainer main, string name) => main.FindControl<Control>(name)!.IsVisible;

    private static void Invoke(TabMaintainer main, string method) =>
        typeof(TabMaintainer).GetMethod(method, TabMaintainerModesTests.Any)!.Invoke(main, null);

    // -----------------------------------------------------------------------------------
    // Which screen is shown
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // Each screen is its OWN list and its OWN panel, and its button says it is the one shown.
    // Admin's panel is its chosen item - "Unused files", the only one since maintainers moved to the
    // Systems screen (2026-09-27).
    // ###########################################################################################
    [Fact]
    public void Each_screen_shows_its_own_list_and_panel_and_marks_its_button()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            TabMaintainerModesTests.Queue(main, isAdministrator: true);

            (MaintainerMode Mode, string Button, string List, string Panel)[] screens =
            [
                (MaintainerMode.Beta, "BetaModeButton", "BetaListPanel", "BetaDetailView"),
                (MaintainerMode.Systems, "SystemsModeButton", "SystemsListPanel", "SystemDetailView"),
                (MaintainerMode.Admin, "AdminModeButton", "AdminList", "UnusedFilesAdminView"),
                (MaintainerMode.Review, "ReviewModeButton", "ReviewList", "ReviewPanel")
            ];

            foreach ((MaintainerMode mode, string button, string list, string panel) in screens)
            {
                TabMaintainerModesTests.Mode(main, mode);

                Assert.Equal(mode, main.ShownMode);

                foreach ((_, string otherButton, string otherList, string otherPanel) in screens)
                {
                    bool mine = otherButton == button;

                    Assert.Equal(mine, main.FindControl<Button>(otherButton)!.Classes.Contains("Selected"));
                    Assert.Equal(mine, TabMaintainerModesTests.Shown(main, otherList));
                    Assert.Equal(mine, TabMaintainerModesTests.Shown(main, otherPanel));
                }
            }
        });
    }

    // ###########################################################################################
    // *** MAINTAINERS ARE SET ON THE SYSTEMS SCREEN, NOT UNDER ADMIN (owner request, 2026-09-27: "so
    // that should be moved from 'Admin' section"). *** Admin opens on "Unused files", its one item,
    // and has no "Set maintainers" left; the Systems panel carries the controls for an
    // administrator instead.
    // ###########################################################################################
    [Fact]
    public void Admin_opens_on_unused_files_and_setting_maintainers_is_on_the_systems_screen()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            TabMaintainerModesTests.Mode(main, MaintainerMode.Admin);

            Assert.True(TabMaintainerModesTests.Shown(main, "UnusedFilesAdminView"));
            Assert.Null(main.FindControl<ListBoxItem>("SetMaintainersItem"));
            Assert.Null(main.FindControl<Control>("MaintainersAdminView"));

            Assert.Single(main.FindControl<ListBox>("AdminList")!.Items);
        });
    }

    // The Systems panel knows who is looking: an administrator gets the maintainer controls, and
    // they go again when the account stops being one.
    [Fact]
    public void The_systems_panel_follows_whether_the_account_is_an_administrator()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            SystemView view = main.FindControl<SystemView>("SystemDetailView")!;

            view.ShowDetailForTests(new SystemDetailAnswer(TabMaintainerModesTests.System(), [], [], []));

            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            Assert.True(view.FindControl<StackPanel>("MaintainerAdminPanel")!.IsVisible);

            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            Assert.False(view.FindControl<StackPanel>("MaintainerAdminPanel")!.IsVisible);
        });
    }

    // ###########################################################################################
    // *** ADMIN IS THE ADMINISTRATOR'S ONLY. *** Offered when the server says so - the server
    // refuses everybody else regardless - and an account that stops being one while on it goes
    // back to Review.
    // ###########################################################################################
    [Fact]
    public void The_admin_screen_is_offered_only_to_an_administrator()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            TabMaintainerModesTests.Queue(main, isAdministrator: false);

            Assert.False(TabMaintainerModesTests.Shown(main, "AdminModeButton"));
            TabMaintainerModesTests.Mode(main, MaintainerMode.Admin);
            Assert.Equal(MaintainerMode.Review, main.ShownMode);

            // BETA and Systems are every maintainer's.
            Assert.True(TabMaintainerModesTests.Shown(main, "BetaModeButton"));
            Assert.True(TabMaintainerModesTests.Shown(main, "SystemsModeButton"));

            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            Assert.True(TabMaintainerModesTests.Shown(main, "AdminModeButton"));

            TabMaintainerModesTests.Mode(main, MaintainerMode.Admin);
            Assert.Equal(MaintainerMode.Admin, main.ShownMode);

            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.True(TabMaintainerModesTests.Shown(main, "ReviewList"));
            Assert.False(TabMaintainerModesTests.Shown(main, "AdminList"));
        });
    }

    // ###########################################################################################
    // *** SWITCHING SCREEN HIDES, IT NEVER CLOSES. *** The submission open on Review - its table,
    // unsaved changes and all - is exactly where it was on coming back, so leaving for BETA needs
    // no question about unsaved changes.
    // ###########################################################################################
    [Fact]
    public void Coming_back_to_review_finds_the_open_submission_where_it_was()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = TabMaintainerModesTests.Row(42);

            TabMaintainerModesTests.Queue(main, isAdministrator: false, row);
            typeof(TabMaintainer).GetMethod("SelectQueueRow", TabMaintainerModesTests.Any)!.Invoke(main, [42L]);

            Assert.Equal(42, main.SelectedQueueRowForTests!.Id);

            TabMaintainerModesTests.Mode(main, MaintainerMode.Beta);
            Assert.False(TabMaintainerModesTests.Shown(main, "ReviewPanel"));

            TabMaintainerModesTests.Mode(main, MaintainerMode.Review);

            Assert.Equal(42, main.SelectedQueueRowForTests!.Id);
            Assert.True(TabMaintainerModesTests.Shown(main, "ReviewPanel"));
        });
    }

    // -----------------------------------------------------------------------------------
    // The badges
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** REVIEW COUNTS SYSTEMS THAT WAIT FOR YOU. *** Two submissions to one board are one; a board
    // whose only submission is with the other approver is none; nothing waiting is no badge at all.
    // ###########################################################################################
    [Fact]
    public void The_review_badge_counts_the_boards_waiting_for_you()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            TabMaintainerModesTests.Queue(
                main,
                isAdministrator: false,
                TabMaintainerModesTests.Row(1),
                TabMaintainerModesTests.Row(2),
                TabMaintainerModesTests.Row(3, "Commodore/C128/310378", awaitsYou: false));

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Review));

            TabMaintainerModesTests.Queue(main, isAdministrator: false, TabMaintainerModesTests.Row(3, "Commodore/C128/310378", awaitsYou: false));

            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Review));
        });
    }

    // BETA counts the systems that wait for you - none shown until the list has been read.
    [Fact]
    public void The_beta_badge_counts_the_systems_waiting_for_you_once_the_list_is_read()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Beta));

            TabMaintainerModesTests.BetaList(
                main,
                background: true,
                TabMaintainerModesTests.Beta("Commodore/C64/250407", awaitsYou: true),
                TabMaintainerModesTests.Beta("Commodore/C128/310378", awaitsYou: false),
                TabMaintainerModesTests.Beta("Commodore/VIC-20/250403", awaitsYou: null));

            Assert.Equal("2", main.ModeBadgeForTests(MaintainerMode.Beta));
        });
    }

    // The discreet Systems count: hidden until the list is read, then the number of systems.
    [Fact]
    public void The_systems_badge_is_the_count_of_systems_once_known()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Systems));

            main.ApplySystemsListAsync(new SystemOverviewAnswer(
            [
                TabMaintainerModesTests.System("Commodore/C64/250407"),
                TabMaintainerModesTests.System("Commodore/C128/310378"),
                TabMaintainerModesTests.System("Commodore/VIC-20/250403")
            ]), background: true).GetAwaiter().GetResult();

            Assert.Equal("3", main.ModeBadgeForTests(MaintainerMode.Systems));
            Assert.Contains("Count", main.FindControl<Border>("SystemsBadge")!.Classes);
            Assert.Contains("Attention", main.FindControl<Border>("ReviewBadge")!.Classes);
        });
    }

    // -----------------------------------------------------------------------------------
    // The BETA list
    // -----------------------------------------------------------------------------------

    // A system this account already approved waits for the other approver - dimmed, and saying so.
    [Fact]
    public void A_beta_system_waiting_for_the_other_approver_is_dimmed_and_says_why()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            TabMaintainerModesTests.BetaList(main, background: true, TabMaintainerModesTests.Beta(awaitsYou: false));

            Assert.Equal(
                ["Commodore / C64 / 250407", "BETA 2026-September-25, production 2026-May-14 - with the other approver"],
                main.BetaTextsForTests());

            ListBoxItem item = ((IEnumerable<ListBoxItem>)main.FindControl<ListBox>("BetaList")!.ItemsSource!).Single();
            Assert.True(((Control)item.Content!).Opacity < 1);
        });
    }

    // ###########################################################################################
    // *** A CHECK NOBODY ASKED FOR NEVER THROWS AWAY A TICK. *** The list is read every minute; the
    // "I have checked this in BETA" tick on the system shown survives it unless the system's row
    // CHANGED (BETA moved) - when the plan is worked out again and the tick, given to other
    // content, goes. Fails against a refresh that re-plans the shown system every time.
    // ###########################################################################################
    [Fact]
    public void A_background_refresh_keeps_the_tick_until_beta_moves()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ProductionSystemRow row = TabMaintainerModesTests.Beta();

            TabMaintainerModesTests.BetaList(main, background: false, row);
            main.SelectBetaRowForTests(row.SystemId);
            Dispatcher.UIThread.RunJobs();

            BetaView view = main.FindControl<BetaView>("BetaDetailView")!;
            view.ShowPlanForTests(
                new ProductionPlanView(row.SystemId, row.BetaRevision, row.BetaContentHash, false, true, null, 10, [], []),
                row);

            CheckBox tick = view.FindControl<CheckBox>("CheckedInBetaCheckBox")!;
            tick.IsChecked = true;

            // The same row again: nothing moved, so nothing is re-read and the tick stays.
            TabMaintainerModesTests.BetaList(main, background: true, TabMaintainerModesTests.Beta());
            Assert.True(tick.IsChecked);

            // BETA moved under it: worked out again, the tick gone.
            TabMaintainerModesTests.BetaList(main, background: true, TabMaintainerModesTests.Beta(hash: "hash-2"));
            Assert.False(tick.IsChecked);
            Assert.Equal("hash-2", main.SelectedBetaRowForTests!.BetaContentHash);
        });
    }

    // Published or pushed back by somebody else meanwhile: the panel empties and says so.
    [Fact]
    public void A_beta_system_that_left_the_list_meanwhile_says_so()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ProductionSystemRow row = TabMaintainerModesTests.Beta();

            TabMaintainerModesTests.BetaList(main, background: false, row);
            main.SelectBetaRowForTests(row.SystemId);
            Dispatcher.UIThread.RunJobs();

            TabMaintainerModesTests.BetaList(main, background: true);

            BetaView view = main.FindControl<BetaView>("BetaDetailView")!;

            Assert.Null(view.ShownRow);
            Assert.Contains("no longer waiting for production", view.FindControl<TextBlock>("MessageText")!.Text, StringComparison.Ordinal);
            Assert.Equal("Production is up to date with BETA for every system you review.", main.FindControl<TextBlock>("BetaListMessageText")!.Text);
        });
    }

    // -----------------------------------------------------------------------------------
    // The Systems list
    // -----------------------------------------------------------------------------------

    [Fact]
    public void Each_system_is_listed_by_name_with_where_its_data_is_and_who_maintains_it()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.ApplySystemsListAsync(new SystemOverviewAnswer(
            [
                TabMaintainerModesTests.System("Commodore/C64/250407", maintainers: 2)
            ]), background: true).GetAwaiter().GetResult();

            Assert.Equal(["Commodore / C64 / 250407", "In BETA and production - 2 maintainers"], main.SystemsTextsForTests());
        });
    }

    // -----------------------------------------------------------------------------------
    // Signing out
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** NOTHING OF THE PREVIOUS ACCOUNT'S STAYS. *** The next sign-in may be somebody else on the
    // same machine: every list and badge goes, Admin is hidden again, and it starts on Review.
    // (ResetScreens directly - ShowSignInPanel also forgets the stored session file, which is not
    // this test's to touch.)
    // ###########################################################################################
    [Fact]
    public void Signing_out_clears_every_screen_and_starts_again_on_review()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            TabMaintainerModesTests.BetaList(main, background: true, TabMaintainerModesTests.Beta());
            main.ApplySystemsListAsync(new SystemOverviewAnswer([TabMaintainerModesTests.System()]), background: true).GetAwaiter().GetResult();
            TabMaintainerModesTests.Mode(main, MaintainerMode.Admin);

            TabMaintainerModesTests.Invoke(main, "ResetScreens");

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.False(TabMaintainerModesTests.Shown(main, "AdminModeButton"));
            Assert.Empty(main.BetaTextsForTests());
            Assert.Empty(main.SystemsTextsForTests());
            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Beta));
            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Systems));
            Assert.Null(main.FindControl<ListBox>("AdminList")!.SelectedItem);
        });
    }

    // -----------------------------------------------------------------------------------
    // Layout
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE BUTTONS NEVER RUN OUT OF THEIR COLUMN, WHATEVER THE BADGES SAY. *** Two-digit badges
    // on all three counted buttons, the widest they ordinarily get: every button still ends inside
    // the column (the old buttons once ran over the panel beside it - reported).
    //
    // WHETHER THEY SHARE ONE ROW IS NOT ASSERTED HERE, deliberately: this suite draws with
    // Avalonia's stub headless font, whose glyphs are wider than the real one, so the row wraps here
    // and not on screen. That was checked by rendering with the real font - at 340 wide the Admin
    // button wrapped onto a row of its own, at 370 all four fit with "12", "12" and "42".
    // ###########################################################################################
    [Fact]
    public void The_screen_buttons_stay_inside_the_column_with_two_digit_badges()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Queue(main, true, Enumerable.Range(1, 12).Select(id => TabMaintainerModesTests.Row(id, $"Commodore/C{id}/1")).ToArray());
            TabMaintainerModesTests.BetaList(main, true, Enumerable.Range(1, 12).Select(id => TabMaintainerModesTests.Beta($"Commodore/C{id}/1")).ToArray());
            main.ApplySystemsListAsync(new SystemOverviewAnswer(
                Enumerable.Range(1, 42).Select(id => TabMaintainerModesTests.System($"Commodore/C{id}/1")).ToList()), background: true).GetAwaiter().GetResult();

            Invoke(main, "ShowQueuePanel");

            // The tab in a plain window, for a layout pass. (It moved its CONTENT into one while it
            // was a Window, whose own showing signed in to the server; a tab restores a session only
            // from ReviewSessionStore, which no test points at a real file.)

            var window = new Window { Content = main, Width = 1100, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("12", main.ModeBadgeForTests(MaintainerMode.Review));
            Assert.Equal("12", main.ModeBadgeForTests(MaintainerMode.Beta));
            Assert.Equal("42", main.ModeBadgeForTests(MaintainerMode.Systems));

            ListBox queue = main.FindControl<ListBox>("QueueList")!;
            double columnRight = queue.TranslatePoint(new Point(queue.Bounds.Width, 0), window)!.Value.X;

            foreach (string name in new[] { "ReviewModeButton", "BetaModeButton", "SystemsModeButton", "AdminModeButton" })
            {
                Button button = main.FindControl<Button>(name)!;
                double right = button.TranslatePoint(new Point(button.Bounds.Width, 0), window)!.Value.X;

                Assert.True(button.IsEffectivelyVisible, $"{name} is not shown.");
                Assert.True(right <= columnRight + 0.5, $"{name} ends at {right}, past the column's edge at {columnRight}.");
            }

            window.Close();
        });
    }

    // ###########################################################################################
    // A new system waiting for a place in CRT's drop-down lists (2026-09-27): it leads the Systems
    // list, marked, and the Systems button's discreet count becomes an attention badge counting the
    // systems this account can place - since nothing it approves can be published until then.
    // ###########################################################################################
    [Fact]
    public void A_system_waiting_for_a_place_leads_the_list_and_turns_the_systems_badge_to_attention()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            var listing = new SystemListingAnswer(
                true,
                [
                    new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                    new SystemListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                ],
                [
                    new UnlistedSystemEntry(
                        "Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, true, null,
                        new SystemPlacement("Commodore 128", "310378 Open128", string.Empty, null)),
                ]);

            main.ApplySystemsListAsync(
                new SystemOverviewAnswer(
                [
                    TabMaintainerModesTests.System("Commodore/C128/310378"),
                    TabMaintainerModesTests.System("Commodore/C64/250407"),
                    TabMaintainerModesTests.System("Commodore/C128/310378 Open128"),
                ]),
                background: true,
                listing).GetAwaiter().GetResult();

            IReadOnlyList<string> texts = main.SystemsTextsForTests();

            Assert.Equal("Commodore / C128 / 310378 Open128", texts[0]);
            Assert.Equal("Needs a place in the drop-down lists", texts[2]);
            Assert.Equal("Commodore / C64 / 250407", texts[3]);
            Assert.Equal("Commodore / C128 / 310378", texts[5]);

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Systems));
            Assert.Contains("Attention", main.FindControl<Border>("SystemsBadge")!.Classes);
            Assert.DoesNotContain("Count", main.FindControl<Border>("SystemsBadge")!.Classes);

            // A check that could not read the lists keeps what was known - never "nothing waits".
            main.ApplySystemsListAsync(
                new SystemOverviewAnswer([TabMaintainerModesTests.System("Commodore/C128/310378 Open128")]),
                background: true).GetAwaiter().GetResult();

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Systems));

            // Placed: back to the discreet count of all systems.
            main.ApplySystemsListAsync(
                new SystemOverviewAnswer([TabMaintainerModesTests.System("Commodore/C128/310378 Open128")]),
                background: true,
                listing with
                {
                    Unlisted = [listing.Unlisted[0] with { Placement = new SystemPlacement("Commodore 128", "Open128", string.Empty, null) }]
                }).GetAwaiter().GetResult();

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Systems));
            Assert.Contains("Count", main.FindControl<Border>("SystemsBadge")!.Classes);
            Assert.Equal("Placed - listed when published to BETA", main.SystemsTextsForTests()[2]);
        });
    }

    // The buttons' names (owner request, 2026-09-27: "'Review' should be renamed to 'Contributor
    // Submissions', 'BETA' button should be renamed to 'Beta > Prod'").
    [Fact]
    public void The_screen_buttons_are_named_as_the_owner_asked()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            string Label(string button) =>
                ((StackPanel)main.FindControl<Button>(button)!.Content!).Children.OfType<TextBlock>().First().Text!;

            Assert.Equal("Systems", Label("SystemsModeButton"));
            Assert.Equal("Contributor Submissions", Label("ReviewModeButton"));
            Assert.Equal("Beta > Prod", Label("BetaModeButton"));
        });
    }

    // Systems is the FIRST screen button (owner request, 2026-09-27); a sign-in still opens on Review.
    [Fact]
    public void Systems_is_the_first_screen_button_and_Review_still_opens_first()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.Equal(
                ["SystemsModeButton", "ReviewModeButton", "BetaModeButton", "AdminModeButton"],
                main.FindControl<WrapPanel>("ModeBar")!.Children.Select(child => child.Name));

            Assert.Contains("Selected", main.FindControl<Button>("ReviewModeButton")!.Classes);
            Assert.DoesNotContain("Selected", main.FindControl<Button>("SystemsModeButton")!.Classes);
        });
    }

    // In the Systems list, the part of a system's grey line still to be done is BOLD - and only it.
    [Fact]
    public void In_the_systems_list_nobody_assigned_is_bold_and_nothing_else_is()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.ApplySystemsListAsync(
                new SystemOverviewAnswer([TabMaintainerModesTests.System("Commodore/C64/250407", maintainers: 0)]),
                background: true).GetAwaiter().GetResult();

            ListBoxItem item = ((IEnumerable<ListBoxItem>)main.FindControl<ListBox>("SystemsList")!.ItemsSource!).Single();
            TextBlock line = ((StackPanel)item.Content!).Children.OfType<TextBlock>().ElementAt(1);
            List<Avalonia.Controls.Documents.Run> runs = line.Inlines!.OfType<Avalonia.Controls.Documents.Run>().ToList();

            Assert.Equal("In BETA and production - nobody assigned", string.Concat(runs.Select(run => run.Text)));
            Assert.Equal(["nobody assigned"], runs.Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold).Select(run => run.Text));
        });
    }

    // ###########################################################################################
    // "I HAVE AN INVITATION" (2026-09-27): the code, a name and a password go to the server; on
    // success the sign-in fields are filled with the invited address and the password - the
    // password reset's rule, since accepting opens no session - and the panel closes.
    // ###########################################################################################
    [Fact]
    public async Task Accepting_an_invitation_sends_the_code_name_and_password_and_fills_the_sign_in_fields()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            var sent = new List<AcceptInvitationRequest>();

            main.AcceptInvitationOverrideForTests = request =>
            {
                sent.Add(request);
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Ok(
                    new AcceptInvitationAnswer("anna@example.com", ["Commodore/C64/250407"], "Your account is ready.")));
            };

            main.FindControl<Button>("HaveInvitationButton")!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Assert.True(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);

            main.FindControl<TextBox>("InvitationCodeTextBox")!.Text = "  the-code  ";
            main.FindControl<TextBox>("InvitationNameTextBox")!.Text = "Anna";
            main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text = "correct horse battery staple";

            await main.AcceptInvitationAsync();

            Assert.Equal([new AcceptInvitationRequest("the-code", "Anna", "correct horse battery staple")], sent);
            Assert.Equal("anna@example.com", main.FindControl<TextBox>("EmailTextBox")!.Text);
            Assert.Equal("correct horse battery staple", main.FindControl<TextBox>("PasswordTextBox")!.Text);
            Assert.False(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.Equal("Your account is ready.", main.FindControl<TextBlock>("SignInMessageText")!.Text);
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationCodeTextBox")!.Text);
        });
    }

    // A refusal keeps everything typed and shows the server's words; an empty box asks nobody.
    [Fact]
    public async Task A_refused_or_incomplete_invitation_keeps_the_panel_and_says_why()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            int asked = 0;

            main.AcceptInvitationOverrideForTests = _ =>
            {
                asked++;
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Failed(
                    ReviewApiFailure.Refused, "That invitation has expired. Ask the administrator for a new one."));
            };

            main.FindControl<Button>("HaveInvitationButton")!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            await main.AcceptInvitationAsync();
            Assert.Equal(0, asked);
            Assert.Equal("Paste the code from the invitation email first.", main.FindControl<TextBlock>("SignInMessageText")!.Text);

            main.FindControl<TextBox>("InvitationCodeTextBox")!.Text = "old-code";
            main.FindControl<TextBox>("InvitationNameTextBox")!.Text = "Anna";
            main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text = "correct horse battery staple";

            await main.AcceptInvitationAsync();

            Assert.Equal(1, asked);
            Assert.Equal("That invitation has expired. Ask the administrator for a new one.", main.FindControl<TextBlock>("SignInMessageText")!.Text);
            Assert.True(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.Equal("old-code", main.FindControl<TextBox>("InvitationCodeTextBox")!.Text);
        });
    }

    // ###########################################################################################
    // *** "PLEASE WAIT" OVER THE WHOLE WINDOW (owner request, 2026-09-27: "When pushing a system back
    // from Beta then please dim everything"). *** The Beta > Prod panel finds the main window's one
    // BusyOverlay (every wait's since 2026-09-28); it is up, with the sentence, while the work
    // runs - the rest of the window faded once revealed - and down again after, even when the work
    // throws.
    // ###########################################################################################
    [Fact]
    public async Task The_window_is_dimmed_while_a_push_back_runs_and_always_comes_back()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            TabMaintainerModesTests.Invoke(main, "InitialiseScreens");

            BetaView beta = main.FindControl<BetaView>("BetaDetailView")!;
            BusyOverlay overlay = MaintainerTabHost.AddOverlay(main);

            Assert.Same(overlay, BusyOverlay.For(beta));
            Assert.False(overlay.IsVisible);

            bool seen = false;

            await BusyOverlay.HoldAsync(beta, ProductionDisplay.PushingBackWait("Commodore/C64/250407"), () =>
            {
                overlay.RevealNowForTests();
                // The overlay fades what sits beside it in its grid. In the separate application
                // that was the queue panel; in CRT's window it is everything around the overlay -
                // the tab included (here, the tab itself: MaintainerTabHost).
                seen = overlay.IsVisible && main.Opacity < 1;
                Assert.Equal(
                    "Pushing Commodore/C64/250407 back to the queue. BETA's data is being put back as production has it - please wait until it is done.",
                    overlay.Message);
                return Task.CompletedTask;
            });

            Assert.True(seen);
            Assert.False(overlay.IsVisible);
            Assert.Equal(1, main.Opacity);

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                BusyOverlay.HoldAsync(beta, "Working", () => throw new InvalidOperationException("lost")));

            Assert.False(overlay.IsVisible);
        });
    }
}
