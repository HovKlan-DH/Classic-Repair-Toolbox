using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The "Systems" screen's right-hand panel, SystemView (owner request, 2026-09-27: "Right-side
// will show data per system - e.g. who has contributed to it and who is set as maintainer etc.").
//
// Drawn from an answer without a server; with no client the panel asks nobody. The words are
// SystemsDisplay's and tested there - these pin what the panel puts where, and that every empty
// section SAYS it is empty.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemViewTests
{
    private static readonly DateTimeOffset Noon = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

    private static SystemOverviewEntry System() =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, true, true, "2026-September-25", null, null, 1);

    [Fact]
    public void With_nothing_chosen_the_panel_says_what_to_do()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            Assert.Equal(["Select a system"], view.TextsForTests());
        });
    }

    // ###########################################################################################
    // The order a maintainer asks in: which system and where its data is, who maintains it, who has
    // contributed and how it went, then the submissions - each with what the contributor was told.
    // ###########################################################################################
    [Fact]
    public void A_system_reads_as_its_state_then_maintainers_contributors_and_submissions()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System(),
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
                [new SystemContributorEntry("hest@mailscan.dk", null, 2, 0, 0, 0, SystemViewTests.Noon)],
                [new SystemSubmissionEntry(41, "hest@mailscan.dk", "Corrected U8.", "returned", SystemViewTests.Noon, null, "U7 is the wrong revision.")]));

            Assert.Equal(
                [
                    "Commodore / C64 / 250407",
                    "BETA ahead of production - 1 maintainer",
                    "BETA revision 2026-September-25",
                    "Maintainer",
                    "Anna (anna@example.com)",
                    "Contributor",
                    "hest@mailscan.dk",
                    "2 accepted - last sent 2026-September-25",
                    "Recent submissions",
                    "Corrected U8.",
                    "#41 - Taken back out of BETA - waiting for review again - sent 2026-September-25 - hest@mailscan.dk",
                    "Told the contributor: U7 is the wrong revision."
                ],
                view.TextsForTests());

            Assert.Equal("Commodore/C64/250407", view.ShownSystem!.SystemId);
        });
    }

    // ###########################################################################################
    // Board views (owner request, 2026-09-27): after the maintainers, before the contributors - and
    // no section at all when the server sent no counts, since "no views" would be a claim about the
    // board the server never made.
    // ###########################################################################################
    [Fact]
    public void A_systems_views_sit_between_its_maintainers_and_its_contributors()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System() with { ViewsLast30Days = 48 },
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
                [],
                [],
                Views: new BoardViewStatistics(12, 48, 310, 3, [new("DK", "Denmark", 120)])));

            IReadOnlyList<string> texts = view.TextsForTests();

            int maintainer = texts.ToList().IndexOf("Anna (anna@example.com)");
            int heading = texts.ToList().IndexOf("Views in CRT");

            Assert.True(maintainer >= 0 && heading == maintainer + 1, string.Join(" | ", texts));
            Assert.Equal(
                [
                    "Views in CRT",
                    "[12] in the last 7 days - [48] in 30 days - [310] in 12 months",
                    "Most views in the last 12 months: Denmark (120).",
                    "Not counted above: [3] from CRTs downloading BETA data in the last 30 days.",
                    "A view is this board on screen in CRT for at least 10 seconds, counted every time."
                ],
                texts.Skip(heading).Take(5));
            Assert.Contains("BETA ahead of production - 1 maintainer - 48 views in 30 days", texts);
        });
    }

    [Fact]
    public void Without_counts_from_the_server_there_is_no_views_section()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(SystemViewTests.System(), [], [], []));

            IReadOnlyList<string> texts = view.TextsForTests();

            Assert.DoesNotContain("Views in CRT", texts);
            Assert.DoesNotContain(texts, text => text.StartsWith("No views", StringComparison.Ordinal) || text.Contains(" in 30 days", StringComparison.Ordinal));
        });
    }

    // Nobody assigned, nobody contributed, nothing submitted: each section says so rather than
    // leaving a blank that reads as not loaded.
    [Fact]
    public void Every_empty_section_says_it_is_empty()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(SystemViewTests.System() with { MaintainerCount = 0 }, [], [], []));

            IReadOnlyList<string> texts = view.TextsForTests();

            Assert.Contains("Nobody maintains this system - its submissions go to the administrator.", texts);
            Assert.Contains("Nobody has contributed to this system through CRT yet.", texts);
            Assert.Contains("No submissions yet.", texts);
        });
    }

    // Signed out: nothing of the previous account's system stays.
    [Fact]
    public void Clearing_the_panel_leaves_nothing_of_the_system()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(SystemViewTests.System(), [], [], []));
            view.Clear();

            Assert.Null(view.ShownSystem);
            Assert.Equal(["Select a system"], view.TextsForTests());
        });
    }

    // ###########################################################################################
    // A system not in CRT's drop-down lists gets the placement panel above its sections (owner
    // request, 2026-09-27); a listed one names the entry CRT shows it under instead.
    // ###########################################################################################
    private static SystemListingAnswer Listing() =>
        new(
            true,
            [new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx")],
            [
                new UnlistedSystemEntry(
                    "Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, true, null,
                    new SystemPlacement("Commodore 128", "310378 Open128", string.Empty, null)),
            ]);

    private static SystemDetailAnswer Detail(SystemOverviewEntry system) => new(system, [], [], []);

    [Fact]
    public void A_system_not_in_the_lists_gets_the_placement_panel_and_a_listed_one_names_its_entry()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.UseListing(SystemViewTests.Listing());

            view.ShowDetailForTests(SystemViewTests.Detail(
                new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, null, false, true, null, null, null, 0)));

            Assert.True(view.PlacementForTests.IsVisible);
            Assert.Equal("Commodore/C128/310378 Open128", view.PlacementForTests.ShownSystemId);
            Assert.DoesNotContain(view.TextsForTests(), text => text.StartsWith("In CRT's drop-down lists", StringComparison.Ordinal));

            view.ShowDetailForTests(SystemViewTests.Detail(SystemViewTests.System()));

            Assert.False(view.PlacementForTests.IsVisible);
            Assert.Contains("In CRT's drop-down lists as Commodore 64 / 250407 (long board)", view.TextsForTests());
        });
    }

    // Signed out: the listing goes with the session - nothing of it is left for the next account.
    [Fact]
    public void Clearing_the_panel_forgets_the_listing()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.UseListing(SystemViewTests.Listing());

            view.Clear();
            view.ShowDetailForTests(SystemViewTests.Detail(
                new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, null, false, true, null, null, null, 0)));

            Assert.False(view.PlacementForTests.IsVisible);
        });
    }

    // The chosen system's own status line carries the same bold as its row in the list.
    [Fact]
    public void The_chosen_systems_status_line_has_what_is_still_to_be_done_in_bold()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(SystemViewTests.Detail(
                new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", true, false, true, true, null, null, null, 0)));

            TextBlock state = view.FindControl<TextBlock>("SystemStateText")!;
            List<Avalonia.Controls.Documents.Run> runs = state.Inlines!.OfType<Avalonia.Controls.Documents.Run>().ToList();

            Assert.Equal(
                ["not in production yet", "nobody assigned"],
                runs.Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold).Select(run => run.Text));
            Assert.Contains("In BETA - not in production yet - nobody assigned", view.TextsForTests());
        });
    }

    // -----------------------------------------------------------------------------------
    // Setting maintainers (owner request, 2026-09-27: moved here from Admin - "either select an
    // existing maintainer or invite a new maintainer via email").
    // -----------------------------------------------------------------------------------

    private static SystemDetailAnswer PoolDetail(IReadOnlyList<MaintainerInvitationEntry>? invitations = null) =>
        new(
            SystemViewTests.System(),
            [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
            [],
            [],
            invitations);

    private static readonly IReadOnlyList<ReviewAccountRow> Accounts =
    [
        new(7, "anna@example.com", "Anna", false, true, false),
        new(8, "bo@example.com", "Bo", false, true, false),
        new(9, "cy@example.com", "Cy", false, false, false),
    ];

    // The administrator's view, with every change answered by the test and remembered.
    private static (SystemView View, List<PoolAction> Sent) AdminView(IReadOnlyList<MaintainerInvitationEntry>? invitations = null)
    {
        var view = new SystemView();
        var sent = new List<PoolAction>();

        view.PoolActionOverrideForTests = action =>
        {
            sent.Add(action);
            return Task.FromResult(ReviewApiResult<string>.Ok(action.Kind == PoolActionKind.Invite ? "An invitation is on its way to x." : "Done."));
        };

        view.SetAdministrator(true);
        view.ShowDetailForTests(SystemViewTests.PoolDetail(invitations));
        view.UseAccountsForTests(SystemViewTests.Accounts);

        return (view, sent);
    }

    private static Button[] Buttons(SystemView view, string buttonClass) =>
        view.FindControl<StackPanel>("MaintainersSection")!.Children
            .OfType<Grid>()
            .SelectMany(grid => grid.Children.OfType<Button>())
            .Where(button => button.Classes.Contains(buttonClass))
            .ToArray();

    private static void Click(Button button) =>
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

    // A maintainer who is not an administrator reads the section as before - no buttons, no fields.
    [Fact]
    public void Only_an_administrator_gets_the_maintainer_controls()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.ShowDetailForTests(SystemViewTests.PoolDetail([new MaintainerInvitationEntry(3, "new@example.com", SystemViewTests.Noon, SystemViewTests.Noon.AddDays(14))]));

            Assert.False(view.FindControl<StackPanel>("MaintainerAdminPanel")!.IsVisible);
            Assert.Empty(SystemViewTests.Buttons(view, "RemoveMaintainer"));
            Assert.DoesNotContain(view.TextsForTests(), text => text.StartsWith("Invited", StringComparison.Ordinal));

            view.SetAdministrator(true);

            Assert.True(view.FindControl<StackPanel>("MaintainerAdminPanel")!.IsVisible);
            Assert.Single(SystemViewTests.Buttons(view, "RemoveMaintainer"));
            Assert.Contains("Invited: new@example.com", view.TextsForTests());
        });
    }

    // The list offers everybody not already maintaining this system; choosing one and pressing Add
    // sends that account for this system.
    [Fact]
    public async Task Adding_sends_the_chosen_account_for_this_system()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, List<PoolAction> sent) = SystemViewTests.AdminView();

            ComboBox combo = view.FindControl<ComboBox>("AddAccountCombo")!;
            Assert.Equal(2, combo.ItemCount);

            combo.SelectedIndex = 0;
            await view.AddChosenAccountAsync();

            Assert.Equal([new PoolAction(PoolActionKind.Add, "Commodore/C64/250407", AccountId: 8, Who: "Bo")], sent);
            Assert.Equal("Bo now maintains Commodore/C64/250407.", view.FindControl<TextBlock>("PoolMessageText")!.Text);
        });
    }

    // An account that cannot be granted is refused here, in the list's own words, before anything is sent.
    [Fact]
    public async Task An_account_that_cannot_be_a_maintainer_is_refused_before_anything_is_sent()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, List<PoolAction> sent) = SystemViewTests.AdminView();

            view.FindControl<ComboBox>("AddAccountCombo")!.SelectedIndex = 1;
            await view.AddChosenAccountAsync();

            Assert.Empty(sent);
            Assert.StartsWith("cy@example.com cannot be a maintainer:", view.FindControl<TextBlock>("PoolMessageText")!.Text, StringComparison.Ordinal);

            view.FindControl<ComboBox>("AddAccountCombo")!.SelectedIndex = -1;
            await view.AddChosenAccountAsync();
            Assert.Equal("Choose an account first.", view.FindControl<TextBlock>("PoolMessageText")!.Text);
        });
    }

    // ###########################################################################################
    // Inviting sends the typed address and shows the server's sentence; an address that already
    // has an account is pointed at the list instead, without a round trip.
    // ###########################################################################################
    [Fact]
    public async Task Inviting_sends_the_typed_address_and_an_existing_account_is_pointed_at_the_list()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, List<PoolAction> sent) = SystemViewTests.AdminView();
            TextBox box = view.FindControl<TextBox>("InviteEmailBox")!;

            box.Text = "BO@example.com";
            await view.InviteTypedAddressAsync();

            Assert.Empty(sent);
            Assert.Equal("bo@example.com already has an account - choose it in the list above instead.", view.FindControl<TextBlock>("PoolMessageText")!.Text);

            box.Text = "  new@example.com ";
            await view.InviteTypedAddressAsync();

            Assert.Equal([new PoolAction(PoolActionKind.Invite, "Commodore/C64/250407", Email: "new@example.com", Who: "new@example.com")], sent);
            Assert.Equal("An invitation is on its way to x.", view.FindControl<TextBlock>("PoolMessageText")!.Text);
            Assert.Equal(string.Empty, box.Text);
        });
    }

    [Fact]
    public void Remove_and_withdraw_send_the_maintainer_and_the_invitation_they_sit_beside()
    {
        UiTest.Run(() =>
        {
            (SystemView view, List<PoolAction> sent) = SystemViewTests.AdminView(
                [new MaintainerInvitationEntry(3, "new@example.com", SystemViewTests.Noon, SystemViewTests.Noon.AddDays(14))]);

            SystemViewTests.Click(Assert.Single(SystemViewTests.Buttons(view, "RemoveMaintainer")));
            SystemViewTests.Click(Assert.Single(SystemViewTests.Buttons(view, "WithdrawInvitation")));

            Assert.Equal(
                [
                    new PoolAction(PoolActionKind.Remove, "Commodore/C64/250407", AccountId: 7, Who: "Anna"),
                    new PoolAction(PoolActionKind.Withdraw, "Commodore/C64/250407", InvitationId: 3, Who: "new@example.com"),
                ],
                sent);
        });
    }

    // A refusal is the server's sentence, in red, and nothing is re-read.
    [Fact]
    public async Task A_refused_change_shows_the_servers_words()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, _) = SystemViewTests.AdminView();

            view.PoolActionOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<string>.Failed(ReviewApiFailure.Refused, "That does not look like an email address."));

            view.FindControl<TextBox>("InviteEmailBox")!.Text = "nonsense";
            await view.InviteTypedAddressAsync();

            Assert.Equal("That does not look like an email address.", view.FindControl<TextBlock>("PoolMessageText")!.Text);
            Assert.Equal("nonsense", view.FindControl<TextBox>("InviteEmailBox")!.Text);
        });
    }

    // ###########################################################################################
    // THE HISTORY REPLACES THE BARE SUBMISSION LIST when the server sends one (2026-09-27): date
    // first, newest first, as the server ordered it.
    // ###########################################################################################
    [Fact]
    public void A_systems_history_is_shown_date_first_in_the_servers_order()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System(),
                [],
                [],
                [new SystemSubmissionEntry(9, "hest@mailscan.dk", "Corrected U8.", "merged", SystemViewTests.Noon, SystemViewTests.Noon, null)],
                null,
                [
                    new SystemHistoryEntry(SystemViewTests.Noon, SystemHistoryEvents.Decided, "Anna", 9, "merged"),
                    new SystemHistoryEntry(SystemViewTests.Noon.AddDays(-1), SystemHistoryEvents.Sent, "hest@mailscan.dk", 9, "Corrected U8."),
                ]));

            IReadOnlyList<string> texts = view.TextsForTests();

            int heading = texts.ToList().IndexOf("History");
            Assert.True(heading >= 0);
            Assert.Equal(
                ["History", "2026-September-25 - #9 - Published to BETA source", "by Anna", "2026-September-24 - #9 sent", "from hest@mailscan.dk - Corrected U8."],
                texts.Skip(heading).Take(5));
            Assert.DoesNotContain("Recent submissions", texts);
        });
    }

    // *** SEVERAL MAINTAINERS PER SYSTEM (owner request, 2026-09-27) *** - one line and one Remove
    // each, the list offering everybody not already among them, and the hint saying so.
    [Fact]
    public void A_system_can_show_and_keep_adding_several_maintainers()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.SetAdministrator(true);

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System(),
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com"), new PoolMaintainerEntry(8, "Bo", "bo@example.com")],
                [],
                []));
            view.UseAccountsForTests(SystemViewTests.Accounts);

            Assert.Equal(2, SystemViewTests.Buttons(view, "RemoveMaintainer").Length);
            Assert.Contains("Maintainers (2)", view.TextsForTests());

            // Only Cy is left to add - and adding is the same button however many there are.
            Assert.Equal(1, view.FindControl<ComboBox>("AddAccountCombo")!.ItemCount);
            Assert.True(view.FindControl<Button>("AddAccountButton")!.IsEnabled);
            Assert.True(view.FindControl<TextBlock>("SeveralMaintainersHint")!.IsVisible);
        });
    }
}
