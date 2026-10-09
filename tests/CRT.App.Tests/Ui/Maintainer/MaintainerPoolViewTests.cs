using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// Account > Maintainers, MaintainerPoolView (owner request, 2026-10-04: "please make sure that the
// actual 'Send invitation' and 'Add as maintainer' gets moved to the 'Admin' tab (in top menu), as
// this is something only the admin should be able to do"). These are the Boards screen's pool tests
// of 2026-09-27, moved with the controls - plus choosing the board, which the Account screen needs
// and the Boards screen did not. Drawn without a server: the lists through UseListsForTests, every
// read of a board and every change answered by the test.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MaintainerPoolViewTests
{
    private static readonly DateTimeOffset Noon = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

    private const string C64 = "Commodore/C64/250407";
    private const string C128 = "Commodore/C128/310378";

    private static readonly IReadOnlyList<ReviewBoardRow> Boards =
    [
        new(C64, "Commodore", "C64", "250407", null, true, [new MaintainerRow(7, "Anna", "anna@example.com")]),
        new(C128, "Commodore", "C128", "310378", null, true, []),
    ];

    private static readonly IReadOnlyList<ReviewAccountRow> Accounts =
    [
        new(7, "anna@example.com", "Anna", false, true, false),
        new(8, "bo@example.com", "Bo", false, true, false),
        new(9, "cy@example.com", "Cy", false, false, false),
    ];

    private static BoardOverviewEntry Board(string id) =>
        new(id, "Commodore", id.Split('/')[1], id.Split('/')[2], true, true, false, true, null, null, null, 1);

    private static BoardDetailAnswer Detail(string id, IReadOnlyList<PoolMaintainerEntry>? maintainers = null, IReadOnlyList<MaintainerInvitationEntry>? invitations = null) =>
        new(MaintainerPoolViewTests.Board(id), maintainers ?? [new PoolMaintainerEntry(7, "Anna", "anna@example.com")], [], [], invitations ?? []);

    // The view with both lists read, the C64 chosen, and every change answered and remembered.
    private static async Task<(MaintainerPoolView View, List<PoolAction> Sent)> ChosenAsync(
        IReadOnlyList<MaintainerInvitationEntry>? invitations = null,
        IReadOnlyList<PoolMaintainerEntry>? maintainers = null)
    {
        var view = new MaintainerPoolView();
        var sent = new List<PoolAction>();

        view.PoolActionOverrideForTests = action =>
        {
            sent.Add(action);
            return Task.FromResult(ReviewApiResult<string>.Ok(action.Kind == PoolActionKind.Invite ? "An invitation is on its way to x." : "Done."));
        };

        view.PoolDetailOverrideForTests = id => Task.FromResult(MaintainerPoolViewTests.Detail(id, id == C64 ? maintainers : [], invitations));

        await view.UseListsForTests(MaintainerPoolViewTests.Boards, MaintainerPoolViewTests.Accounts);
        await view.ChooseBoardForTests(C64);

        return (view, sent);
    }

    private static Button[] Buttons(MaintainerPoolView view, string buttonClass) =>
        view.FindControl<StackPanel>("MaintainersSection")!.Children
            .OfType<Grid>()
            .SelectMany(grid => grid.Children.OfType<Button>())
            .Where(button => button.Classes.Contains(buttonClass))
            .ToArray();

    private static void Click(Button button) =>
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

    private static string PoolMessage(MaintainerPoolView view) => view.FindControl<TextBlock>("PoolMessageText")!.Text ?? string.Empty;

    // ###########################################################################################
    // Every board in the list, by name, each with how many maintain it - and nothing of a pool
    // until one is chosen.
    // ###########################################################################################
    [Fact]
    public async Task Every_board_is_offered_with_its_count_and_nothing_shows_until_one_is_chosen()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new MaintainerPoolView();
            await view.UseListsForTests(MaintainerPoolViewTests.Boards, MaintainerPoolViewTests.Accounts);

            ComboBox boards = view.FindControl<ComboBox>("BoardCombo")!;

            Assert.Equal(
                ["Commodore / C128 / 310378  -  nobody assigned", "Commodore / C64 / 250407  -  1 maintainer"],
                boards.Items.Cast<string>());
            Assert.False(view.FindControl<StackPanel>("PoolPanel")!.IsVisible);
            Assert.Null(view.ShownBoardId);
        });
    }

    // Choosing a board shows its maintainers, each with Remove, and its open invitations, each with
    // Withdraw; the list to add from offers everybody not already in the pool.
    [Fact]
    public async Task A_chosen_board_shows_its_pool_with_its_buttons()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, _) = await MaintainerPoolViewTests.ChosenAsync(
                [new MaintainerInvitationEntry(3, "new@example.com", MaintainerPoolViewTests.Noon, MaintainerPoolViewTests.Noon.AddDays(14))]);

            Assert.True(view.FindControl<StackPanel>("PoolPanel")!.IsVisible);
            Assert.Equal(C64, view.ShownBoardId);
            Assert.Equal(
                [
                    "Maintainer",
                    "Anna (anna@example.com)",
                    "Invited: new@example.com",
                    "Sent 2026-September-25 - the code works until 2026-October-9, and they become a maintainer when they use it"
                ],
                view.PoolTextsForTests());
            Assert.Single(MaintainerPoolViewTests.Buttons(view, "RemoveMaintainer"));
            Assert.Single(MaintainerPoolViewTests.Buttons(view, "WithdrawInvitation"));
            Assert.Equal(2, view.FindControl<ComboBox>("AddAccountCombo")!.ItemCount);
        });
    }

    // The list offers everybody not already maintaining this board; choosing one and pressing Add
    // sends that account for this board.
    [Fact]
    public async Task Adding_sends_the_chosen_account_for_the_chosen_board()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, List<PoolAction> sent) = await MaintainerPoolViewTests.ChosenAsync();

            view.FindControl<ComboBox>("AddAccountCombo")!.SelectedIndex = 0;
            await view.AddChosenAccountAsync();

            Assert.Equal([new PoolAction(PoolActionKind.Add, C64, AccountId: 8, Who: "Bo")], sent);
            Assert.Equal("Bo now maintains Commodore/C64/250407.", MaintainerPoolViewTests.PoolMessage(view));
        });
    }

    // An account that cannot be granted is refused here, in the list's own words, before anything is sent.
    [Fact]
    public async Task An_account_that_cannot_be_a_maintainer_is_refused_before_anything_is_sent()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, List<PoolAction> sent) = await MaintainerPoolViewTests.ChosenAsync();

            view.FindControl<ComboBox>("AddAccountCombo")!.SelectedIndex = 1;
            await view.AddChosenAccountAsync();

            Assert.Empty(sent);
            Assert.StartsWith("cy@example.com cannot be a maintainer:", MaintainerPoolViewTests.PoolMessage(view), StringComparison.Ordinal);

            view.FindControl<ComboBox>("AddAccountCombo")!.SelectedIndex = -1;
            await view.AddChosenAccountAsync();
            Assert.Equal("Choose an account first.", MaintainerPoolViewTests.PoolMessage(view));
        });
    }

    // ###########################################################################################
    // Inviting sends the typed address and shows the server's sentence; an address that already has
    // an account is pointed at the list instead, without a round trip.
    // ###########################################################################################
    [Fact]
    public async Task Inviting_sends_the_typed_address_and_an_existing_account_is_pointed_at_the_list()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, List<PoolAction> sent) = await MaintainerPoolViewTests.ChosenAsync();
            TextBox box = view.FindControl<TextBox>("InviteEmailBox")!;

            box.Text = "BO@example.com";
            await view.InviteTypedAddressAsync();

            Assert.Empty(sent);
            Assert.Equal("bo@example.com already has an account - choose it in the list above instead.", MaintainerPoolViewTests.PoolMessage(view));

            box.Text = "  new@example.com ";
            await view.InviteTypedAddressAsync();

            Assert.Equal([new PoolAction(PoolActionKind.Invite, C64, Email: "new@example.com", Who: "new@example.com")], sent);
            Assert.Equal("An invitation is on its way to x.", MaintainerPoolViewTests.PoolMessage(view));
            Assert.Equal(string.Empty, box.Text);
        });
    }

    [Fact]
    public async Task Remove_and_withdraw_send_the_maintainer_and_the_invitation_they_sit_beside()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, List<PoolAction> sent) = await MaintainerPoolViewTests.ChosenAsync(
                [new MaintainerInvitationEntry(3, "new@example.com", MaintainerPoolViewTests.Noon, MaintainerPoolViewTests.Noon.AddDays(14))]);

            MaintainerPoolViewTests.Click(Assert.Single(MaintainerPoolViewTests.Buttons(view, "RemoveMaintainer")));
            MaintainerPoolViewTests.Click(Assert.Single(MaintainerPoolViewTests.Buttons(view, "WithdrawInvitation")));

            Assert.Equal(
                [
                    new PoolAction(PoolActionKind.Remove, C64, AccountId: 7, Who: "Anna"),
                    new PoolAction(PoolActionKind.Withdraw, C64, InvitationId: 3, Who: "new@example.com"),
                ],
                sent);
        });
    }

    // A refusal is the server's sentence, in red, and the typed address stays to be corrected.
    [Fact]
    public async Task A_refused_change_shows_the_servers_words()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, _) = await MaintainerPoolViewTests.ChosenAsync();

            view.PoolActionOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<string>.Failed(ReviewApiFailure.Refused, "That does not look like an email address."));

            view.FindControl<TextBox>("InviteEmailBox")!.Text = "nonsense";
            await view.InviteTypedAddressAsync();

            Assert.Equal("That does not look like an email address.", MaintainerPoolViewTests.PoolMessage(view));
            Assert.Equal("nonsense", view.FindControl<TextBox>("InviteEmailBox")!.Text);
        });
    }

    // ###########################################################################################
    // *** SEVERAL MAINTAINERS PER BOARD (owner request, 2026-09-27) *** - one line and one Remove
    // each, the list offering everybody not already among them, and the hint saying so.
    // ###########################################################################################
    [Fact]
    public async Task A_board_can_show_and_keep_adding_several_maintainers()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, _) = await MaintainerPoolViewTests.ChosenAsync(
                maintainers: [new PoolMaintainerEntry(7, "Anna", "anna@example.com"), new PoolMaintainerEntry(8, "Bo", "bo@example.com")]);

            Assert.Equal(2, MaintainerPoolViewTests.Buttons(view, "RemoveMaintainer").Length);
            Assert.Contains("Maintainers (2)", view.PoolTextsForTests());

            // Only Cy is left to add - and adding is the same button however many there are.
            Assert.Equal(1, view.FindControl<ComboBox>("AddAccountCombo")!.ItemCount);
            Assert.True(view.FindControl<Button>("AddAccountButton")!.IsEnabled);
            Assert.True(view.FindControl<TextBlock>("SeveralMaintainersHint")!.IsVisible);
        });
    }

    // Another board chosen: its own pool, nothing of the first one's.
    [Fact]
    public async Task Choosing_another_board_shows_that_boards_pool()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, _) = await MaintainerPoolViewTests.ChosenAsync();

            await view.ChooseBoardForTests(C128);

            Assert.Equal(C128, view.ShownBoardId);
            Assert.Equal(["Nobody maintains this board - its submissions go to the administrator."], view.PoolTextsForTests());
            Assert.Empty(MaintainerPoolViewTests.Buttons(view, "RemoveMaintainer"));
            Assert.Equal(3, view.FindControl<ComboBox>("AddAccountCombo")!.ItemCount);
        });
    }

    // Signed out: nothing of the previous account's boards or people stays.
    [Fact]
    public async Task Clearing_leaves_nothing_of_the_lists_or_the_pool()
    {
        await UiTest.RunAsync(async () =>
        {
            (MaintainerPoolView view, _) = await MaintainerPoolViewTests.ChosenAsync();

            view.Clear();

            Assert.Null(view.ShownBoardId);
            Assert.Equal(0, view.FindControl<ComboBox>("BoardCombo")!.ItemCount);
            Assert.False(view.FindControl<StackPanel>("PoolPanel")!.IsVisible);
            Assert.Empty(view.PoolTextsForTests());
        });
    }
}
