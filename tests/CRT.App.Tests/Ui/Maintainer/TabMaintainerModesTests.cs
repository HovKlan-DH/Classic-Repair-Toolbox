using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.LogicalTree;
using Avalonia.Styling;
using Avalonia.Threading;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// THE FOUR SCREENS (owner request, 2026-09-27): Review, BETA, Boards and Account (Admin until
// 2026-10-04), chosen by the buttons at the top left - each its own list on the left and its own panel on the right, where
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

    private static ReviewQueueRow Row(long id, string board = "Commodore/C64/250407", bool? awaitsYou = true) =>
        new(id, board, "pending", $"Submission {id}.", "c@example.com", DateTimeOffset.UtcNow.AddDays(-1), false, false, awaitsYou);

    private static ProductionBoardRow Beta(string board = "Commodore/C64/250407", string hash = "hash-1", bool? awaitsYou = true) =>
        new(board, "Commodore", board.Split('/')[1], board.Split('/')[2], "2026-September-25", hash, "2026-May-14", null, awaitsYou);

    private static BoardOverviewEntry Board(string board = "Commodore/C64/250407", int maintainers = 1) =>
        new(board, "Commodore", board.Split('/')[1], board.Split('/')[2], true, true, false, true, null, null, null, maintainers);

    private static void Queue(TabMaintainer main, bool isAdministrator, params ReviewQueueRow[] rows) =>
        main.ApplyQueueResponse(new ReviewQueueResponse(CanPublish: true, Submissions: rows, IsAdministrator: isAdministrator));

    private static void Mode(TabMaintainer main, MaintainerMode mode) =>
        main.ShowModeAsync(mode).GetAwaiter().GetResult();

    private static void BetaList(TabMaintainer main, bool background, params ProductionBoardRow[] rows) =>
        main.ApplyBetaListAsync(new ProductionListResponse(true, rows), background).GetAwaiter().GetResult();

    private static bool Shown(TabMaintainer main, string name) => main.FindControl<Control>(name)!.IsVisible;

    private static void Invoke(TabMaintainer main, string method) =>
        typeof(TabMaintainer).GetMethod(method, TabMaintainerModesTests.Any)!.Invoke(main, null);

    // -----------------------------------------------------------------------------------
    // Which screen is shown
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // Each screen is its OWN list and its OWN panel, and its button says it is the one shown.
    // Account's panel is its chosen item - "My account", the first (2026-10-04).
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
                (MaintainerMode.Boards, "BoardsModeButton", "BoardsListPanel", "BoardDetailPane"),
                (MaintainerMode.Account, "AccountModeButton", "AccountList", "MyAccountPanel"),
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
    // *** ACCOUNT OPENS ON "MY ACCOUNT" (owner request, 2026-10-04: "Move the current account
    // functionality from the bottom-left corner to a new left-side entry named "My account""). ***
    // Then "Server version", every maintainer's too, and below them the administrator's own:
    // "Maintainers" (back from the Boards screen, 2026-10-04), "Order of boards", "Unused files",
    // "Rebuild checksum manifests", "Delete a board", "API usage" and - LAST, used once at go-live
    // and deleting the most - "Reset contribution data".
    //
    // The panel chosen first reads nothing and writes nothing - and only that panel is shown.
    // ###########################################################################################
    [Fact]
    public void Account_opens_on_my_account_and_lists_every_entry_in_order()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);

            Assert.True(TabMaintainerModesTests.Shown(main, "MyAccountPanel"));

            foreach (string other in TabMaintainerModesTests.AccountPanels.Where(panel => panel != "MyAccountPanel"))
                Assert.False(TabMaintainerModesTests.Shown(main, other), other);

            Assert.Equal(
                ["MyAccountItem", "ServerVersionItem", "MaintainersItem", "BoardOrderItem", "UnusedFilesItem", "RebuildManifestsItem", "DeleteBoardItem", "ApiUsageItem", "ResetDataItem"],
                main.FindControl<ListBox>("AccountList")!.Items.OfType<ListBoxItem>().Select(item => item.Name));

            main.FindControl<ListBox>("AccountList")!.SelectedItem = main.FindControl<ListBoxItem>("BoardOrderItem");

            Assert.True(TabMaintainerModesTests.Shown(main, "BoardOrderAdminView"));
            Assert.False(TabMaintainerModesTests.Shown(main, "MyAccountPanel"));
        });
    }

    // Every panel the Account screen shows on the right.
    private static readonly string[] AccountPanels =
    [
        "MyAccountPanel", "ServerVersionView", "MaintainerPoolAdminView", "BoardOrderAdminView", "UnusedFilesAdminView",
        "RebuildManifestsAdminView", "BoardDeletionAdminView", "ApiUsageAdminView", "DataResetAdminView"
    ];

    // ###########################################################################################
    // The badge's number is trusted - and so may hide the tab - only while somebody is SIGNED IN
    // and both lists were read (code review, 2026-10-01). Lists applied with no session, which is
    // what a tab shown on its own is, are not an answer.
    // ###########################################################################################
    [Fact]
    public void The_badge_is_not_known_without_a_signed_in_session_whatever_was_read()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            Assert.False(main.BadgeKnown);

            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            TabMaintainerModesTests.BetaList(main, background: false);

            Assert.False(main.BadgeKnown);
        });
    }

    // ###########################################################################################
    // "REBUILD CHECKSUM MANIFESTS" ON THE ACCOUNT SCREEN (owner request, 2026-10-01: "I need a way
    // ... to press a button, and then it will generate a new online manifest, both for production
    // and BETA"). Choosing it shows its panel and hides the other - the Account screen is one panel
    // at a time, like the rest of the tab.
    // ###########################################################################################
    [Fact]
    public void Choosing_rebuild_manifests_shows_its_panel_instead_of_the_unused_files_one()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);

            main.FindControl<ListBox>("AccountList")!.SelectedItem =
                main.FindControl<ListBoxItem>("RebuildManifestsItem");

            Assert.True(TabMaintainerModesTests.Shown(main, "RebuildManifestsAdminView"));
            Assert.False(TabMaintainerModesTests.Shown(main, "UnusedFilesAdminView"));
            Assert.False(TabMaintainerModesTests.Shown(main, "BoardDeletionAdminView"));
        });
    }

    // ###########################################################################################
    // "DELETE A BOARD" ON THE ACCOUNT SCREEN (owner request, 2026-10-03: "maybe this should be a
    // part of the 'Admin' menu, that only I do have access to"). Choosing it shows its panel alone,
    // and leaving Account hides it with the rest.
    // ###########################################################################################
    [Fact]
    public void Choosing_delete_a_board_shows_its_panel_alone()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);

            main.FindControl<ListBox>("AccountList")!.SelectedItem =
                main.FindControl<ListBoxItem>("DeleteBoardItem");

            Assert.True(TabMaintainerModesTests.Shown(main, "BoardDeletionAdminView"));
            Assert.False(TabMaintainerModesTests.Shown(main, "UnusedFilesAdminView"));
            Assert.False(TabMaintainerModesTests.Shown(main, "RebuildManifestsAdminView"));

            TabMaintainerModesTests.Mode(main, MaintainerMode.Boards);

            Assert.False(TabMaintainerModesTests.Shown(main, "BoardDeletionAdminView"));
        });
    }

    // ###########################################################################################
    // The Boards screen's Maintainer view is a LIST for everybody, the administrator included
    // (owner request, 2026-10-04: "The stuff that should be visible in here, is just the selected
    // maintainer(s)") - no Remove button, no list to add from, no invitation box.
    // ###########################################################################################
    [Fact]
    public void The_boards_panel_lists_maintainers_without_controls_even_for_an_administrator()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            BoardDetailView view = main.FindControl<BoardDetailView>("BoardDetailPane")!;

            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            view.ShowDetailForTests(new BoardDetailAnswer(
                TabMaintainerModesTests.Board(),
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
                [],
                [],
                [new MaintainerInvitationEntry(3, "new@example.com", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddDays(14))]));

            Assert.Equal(["Maintainer", "Anna (anna@example.com)"], view.SectionTextsForTests("MaintainersSection"));
            Assert.Null(view.FindControl<ComboBox>("AddAccountCombo"));
            Assert.Null(view.FindControl<TextBox>("InviteEmailBox"));
            Assert.Empty(view.FindControl<StackPanel>("MaintainersSection")!.GetLogicalDescendants().OfType<Button>());
        });
    }

    // ###########################################################################################
    // *** ACCOUNT IS EVERY MAINTAINER'S; ITS ADMINISTRATOR'S ENTRIES ARE NOT (owner request,
    // 2026-10-04: "make this tab available to all maintainers ... As admin, I should still be able to
    // see all the special admin entries - no maintainer should have access to those"). *** A
    // maintainer has "My account" and "Server version" and nothing else; the entries follow the
    // server's word on each queue answer, and an account that stops being an administrator while one
    // of them is chosen is put back on "My account" - still on the Account screen.
    // ###########################################################################################
    [Fact]
    public void The_account_screen_is_every_maintainers_and_its_administrators_entries_are_not()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            TabMaintainerModesTests.Queue(main, isAdministrator: false);

            Assert.True(TabMaintainerModesTests.Shown(main, "AccountModeButton"));
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);
            Assert.Equal(MaintainerMode.Account, main.ShownMode);
            Assert.Equal(["MyAccountItem", "ServerVersionItem"], main.AccountEntriesShownForTests);
            Assert.True(TabMaintainerModesTests.Shown(main, "MyAccountPanel"));

            // BETA and Boards are every maintainer's too.
            Assert.True(TabMaintainerModesTests.Shown(main, "BetaModeButton"));
            Assert.True(TabMaintainerModesTests.Shown(main, "BoardsModeButton"));

            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            Assert.Equal(TabMaintainerModesTests.AllAccountEntries, main.AccountEntriesShownForTests);

            main.FindControl<ListBox>("AccountList")!.SelectedItem = main.FindControl<ListBoxItem>("MaintainersItem");
            Assert.True(TabMaintainerModesTests.Shown(main, "MaintainerPoolAdminView"));

            TabMaintainerModesTests.Queue(main, isAdministrator: false);

            Assert.Equal(["MyAccountItem", "ServerVersionItem"], main.AccountEntriesShownForTests);
            Assert.Equal(MaintainerMode.Account, main.ShownMode);
            Assert.Same(main.FindControl<ListBoxItem>("MyAccountItem"), main.FindControl<ListBox>("AccountList")!.SelectedItem);
            Assert.True(TabMaintainerModesTests.Shown(main, "MyAccountPanel"));
            Assert.False(TabMaintainerModesTests.Shown(main, "MaintainerPoolAdminView"));
        });
    }

    // ###########################################################################################
    // *** AS DRAWN, NOT ONLY AS BUILT. *** The first version hid the administrator's entries with
    // IsVisible - which held in a tab never shown (the test above) while a REAL window showed every
    // one of them to a plain maintainer: the list's VirtualizingStackPanel sets IsVisible back to
    // true on each container it realises (seen in a render, 2026-10-04). So they are taken out of the
    // list, and this reads the containers the list actually drew, in a shown window - it fails
    // against the IsVisible version.
    // ###########################################################################################
    [Fact]
    public void In_a_shown_window_a_maintainer_is_drawn_only_the_two_entries_that_are_theirs()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            var window = new Window { Content = main, Width = 1200, Height = 800 };
            window.Show();

            Invoke(main, "ShowQueuePanel");
            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);
            Dispatcher.UIThread.RunJobs();

            ListBox list = main.FindControl<ListBox>("AccountList")!;

            string?[] Drawn() => list.GetRealizedContainers().Where(container => container.IsEffectivelyVisible).Select(container => container.Name).ToArray();

            Assert.Equal(["MyAccountItem", "ServerVersionItem"], Drawn());

            // Made an administrator, every entry is drawn, in the markup's order; and taken away again.
            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(TabMaintainerModesTests.AllAccountEntries, Drawn());

            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["MyAccountItem", "ServerVersionItem"], Drawn());

            window.Close();
        });
    }

    private static readonly string[] AllAccountEntries =
        ["MyAccountItem", "ServerVersionItem", "MaintainersItem", "BoardOrderItem", "UnusedFilesItem", "RebuildManifestsItem", "DeleteBoardItem", "ApiUsageItem", "ResetDataItem"];

    // ###########################################################################################
    // *** EVERY ADMINISTRATOR'S ENTRY CARRIES THE PADLOCK, AND NOTHING ELSE DOES (owner request,
    // 2026-10-04: "Mark all admin locked things ... with the Font Awesome "fa-lock" icon, so it is
    // clear which things are applicable only for the admin. No title/helper text on this icon"). ***
    //
    // Which entries are the administrator's is written out here, so a new entry has to be decided
    // one way or the other - and the AdminOnly class that hides it must agree with its padlock. No
    // tooltip on the padlock or its entry. Drawn in a shown window, so the AdminLock style has been
    // applied: Font Awesome's solid face, with the top room the padlock's outline needs
    // (FontAwesomeGlyphMetrics) - without it the top of the shackle is clipped.
    // ###########################################################################################
    [Fact]
    public void Every_administrators_entry_carries_the_padlock_with_no_tooltip_and_nothing_else_does()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            var window = new Window { Content = main, Width = 1200, Height = 800 };
            window.Show();

            Invoke(main, "ShowQueuePanel");
            TabMaintainerModesTests.Queue(main, isAdministrator: true);
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);
            Dispatcher.UIThread.RunJobs();

            string[] administrators = ["MaintainersItem", "BoardOrderItem", "UnusedFilesItem", "RebuildManifestsItem", "DeleteBoardItem", "ApiUsageItem", "ResetDataItem"];

            foreach (ListBoxItem item in main.FindControl<ListBox>("AccountList")!.Items.OfType<ListBoxItem>())
            {
                bool administrator = administrators.Contains(item.Name);
                TextBlock[] locks = item.GetLogicalDescendants().OfType<TextBlock>().Where(text => text.Classes.Contains("AdminLock")).ToArray();

                Assert.Equal(administrator, item.Classes.Contains("AdminOnly"));
                Assert.Equal(administrator ? 1 : 0, locks.Length);
                Assert.Null(ToolTip.GetTip(item));

                foreach (TextBlock padlock in locks)
                {
                    Assert.Equal("\uF023", padlock.Text);
                    Assert.Null(ToolTip.GetTip(padlock));
                    Assert.Same(Application.Current!.FindResource("FontAwesomeSolid"), padlock.FontFamily);
                    Assert.Equal(Handlers.Geometry.FontAwesomeGlyphMetrics.GetTopOverflowThicknessForText(padlock.Text, padlock.FontSize), padlock.Padding);
                    Assert.True(padlock.Padding.Top > 0, $"{item.Name}'s padlock has no room for its shackle.");
                }
            }

            window.Close();
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
    // *** REVIEW COUNTS BOARDS THAT WAIT FOR YOU. *** Two submissions to one board are one; a board
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

    // BETA counts the boards that wait for you - none shown until the list has been read.
    [Fact]
    public void The_beta_badge_counts_the_boards_waiting_for_you_once_the_list_is_read()
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

    // The discreet Boards count: hidden until the list is read, then the number of boards.
    [Fact]
    public void The_boards_badge_is_the_count_of_boards_once_known()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Boards));

            main.ApplyBoardsListAsync(new BoardOverviewAnswer(
            [
                TabMaintainerModesTests.Board("Commodore/C64/250407"),
                TabMaintainerModesTests.Board("Commodore/C128/310378"),
                TabMaintainerModesTests.Board("Commodore/VIC-20/250403")
            ]), background: true).GetAwaiter().GetResult();

            Assert.Equal("3", main.ModeBadgeForTests(MaintainerMode.Boards));
            Assert.Contains("Count", main.FindControl<Border>("BoardsBadge")!.Classes);
            Assert.Contains("Attention", main.FindControl<Border>("ReviewBadge")!.Classes);
        });
    }

    // -----------------------------------------------------------------------------------
    // The BETA list
    // -----------------------------------------------------------------------------------

    // A board this account already approved waits for the other approver - dimmed, and saying so.
    [Fact]
    public void A_beta_board_waiting_for_the_other_approver_is_dimmed_and_says_why()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            TabMaintainerModesTests.BetaList(main, background: true, TabMaintainerModesTests.Beta(awaitsYou: false));

            Assert.Equal(
                ["Commodore / C64 / 250407", "BETA 2026-September-25, stable 2026-May-14 - with the other approver"],
                main.BetaTextsForTests());

            ListBoxItem item = ((IEnumerable<ListBoxItem>)main.FindControl<ListBox>("BetaList")!.ItemsSource!).Single();
            Assert.True(((Control)item.Content!).Opacity < 1);
        });
    }

    // ###########################################################################################
    // *** A CHECK NOBODY ASKED FOR NEVER THROWS AWAY A TICK. *** The list is read every minute; the
    // "I have checked this in BETA" tick on the board shown survives it unless the board's row
    // CHANGED (BETA moved) - when the plan is worked out again and the tick, given to other
    // content, goes. Fails against a refresh that re-plans the shown board every time.
    // ###########################################################################################
    [Fact]
    public void A_background_refresh_keeps_the_tick_until_beta_moves()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ProductionBoardRow row = TabMaintainerModesTests.Beta();

            TabMaintainerModesTests.BetaList(main, background: false, row);
            main.SelectBetaRowForTests(row.BoardId);
            Dispatcher.UIThread.RunJobs();

            BetaView view = main.FindControl<BetaView>("BetaDetailView")!;
            view.ShowPlanForTests(
                new ProductionPlanView(row.BoardId, row.BetaRevision, row.BetaContentHash, false, true, null, 10, [], []),
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
    public void A_beta_board_that_left_the_list_meanwhile_says_so()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ProductionBoardRow row = TabMaintainerModesTests.Beta();

            TabMaintainerModesTests.BetaList(main, background: false, row);
            main.SelectBetaRowForTests(row.BoardId);
            Dispatcher.UIThread.RunJobs();

            TabMaintainerModesTests.BetaList(main, background: true);

            BetaView view = main.FindControl<BetaView>("BetaDetailView")!;

            Assert.Null(view.ShownRow);
            Assert.Contains("no longer waiting to go to stable", view.FindControl<TextBlock>("MessageText")!.Text, StringComparison.Ordinal);
            Assert.Equal("The stable source is up to date with BETA for every board you review.", main.FindControl<TextBlock>("BetaListMessageText")!.Text);
        });
    }

    // -----------------------------------------------------------------------------------
    // The Boards list
    // -----------------------------------------------------------------------------------

    // The ordinary "in BETA and the stable source" is not said (owner request, 2026-10-03).
    [Fact]
    public void Each_board_is_listed_by_name_with_who_maintains_it()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.ApplyBoardsListAsync(new BoardOverviewAnswer(
            [
                TabMaintainerModesTests.Board("Commodore/C64/250407", maintainers: 2)
            ]), background: true).GetAwaiter().GetResult();

            Assert.Equal(["Commodore / C64 / 250407", "2 maintainers"], main.BoardsTextsForTests());
        });
    }

    // -----------------------------------------------------------------------------------
    // Signing out
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** NOTHING OF THE PREVIOUS ACCOUNT'S STAYS. *** The next sign-in may be somebody else on the
    // same machine: every list and badge goes, the administrator's entries are hidden again, and it
    // starts on Review.
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
            main.ApplyBoardsListAsync(new BoardOverviewAnswer([TabMaintainerModesTests.Board()]), background: true).GetAwaiter().GetResult();
            TabMaintainerModesTests.Mode(main, MaintainerMode.Account);

            TabMaintainerModesTests.Invoke(main, "ResetScreens");

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(["MyAccountItem", "ServerVersionItem"], main.AccountEntriesShownForTests);
            Assert.Empty(main.BetaTextsForTests());
            Assert.Empty(main.BoardsTextsForTests());
            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Beta));
            Assert.Null(main.ModeBadgeForTests(MaintainerMode.Boards));
            Assert.Null(main.FindControl<ListBox>("AccountList")!.SelectedItem);
        });
    }

    // -----------------------------------------------------------------------------------
    // Layout
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE SCREEN TABS NEVER RUN OUT OF THE TAB, WHATEVER THE BADGES SAY. *** Two-digit badges on
    // all three counted tabs, the widest they ordinarily get: all four are shown, on one line, and
    // end inside the tab. They were buttons wrapped onto two rows of the list column until
    // 2026-10-01, and the old ones once ran over the panel beside it - reported.
    //
    // 1700 wide, not CRT's 800 minimum: this suite draws with Avalonia's stub headless font, whose
    // glyphs are much WIDER than the real one. Rendered with the real font and "12", "12" and "42",
    // the four needed about 950 wide after the owner's longer names of 2026-10-04 ("Queue: Awaiting
    // push from BETA to stable" and the rest); the last one's rename to "Account" the same day saved
    // some, but at 800 it still wraps onto a second line (rendered), and with no badges at all the
    // four fit at 800. Narrower, the last wraps, still inside the tab - the next test.
    // ###########################################################################################
    [Fact]
    public void The_screen_tabs_stay_on_one_line_inside_the_tab_with_two_digit_badges()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = TabMaintainerModesTests.WithTwoDigitBadges();
            var window = new Window { Content = main, Width = 1700, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("12", main.ModeBadgeForTests(MaintainerMode.Review));
            Assert.Equal("12", main.ModeBadgeForTests(MaintainerMode.Beta));
            Assert.Equal("42", main.ModeBadgeForTests(MaintainerMode.Boards));

            List<Rect> tabs = TabMaintainerModesTests.ScreenTabsInside(main, window);

            Assert.All(tabs, bounds => Assert.Equal(tabs[0].Y, bounds.Y, 0.5));

            window.Close();
        });
    }

    // ###########################################################################################
    // NARROWER THAN THE FOUR NEED - CRT's 800 minimum, where the real font still wraps the last one
    // with two-digit badges (rendered 2026-10-04, "Account" included):
    // every tab is still shown and ends inside the tab, wrapped onto a line of its own rather than
    // clipped or run over the edge (the WrapPanel's whole reason).
    // ###########################################################################################
    [Fact]
    public void In_a_narrow_window_the_screen_tabs_wrap_inside_the_tab()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = TabMaintainerModesTests.WithTwoDigitBadges();
            var window = new Window { Content = main, Width = 800, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            List<Rect> tabs = TabMaintainerModesTests.ScreenTabsInside(main, window);

            Assert.Contains(tabs, bounds => bounds.Y > tabs[0].Y + 0.5);

            window.Close();
        });
    }

    // Every screen tab shown and ending inside the tab, read as laid out in `window`.
    private static List<Rect> ScreenTabsInside(TabMaintainer main, Window window)
    {
        Grid queuePanel = main.FindControl<Grid>("QueuePanel")!;
        double tabRight = queuePanel.TranslatePoint(new Point(queuePanel.Bounds.Width, 0), window)!.Value.X;
        var tabs = new List<Rect>();

        foreach (string name in new[] { "BoardsModeButton", "ReviewModeButton", "BetaModeButton", "AccountModeButton" })
        {
            Button button = main.FindControl<Button>(name)!;
            Point origin = button.TranslatePoint(default, window)!.Value;
            double right = origin.X + button.Bounds.Width;

            Assert.True(button.IsEffectivelyVisible, $"{name} is not shown.");
            Assert.True(right <= tabRight + 0.5, $"{name} ends at {right}, past the tab's edge at {tabRight}.");
            tabs.Add(new Rect(origin, button.Bounds.Size));
        }

        return tabs;
    }

    // An administrator's tab with two-digit badges on all three counted screens.
    // The tab goes into a plain window for a layout pass: a tab restores a session only from
    // ReviewSessionStore, which no test points at a real file.
    private static TabMaintainer WithTwoDigitBadges()
    {
        var main = new TabMaintainer();

        Queue(main, true, Enumerable.Range(1, 12).Select(id => TabMaintainerModesTests.Row(id, $"Commodore/C{id}/1")).ToArray());
        TabMaintainerModesTests.BetaList(main, true, Enumerable.Range(1, 12).Select(id => TabMaintainerModesTests.Beta($"Commodore/C{id}/1")).ToArray());
        main.ApplyBoardsListAsync(new BoardOverviewAnswer(
            Enumerable.Range(1, 42).Select(id => TabMaintainerModesTests.Board($"Commodore/C{id}/1")).ToList()), background: true).GetAwaiter().GetResult();

        Invoke(main, "ShowQueuePanel");
        typeof(TabMaintainer).GetMethod("SetAdministrator", TabMaintainerModesTests.Any)!.Invoke(main, [true]);

        return main;
    }

    // ###########################################################################################
    // A new board waiting for a place in CRT's drop-down lists (2026-09-27): it leads the Boards
    // list, marked, and the Boards button's discreet count becomes an attention badge counting the
    // boards this account can place - since nothing it approves can be published until then.
    // ###########################################################################################
    [Fact]
    public void A_board_waiting_for_a_place_leads_the_list_and_turns_the_boards_badge_to_attention()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            var listing = new BoardListingAnswer(
                true,
                [
                    new BoardListingRow("Commodore/C64/250407", "Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"),
                    new BoardListingRow("Commodore/C128/310378", "Commodore 128", "310378", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                ],
                [
                    new UnlistedBoardEntry(
                        "Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, true, null,
                        new BoardPlacement("Commodore 128", "310378 Open128", string.Empty, null)),
                ]);

            main.ApplyBoardsListAsync(
                new BoardOverviewAnswer(
                [
                    TabMaintainerModesTests.Board("Commodore/C128/310378"),
                    TabMaintainerModesTests.Board("Commodore/C64/250407"),
                    TabMaintainerModesTests.Board("Commodore/C128/310378 Open128"),
                ]),
                background: true,
                listing).GetAwaiter().GetResult();

            IReadOnlyList<string> texts = main.BoardsTextsForTests();

            Assert.Equal("Commodore / C128 / 310378 Open128", texts[0]);
            Assert.Equal("Needs a place in the drop-down lists", texts[2]);
            Assert.Equal("Commodore / C64 / 250407", texts[3]);
            Assert.Equal("Commodore / C128 / 310378", texts[5]);

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Boards));
            Assert.Contains("Attention", main.FindControl<Border>("BoardsBadge")!.Classes);
            Assert.DoesNotContain("Count", main.FindControl<Border>("BoardsBadge")!.Classes);

            // A check that could not read the lists keeps what was known - never "nothing waits".
            main.ApplyBoardsListAsync(
                new BoardOverviewAnswer([TabMaintainerModesTests.Board("Commodore/C128/310378 Open128")]),
                background: true).GetAwaiter().GetResult();

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Boards));

            // Placed: back to the discreet count of all boards.
            main.ApplyBoardsListAsync(
                new BoardOverviewAnswer([TabMaintainerModesTests.Board("Commodore/C128/310378 Open128")]),
                background: true,
                listing with
                {
                    Unlisted = [listing.Unlisted[0] with { Placement = new BoardPlacement("Commodore 128", "Open128", string.Empty, null) }]
                }).GetAwaiter().GetResult();

            Assert.Equal("1", main.ModeBadgeForTests(MaintainerMode.Boards));
            Assert.Contains("Count", main.FindControl<Border>("BoardsBadge")!.Classes);
            Assert.Equal("Placed - listed when published to BETA", main.BoardsTextsForTests()[2]);
        });
    }

    // ###########################################################################################
    // "SERVER VERSION" UNDER "Account" (owner requests, 2026-10-04: "I would like to see the server
    // version listed, so it is clear to me what has been deployed", then as a list entry of its own,
    // worded as three lines - it was one line above the account row - and "do have the "Server
    // version" entry in the list available for all maintainers"). Asked of GET /api/health when the
    // entry is chosen, shown in its panel and nowhere else, and asked again when Account is shown
    // again with it chosen - so a deploy since shows. A server that does not answer is said, never a
    // blank line. All of it as a maintainer who is NOT an administrator.
    // ###########################################################################################
    [Fact]
    public async Task The_server_version_entry_shows_the_three_lines_and_asks_again_after_a_deploy()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            string version = "4.6.0";
            int asked = 0;

            main.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(request =>
            {
                if (request.RequestUri!.AbsolutePath != "/api/health")
                    return AnsweringHttpHandler.Refused();

                asked++;
                return AnsweringHttpHandler.Json($"{{\"status\":\"ok\",\"version\":\"{version}\",\"utc\":\"2026-10-04T12:00:00+00:00\",\"apiRevision\":1}}");
            }))));

            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            await main.ShowModeAsync(MaintainerMode.Account);

            // Account opens on "My account": the version is neither shown nor asked.
            Assert.Null(main.ServerVersionLineForTests);
            Assert.Equal(0, asked);

            await TabMaintainerModesTests.ChooseAccountEntryAsync(main, "ServerVersionItem");

            string application = ClientVersionContract.ApiRevision.ToString(global::System.Globalization.CultureInfo.InvariantCulture);
            Assert.Equal($"Server version [4.6.0]\nAPI server version [1]\nAPI application version [{application}]", main.ServerVersionLineForTests);
            Assert.False(TabMaintainerModesTests.Shown(main, "MyAccountPanel"));

            // Gone with another entry, and with another screen.
            await TabMaintainerModesTests.ChooseAccountEntryAsync(main, "MyAccountItem");
            Assert.Null(main.ServerVersionLineForTests);

            await TabMaintainerModesTests.ChooseAccountEntryAsync(main, "ServerVersionItem");
            await main.ShowModeAsync(MaintainerMode.Boards);
            Assert.Null(main.ServerVersionLineForTests);

            // A deploy since: showing Account again, the entry still chosen, asks again.
            version = "4.6.1";
            int askedBefore = asked;
            await main.ShowModeAsync(MaintainerMode.Account);

            Assert.Equal(askedBefore + 1, asked);
            Assert.StartsWith("Server version [4.6.1]", main.ServerVersionLineForTests);
        });
    }

    // The line above the account row is gone: the version has its own entry now.
    [Fact]
    public void The_server_version_is_no_longer_shown_above_the_account_row()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.Null(main.FindControl<TextBlock>("ServerVersionText"));
        });
    }

    [Fact]
    public async Task A_server_that_does_not_answer_is_said_under_Server_version()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            main.UseClientForTests(new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(_ =>
                new HttpResponseMessage(global::System.Net.HttpStatusCode.BadGateway)))));

            TabMaintainerModesTests.Queue(main, isAdministrator: false);
            await main.ShowModeAsync(MaintainerMode.Account);
            await TabMaintainerModesTests.ChooseAccountEntryAsync(main, "ServerVersionItem");

            Assert.StartsWith("Server version: the server did not answer\nAPI application version [", main.ServerVersionLineForTests);
        });
    }

    // Chooses an entry in the Account list and lets what it reads finish - its SelectionChanged
    // handler is async, and the test's client answers at once.
    private static async Task ChooseAccountEntryAsync(TabMaintainer main, string item)
    {
        main.FindControl<ListBox>("AccountList")!.SelectedItem = main.FindControl<ListBoxItem>(item);

        for (int pass = 0; pass < 5; pass++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
    }

    // The buttons' names (owner request, 2026-09-27: "'Review' should be renamed to 'Contributor
    // Submissions', 'BETA' button should be renamed to 'Beta > Prod'"; and 2026-10-04: "rename the
    // "Contributor Submissions" to "Queue: Contributor submissions" and "BETA > Stable" to "Queue:
    // Awaiting push from BETA to stable" and "Admin" to "Administrator activities""; and later the
    // same day: "rename the "Administrator activities" tab to "Account""). Written out here rather
    // than read from MaintainerScreenWording, so a change to the shared words fails.
    [Fact]
    public void The_screen_buttons_are_named_as_the_owner_asked()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            string Label(string button) =>
                main.FindControl<Button>(button)!.Content switch
                {
                    StackPanel panel => panel.Children.OfType<TextBlock>().First().Text!,
                    TextBlock text => text.Text!,
                    var other => throw new InvalidOperationException($"{button} holds {other}")
                };

            Assert.Equal("Boards", Label("BoardsModeButton"));
            Assert.Equal("Queue: Contributor submissions", Label("ReviewModeButton"));
            Assert.Equal("Queue: Awaiting push from BETA to stable", Label("BetaModeButton"));
            Assert.Equal("Account", Label("AccountModeButton"));

            // The words every message naming a screen is built from are the tab strip's own.
            Assert.Equal(MaintainerScreenWording.ContributorQueue, Label("ReviewModeButton"));
            Assert.Equal(MaintainerScreenWording.BetaQueue, Label("BetaModeButton"));
            Assert.Equal(MaintainerScreenWording.Account, Label("AccountModeButton"));
        });
    }

    // Boards is the FIRST screen button (owner request, 2026-09-27); a sign-in still opens on Review.
    [Fact]
    public void Boards_is_the_first_screen_button_and_Review_still_opens_first()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Assert.Equal(
                ["BoardsModeButton", "ReviewModeButton", "BetaModeButton", "AccountModeButton"],
                main.FindControl<WrapPanel>("ModeBar")!.Children.Select(child => child.Name));

            Assert.Contains("Selected", main.FindControl<Button>("ReviewModeButton")!.Classes);
            Assert.DoesNotContain("Selected", main.FindControl<Button>("BoardsModeButton")!.Classes);
        });
    }

    // In the Boards list, the part of a board's grey line still to be done is BOLD - and only it.
    [Fact]
    public void In_the_boards_list_nobody_assigned_is_bold_and_nothing_else_is()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.ApplyBoardsListAsync(
                new BoardOverviewAnswer([TabMaintainerModesTests.Board("Commodore/C64/250407", maintainers: 0)]),
                background: true).GetAwaiter().GetResult();

            ListBoxItem item = ((IEnumerable<ListBoxItem>)main.FindControl<ListBox>("BoardsList")!.ItemsSource!).Single();
            TextBlock line = ((StackPanel)item.Content!).Children.OfType<TextBlock>().ElementAt(1);
            List<Avalonia.Controls.Documents.Run> runs = line.Inlines!.OfType<Avalonia.Controls.Documents.Run>().ToList();

            Assert.Equal("nobody assigned", string.Concat(runs.Select(run => run.Text)));
            Assert.Equal(["nobody assigned"], runs.Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold).Select(run => run.Text));
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
                    "Pushing Commodore/C64/250407 back to the queue. BETA's data is being put back as the stable source has it - please wait until it is done.",
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
