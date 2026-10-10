using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// A board's SIX VIEWS on the Boards screen (owner request, 2026-10-03: "I would like all the same
// functionalities, as the 'Contributor Submissions' has ... Board data (and I should be able to do
// the same edits), Files (should not show changed files - just list all files), Contributor,
// Maintainer, Statistics" - and History, 2026-10-04: "another 'History' tab/button, after the
// 'Maintainer' button").
//
// The Board data view's save goes STRAIGHT TO BETA (owner decision, 2026-10-03: "it should go
// directly to the next queue, 'BETA > Stable'"), after the server's check of what it removes and a
// reason asked for in a dialog ("do ask for a change reason when clicking the 'Save changes'
// button") - so these pin the order (check, reason, publish), what is sent (the fingerprint the table
// was opened on, the reason, the edited rows, the removals shown), that a refusal or a cancelled
// reason keeps the change, that a published change shows BETA as it is now, and that one made but not
// published is handed on to be opened where it waits. Drawn without a server: the table through
// OpenTableForTests, the rest through the ...OverrideForTests / ...AnswerForTests seams.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardDetailViewSectionsTests
{
    private const string BoardId = "Commodore/C64/250407";

    private static BoardOverviewEntry Board(string boardId = BoardDetailViewSectionsTests.BoardId) =>
        new(boardId, "Commodore", "C64", boardId.Split('/')[2], true, true, false, true, null, null, null, 1);

    private static BoardDetailAnswer Detail(string boardId = BoardDetailViewSectionsTests.BoardId) =>
        new(BoardDetailViewSectionsTests.Board(boardId), [], [], []);

    private static SubmissionRows Rows() => new()
    {
        RevisionDate = "2026-August-21",
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
        KiCadCalibrations = [new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "board" }]
    };

    private static BoardTableAnswer Table(bool mayEdit = true, string? reason = null) =>
        new(BoardDetailViewSectionsTests.BoardId, new string('f', 64), BoardDetailViewSectionsTests.Rows(), mayEdit, reason);

    // A board on screen with its table open on BETA's board.
    private static BoardDetailView WithTable(bool mayEdit = true, string? reason = null)
    {
        var view = new BoardDetailView();
        view.ShowDetailForTests(BoardDetailViewSectionsTests.Detail());
        view.OpenTableForTests(BoardDetailViewSectionsTests.Table(mayEdit, reason));
        return view;
    }

    // One edit, as typing it would make: U8's friendly name.
    private static void EditFriendlyName(BoardDetailView view, string text)
    {
        BoardTableSheet components = view.BoardTableForTests.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        components.Rows.Single().Cells[friendly].Text = text;
    }

    // The line above the table, under the BETA / Stable switch (owner request, 2026-10-04: it was
    // the table's own line under its search box).
    private static string Status(BoardDetailView view) => view.TableNoteForTests;

    private static bool Shown(Control root, string name) => root.FindControl<Control>(name)!.IsVisible;

    // -----------------------------------------------------------------------------------
    // The views
    // -----------------------------------------------------------------------------------

    [Fact]
    public void With_nothing_chosen_there_are_no_views()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();

            Assert.False(BoardDetailViewSectionsTests.Shown(view, "BoardSectionBar"));
            Assert.False(BoardDetailViewSectionsTests.Shown(view, "SectionsPanel"));
        });
    }

    // The six, in the order named, as one joined switch - opening on Board data.
    [Fact]
    public void A_chosen_board_offers_its_six_views_and_opens_on_Board_data()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewSectionsTests.Detail());

            StackPanel bar = view.FindControl<StackPanel>("BoardSectionBar")!;

            Assert.True(bar.IsVisible);
            Assert.Equal(
                ["Board data", "Files", "Contributor", "Maintainer", "History", "Statistics"],
                bar.Children.OfType<Button>().Select(button => button.Content as string));
            Assert.All(bar.Children.OfType<Button>(), button => Assert.Contains("Segment", button.Classes));

            Assert.Equal(BoardSection.BoardData, view.ShownSection);
            Assert.Contains("Selected", view.FindControl<Button>("BoardDataSectionButton")!.Classes);
            Assert.True(BoardDetailViewSectionsTests.Shown(view, "BoardDataSection"));
            Assert.False(BoardDetailViewSectionsTests.Shown(view, "StatisticsSection"));
        });
    }

    // ###########################################################################################
    // A BOARD THE ACCOUNT DOES NOT MAINTAIN comes without addresses (owner request, 2026-10-05), and
    // its Contributor and Maintainer views say so once each - so the names alone read as the rule,
    // and a contributor with no account as one, not as somebody who gave no address. A board the
    // account maintains says nothing of the kind.
    // ###########################################################################################
    [Fact]
    public void A_board_sent_without_addresses_says_why_in_its_contributor_and_maintainer_views()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();

            BoardDetailAnswer detail = BoardDetailViewSectionsTests.Detail() with
            {
                Maintainers = [new PoolMaintainerEntry(1, "Dennis", string.Empty)],
                Contributors =
                [
                    new BoardContributorEntry(null, null, 1, 0, 0, 0, null),
                    new BoardContributorEntry(null, "Dora", 0, 1, 0, 0, null)
                ],
                AddressesHidden = true
            };

            view.ShowDetailForTests(detail);

            List<string> contributors = BoardDetailViewSectionsTests.TextsIn(view, "ContributorsSection");
            List<string> maintainers = BoardDetailViewSectionsTests.TextsIn(view, "MaintainersSection");

            Assert.Single(contributors, text => text == BoardsDisplay.AddressesHiddenLine);
            Assert.Contains("A contributor without an account", contributors);
            Assert.Contains("Dora", contributors);
            Assert.Single(maintainers, text => text == BoardsDisplay.AddressesHiddenLine);
            Assert.Contains("Dennis", maintainers);

            view.ShowDetailForTests(detail with { AddressesHidden = false });

            Assert.DoesNotContain(BoardsDisplay.AddressesHiddenLine, BoardDetailViewSectionsTests.TextsIn(view, "ContributorsSection"));
            Assert.DoesNotContain(BoardsDisplay.AddressesHiddenLine, BoardDetailViewSectionsTests.TextsIn(view, "MaintainersSection"));
        });
    }

    // Every text in one of the view's sections, top to bottom.
    private static List<string> TextsIn(BoardDetailView view, string section) =>
        Avalonia.LogicalTree.LogicalExtensions.GetLogicalDescendants(view.FindControl<StackPanel>(section)!)
            .OfType<TextBlock>()
            .Select(TabMaintainer.TextOf)
            .ToList();

    // ###########################################################################################
    // A view shows alone, and STAYS CHOSEN for the next board - looking at each board's statistics
    // in turn is the ordinary use (BoardSections' header).
    // ###########################################################################################
    [Fact]
    public async Task A_chosen_view_shows_alone_and_stays_chosen_for_the_next_board()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewSectionsTests.Detail());

            await view.ShowSectionAsync(BoardSection.Statistics);

            Assert.True(BoardDetailViewSectionsTests.Shown(view, "StatisticsSection"));
            Assert.False(BoardDetailViewSectionsTests.Shown(view, "BoardDataSection"));
            Assert.Contains("Selected", view.FindControl<Button>("StatisticsSectionButton")!.Classes);
            Assert.DoesNotContain("Selected", view.FindControl<Button>("BoardDataSectionButton")!.Classes);

            await view.ShowBoardAsync(BoardDetailViewSectionsTests.Board("Commodore/C128/310378"));

            Assert.Equal(BoardSection.Statistics, view.ShownSection);
            Assert.True(BoardDetailViewSectionsTests.Shown(view, "StatisticsSection"));
        });
    }

    // -----------------------------------------------------------------------------------
    // Board data
    // -----------------------------------------------------------------------------------

    // Coloured against BETA as opened: nothing marked until the maintainer changes something - and
    // no description box any more: the reason is asked for when Save is pressed.
    [Fact]
    public void An_editable_table_opens_unmarked_with_Save_and_no_description_box()
    {
        UiTest.Run(() =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            BoardTableEditor editor = view.BoardTableForTests;

            Assert.False(editor.IsReadOnly);
            Assert.False(editor.HasUnsavedChanges);
            Assert.Equal(0, editor.CommitAndGetDocument()!.ModifiedCount);
            Assert.Null(view.FindControl<TextBox>("ChangeDescriptionBox"));
            Assert.True(BoardDetailViewSectionsTests.Shown(editor, "SaveButton"));
            Assert.Equal(BoardSections.TableNote(BoardDetailViewSectionsTests.Table(), comparedWithStable: false), BoardDetailViewSectionsTests.Status(view));
        });
    }

    // ###########################################################################################
    // A BOARD THIS ACCOUNT MAY NOT CHANGE - or one waiting under BETA > Stable: there to look at,
    // nothing to change - no typing, no row buttons, no Save - and the server's reason said.
    // ###########################################################################################
    [Fact]
    public void A_table_this_account_may_not_change_is_read_only_and_says_why()
    {
        UiTest.Run(() =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable(mayEdit: false, reason: "Only this board's maintainers can.");
            BoardTableEditor editor = view.BoardTableForTests;

            Assert.True(editor.IsReadOnly);
            Assert.True(editor.FindControl<DataGrid>("TableGrid")!.IsReadOnly);
            Assert.False(BoardDetailViewSectionsTests.Shown(editor, "SaveButton"));
            Assert.False(BoardDetailViewSectionsTests.Shown(editor, "InsertRowBelowButton"));
            Assert.False(BoardDetailViewSectionsTests.Shown(editor, "DeleteRowButton"));
            // Why it cannot be changed is a panel of its own above the table (owner request,
            // 2026-10-09), not the ordinary line.
            Assert.Equal("Only this board's maintainers can.", view.ReadOnlyNoticeForTests);
            Assert.Equal(string.Empty, BoardDetailViewSectionsTests.Status(view));
        });
    }

    // ###########################################################################################
    // *** WHY IT CANNOT BE CHANGED, ABOVE EVERY VIEW FROM THE START (code review, 2026-10-09). *** The
    // panel was said only once BETA's table was read, so a board opened on Statistics showed none
    // until Board data had been visited - the same board looked different by the views visited. The
    // detail carries the table's answer now, so the panel is there before any table is read. An
    // older server's detail says nothing either way and leaves the panel to the table.
    // ###########################################################################################
    [Fact]
    public async Task The_board_detail_says_why_BETA_cannot_be_changed_before_any_table_is_read()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            await view.ShowSectionAsync(BoardSection.Statistics);

            BoardOverviewEntry board = BoardDetailViewSectionsTests.Board();

            view.ShowDetailForTests(new BoardDetailAnswer(board, [], [], [], MayEdit: false, MayNotEditReason: "It waits in BETA for the stable source."));

            Assert.Equal(BoardSection.Statistics, view.ShownSection);
            Assert.Equal("It waits in BETA for the stable source.", view.ReadOnlyNoticeForTests);

            // Published to stable meanwhile: the minute check's detail takes the panel away.
            view.ShowDetailForTests(new BoardDetailAnswer(board, [], [], [], MayEdit: true));
            Assert.Equal(string.Empty, view.ReadOnlyNoticeForTests);

            // An older server: nothing in the detail, so what the table said stays.
            view.OpenTableForTests(BoardDetailViewSectionsTests.Table(mayEdit: false, reason: "Only this board's maintainers can."));
            view.ShowDetailForTests(new BoardDetailAnswer(board, [], [], []));
            Assert.Equal("Only this board's maintainers can.", view.ReadOnlyNoticeForTests);
        });
    }

    // The next board's table is editable again - read-only belongs to the table it was set for.
    [Fact]
    public void Read_only_goes_with_the_table_it_was_set_for()
    {
        UiTest.Run(() =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable(mayEdit: false);

            view.OpenTableForTests(BoardDetailViewSectionsTests.Table(mayEdit: true));

            Assert.False(view.BoardTableForTests.IsReadOnly);
            Assert.True(BoardDetailViewSectionsTests.Shown(view.BoardTableForTests, "SaveButton"));
        });
    }

    // A save wired up the way the screen does it, against no server: what the check answers, the
    // reason typed, and the publish's answer - with every request and every call recorded.
    private sealed class Save
    {
        public List<BoardEditRequest> Checked { get; } = [];

        public List<BoardEditRequest> Sent { get; } = [];

        public List<IReadOnlyList<string>> AskedWith { get; } = [];

        public List<long> Opened { get; } = [];

        public int Published { get; set; }

        public static Save On(
            BoardDetailView view,
            ReviewApiResult<BoardEditCheckAnswer>? check = null,
            string? reason = "U8 is the CPU.",
            ReviewApiResult<BoardEditResult>? sent = null,
            ReviewApiResult<BoardTableAnswer>? reread = null)
        {
            var save = new Save();

            view.CheckOverrideForTests = request =>
            {
                save.Checked.Add(request);
                return Task.FromResult(check ?? ReviewApiResult<BoardEditCheckAnswer>.Ok(new BoardEditCheckAnswer([])));
            };

            view.ReasonAnswerForTests = removals =>
            {
                save.AskedWith.Add(removals);
                return reason;
            };

            view.SendOverrideForTests = request =>
            {
                save.Sent.Add(request);
                return Task.FromResult(sent ?? ReviewApiResult<BoardEditResult>.Ok(
                    new BoardEditResult(57, [], Published: true, Revision: "2026-October-03", RemovedFiles: [])));
            };

            view.ReadTableOverrideForTests = _ => Task.FromResult(reread ?? ReviewApiResult<BoardTableAnswer>.Ok(
                BoardDetailViewSectionsTests.Table(mayEdit: false, reason: OneSubmissionInBeta.NoChangeMessage(BoardDetailViewSectionsTests.BoardId))));

            view.AfterPublished = () =>
            {
                save.Published++;
                return Task.CompletedTask;
            };

            view.AfterSent = id =>
            {
                save.Opened.Add(id);
                return Task.CompletedTask;
            };

            return save;
        }
    }

    // ###########################################################################################
    // PUBLISHED (owner decision, 2026-10-03: straight to BETA): the check is asked first, then the
    // reason, then the publish carries the fingerprint the table was opened on, the reason, BETA's
    // rows with the edit - calibrations included - and the (empty) list of removals the check gave.
    // The table then shows BETA as it is NOW - read-only while the board waits under BETA > Stable,
    // what the publish did said above it and the server's reason in the panel (code review,
    // 2026-10-09: the panel says why, always) - and the lists are read again (AfterPublished).
    // Nothing is opened under Contributor Submissions.
    // ###########################################################################################
    [Fact]
    public async Task A_change_is_checked_then_published_with_its_reason_and_BETA_is_shown_as_it_is_now()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save save = Save.On(view);

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.True(await view.SendTableAsync());

            BoardEditRequest check = Assert.Single(save.Checked);
            Assert.Equal("CPU 6510", Assert.Single(check.Rows!.Components).FriendlyName);
            Assert.Empty(Assert.Single(save.AskedWith));

            BoardEditRequest sent = Assert.Single(save.Sent);
            Assert.Equal(BoardDetailViewSectionsTests.BoardId, sent.BoardId);
            Assert.Equal(new string('f', 64), sent.Fingerprint);
            Assert.Equal("U8 is the CPU.", sent.Summary);
            Assert.Equal("CPU 6510", Assert.Single(sent.Rows!.Components).FriendlyName);
            Assert.Equal("board", Assert.Single(sent.Rows.KiCadCalibrations).CadName);
            Assert.Empty(sent.ExpectedRemovals!);

            Assert.Equal(1, save.Published);
            Assert.Empty(save.Opened);
            Assert.False(view.HasUnsavedTableEdits);
            Assert.True(view.BoardTableForTests.IsReadOnly);
            Assert.Equal(
                BoardSections.Published(new BoardEditResult(57, [], Published: true, Revision: "2026-October-03")),
                BoardDetailViewSectionsTests.Status(view));
            Assert.Equal(OneSubmissionInBeta.NoChangeMessage(BoardDetailViewSectionsTests.BoardId), view.ReadOnlyNoticeForTests);
        });
    }

    // The files the check says the publish removes are what the reason dialog is shown, and what the
    // publish sends back - the server refuses any other list.
    [Fact]
    public async Task The_files_the_check_names_are_shown_with_the_reason_and_sent_back()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save save = Save.On(view, check: ReviewApiResult<BoardEditCheckAnswer>.Ok(new BoardEditCheckAnswer(["Commodore/C64/250407/manual.pdf"])));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.True(await view.SendTableAsync());

            Assert.Equal(["Commodore/C64/250407/manual.pdf"], Assert.Single(save.AskedWith));
            Assert.Equal(["Commodore/C64/250407/manual.pdf"], Assert.Single(save.Sent).ExpectedRemovals);
        });
    }

    // Cancel on the reason: nothing published, the change stays in the table.
    [Fact]
    public async Task Cancelling_the_reason_publishes_nothing_and_keeps_the_change()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save save = Save.On(view, reason: null);

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.False(await view.SendTableAsync());
            Assert.Empty(save.Sent);
            Assert.Equal(0, save.Published);
            Assert.True(view.HasUnsavedTableEdits);
        });
    }

    // ###########################################################################################
    // The CHECK refuses - the board began waiting under BETA > Stable, say: the maintainer is not
    // asked for a reason only to be refused, nothing is published, and the change stays with the
    // server's words.
    // ###########################################################################################
    [Fact]
    public async Task A_change_the_check_refuses_asks_no_reason_and_stays_in_the_table()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            string waiting = OneSubmissionInBeta.NoChangeMessage(BoardDetailViewSectionsTests.BoardId);
            Save save = Save.On(view, check: ReviewApiResult<BoardEditCheckAnswer>.Failed(ReviewApiFailure.NotPermitted, waiting));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.False(await view.SendTableAsync());
            Assert.Empty(save.AskedWith);
            Assert.Empty(save.Sent);
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal(BoardSections.NotPublished(waiting), BoardDetailViewSectionsTests.Status(view));
        });
    }

    // The PUBLISH refuses - BETA changed since, say: the change stays in the table, with the server's
    // words, and nothing is read again or opened.
    [Fact]
    public async Task A_refused_publish_stays_in_the_table_with_the_servers_words()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save save = Save.On(view, sent: ReviewApiResult<BoardEditResult>.Failed(
                ReviewApiFailure.Conflict, "BETA's board changed after you opened the table."));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.False(await view.SendTableAsync());
            Assert.Empty(save.Opened);
            Assert.Equal(0, save.Published);
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal("Not published: BETA's board changed after you opened the table.", BoardDetailViewSectionsTests.Status(view));
        });
    }

    // ###########################################################################################
    // MADE BUT NOT PUBLISHED - something changed between the check and the publish: nothing is lost.
    // The table goes back to BETA (the change lives in its submission now), says so, and the
    // submission is handed on to be opened under Contributor Submissions (AfterSent).
    //
    // *** READ AGAIN, SO IT IS READ-ONLY (code review, 2026-10-04). *** The server refuses another
    // change while that submission waits; reopening the answer the table was read from left it
    // editable, and a second change was refused only at Save. It now shows the server's answer.
    // ###########################################################################################
    [Fact]
    public async Task A_change_made_but_not_published_says_so_opens_its_submission_and_is_read_only()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            var notPublished = new BoardEditResult(58, [], NotPublishedReason: "BETA moved.");
            const string Waiting = "Your earlier change of this board is still waiting in the queue.";
            Save save = Save.On(
                view,
                sent: ReviewApiResult<BoardEditResult>.Ok(notPublished),
                reread: ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewSectionsTests.Table(mayEdit: false, reason: Waiting)));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.True(await view.SendTableAsync());
            Assert.Equal([58L], save.Opened);
            Assert.Equal(0, save.Published);
            Assert.False(view.HasUnsavedTableEdits);
            Assert.True(view.BoardTableForTests.IsReadOnly);
            Assert.Equal(BoardSections.SavedNotPublished(notPublished), BoardDetailViewSectionsTests.Status(view));
            Assert.Equal(Waiting, view.ReadOnlyNoticeForTests);
        });
    }

    // ###########################################################################################
    // Not read again: the table as it was read, but not to be edited, saying why - and the amber
    // panel says WHY in the server's own sentence for a change still waiting, naming its submission
    // (code review, 2026-10-10: it said only "You can look at this board's data, but not send a
    // change to it", the real reason left to the grey line below).
    // ###########################################################################################
    [Fact]
    public async Task A_change_made_but_not_published_whose_table_cannot_be_read_again_is_still_read_only()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            var notPublished = new BoardEditResult(58, [], NotPublishedReason: "BETA moved.");
            Save save = Save.On(
                view,
                sent: ReviewApiResult<BoardEditResult>.Ok(notPublished),
                reread: ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, "No answer."));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.True(await view.SendTableAsync());
            Assert.True(view.BoardTableForTests.HasTable);
            Assert.True(view.BoardTableForTests.IsReadOnly);
            Assert.Equal(BoardSections.SavedNotPublished(notPublished), BoardDetailViewSectionsTests.Status(view));
            Assert.Equal(BoardEditWording.AlreadyWaitingMessage(58), view.ReadOnlyNoticeForTests);
            Assert.Contains("submission #58", view.ReadOnlyNoticeForTests, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // The reason typed is kept when the publish is refused - the server's "save again to see the
    // list as it is now", say - and offered again by the next Save (code review, 2026-10-04: it was
    // typed again every time). Gone once the change has left the table.
    // ###########################################################################################
    [Fact]
    public async Task A_reason_typed_for_a_refused_publish_is_offered_again_and_forgotten_once_published()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save.On(view, sent: ReviewApiResult<BoardEditResult>.Failed(ReviewApiFailure.Conflict, "Save again to see the list as it is now."));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.False(await view.SendTableAsync());
            Assert.Null(view.ReasonOfferedForTests);

            Save.On(view);
            Assert.True(await view.SendTableAsync());
            Assert.Equal("U8 is the CPU.", view.ReasonOfferedForTests);

            // Published: the next change starts with an empty reason.
            view.OpenTableForTests(BoardDetailViewSectionsTests.Table());
            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510 again");
            Save.On(view, reason: null);
            await view.SendTableAsync();
            Assert.Null(view.ReasonOfferedForTests);
        });
    }

    // ###########################################################################################
    // One maintainer swapped for another keeps the count on the board's row, but changes who may
    // edit - so the table is read again when the board's detail names different maintainers (code
    // review, 2026-10-04).
    // ###########################################################################################
    [Fact]
    public async Task A_maintainer_swapped_for_another_reads_the_boards_table_again()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            BoardOverviewEntry board = BoardDetailViewSectionsTests.Board();
            var detail = new BoardDetailAnswer(board, [new PoolMaintainerEntry(1, "Ann", "ann@example.com")], [], []);

            view.ShowDetailForTests(detail);
            view.OpenTableForTests(BoardDetailViewSectionsTests.Table(), readAt: board);

            int reads = 0;
            view.ReadTableOverrideForTests = _ =>
            {
                reads++;
                return Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewSectionsTests.Table(mayEdit: false)));
            };

            // The same people: nothing read.
            view.ShowDetailForTests(detail);
            await view.CatchUpWithBetaForTests(board);
            Assert.Equal(0, reads);

            // Ann swapped for Bob - the count is the same.
            view.ShowDetailForTests(detail with { Maintainers = [new PoolMaintainerEntry(2, "Bob", "bob@example.com")] });
            await view.CatchUpWithBetaForTests(board);

            Assert.Equal(1, reads);
            Assert.True(view.BoardTableForTests.IsReadOnly);
        });
    }

    // ###########################################################################################
    // Published, but BETA's table cannot be read again: the table is CLOSED, never left on the board
    // as it was before the change - which would look editable and be refused - and the line says
    // what the publish did and how to read the table again.
    // ###########################################################################################
    [Fact]
    public async Task A_table_that_cannot_be_read_again_after_a_publish_is_closed_rather_than_left_stale()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save save = Save.On(view, reread: ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, "No answer."));

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.True(await view.SendTableAsync());
            Assert.Equal(1, save.Published);
            Assert.False(view.BoardTableForTests.HasTable);
            Assert.False(view.HasUnsavedTableEdits);

            string line = view.FindControl<TextBlock>("TableLoadMessageText")!.Text ?? string.Empty;
            Assert.StartsWith(BoardSections.Published(new BoardEditResult(57, [], Published: true, Revision: "2026-October-03")), line, StringComparison.Ordinal);
            Assert.EndsWith(BoardSections.NotReadAgain("No answer."), line, StringComparison.Ordinal);
        });
    }

    // Leaving a change not published asks - in the board's own words - and Cancel keeps it.
    [Fact]
    public async Task Leaving_a_change_not_published_asks_with_the_boards_wording_and_Cancel_keeps_it()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            var asked = new List<UnsavedTableEditsPrompt>();
            UnsavedTableEditsChoice answer = UnsavedTableEditsChoice.Cancel;

            view.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return answer;
            };

            Assert.True(await view.MayLeaveTableAsync());
            Assert.Empty(asked);

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.False(await view.MayLeaveTableAsync());
            Assert.True(view.HasUnsavedTableEdits);

            answer = UnsavedTableEditsChoice.Discard;
            Assert.True(await view.MayLeaveTableAsync());

            Assert.Equal([UnsavedTableEditsPrompt.LeavingBoard, UnsavedTableEditsPrompt.LeavingBoard], asked);
        });
    }

    // Save on the leaving prompt is the button's own save: the reason is asked for there too.
    [Fact]
    public async Task Save_on_the_leaving_prompt_asks_for_the_reason_and_publishes()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();
            Save save = Save.On(view);
            view.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Save;

            BoardDetailViewSectionsTests.EditFriendlyName(view, "CPU 6510");

            Assert.True(await view.MayLeaveTableAsync());
            Assert.Single(save.AskedWith);
            Assert.Equal("U8 is the CPU.", Assert.Single(save.Sent).Summary);
        });
    }

    // Another board: nothing of the previous one's table stays to be read as this one's.
    [Fact]
    public async Task Choosing_another_board_empties_the_table()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardDetailView view = BoardDetailViewSectionsTests.WithTable();

            await view.ShowBoardAsync(BoardDetailViewSectionsTests.Board("Commodore/C128/310378"));

            Assert.False(view.BoardTableForTests.HasTable);
        });
    }

    // -----------------------------------------------------------------------------------
    // Files
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // "Just list all files": no "Show only changed files", a count of files rather than of changes,
    // and the board's own folder open with the shared folder beside it closed.
    // ###########################################################################################
    [Fact]
    public void A_boards_files_are_a_listing_opened_on_its_own_folder()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewSectionsTests.Detail());

            view.ShowFilesForTests(new BoardFilesAnswer(
                BoardDetailViewSectionsTests.BoardId,
                [
                    new BoardFileEntry("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", BoardFileChange.Unchanged, BoardFileSource.Beta),
                    new BoardFileEntry("Commodore/C64/250407/manual.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta),
                    new BoardFileEntry("Commodore/Shared files/74LS08.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta)
                ]));

            FileTreeView tree = view.FileTreeForTests;

            Assert.False(tree.FindControl<CheckBox>("OnlyChangedCheckBox")!.IsVisible);
            Assert.Equal("3 files.", tree.FindControl<TextBlock>("SummaryText")!.Text);
            Assert.Equal(BoardSections.FilesHeading, view.FindControl<TextBlock>("FilesHeadingText")!.Text);

            Assert.Equal(
                ["Commodore", "C64", "250407", "Data C64 250407 v2.0.0.xlsx", "manual.pdf", "Shared files"],
                tree.RowsForTests.Select(row => row.Name));
        });
    }
}
