using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// "COMPARE SOURCES" ON A BOARD'S BOARD DATA (owner request, 2026-10-09: "a "Compare sources"
// checkbox shown right after the two radio buttons, "BETA" and "Stable" and if checked, then the
// below table should show the changes as-if this was a normal board submission ... if I have
// selected the "BETA" source ... "Stable source value:" and likewise if I have chosen "Stable" then
// it would show "BETA source value:"" - and "do keep the checkbox as last state").
// BoardDetailView.Compare.cs. Drawn without a server: both tables through their read seams.
//
// The two boards: BETA has U8 renamed "PLA (82S100)" and a new U9; the stable source still has U8
// as "PLA" and a U7 that BETA no longer has. Compared, BETA's table shows U8 changed, U9 added and
// U7 deleted - and the stable source's the same three the other way round.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardDetailViewCompareTests
{
    private const string BoardId = "Commodore/C64/250407";

    private static BoardOverviewEntry Board(bool inBeta = true, bool? inStable = true) =>
        new(BoardDetailViewCompareTests.BoardId, "Commodore", "C64", "250407", inBeta, inStable, false, true,
            inBeta ? "2026-October-4" : null, inStable == true ? "2026-September-25" : null, null, 1);

    private static BoardDetailAnswer Detail(BoardOverviewEntry board) => new(board, [], [], []);

    private static ComponentEntry Component(string label, string friendly) =>
        new() { BoardLabel = label, FriendlyName = friendly, TechnicalNameOrValue = "906114-01" };

    private static readonly BoardTableAnswer Beta = new(
        BoardDetailViewCompareTests.BoardId,
        new string('f', 64),
        new SubmissionRows { Components = [Component("U8", "PLA (82S100)"), Component("U9", "CIA")] },
        MayEdit: true);

    // The stable source's answer as the server writes it: no BETA address, never editable.
    private static readonly BoardTableAnswer Stable = new(
        BoardDetailViewCompareTests.BoardId,
        new string('e', 64),
        new SubmissionRows { Components = [Component("U7", "SID"), Component("U8", "PLA")] },
        MayEdit: false,
        MayNotEditReason: "Only a publish from BETA changes it.");

    // A view reading both tables from the test, counting how often each is asked for.
    private static (BoardDetailView View, Func<int> BetaReads, Func<int> StableReads, List<bool> Remembered) View(bool compare = false)
    {
        var view = new BoardDetailView();
        int betaReads = 0;
        int stableReads = 0;
        var remembered = new List<bool>();

        view.UseRememberedComparison(compare, remembered.Add);

        view.ReadTableOverrideForTests = _ =>
        {
            betaReads++;
            return Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewCompareTests.Beta));
        };

        view.ReadStableTableOverrideForTests = _ =>
        {
            stableReads++;
            return Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewCompareTests.Stable));
        };

        return (view, () => betaReads, () => stableReads, remembered);
    }

    private static BoardTableSheet Components(BoardTableEditor editor) =>
        editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static BoardTableRow Row(BoardTableEditor editor, string label) =>
        Components(editor).Rows.Single(row => !row.IsDeleted && row.Cells[LabelColumn].Text == label);

    private static int LabelColumn => BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);

    private static int FriendlyColumn => BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);

    // ###########################################################################################
    // RIGHT AFTER THE BETA / STABLE BUTTONS, and with Board data only - the Files view has the
    // switch but no table to colour.
    // ###########################################################################################
    [Fact]
    public async Task The_box_sits_right_after_the_source_buttons_and_shows_with_Board_data_only()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, _, _) = BoardDetailViewCompareTests.View();
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));

            var bar = view.FindControl<StackPanel>("TreeSwitchBar")!;
            CheckBox box = view.CompareBoxForTests;

            Assert.Same(box, bar.Children[^1]);
            Assert.Contains(((StackPanel)bar.Children[^2]).Children.OfType<Button>(), button => button.Name == "StableTreeButton");
            Assert.Equal(BoardSections.CompareSourcesLabel, box.Content);
            Assert.True(box.IsVisible);
            Assert.True(box.IsEnabled);

            await view.ShowSectionAsync(BoardSection.Files);
            Assert.False(box.IsVisible);

            await view.ShowSectionAsync(BoardSection.BoardData);
            Assert.True(box.IsVisible);
        });
    }

    // ###########################################################################################
    // *** TICKED ON BETA: BETA'S TABLE AGAINST THE STABLE SOURCE. *** The stable source's board is
    // read for it, and every difference is coloured as in a submission - a changed cell's tooltip
    // naming "Stable source value". The tick is remembered.
    // ###########################################################################################
    [Fact]
    public async Task Ticked_on_BETA_the_table_is_coloured_against_the_stable_source()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, stableReads, remembered) = BoardDetailViewCompareTests.View();
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);

            BoardTableEditor beta = view.BoardTableForTests;

            // Unticked: coloured against BETA as opened - nothing marked, the stable source not read.
            Assert.Equal(0, Components(beta).ChangeCount);
            Assert.Equal(0, stableReads());

            await view.CompareSourcesAsync(true);

            Assert.Equal([true], remembered);
            Assert.Equal(1, stableReads());
            Assert.False(view.ShowsStableForTests);
            Assert.Same(BoardDetailViewCompareTests.Stable, view.TableComparedWithForTests);

            BoardTableSheet components = Components(beta);
            Assert.Equal(3, components.ChangeCount);
            Assert.Equal("Stable source value", components.ReplacedValueLabel);
            Assert.Equal("Stable source value: PLA", Row(beta, "U8").Cells[FriendlyColumn].ToolTip);
            Assert.Contains(components.Rows, row => row.IsDeleted && row.Cells[LabelColumn].Text == "U7");

            // Still BETA's own table: editable, and it says what is marked now.
            Assert.False(beta.IsReadOnly);
            Assert.Equal(BoardSections.TableNote(BoardDetailViewCompareTests.Beta, comparedWithStable: true), view.TableNoteForTests);
        });
    }

    // ###########################################################################################
    // *** TICKED ON STABLE: THE STABLE SOURCE'S TABLE AGAINST BETA *** - "BETA source value" - and
    // unticked, both go back to being coloured against themselves.
    // ###########################################################################################
    [Fact]
    public async Task On_Stable_the_table_is_coloured_against_BETA_and_unticking_puts_both_back()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, _, remembered) = BoardDetailViewCompareTests.View(compare: true);
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowTreeAsync(stable: true);

            BoardTableEditor stable = view.StableTableForTests;

            Assert.True(view.ShowsStableForTests);
            Assert.Same(BoardDetailViewCompareTests.Beta, view.StableComparedWithForTests);
            Assert.Equal(3, Components(stable).ChangeCount);
            Assert.Equal("BETA source value", Components(stable).ReplacedValueLabel);
            Assert.Equal("BETA source value: PLA (82S100)", Row(stable, "U8").Cells[FriendlyColumn].ToolTip);
            Assert.True(stable.IsReadOnly);

            await view.CompareSourcesAsync(false);

            Assert.Equal([false], remembered);
            Assert.Null(view.StableComparedWithForTests);
            Assert.Null(view.TableComparedWithForTests);
            Assert.Equal(0, Components(stable).ChangeCount);
            Assert.Equal(BoardSections.StableBaselineLabel, Components(stable).ReplacedValueLabel);
            Assert.Equal(0, Components(view.BoardTableForTests).ChangeCount);
        });
    }

    // Remembered ticked, the screen opens compared: both sources read once, for the board chosen.
    [Fact]
    public async Task A_remembered_tick_opens_the_table_compared()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, betaReads, stableReads, _) = BoardDetailViewCompareTests.View(compare: true);
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.True(view.CompareBoxForTests.IsChecked);
            Assert.Equal(1, betaReads());
            Assert.Equal(1, stableReads());
            Assert.Same(BoardDetailViewCompareTests.Stable, view.TableComparedWithForTests);
            Assert.Equal(3, Components(view.BoardTableForTests).ChangeCount);
        });
    }

    // ###########################################################################################
    // *** NEVER UNDER A CHANGE NOT SAVED. *** Comparing builds BETA's table again, so while it holds a
    // change the box is off, saying why, and a tick changes nothing - the change stays. Built again
    // as a save does, the box is on again.
    // ###########################################################################################
    [Fact]
    public async Task A_change_waiting_in_BETAs_table_turns_the_box_off_and_keeps_the_change()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, stableReads, remembered) = BoardDetailViewCompareTests.View();
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            view.OpenTableForTests(BoardDetailViewCompareTests.Beta);

            CheckBox box = view.CompareBoxForTests;
            Assert.True(box.IsEnabled);

            Row(view.BoardTableForTests, "U9").Cells[FriendlyColumn].Text = "CIA 6526";

            Assert.True(view.HasUnsavedTableEdits);
            Assert.False(box.IsEnabled);
            Assert.Equal(BoardSections.CompareNeedsSavedTable, ToolTip.GetTip(box));

            await view.CompareSourcesAsync(true);

            Assert.Empty(remembered);
            Assert.Equal(0, stableReads());
            Assert.Null(view.TableComparedWithForTests);
            Assert.False(box.IsChecked);
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal("CIA 6526", Row(view.BoardTableForTests, "U9").Cells[FriendlyColumn].Text);

            view.OpenTableForTests(BoardDetailViewCompareTests.Beta);

            Assert.True(box.IsEnabled);
            Assert.Equal(BoardSections.CompareSourcesTip, ToolTip.GetTip(box));
        });
    }

    // A board only one source holds cannot be compared: the box is off, saying why, and nothing of
    // the other source is asked for - even with the tick remembered.
    [Fact]
    public async Task A_board_in_one_source_only_cannot_be_compared()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, stableReads, _) = BoardDetailViewCompareTests.View(compare: true);
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board(inStable: false)));
            await view.ShowSectionAsync(BoardSection.BoardData);

            CheckBox box = view.CompareBoxForTests;

            Assert.False(box.IsEnabled);
            Assert.Equal(BoardSections.CompareNeedsBothSources, ToolTip.GetTip(box));
            Assert.Equal(0, stableReads());
            Assert.Null(view.TableComparedWithForTests);
            Assert.Equal(0, Components(view.BoardTableForTests).ChangeCount);
        });
    }

    // ###########################################################################################
    // *** A SAVE SENDS BETA'S ROWS - NEVER THE STABLE SOURCE'S. *** Compared, the stable source's
    // U7 is drawn as a deleted row in BETA's table; a change saved from it carries U8 and U9 as BETA
    // holds them, the change made, and no U7 - publishing it can neither bring U7 back nor drop
    // anything BETA has.
    // ###########################################################################################
    [Fact]
    public async Task A_change_saved_from_a_compared_table_sends_BETAs_rows_only()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, _, _) = BoardDetailViewCompareTests.View(compare: true);
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);

            var sent = new List<BoardEditRequest>();
            view.CheckOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardEditCheckAnswer>.Ok(new BoardEditCheckAnswer([])));
            view.ReasonAnswerForTests = _ => "U9 is a CIA 6526.";
            view.SendOverrideForTests = request =>
            {
                sent.Add(request);
                return Task.FromResult(ReviewApiResult<BoardEditResult>.Ok(
                    new BoardEditResult(57, [], Published: true, Revision: "2026-October-09", RemovedFiles: [])));
            };

            Row(view.BoardTableForTests, "U9").Cells[FriendlyColumn].Text = "CIA 6526";

            Assert.True(await view.SendTableAsync());

            BoardEditRequest request = Assert.Single(sent);
            Assert.Equal(
                [("U8", "PLA (82S100)"), ("U9", "CIA 6526")],
                request.Rows!.Components.Select(component => (component.BoardLabel, component.FriendlyName)));
            Assert.Equal(BoardDetailViewCompareTests.Beta.Fingerprint, request.Fingerprint);
        });
    }

    // ###########################################################################################
    // *** EACH TABLE BUILT ONCE (code review, 2026-10-09). *** Both answers are taken before either
    // table is built, so a compared board chosen builds BETA's table and the stable source's once
    // each - the stable table was built twice, once against nothing and again once BETA's came.
    // ###########################################################################################
    [Fact]
    public async Task A_compared_board_builds_each_table_once()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, betaReads, stableReads, _) = BoardDetailViewCompareTests.View(compare: true);
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.Equal((1, 1), (betaReads(), stableReads()));
            Assert.Equal((1, 1), (view.TableBuildsForTests, view.StableTableBuildsForTests));
            Assert.Same(BoardDetailViewCompareTests.Stable, view.TableComparedWithForTests);
            Assert.Same(BoardDetailViewCompareTests.Beta, view.StableComparedWithForTests);
        });
    }

    // ###########################################################################################
    // *** AFTER A PUBLISH THE STABLE TABLE FOLLOWS BETA (code review, 2026-10-09). *** BETA's table is
    // read again after the publish; the stable table compared with BETA was left compared with BETA
    // as it was, its colours and "BETA source value:" naming what BETA no longer holds. Now it is
    // built again against BETA as published - and the read-only BETA table says why it cannot be
    // changed in the panel, and that it is compared in the line under what the publish did.
    // ###########################################################################################
    [Fact]
    public async Task After_a_publish_the_stable_table_is_compared_with_BETA_as_published()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.UseRememberedComparison(true, _ => { });

            BoardTableAnswer published = BoardDetailViewCompareTests.Beta with
            {
                Fingerprint = new string('a', 64),
                Rows = new SubmissionRows { Components = [Component("U8", "PLA (82S100)"), Component("U9", "CIA 6526")] },
                MayEdit = false,
                MayNotEditReason = "It waits in BETA for the stable source."
            };

            BoardTableAnswer betaNow = BoardDetailViewCompareTests.Beta;
            view.ReadTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(betaNow));
            view.ReadStableTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewCompareTests.Stable));

            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);
            Assert.Same(BoardDetailViewCompareTests.Beta, view.StableComparedWithForTests);

            view.CheckOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardEditCheckAnswer>.Ok(new BoardEditCheckAnswer([])));
            view.ReasonAnswerForTests = _ => "U9 is a CIA 6526.";
            view.SendOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardEditResult>.Ok(
                new BoardEditResult(57, [], Published: true, Revision: "2026-October-09", RemovedFiles: [])));

            Row(view.BoardTableForTests, "U9").Cells[FriendlyColumn].Text = "CIA 6526";
            betaNow = published;

            Assert.True(await view.SendTableAsync());

            Assert.Same(published, view.StableComparedWithForTests);
            Assert.Contains(Components(view.StableTableForTests).Rows, row => row.IsDeleted && row.Cells[FriendlyColumn].Text == "CIA 6526");

            Assert.Equal("It waits in BETA for the stable source.", view.ReadOnlyNoticeForTests);
            Assert.Equal("Published to BETA as revision 2026-October-09. " + BoardSections.ComparedReadOnlyLine, view.TableNoteForTests);
        });
    }

    // ###########################################################################################
    // *** A CELL STILL BEING TYPED IN IS A CHANGE (code review, 2026-10-09). *** The grid does not
    // commit when the box takes the click, so the box saw no change, built BETA's table again and
    // threw the typing away. It commits first now: the change is kept, and the box stays off.
    // ###########################################################################################
    [Fact]
    public async Task Ticking_the_box_under_a_cell_being_typed_in_keeps_what_was_typed()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, _, stableReads, remembered) = BoardDetailViewCompareTests.View();
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            view.OpenTableForTests(BoardDetailViewCompareTests.Beta);

            var window = new Window { Content = view, Width = 1400, Height = 900 };
            window.Show();
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            BoardTableEditor editor = view.BoardTableForTests;
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");

            editor.SelectSheet(Components(editor));
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            editor.SelectCell(Row(editor, "U9"), FriendlyColumn);
            grid.Focus();
            Assert.True(grid.BeginEdit());
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            TextBox typing = Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(grid).OfType<TextBox>().Single(box => box.IsFocused);
            typing.Text = "CIA 6526";

            // Typed, not committed: the table holds no change yet.
            Assert.True(editor.IsEditingCell);
            Assert.False(view.HasUnsavedTableEdits);

            await view.CompareSourcesAsync(true);

            Assert.Empty(remembered);
            Assert.Equal(0, stableReads());
            Assert.Null(view.TableComparedWithForTests);
            Assert.True(view.HasUnsavedTableEdits);
            Assert.False(view.CompareBoxForTests.IsEnabled);
            Assert.Equal("CIA 6526", Row(editor, "U9").Cells[FriendlyColumn].Text);

            window.Close();
        });
    }

    // A board waiting under BETA > Stable, compared: its read-only table says what its colours are
    // (code review, 2026-10-09 - nothing near it did), and why it cannot be changed is the panel.
    [Fact]
    public async Task A_read_only_BETA_table_compared_says_what_its_colours_are()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.UseRememberedComparison(true, _ => { });

            BoardTableAnswer waiting = BoardDetailViewCompareTests.Beta with { MayEdit = false, MayNotEditReason = "It waits in BETA for the stable source." };
            view.ReadTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(waiting));
            view.ReadStableTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewCompareTests.Stable));

            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.True(view.BoardTableForTests.IsReadOnly);
            Assert.Equal(3, Components(view.BoardTableForTests).ChangeCount);
            Assert.Equal(BoardSections.ComparedReadOnlyLine, view.TableNoteForTests);
            Assert.Equal("It waits in BETA for the stable source.", view.ReadOnlyNoticeForTests);

            // Unticked, the line has nothing of its own to say; the panel still says why.
            await view.CompareSourcesAsync(false);

            Assert.Equal(string.Empty, view.TableNoteForTests);
            Assert.Equal("It waits in BETA for the stable source.", view.ReadOnlyNoticeForTests);
        });
    }

    // ###########################################################################################
    // *** A STABLE SOURCE THAT CANNOT BE READ: BETA'S TABLE SAYS IT IS NOT COMPARED (code review,
    // 2026-10-10). *** Ticked, with BETA on screen, a failed stable read was said only on the stable
    // half, hidden behind the switch - and BETA's table, coloured against itself, showed every row
    // white under a box still ticked: "BETA and stable are identical". Now BETA's line says why it
    // is not compared, and Board data shown again tries again - and, read, compares.
    //
    // Read-only, so the line has nothing else to say: what it says is all the failure's.
    // ###########################################################################################
    [Fact]
    public async Task Ticked_with_the_stable_source_unreadable_BETAs_table_says_it_is_not_compared()
    {
        await UiTest.RunAsync(async () =>
        {
            const string NoAnswer = "The server could not be reached.";
            var view = new BoardDetailView();
            view.UseRememberedComparison(true, _ => { });

            BoardTableAnswer waiting = BoardDetailViewCompareTests.Beta with { MayEdit = false, MayNotEditReason = "It waits in BETA for the stable source." };
            bool stableAnswers = false;

            view.ReadTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(waiting));
            view.ReadStableTableOverrideForTests = _ => Task.FromResult(stableAnswers
                ? ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewCompareTests.Stable)
                : ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, NoAnswer));

            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.False(view.ShowsStableForTests);
            Assert.True(view.CompareBoxForTests.IsChecked);
            Assert.Null(view.TableComparedWithForTests);
            Assert.Equal(0, Components(view.BoardTableForTests).ChangeCount);
            Assert.Equal(BoardSections.NotComparedLine(NoAnswer), view.TableNoteForTests);

            // Board data shown again reads what is missing - and now compares.
            stableAnswers = true;
            await view.ShowSectionAsync(BoardSection.Files);
            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.Same(BoardDetailViewCompareTests.Stable, view.TableComparedWithForTests);
            Assert.Equal(3, Components(view.BoardTableForTests).ChangeCount);
            Assert.Equal(BoardSections.ComparedReadOnlyLine, view.TableNoteForTests);
        });
    }

    // ###########################################################################################
    // The same when BETA's table is already on screen and the box is ticked: BETA's table is not read
    // again, so it is built again for the line to say so - and unticked, the line goes.
    // ###########################################################################################
    [Fact]
    public async Task Ticking_with_BETAs_table_open_and_the_stable_source_unreadable_says_it_is_not_compared()
    {
        await UiTest.RunAsync(async () =>
        {
            const string NoAnswer = "The server could not be reached.";
            var (view, betaReads, stableReads, _) = BoardDetailViewCompareTests.View();

            view.ReadStableTableOverrideForTests = _ =>
                Task.FromResult(ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, NoAnswer));

            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));
            await view.ShowSectionAsync(BoardSection.BoardData);
            Assert.DoesNotContain("Not compared", view.TableNoteForTests, StringComparison.Ordinal);

            await view.CompareSourcesAsync(true);

            Assert.Equal(1, betaReads());
            Assert.Null(view.TableComparedWithForTests);
            Assert.StartsWith(BoardSections.NotComparedLine(NoAnswer), view.TableNoteForTests, StringComparison.Ordinal);

            await view.CompareSourcesAsync(false);

            Assert.DoesNotContain("Not compared", view.TableNoteForTests, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // *** THE TWO TABLES ARE ASKED FOR TOGETHER (code review, 2026-10-10). *** They are independent;
    // asked one after the other, every compared board chosen paid two round trips back to back
    // under the "please wait". BETA's is asked for while the stable source's is still out - and
    // each is still built once, against the other.
    // ###########################################################################################
    [Fact]
    public async Task Compared_both_tables_are_asked_for_at_once_and_each_is_built_once()
    {
        await UiTest.RunAsync(async () =>
        {
            var (view, betaReads, _, _) = BoardDetailViewCompareTests.View(compare: true);
            var stableAnswer = new TaskCompletionSource<ReviewApiResult<BoardTableAnswer>>(TaskCreationOptions.RunContinuationsAsynchronously);

            view.ReadStableTableOverrideForTests = _ => stableAnswer.Task;
            view.ShowDetailForTests(BoardDetailViewCompareTests.Detail(BoardDetailViewCompareTests.Board()));

            Task shown = view.ShowSectionAsync(BoardSection.BoardData);

            // The stable source has not answered, and BETA's table has been asked for already.
            Assert.Equal(1, betaReads());
            Assert.Equal(0, view.TableBuildsForTests);

            stableAnswer.SetResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewCompareTests.Stable));
            await shown;

            Assert.Equal(1, view.TableBuildsForTests);
            Assert.Equal(1, view.StableTableBuildsForTests);
            Assert.Same(BoardDetailViewCompareTests.Stable, view.TableComparedWithForTests);
            Assert.Same(BoardDetailViewCompareTests.Beta, view.StableComparedWithForTests);
        });
    }
}
