using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// A BOARD'S TABLE AND FILES FOLLOW BETA (owner report, 2026-10-04: "I did push a change through the
// system, fixing the very much highlighted issue in this one ... then it continues to show the error
// in the "Maintainer" tab on "Systems" ... After an app restart ... it did not show as an error any
// more").
//
// The Boards screen read a board's Board data and Files once, and kept them until another board
// was chosen. The board's row now carries BETA's content hash, and every time the row is read again
// (BoardDetailView.ShowBoardAsync -> CatchUpWithBetaAsync) a table or a file list read before BETA moved
// is read again - never a table holding a change not published yet. Drawn without a server: the
// table through OpenTableForTests and ReadTableOverrideForTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardDetailViewBetaCatchUpTests
{
    private const string BoardId = "Commodore/C64/250407";

    private static BoardOverviewEntry Board(string? hash, bool awaiting = false) =>
        new(BoardId, "Commodore", "C64", "250407", true, true, awaiting, true, "2026-October-3", null, null, 1, BetaContentHash: hash);

    // The owner's board before the fix: a highlight on "Board layout" naming "hest", which is in no
    // row of the Components sheet - the warning above the table.
    private static SubmissionRows Before() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Board layout", SchematicImageFile = "a/board.png" }],
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" },
            new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "hest", X = "9", Y = "9", Width = "5", Height = "5" }
        ]
    };

    // ...and after it: the "hest" highlight gone.
    private static SubmissionRows After() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Board layout", SchematicImageFile = "a/board.png" }],
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" }
        ]
    };

    private static BoardTableAnswer Table(SubmissionRows rows, bool mayEdit = true) =>
        new(BoardId, new string('f', 64), rows, mayEdit, mayEdit ? null : "It waits under \"Queue: Awaiting push from BETA to stable\".");

    private static void Show(BoardDetailView view, BoardOverviewEntry board) =>
        view.ShowDetailForTests(new BoardDetailAnswer(board, [], [], []));

    // The board's row read again - as ShowBoardAsync does, the detail drawn first.
    private static Task ReadAgain(BoardDetailView view, BoardOverviewEntry now)
    {
        Show(view, now);
        return view.CatchUpWithBetaForTests(now);
    }

    private static TextBlock OutsideProblems(BoardDetailView view) =>
        view.BoardTableForTests.FindControl<TextBlock>("OutsideProblemsText")!;

    // The line above the table, under the BETA / Stable switch (owner request, 2026-10-04: it was
    // the table's own line under its search box).
    private static string Status(BoardDetailView view) => view.TableNoteForTests;

    // A view whose table was read at `readAt`, and a count of the tables read since.
    private static (BoardDetailView View, Func<int> Reads) WithTable(BoardOverviewEntry readAt, SubmissionRows rows, SubmissionRows next, bool nextMayEdit = true)
    {
        var view = new BoardDetailView();
        Show(view, readAt);
        view.OpenTableForTests(Table(rows), readAt);

        int reads = 0;
        view.ReadTableOverrideForTests = _ =>
        {
            reads++;
            return Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(Table(next, nextMayEdit)));
        };

        return (view, () => reads);
    }

    // The owner's case, end to end: the warning about "hest" goes once BETA no longer has it.
    [Fact]
    public async Task A_table_read_before_BETA_moved_is_read_again_and_a_fixed_warning_goes()
    {
        await UiTest.RunAsync(async () =>
        {
            (BoardDetailView view, Func<int> reads) = WithTable(Board("aaa"), Before(), After());

            Assert.True(OutsideProblems(view).IsVisible);
            Assert.Contains("hest", OutsideProblems(view).Text, StringComparison.Ordinal);

            await ReadAgain(view, Board("bbb"));

            Assert.Equal(1, reads());
            Assert.False(OutsideProblems(view).IsVisible);
        });
    }

    // Read again only when it moved - not on every check of the board.
    [Fact]
    public async Task A_table_BETA_has_not_moved_past_is_not_read_again()
    {
        await UiTest.RunAsync(async () =>
        {
            (BoardDetailView view, Func<int> reads) = WithTable(Board("aaa"), Before(), After());

            await ReadAgain(view, Board("aaa") with { ViewsLast30Days = 40 });

            Assert.Equal(0, reads());
            Assert.True(OutsideProblems(view).IsVisible);
        });
    }

    // Read again on the sheet it was on - the maintainer is looking at it.
    [Fact]
    public async Task A_table_read_again_stays_on_the_sheet_it_was_on()
    {
        await UiTest.RunAsync(async () =>
        {
            (BoardDetailView view, _) = WithTable(Board("aaa"), Before(), After());
            BoardTableEditor editor = view.BoardTableForTests;

            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetCredits)!);
            await ReadAgain(view, Board("bbb"));

            Assert.Equal(BoardWorkbookSchema.SheetCredits, editor.CurrentSheet!.Name);
        });
    }

    // ###########################################################################################
    // A promotion leaves BETA as it is but lets the table be changed again: read while the board
    // waited under BETA > Stable, it was read-only with that reason, and stayed so after it no
    // longer waited.
    // ###########################################################################################
    [Fact]
    public async Task A_table_read_while_the_board_waited_is_read_again_once_it_no_longer_does()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardOverviewEntry waiting = Board("aaa", awaiting: true);
            var view = new BoardDetailView();
            Show(view, waiting);
            view.OpenTableForTests(Table(Before(), mayEdit: false), waiting);
            view.BoardTableForTests.IsReadOnly = true;
            view.ReadTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(Table(Before(), mayEdit: true)));

            await ReadAgain(view, Board("aaa", awaiting: false));

            Assert.False(view.BoardTableForTests.IsReadOnly);
        });
    }

    // ###########################################################################################
    // *** NEVER UNDER A CHANGE NOT PUBLISHED YET *** - reading it again would throw the change away.
    // It stays, and says it can no longer go (the server would refuse it). Undone, the next look at
    // Board data reads BETA as it is.
    // ###########################################################################################
    [Fact]
    public async Task A_table_holding_a_change_is_not_read_again_but_says_the_change_can_no_longer_go()
    {
        await UiTest.RunAsync(async () =>
        {
            (BoardDetailView view, Func<int> reads) = WithTable(Board("aaa"), Before(), After());
            BoardTableDocument document = view.BoardTableForTests.CommitAndGetDocument()!;
            BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            components.Rows.Single().Cells[BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName)].Text = "Mine";

            await ReadAgain(view, Board("bbb"));

            Assert.Equal(0, reads());
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal(BoardSections.BetaMovedUnderChange, Status(view));

            document.History.Undo();
            Assert.False(view.HasUnsavedTableEdits);

            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.Equal(1, reads());
            Assert.False(OutsideProblems(view).IsVisible);
        });
    }

    // ###########################################################################################
    // A table read just after a publish of its own (thisTableReadAt null) takes the next reading as
    // the state it was read at - reading it again would only replace the line saying what the publish
    // did. A later move still reads it again.
    // ###########################################################################################
    [Fact]
    public async Task A_table_read_after_its_own_publish_is_not_read_again_for_that_publish()
    {
        await UiTest.RunAsync(async () =>
        {
            (BoardDetailView view, Func<int> reads) = WithTable(Board("aaa"), Before(), After());
            view.OpenTableForTests(Table(After()), readAt: null);

            await ReadAgain(view, Board("bbb"));
            Assert.Equal(0, reads());

            await ReadAgain(view, Board("ccc"));
            Assert.Equal(1, reads());
        });
    }

    // A board BETA no longer holds (pushed back out of it) is not left on screen as if it were.
    [Fact]
    public async Task A_table_whose_board_left_BETA_is_closed()
    {
        await UiTest.RunAsync(async () =>
        {
            (BoardDetailView view, _) = WithTable(Board("aaa"), Before(), After());
            view.ReadTableOverrideForTests = _ =>
                Task.FromResult(ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.NotFound, "BETA does not hold this board."));

            await ReadAgain(view, Board(null) with { InBeta = false });

            Assert.False(view.BoardTableForTests.HasTable);
        });
    }

    // ###########################################################################################
    // The Files view follows the board alone: a list not on screen is dropped when BETA moves, to be
    // read when next shown - and kept through a promotion, which moves nothing in BETA.
    // ###########################################################################################
    [Fact]
    public async Task A_file_list_not_on_screen_is_dropped_when_BETA_moves_and_kept_when_it_does_not()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            BoardOverviewEntry readAt = Board("aaa", awaiting: true);
            Show(view, readAt);
            view.ShowFilesForTests(new BoardFilesAnswer(BoardId, []), readAt);

            await ReadAgain(view, Board("aaa", awaiting: false));
            Assert.True(view.HoldsFilesForTests(BoardId));

            await ReadAgain(view, Board("bbb"));
            Assert.False(view.HoldsFilesForTests(BoardId));
        });
    }

    // ###########################################################################################
    // *** A DETAIL THAT SAYS OTHERWISE THAN THE TABLE READS IT AGAIN (code review, 2026-10-10). ***
    // After a publish here the table opens read-only - the board waits under BETA > Stable - and the
    // next reading of the board is taken as the state it was read at. Promoted before that reading,
    // the board's detail says it may be changed: the panel went, the table stayed read-only, and
    // nothing read it again - not even Board data shown again. Now the table is read again; and if
    // that fails, Board data shown again tries again.
    // ###########################################################################################
    [Fact]
    public async Task A_read_only_table_is_read_again_when_the_boards_detail_says_it_may_now_be_changed()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardOverviewEntry waiting = Board("bbb", awaiting: true);
            BoardOverviewEntry promoted = Board("bbb", awaiting: false);
            var view = new BoardDetailView();
            Show(view, waiting);

            // As after a publish here: read-only, read at the board's next reading.
            view.OpenTableForTests(Table(After(), mayEdit: false), readAt: null);
            Assert.True(view.BoardTableForTests.IsReadOnly);

            int reads = 0;
            view.ReadTableOverrideForTests = _ => Task.FromResult(++reads == 1
                ? ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, "No answer.")
                : ReviewApiResult<BoardTableAnswer>.Ok(Table(After(), mayEdit: true)));

            // The first reading after the publish: the board promoted meanwhile.
            view.ShowDetailForTests(new BoardDetailAnswer(promoted, [], [], [], MayEdit: true));
            await view.CatchUpWithBetaForTests(promoted);

            // Read again - which failed: still read-only, and Board data shown again tries again.
            Assert.Equal(1, reads);
            Assert.True(view.BoardTableForTests.IsReadOnly);

            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.Equal(2, reads);
            Assert.False(view.BoardTableForTests.IsReadOnly);
            Assert.Equal(string.Empty, view.ReadOnlyNoticeForTests);
        });
    }

    // A detail that agrees with the table read after a publish reads nothing - the publish's own rule.
    [Fact]
    public async Task A_read_only_table_is_not_read_again_when_the_boards_detail_agrees()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardOverviewEntry waiting = Board("bbb", awaiting: true);
            var view = new BoardDetailView();
            Show(view, waiting);
            view.OpenTableForTests(Table(After(), mayEdit: false), readAt: null);

            int reads = 0;
            view.ReadTableOverrideForTests = _ =>
            {
                reads++;
                return Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(Table(After(), mayEdit: false)));
            };

            view.ShowDetailForTests(new BoardDetailAnswer(waiting, [], [], [], MayEdit: false, MayNotEditReason: "It waits."));
            await view.CatchUpWithBetaForTests(waiting);
            await view.ShowSectionAsync(BoardSection.BoardData);

            Assert.Equal(0, reads);
            Assert.True(view.BoardTableForTests.IsReadOnly);
        });
    }
}
